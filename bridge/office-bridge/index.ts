// Jo Office bridge: lets the Jo Office web app reach Zoho Mail, Zoho CRM, Zoho Expense and
// Zoho Calendar (browsers can't call them directly), plus a free web search.
//
// One Zoho login covers all four apps: make a Self Client grant code with every scope in
// ZOHO_SCOPES below and paste it in Jo Office's settings. The refresh token stays here.
//
// Deploy in the Supabase dashboard: Edge Functions → Deploy a new function → Via editor,
// name it "office-bridge", paste this file, and turn OFF "Verify JWT" (Jo Office uses its own key).
//
// Secrets (Edge Functions → Secrets):
//   OFFICE_BRIDGE_KEY     any long random text; the same value goes in Jo Office's settings
//   ZOHO_CLIENT_ID        from api-console.zoho.com → Self Client
//   ZOHO_CLIENT_SECRET
//   ZOHO_REGION           com (default), in, eu, com.au, jp
//   ZOHO_EXPENSE_ORG_ID   optional; the first Zoho Expense organization is used otherwise
// SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are provided by Supabase automatically.
//
// One table stores the Zoho refresh token (run once in the SQL editor):
//   create table if not exists public.office_bridge_settings (key text primary key, value text not null);
//   alter table public.office_bridge_settings enable row level security;  -- no policies: service role only
// Jo's memory and background thinking need two more tables and a cron job: supabase/memory.sql.

export const ZOHO_SCOPES = [
  "ZohoMail.accounts.READ", "ZohoMail.folders.READ", "ZohoMail.messages.ALL",
  "ZohoCRM.modules.ALL", "ZohoCRM.users.READ", "ZohoCRM.coql.READ", "ZohoCRM.org.READ",
  "ZohoExpense.fullaccess.ALL",
  "ZohoCalendar.calendar.READ", "ZohoCalendar.event.ALL",
].join(",");

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "content-type, x-office-key, authorization, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

// CRM modules Jo Office may write to.
const CRM_MODULES = new Set(["Leads", "Contacts", "Accounts", "Deals", "Products", "Tasks", "Events", "Calls"]);

type Env = (name: string) => string | undefined;
// deno-lint-ignore no-explicit-any
type Json = any;

export function createHandler(env: Env = (n) => Deno.env.get(n), http: typeof fetch = fetch) {
  const region = () => (env("ZOHO_REGION") || "com").trim();
  let accessToken = "";
  let accessExpiry = 0;
  let mailAccount: { id: string; email: string; name: string } | null = null;
  let crmUser: { id: string; email: string; name: string } | null = null;
  let expenseOrg = "";
  type Folder = { id: string; name: string; type: string };
  let folderCache: { at: number; list: Folder[] } = { at: 0, list: [] };

  // ---- settings table (service role) ----
  const restBase = () => `${env("SUPABASE_URL")}/rest/v1/office_bridge_settings`;
  const serviceHeaders = () => {
    const key = env("SUPABASE_SERVICE_ROLE_KEY") ?? "";
    const h: Record<string, string> = { apikey: key, "Content-Type": "application/json" };
    if (key.startsWith("eyJ")) h.Authorization = `Bearer ${key}`;
    return h;
  };
  async function getSetting(key: string): Promise<string> {
    const r = await http(`${restBase()}?key=eq.${encodeURIComponent(key)}&select=value`, { headers: serviceHeaders() });
    if (!r.ok) throw new Error(`Settings table not readable (${r.status}). Did you create office_bridge_settings?`);
    const rows = await r.json();
    return rows[0]?.value ?? "";
  }
  async function setSetting(key: string, value: string) {
    const r = await http(restBase(), {
      method: "POST",
      headers: { ...serviceHeaders(), Prefer: "resolution=merge-duplicates" },
      body: JSON.stringify({ key, value }),
    });
    if (!r.ok) throw new Error(`Could not save setting (${r.status}): ${(await r.text()).slice(0, 200)}`);
  }
  /** Several settings in one read. */
  async function getSettings(keys: string[]): Promise<Record<string, string>> {
    const r = await http(`${restBase()}?key=in.(${keys.map(encodeURIComponent).join(",")})&select=key,value`, { headers: serviceHeaders() });
    if (!r.ok) throw new Error(`Settings table not readable (${r.status}). Did you create office_bridge_settings?`);
    return Object.fromEntries((await r.json()).map((row: { key: string; value: string }) => [row.key, row.value]));
  }

  // ---- account details kept between restarts ----
  // Supabase starts a fresh copy of this function for most requests, and each copy used to look up
  // the mail account, CRM user, folders, team and so on again: several slow trips to Zoho before the
  // real work. They are saved here once and refreshed after a day (lists after an hour).
  type Saved = Record<string, Json> & { at?: number };
  let saved: Saved | null = null;
  const DAY = 86_400_000, HOUR = 3_600_000;
  function loadSaved(raw: string | undefined) {
    try { saved = JSON.parse(raw || "{}") ?? {}; } catch { saved = {}; }
    if (!saved!.at || Date.now() - saved!.at > DAY) saved = { at: Date.now() };
  }
  /** On a fresh start: the shared Zoho login and the saved details, in one database read. */
  let starting: Promise<string> | null = null;
  function startUp(): Promise<string> {
    return starting ??= getSettings(["zoho_access_token", "zoho_ids"]).then((both) => {
      if (!saved) loadSaved(both.zoho_ids);
      return both.zoho_access_token ?? "";
    }).finally(() => { starting = null; });
  }
  /** A saved detail, or [fetch] it once and save it. [maxAge] for lists that change more often. */
  async function remembered<T>(name: string, fetch: () => Promise<T>, maxAge = DAY): Promise<T> {
    if (!saved) {
      const shared = await startUp().catch(() => "");
      const [token, until] = shared.split("|");
      if (!accessToken && token && Number(until) > Date.now()) { accessToken = token; accessExpiry = Number(until); }
      if (!saved) saved = { at: Date.now() };
    }
    const hit = saved![name];
    if (hit && Date.now() - hit.at < maxAge) return hit.value as T;
    const value = await fetch();
    saved = { ...saved, [name]: { value, at: Date.now() } };
    setSetting("zoho_ids", JSON.stringify(saved)).catch(() => {}); // no need to wait for the save
    return value;
  }

  // ---- Zoho login (one refresh token for Mail, CRM, Expense and Calendar) ----
  async function zohoToken(params: Record<string, string>) {
    const form = new URLSearchParams({
      client_id: (env("ZOHO_CLIENT_ID") ?? "").trim(),
      client_secret: (env("ZOHO_CLIENT_SECRET") ?? "").trim(),
      ...params,
    });
    const r = await http(`https://accounts.zoho.${region()}/oauth/v2/token`, { method: "POST", body: form });
    return await r.json();
  }

  // Zoho allows only a few new access tokens per refresh token every 10 minutes, and Supabase runs
  // several short-lived copies of this function. So the access token is shared through the settings
  // table, and each copy refreshes at most once at a time.
  let refreshing: Promise<string> | null = null;
  async function zohoAccess(): Promise<string> {
    if (accessToken && Date.now() < accessExpiry) return accessToken;
    return refreshing ??= (async () => {
      try {
        // The login and the saved account details come back in one read.
        const [token, until] = (await startUp()).split("|");
        if (token && Number(until) > Date.now()) {
          accessToken = token;
          accessExpiry = Number(until);
          return accessToken;
        }
        const refresh = await getSetting("zoho_refresh_token");
        if (!refresh) throw new Error("Zoho is not connected yet. Use Connect Zoho in Jo Office's settings.");
        const json = await zohoToken({ refresh_token: refresh, grant_type: "refresh_token" });
        if (!json.access_token) {
          const reason = String(json.error_description ?? json.error ?? "no token");
          console.error("zoho token refresh failed:", reason);
          if (/too many|access denied|rate/i.test(reason)) {
            throw new Error("Zoho is limiting logins for a few minutes (too many requests). Wait 5 to 10 minutes and try again; no need to reconnect.");
          }
          throw new Error(`Zoho login expired (${reason}). Connect Zoho again in Jo Office's settings.`);
        }
        accessToken = json.access_token;
        accessExpiry = Date.now() + ((json.expires_in ?? 3600) - 300) * 1000;
        await setSetting("zoho_access_token", `${accessToken}|${accessExpiry}`).catch(() => {});
        return accessToken;
      } finally {
        refreshing = null;
      }
    })();
  }

  /** Forget the access token everywhere, e.g. after Zoho rejects it. */
  async function dropAccess() {
    accessToken = "";
    accessExpiry = 0;
    await setSetting("zoho_access_token", "").catch(() => {});
  }

  const BASES: Record<string, () => string> = {
    Mail: () => `https://mail.zoho.${region()}/api`,
    CRM: () => `https://www.zohoapis.${region()}/crm/v7`,
    Expense: () => `https://www.zohoapis.${region()}/expense/v1`,
    Calendar: () => `https://calendar.zoho.${region()}/api/v1`,
  };

  /** Calls one Zoho app. A missing scope gets a message that says how to fix it. */
  async function zoho(app: keyof typeof BASES, path: string, init: RequestInit & { form?: Record<string, string> } = {}, retried = false): Promise<Json> {
    const headers: Record<string, string> = { Authorization: `Zoho-oauthtoken ${await zohoAccess()}` };
    let body = init.body;
    if (init.form) body = new URLSearchParams(init.form);
    else if (body) headers["Content-Type"] = "application/json";
    if (app === "Expense" && !path.startsWith("/organizations")) headers["X-com-zoho-expense-organizationid"] = await expenseOrgId();
    const r = await http(`${BASES[app]()}${path}`, { method: init.method ?? "GET", headers: { ...headers, ...(init.headers ?? {}) }, body });
    if (r.status === 204) return {};
    const text = await r.text();
    let json: Json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { /* not JSON */ }
    // A token Zoho no longer accepts: get a fresh one and try once more.
    if (!retried && (r.status === 401 || json?.code === "INVALID_TOKEN") && /invalid.?(oauth)?.?token|INVALID_TOKEN|expired/i.test(text)) {
      await dropAccess();
      return await zoho(app, path, init, true);
    }
    if (!r.ok || json?.response?.errors || json?.code === "INVALID_TOKEN") {
      console.error(`zoho ${app} ${path.split("?")[0]} -> ${r.status}: ${text.slice(0, 300)}`);
      const detail = json?.message || json?.data?.errorCode || json?.response?.errors?.message || text.slice(0, 200);
      if (/scope/i.test(text)) {
        throw new Error(`Zoho ${app}: permission missing. Make a new grant code with all the scopes listed in Settings and connect again.`);
      }
      throw new Error(`Zoho ${app} error ${r.status}: ${String(typeof detail === "string" ? detail : JSON.stringify(detail)).slice(0, 200)}`);
    }
    return json;
  }

  // ---- Zoho Mail ----
  async function account() {
    if (mailAccount) return mailAccount;
    return (mailAccount = await remembered("mail", async () => {
      const a = (await zoho("Mail", "/accounts")).data?.[0];
      if (!a) throw new Error("No Zoho Mail account found");
      return { id: String(a.accountId), email: String(a.primaryEmailAddress ?? a.mailboxAddress ?? ""), name: String(a.displayName ?? "") };
    }));
  }

  /** Every mail folder (all types), kept an hour. */
  async function allFolders(): Promise<Folder[]> {
    if (Date.now() - folderCache.at < 10 * 60_000 && folderCache.list.length) return folderCache.list;
    const list = await remembered("allFolders", async () => {
      const data = (await zoho("Mail", `/accounts/${(await account()).id}/folders`)).data ?? [];
      return data.map((f: Json) => ({ id: String(f.folderId), name: String(f.folderName ?? ""), type: String(f.folderType ?? "") })) as Folder[];
    }, HOUR);
    folderCache = { at: Date.now(), list };
    return list;
  }
  /** The Inbox (or Sent) folder: by Zoho's folder type, else by name. */
  async function folderOf(kind: "Inbox" | "Sent"): Promise<Folder> {
    const list = await allFolders();
    const re = new RegExp(`^${kind}$`, "i");
    const f = list.find((x) => re.test(x.type)) ?? list.find((x) => re.test(x.name));
    if (!f) throw new Error(`No ${kind} folder found in Zoho Mail.`);
    return f;
  }
  const folderName = (list: Folder[], id: string) => list.find((f) => f.id === id)?.name ?? "";

  const toMail = (m: Json) => ({
    messageId: String(m.messageId ?? ""),
    folderId: String(m.folderId ?? ""),
    from: m.sender || m.fromAddress || "",
    email: String(m.fromAddress ?? "").replace(/^.*<|>.*$/g, ""),
    subject: m.subject ?? "",
    summary: String(m.summary ?? "").slice(0, 300),
    received: Number(m.receivedTime ?? 0),
    unread: String(m.status) === "0",
  });

  /** Mail in one folder, newest first (unread only when asked). */
  async function folderMail(f: Folder, limit: number, unreadOnly: boolean) {
    const acct = await account();
    const q = new URLSearchParams({ folderId: f.id, limit: String(limit), sortorder: "false" });
    if (unreadOnly) q.set("status", "unread");
    return ((await zoho("Mail", `/accounts/${acct.id}/messages/view?${q}`)).data ?? []).map((m: Json) => ({ ...toMail(m), folder: f.name }));
  }

  /**
   * Unread mail in the Inbox, newest first: one request. (The user's mail filters are off, so
   * everything new lands in the Inbox; other folders are still searched when asked by name.)
   */
  async function unreadInbox(limit: number) {
    const inbox = await folderOf("Inbox");
    return { mails: await folderMail(inbox, limit, true), folders: 1, missed: [] as string[] };
  }

  const emails = (v: unknown) => String(v ?? "").split(/[,;\s]+/).filter((x) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));

  async function sendMail(body: Json) {
    const to = emails(body.to);
    if (!to.length) throw new Error("No valid recipient email address.");
    const subject = String(body.subject ?? "").slice(0, 300);
    const content = String(body.content ?? "").slice(0, 20000);
    if (!content.trim()) throw new Error("The email has no text.");
    const a = await account();
    const msg: Json = { fromAddress: a.email, toAddress: to.join(","), subject, content, mailFormat: "plaintext" };
    const cc = emails(body.cc);
    if (cc.length) msg.ccAddress = cc.join(",");
    // A draft reply: Zoho's reply call can't save drafts, so the draft is saved as a new message that
    // carries the original's Message-ID (inReplyTo / refHeader). Zoho then files it in the same thread.
    if (body.draft && body.replyTo) {
      msg.mode = "draft";
      const thread = await threadOf(String(body.replyTo), String(body.folderId ?? "")).catch((e) => {
        console.error("reply draft: original not found:", e instanceof Error ? e.message : e);
        return null;
      });
      if (thread) {
        msg.inReplyTo = thread.messageIdHeader;
        msg.refHeader = thread.refs;
        if (!/^re:/i.test(subject) && thread.subject) msg.subject = `Re: ${thread.subject.replace(/^re:\s*/i, "")}`.slice(0, 300);
        if (thread.quote) msg.content = `${content}\n\n${thread.quote}`.slice(0, 30000);
      }
      await zoho("Mail", `/accounts/${a.id}/messages`, { method: "POST", body: JSON.stringify(msg) });
      return { drafted: true, to, threaded: !!thread };
    }
    // A draft is saved to the Drafts folder instead of being sent.
    if (body.draft) msg.mode = "draft";
    let path = `/accounts/${a.id}/messages`;
    if (body.replyTo) { msg.action = "reply"; path += `/${encodeURIComponent(String(body.replyTo))}`; }
    await zoho("Mail", path, { method: "POST", body: JSON.stringify(msg) });
    return body.draft ? { drafted: true, to } : { sent: true, to };
  }

  /**
   * What a reply draft needs from the email it answers: its Message-ID and References headers, the
   * subject, and the original text quoted below ("On …, Ravi wrote:" and "> " lines).
   */
  async function threadOf(messageId: string, folderId: string) {
    const acct = await account();
    const folder = folderId || (await folderOf("Inbox")).id;
    const base = `/accounts/${acct.id}/folders/${encodeURIComponent(folder)}/messages/${encodeURIComponent(messageId)}`;
    const [header, details, content] = await Promise.all([
      zoho("Mail", `${base}/header`),
      zoho("Mail", `${base}/details`).catch(() => ({})),
      zoho("Mail", `${base}/content`).catch(() => ({})),
    ]);
    const raw = String(header.data?.headerContent ?? header.data ?? "");
    const unfolded = raw.replace(/\r?\n[ \t]+/g, " ");
    const messageIdHeader = (/^message-id:\s*(<[^>\r\n]+>)/im.exec(unfolded) ?? [])[1];
    if (!messageIdHeader) throw new Error("the original email has no Message-ID header");
    const earlier = ((/^references:\s*(.+)$/im.exec(unfolded) ?? [])[1] ?? "").match(/<[^>]+>/g) ?? [];
    const d = details.data ?? {};
    const from = String(d.sender || d.fromAddress || "").replace(/&lt;|&gt;/g, "").trim();
    const when = Number(d.receivedTime ?? d.sentDateInGMT ?? 0);
    const text = htmlToText(String(content.data?.content ?? "")).slice(0, 6000);
    const said = when ? new Date(when).toLocaleString("en-IN", { timeZone: "Asia/Kolkata", dateStyle: "medium", timeStyle: "short" }) : "";
    const quote = text ? `On ${said || "an earlier date"}, ${from || "they"} wrote:\n${text.split("\n").map((l) => `> ${l}`).join("\n")}` : "";
    return { messageIdHeader, refs: [...earlier, messageIdHeader].join(" "), subject: String(d.subject ?? ""), quote };
  }
  const htmlToText = (html: string) => html
    .replace(/<(br|\/p|\/div|\/li|\/tr)[^>]*>/gi, "\n").replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();

  /**
   * Marks emails as read in Zoho Mail. Zoho wants the ids as numbers, and these ids are too long for a
   * JavaScript number, so the digits go into the JSON as they are (Zoho ignores a quoted id).
   */
  async function markRead(messageIds: string | string[]) {
    const ids = (Array.isArray(messageIds) ? messageIds : [messageIds]).map((x) => String(x).trim()).filter(Boolean);
    if (!ids.length) return false;
    try {
      const list = ids.every((x) => /^\d+$/.test(x)) ? `[${ids.join(",")}]` : JSON.stringify(ids);
      await zoho("Mail", `/accounts/${(await account()).id}/updatemessage`, {
        method: "PUT", body: `{"mode":"markAsRead","messageId":${list}}`,
      });
      return true;
    } catch (e) {
      console.error("mark as read failed:", e instanceof Error ? e.message : e);
      return false;
    }
  }

  // ---- Zoho CRM ----
  async function me() {
    if (crmUser) return crmUser;
    return (crmUser = await remembered("crmUser", async () => {
      const u = (await zoho("CRM", "/users?type=CurrentUser")).users?.[0];
      if (!u) throw new Error("Zoho CRM user not found");
      return { id: String(u.id), email: String(u.email ?? ""), name: String(u.full_name ?? "") };
    }));
  }

  /** Active CRM users (the team), kept for ten minutes here and an hour between restarts. */
  type User = { id: string; name: string; email: string; role: string };
  let usersCache: { at: number; list: User[] } = { at: 0, list: [] };
  async function crmUsers(): Promise<User[]> {
    if (Date.now() - usersCache.at < 10 * 60_000 && usersCache.list.length) return usersCache.list;
    const list = await remembered("team", async () => ((await zoho("CRM", "/users?type=ActiveUsers&per_page=200")).users ?? []).map((u: Json) => ({
      id: String(u.id),
      name: String(u.full_name ?? `${u.first_name ?? ""} ${u.last_name ?? ""}`).trim(),
      email: String(u.email ?? ""),
      role: String(u.role?.name ?? ""),
    })) as User[], HOUR);
    usersCache = { at: Date.now(), list };
    return list;
  }

  function crmModule(name: unknown) {
    const m = String(name ?? "").trim();
    const found = [...CRM_MODULES].find((x) => x.toLowerCase() === m.toLowerCase() || x.toLowerCase() === `${m.toLowerCase()}s`);
    if (!found) throw new Error(`CRM module must be one of ${[...CRM_MODULES].join(", ")}`);
    return found;
  }

  async function coql(query: string) {
    const q = query.trim();
    if (!/^select\s/i.test(q) || /;/.test(q)) throw new Error("Only a single COQL select query is allowed.");
    try {
      return (await zoho("CRM", "/coql", { method: "POST", body: JSON.stringify({ select_query: q }) })).data ?? [];
    } catch (e) {
      console.error(`coql failed: ${q}`); // the query itself, so a bad one can be fixed
      throw e;
    }
  }

  // ---- crm_find: records that match every condition, exactly ----
  // Jo picks the conditions; the query is built here (Gemini's own COQL was often invalid), values
  // like "close won" are matched to the real ones ("Closed Won"), and every record is checked
  // against the conditions again before it's returned, so nothing that doesn't match slips through.
  type Cond = { field: string; op: string; value?: unknown; value2?: unknown };
  const FIND_FIELDS: Record<string, string[]> = {
    Deals: ["Deal_Name", "Stage", "Amount", "Closing_Date", "Account_Name", "Owner"],
    Leads: ["First_Name", "Last_Name", "Company", "Email", "Phone", "Lead_Status", "Owner"],
    Contacts: ["First_Name", "Last_Name", "Email", "Phone", "Account_Name", "Owner"],
    Accounts: ["Account_Name", "Phone", "Website", "Owner"],
    Products: ["Product_Name", "Product_Code", "Unit_Price", "Product_Active"],
    Tasks: ["Subject", "Due_Date", "Status", "Priority", "Owner"],
    Events: ["Event_Title", "Start_DateTime", "End_DateTime", "Venue"],
    Calls: ["Subject", "Call_Type", "Call_Start_Time", "Owner"],
  };
  const FIND_OPS = new Set(["equals", "not_equals", "starts_with", "contains", "in", "greater_than", "less_than", "between", "is_empty", "is_not_empty"]);
  const fieldOk = (f: string) => /^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$/.test(f);
  const asNumber = (v: unknown) => (typeof v === "number" || (typeof v === "string" && /^-?\d+(\.\d+)?$/.test(v.trim())) ? Number(v) : null);
  /** A value in a query: quoted text, or a bare number when comparing sizes (Amount > 50000). */
  const lit = (v: unknown, numeric = false) => (asNumber(v) !== null && (numeric || typeof v === "number") ? String(asNumber(v)) : `'${quote(v)}'`);
  /** Lower case, letters and digits only, word endings like -ed/-s dropped: "Close won" ~ "Closed Won". */
  const loose = (s: unknown) => String(s ?? "").toLowerCase().split(/[^\p{L}\p{N}]+/u).filter(Boolean)
    .map((w) => w.length > 3 ? w.replace(/(ed|es|e|d|s)$/, "") : w).join(" ");
  /** A record's value for a field; lookups ({ name, id }) give their name. */
  const valueOf = (r: Json, field: string) => {
    let v = field in r ? r[field] : field.split(".").reduce((o: Json, k) => o?.[k], r);
    if (v && typeof v === "object") v = v.name ?? v.full_name ?? v.id ?? "";
    return v;
  };

  /** The values a text field really has (from up to 2000 records), kept an hour. */
  const valuesCache = new Map<string, { at: number; list: string[] }>();
  async function knownValues(module: string, field: string): Promise<string[]> {
    const key = `${module}.${field}`;
    const hit = valuesCache.get(key);
    if (hit && Date.now() - hit.at < HOUR) return hit.list;
    const rows = await coql(`select ${field} from ${module} where ${field} is not null limit 2000`);
    const list = [...new Set(rows.map((r: Json) => String(valueOf(r, field) ?? "")).filter(Boolean))] as string[];
    valuesCache.set(key, { at: Date.now(), list });
    return list;
  }
  /** The real value the user meant, or an error that lists the choices. */
  function resolveValue(field: string, asked: string, known: string[]): string {
    if (!known.length) return asked;
    const a = asked.trim().toLowerCase();
    const exact = known.find((k) => k.toLowerCase() === a) ?? known.find((k) => loose(k) === loose(asked));
    if (exact) return exact;
    const near = known.filter((k) => loose(k).includes(loose(asked)) || loose(asked).includes(loose(k)));
    if (near.length === 1) return near[0];
    const list = known.length <= 30 ? known.join(", ") : `${known.slice(0, 30).join(", ")}…`;
    throw new Error(near.length > 1
      ? `"${asked}" could mean several ${field} values: ${near.join(", ")}. Ask the user which one.`
      : `${field} has no value like "${asked}". The values in Zoho CRM are: ${list}.`);
  }

  function condSql(c: Cond): string {
    const f = c.field;
    switch (c.op) {
      case "equals": return `${f} = ${lit(c.value)}`;
      case "not_equals": return `${f} != ${lit(c.value)}`;
      case "starts_with": return `${f} like '${quote(c.value)}%'`;
      case "contains": return `${f} like '%${quote(c.value)}%'`;
      case "in": return `${f} in (${(Array.isArray(c.value) ? c.value : String(c.value ?? "").split(",")).map((v) => lit(typeof v === "string" ? v.trim() : v)).join(", ")})`;
      case "greater_than": return `${f} > ${lit(c.value, true)}`;
      case "less_than": return `${f} < ${lit(c.value, true)}`;
      case "between": return `${f} between ${lit(c.value, true)} and ${lit(c.value2, true)}`;
      case "is_empty": return `${f} is null`;
      case "is_not_empty": return `${f} is not null`;
    }
    throw new Error(`Unknown condition ${c.op}`);
  }
  /** COQL allows two conditions per bracket, so they're nested: ((a and b) and c). */
  const whereOf = (conds: Cond[]) => conds.map(condSql).reduce((acc, sql, i) => (i === 0 ? sql : `(${acc} and ${sql})`), "");

  function matches(r: Json, c: Cond): boolean {
    const raw = valueOf(r, c.field);
    if (raw === undefined) return true; // not in the reply: Zoho's own filter stands
    const v = String(raw ?? "").toLowerCase(), want = String(c.value ?? "").toLowerCase();
    const num = (x: unknown) => asNumber(x) ?? x;
    switch (c.op) {
      case "equals": return v === want;
      case "not_equals": return v !== want;
      case "starts_with": return v.startsWith(want);
      case "contains": return v.includes(want);
      case "in": return (Array.isArray(c.value) ? c.value : String(c.value ?? "").split(",")).some((x) => String(x).trim().toLowerCase() === v);
      case "greater_than": return (num(raw) as number) > (num(c.value) as number);
      case "less_than": return (num(raw) as number) < (num(c.value) as number);
      case "between": return (num(raw) as number) >= (num(c.value) as number) && (num(raw) as number) <= (num(c.value2) as number);
      case "is_empty": return raw === null || raw === "";
      case "is_not_empty": return raw !== null && raw !== "";
    }
    return true;
  }

  async function crmFind(body: Json) {
    const module = crmModule(body.module);
    const conds: Cond[] = (Array.isArray(body.conditions) ? body.conditions : []).map((c: Json) => ({
      field: String(c?.field ?? "").trim(), op: String(c?.op ?? "equals").trim().toLowerCase(), value: c?.value, value2: c?.value2,
    }));
    for (const c of conds) {
      if (!fieldOk(c.field)) throw new Error(`"${c.field}" is not a Zoho field name (use API names like Deal_Name, Stage).`);
      if (!FIND_OPS.has(c.op)) throw new Error(`Unknown condition "${c.op}". Use one of: ${[...FIND_OPS].join(", ")}.`);
    }
    // Text values compared exactly are matched to the values the field really has.
    const resolved: { field: string; asked: string; used: string }[] = [];
    for (const c of conds) {
      if (!["equals", "not_equals", "in"].includes(c.op)) continue;
      const asked = Array.isArray(c.value) ? c.value.map(String) : c.op === "in" ? String(c.value ?? "").split(",").map((x) => x.trim()) : [String(c.value ?? "")];
      if (asked.every((x) => asNumber(x) !== null || /^\d{4}-\d{2}-\d{2}/.test(x))) continue;
      const known = await knownValues(module, c.field).catch(() => [] as string[]);
      const used = asked.map((x) => resolveValue(c.field, x, known));
      used.forEach((u, i) => { if (u !== asked[i]) resolved.push({ field: c.field, asked: asked[i], used: u }); });
      c.value = c.op === "in" ? used : used[0];
    }
    const fields = [...new Set(["id", ...(FIND_FIELDS[module] ?? []), ...conds.map((c) => c.field),
      ...(Array.isArray(body.fields) ? body.fields.map(String).filter(fieldOk) : [])])];
    const limit = Math.min(Math.max(Number(body.limit) || 200, 1), 2000);
    const sortBy = fieldOk(String(body.sort_by ?? "")) ? String(body.sort_by) : "";
    const where = conds.length ? whereOf(conds) : "id is not null";
    const q = `select ${fields.join(", ")} from ${module} where ${where}${sortBy ? ` order by ${sortBy} ${/^desc/i.test(String(body.sort_order)) ? "desc" : "asc"}` : ""} limit ${limit}`;
    const rows = await coql(q);
    const records = rows.filter((r: Json) => conds.every((c) => matches(r, c)));
    if (records.length < rows.length) console.error(`crm_find: dropped ${rows.length - records.length} records that didn't match: ${q}`);
    return await withLinks({ module, records, resolved, limited: rows.length >= limit, query: q });
  }

  // ---- call logs and meeting reports ----
  /** A module's picklist choices from Zoho's field list, kept an hour: { Call_Type: ["Outbound", …] }. */
  const picklistCache = new Map<string, { at: number; lists: Record<string, string[]> }>();
  async function picklists(module: string): Promise<Record<string, string[]>> {
    const hit = picklistCache.get(module);
    if (hit && Date.now() - hit.at < HOUR) return hit.lists;
    const fields = (await zoho("CRM", `/settings/fields?module=${module}`)).fields ?? [];
    const lists = Object.fromEntries(fields.filter((f: Json) => Array.isArray(f.pick_list_values) && f.pick_list_values.length)
      .map((f: Json) => [String(f.api_name), f.pick_list_values.map((v: Json) => String(v.actual_value ?? v.display_value)).filter((v: string) => v && v !== "-None-")]));
    picklistCache.set(module, { at: Date.now(), lists });
    return lists;
  }
  /** The picklist's own spelling of [v], or undefined when it isn't one of the choices (Zoho would refuse it). */
  const choice = (list: string[] | undefined, v: unknown) => {
    const s = String(v ?? "").trim();
    if (!s) return undefined;
    if (!list?.length) return s;
    return list.find((x) => x.toLowerCase() === s.toLowerCase()) ?? list.find((x) => loose(x) === loose(s));
  };
  const LINK_MODULES = new Set(["Contacts", "Leads", "Accounts", "Deals"]);
  const recordLink = (crmUrl: string, module: string, id: string) =>
    crmUrl && id ? `${crmUrl}/${({ Deals: "Potentials", Calls: "Calls", Events: "Events", Tasks: "Tasks" } as Record<string, string>)[module] ?? module}/${id}` : "";

  /**
   * Saves a dictated or recorded call / meeting: the Call or Meeting record, the report as a note,
   * the transcript as a second note, and a task for each team action item. A call with the same
   * contact in the last two hours is returned as { duplicate } instead, unless force is set.
   */
  async function logActivity(body: Json) {
    const kind = body.kind === "meeting" ? "meeting" : "call";
    const pick = (o: Json) => {
      const module = crmModule(o?.module);
      if (!LINK_MODULES.has(module)) throw new Error(`Can't link to ${module}.`);
      const id = String(o?.id ?? "").replace(/\D/g, "");
      if (!id) throw new Error("The linked record has no id.");
      return { module, id };
    };
    const contact = body.contact?.id ? pick(body.contact) : null;
    const related = body.related?.id ? pick(body.related) : null;
    const start = Number(body.start) || Date.now();
    const minutes = Math.min(Math.max(Math.round(Number(body.durationMin) || (kind === "call" ? 5 : 60)), 1), 600);
    const description = String(body.report ?? "").slice(0, 30000);
    const subject = String(body.subject ?? "").trim().slice(0, 250) || (kind === "call" ? "Call" : "Meeting");

    // Who / what it's about. A lead can only be linked through What_Id, so it takes that place.
    const links: Json = {};
    if (contact?.module === "Contacts") links.Who_Id = contact.id;
    const what = contact?.module === "Leads" ? contact : related;
    if (what) { links.What_Id = what.id; links.$se_module = what.module; }

    if (kind === "call" && links.Who_Id && !body.force) {
      const dup = await coql(`select id, Subject, Call_Start_Time from Calls where (Who_Id = '${quote(links.Who_Id)}' and Call_Start_Time >= '${coqlTime(Date.now() - 2 * HOUR)}') limit 5`).catch(() => []);
      if (dup.length) return await withLinks({ duplicate: dup });
    }

    let module: string, record: Json;
    if (kind === "call") {
      const lists = await picklists("Calls").catch(() => ({} as Record<string, string[]>));
      const mm = String(minutes).padStart(2, "0");
      record = cleanFields({
        Subject: subject, Call_Type: choice(lists.Call_Type, body.callType) ?? "Outbound", Call_Start_Time: coqlTime(start),
        Call_Duration: `${mm}:00`, Description: description, Call_Purpose: choice(lists.Call_Purpose, body.purpose),
        Call_Result: choice(lists.Call_Result, body.result), Product: choice(lists.Product, body.product),
        Outgoing_Call_Status: undefined, ...links,
      });
      module = "Calls";
    } else {
      record = cleanFields({
        Event_Title: subject, Start_DateTime: coqlTime(start), End_DateTime: coqlTime(start + minutes * 60_000),
        Venue: String(body.venue ?? "").slice(0, 250), Description: description, ...links,
      });
      module = "Events";
    }
    const { id } = crmResult(await zoho("CRM", `/${module}`, { method: "POST", body: JSON.stringify({ data: [record] }) }));

    // The report as a note on the customer record (or the activity itself), and the transcript.
    const noteOn = related ?? contact ?? { module, id };
    const notes: Json[] = [];
    if (description) notes.push({ Note_Title: `${kind === "call" ? "Call" : "Meeting"} report: ${subject}`.slice(0, 120), Note_Content: description });
    const transcript = String(body.transcript ?? "").trim();
    if (transcript) notes.push({ Note_Title: `Transcript: ${subject}`.slice(0, 120), Note_Content: transcript.slice(0, 30000) });
    let noteError = "";
    if (notes.length) {
      await zoho("CRM", `/${noteOn.module}/${noteOn.id}/Notes`, { method: "POST", body: JSON.stringify({ data: notes }) })
        .catch((e) => { noteError = e instanceof Error ? e.message : String(e); });
    }

    // Action items for the team become tasks, linked to the same customer.
    const items = (Array.isArray(body.tasks) ? body.tasks : []).slice(0, 20).map((t: Json) => cleanFields({
      Subject: String(t?.subject ?? "").trim().slice(0, 250),
      Due_Date: /^\d{4}-\d{2}-\d{2}$/.test(String(t?.due ?? "")) ? String(t.due) : undefined,
      Description: String(t?.description ?? "").slice(0, 2000),
      Owner: /^\d+$/.test(String(t?.ownerId ?? "")) ? { id: String(t.ownerId) } : undefined,
      ...links,
    })).filter((t: Json) => t.Subject);
    const tasks: Json[] = [];
    if (items.length) {
      const reply = await zoho("CRM", "/Tasks", { method: "POST", body: JSON.stringify({ data: items }) }).catch((e) => ({ error: e instanceof Error ? e.message : String(e) }));
      items.forEach((t: Json, i: number) => {
        const d = reply.data?.[i];
        tasks.push(d?.status === "success" ? { subject: t.Subject, id: String(d.details?.id ?? ""), ok: true }
          : { subject: t.Subject, ok: false, error: d?.message ?? reply.error ?? "not saved" });
      });
    }
    const crmUrl = await crmTabUrl().catch(() => "");
    return {
      ok: true, kind, module, id, link: recordLink(crmUrl, module, id),
      noteOn: noteOn.module, notes: noteError ? 0 : notes.length, noteError,
      tasks: tasks.map((t) => ({ ...t, link: t.ok ? recordLink(crmUrl, "Tasks", t.id) : "" })), crmUrl,
    };
  }

  // ---- open deals ----
  // CRMs name their closed stages differently ("Closed-Lost to Competition", "Closed Won FES"), and COQL has no
  // "not like", so the closed ones are read from the Stage values the CRM really has.
  const CLOSED_STAGE = /^closed|\bwon\b|\blost\b/i;
  async function closedStages(): Promise<string[]> {
    const known = await knownValues("Deals", "Stage").catch(() => [] as string[]);
    return [...new Set(["Closed Won", "Closed Lost", ...known.filter((s) => CLOSED_STAGE.test(s))])];
  }
  async function openDeals(ownerId: string, limit = 500): Promise<Json[]> {
    const closed = await closedStages();
    const q = `select id, Deal_Name, Stage, Amount, Closing_Date, Account_Name from Deals where (Stage not in (${closed.map((s) => `'${quote(s)}'`).join(", ")}) and Owner = '${quote(ownerId)}') order by Closing_Date asc limit ${limit}`;
    return (await coql(q)).filter((d: Json) => !CLOSED_STAGE.test(String(d.Stage ?? "")));
  }

  /** One CRM record from a create/update reply, or the error Zoho gave for it. */
  function crmResult(json: Json) {
    const d = json.data?.[0];
    if (!d || d.status !== "success") throw new Error(`Zoho CRM refused it: ${d?.message ?? "unknown error"}${d?.details?.api_name ? ` (${d.details.api_name})` : ""}`);
    return { id: String(d.details?.id ?? ""), ok: true };
  }

  const cleanFields = (f: Json) => Object.fromEntries(Object.entries(f ?? {}).filter(([k, v]) => /^\$?[A-Za-z_][A-Za-z0-9_]*$/.test(k) && v !== undefined && v !== ""));

  // ---- Zoho Expense ----
  async function expenseOrgId() {
    if (expenseOrg) return expenseOrg;
    const fixed = (env("ZOHO_EXPENSE_ORG_ID") ?? "").trim();
    if (fixed) return (expenseOrg = fixed);
    return (expenseOrg = await remembered("expenseOrg", async () => {
      const org = (await zoho("Expense", "/organizations")).organizations?.[0];
      if (!org) throw new Error("No Zoho Expense organization found for this login.");
      return String(org.organization_id);
    }));
  }

  // ---- Zoho Calendar ----
  let calendarUid = "";
  async function defaultCalendar() {
    if (calendarUid) return calendarUid;
    return (calendarUid = await remembered("calendar", async () => {
      const list = (await zoho("Calendar", "/calendars")).calendars ?? [];
      const cal = list.find((c: Json) => c.isdefault === true || c.isdefault === "true") ?? list[0];
      if (!cal) throw new Error("No Zoho Calendar found for this login.");
      return String(cal.uid);
    }));
  }
  /** Zoho Calendar times look like 20261024T150000+0530 (Z for UTC). */
  function calTime(s: string) {
    const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z|[+-]\d{4})?$/.exec(String(s ?? "").trim());
    if (!m) return NaN;
    const tz = !m[7] || m[7] === "Z" ? "Z" : `${m[7].slice(0, 3)}:${m[7].slice(3)}`;
    return Date.parse(`${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}${tz}`);
  }
  const toCalEvent = (e: Json) => {
    const allDay = e.isallday === true || e.isallday === "true";
    const s = String(e.dateandtime?.start ?? ""), en = String(e.dateandtime?.end ?? s);
    return {
      id: String(e.uid ?? ""), title: String(e.title ?? "(no title)"), allDay, location: String(e.location ?? ""),
      start: allDay ? null : calTime(s), end: allDay ? null : calTime(en),
      ...(allDay ? { date: `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}` } : {}),
      link: String(e.viewEventURL ?? ""),
    };
  };
  /** Events between two times (at most a month); all-day events are kept for the app to place. */
  async function calendarEvents(from: number, to: number) {
    const uid = await defaultCalendar();
    const day = (ms: number) => new Date(ms).toISOString().slice(0, 10).replace(/-/g, "");
    const end = Math.min(to, from + 31 * 86400000);
    const range = JSON.stringify({ start: day(from - 86400000), end: day(end + 86400000) });
    const list = ((await zoho("Calendar", `/calendars/${uid}/events?range=${encodeURIComponent(range)}`)).events ?? []).map(toCalEvent);
    return list.filter((e: Json) => e.allDay || (e.end > from && e.start < end))
      .sort((a: Json, b: Json) => (a.start ?? 0) - (b.start ?? 0));
  }

  // ---- links into Zoho CRM ----
  let crmTabs = "";
  /** https://crm.zoho.com/crm/<org>/tab ; record links add /<Tab>/<id>. */
  async function crmTabUrl() {
    if (crmTabs) return crmTabs;
    return (crmTabs = await remembered("crmTabs", async () => {
      let domain = "";
      try { domain = String((await zoho("CRM", "/org")).org?.[0]?.domain_name ?? ""); } catch { /* no org permission: links still open Zoho CRM */ }
      return `https://crm.zoho.${region()}/crm/${domain ? `${domain}/` : ""}tab`;
    }));
  }
  const withLinks = async (o: Json) => ({ ...o, crmUrl: await crmTabUrl().catch(() => "") });

  // ---- web search, no API key needed: DuckDuckGo results, Google News for news, Wikipedia as a backup ----
  const BROWSER = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
    "Accept-Language": "en-IN,en;q=0.9",
  };
  type Hit = { title: string; url: string; snippet: string };
  const text = (html: string) => html.replace(/<[^>]+>/g, " ").replace(/&amp;/g, "&").replace(/&quot;/g, '"')
    .replace(/&#x27;|&#39;|&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&nbsp;/g, " ")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n))).replace(/\s+/g, " ").trim();
  // DuckDuckGo links go through a redirect (…/l/?uddg=<real url>); keep the real address.
  const realUrl = (href: string) => {
    const m = /[?&]uddg=([^&]+)/.exec(href.replace(/&amp;/g, "&"));
    return m ? decodeURIComponent(m[1]) : href.startsWith("//") ? `https:${href}` : href;
  };
  const isAd = (url: string) => /duckduckgo\.com\/y\.js|[?&]ad_provider=/.test(url);

  async function duckDuckGo(q: string): Promise<Hit[]> {
    const res = await http("https://html.duckduckgo.com/html/", {
      method: "POST",
      headers: { ...BROWSER, "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ q, kl: "in-en" }),
    });
    if (!res.ok) throw new Error(`DuckDuckGo ${res.status}`);
    const html = await res.text();
    const hits: Hit[] = [];
    const link = /<a\b([^>]*class="result__a"[^>]*)>([\s\S]*?)<\/a>/g;
    let m: RegExpExecArray | null;
    while ((m = link.exec(html)) && hits.length < 8) {
      const href = /href="([^"]+)"/.exec(m[1])?.[1] ?? "";
      const rest = html.slice(link.lastIndex, link.lastIndex + 4000);
      const end = rest.indexOf('class="result__a"');
      const snippet = /class="result__snippet"[^>]*>([\s\S]*?)<\/(?:a|div|td)>/.exec(end > 0 ? rest.slice(0, end) : rest)?.[1] ?? "";
      const url = realUrl(href);
      if (url && !isAd(url)) hits.push({ title: text(m[2]), url, snippet: text(snippet) });
    }
    return hits;
  }

  // Google News RSS: headlines with their source and time. A plain "today's news" gets the top stories.
  async function googleNews(q: string): Promise<{ title: string; source: string; date: string }[]> {
    const keywords = q.toLowerCase().replace(/['’]s\b/g, "")
      .replace(/\b(what|whats|what's|are|is|the|top|today|today's|todays|latest|news|headlines?|breaking|in|and|of|for|me|tell|give|show|any|about|on)\b/g, " ")
      .replace(/[^\p{L}\p{N} ]/gu, " ").replace(/\s+/g, " ").trim();
    const feed = keywords
      ? `https://news.google.com/rss/search?q=${encodeURIComponent(keywords)}&hl=en-IN&gl=IN&ceid=IN:en`
      : "https://news.google.com/rss?hl=en-IN&gl=IN&ceid=IN:en";
    const res = await http(feed, { headers: BROWSER });
    if (!res.ok) throw new Error(`Google News ${res.status}`);
    const xml = await res.text();
    const tag = (item: string, name: string) => {
      const v = new RegExp(`<${name}[^>]*>([\\s\\S]*?)</${name}>`).exec(item)?.[1] ?? "";
      return text(v.replace(/^<!\[CDATA\[|\]\]>$/g, ""));
    };
    return [...xml.matchAll(/<item>([\s\S]*?)<\/item>/g)].slice(0, 8).map(([, item]) => {
      const source = tag(item, "source");
      return { title: tag(item, "title").replace(new RegExp(` - ${source.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`), ""), source, date: tag(item, "pubDate") };
    });
  }

  async function wikipedia(q: string): Promise<Hit[]> {
    const url = `https://en.wikipedia.org/w/api.php?action=query&list=search&format=json&utf8=1&srlimit=3&srsearch=${encodeURIComponent(q)}`;
    const res = await http(url, { headers: { ...BROWSER, "User-Agent": "JoAssistant/1.0 (personal assistant)" } });
    if (!res.ok) throw new Error(`Wikipedia ${res.status}`);
    const data = await res.json();
    return (data.query?.search ?? []).map((r: { title: string; snippet: string }) => ({
      title: r.title, url: `https://en.wikipedia.org/wiki/${encodeURIComponent(r.title.replace(/ /g, "_"))}`, snippet: text(r.snippet),
    }));
  }

  async function webSearch(query: string) {
    const q = query.trim().slice(0, 300);
    if (!q) throw new Error("Nothing to search for.");
    const errors: string[] = [];
    const wantsNews = /\b(news|headlines?|breaking)\b|செய்தி/i.test(q);
    const [results, news] = await Promise.all([
      duckDuckGo(q).catch((e) => { errors.push(e.message); return [] as Hit[]; }),
      wantsNews ? googleNews(q).catch((e) => { errors.push(e.message); return []; }) : Promise.resolve([]),
    ]);
    let hits = results;
    if (!hits.length && !news.length) hits = await wikipedia(q).catch((e) => { errors.push(e.message); return []; });
    return { query: q, results: hits.slice(0, 6), news: news.slice(0, 6), errors };
  }

  // ---- helpers for compact replies ----
  const pick = (o: Json, keys: string[]) => Object.fromEntries(keys.filter((k) => o?.[k] !== undefined && o?.[k] !== null && o?.[k] !== "").map((k) => [k, o[k]]));
  const firstArray = (o: Json): Json[] => (Object.values(o ?? {}).find((v) => Array.isArray(v)) as Json[]) ?? [];
  /** COQL wants date-times like 2026-10-21T00:00:00+00:00. */
  const coqlTime = (ms: number) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, "+00:00");
  const quote = (s: unknown) => String(s ?? "").replace(/['\\]/g, "");

  async function expenseCategories() {
    return firstArray(await zoho("Expense", "/expensecategories?per_page=200"))
      .map((c: Json) => ({ id: String(c.category_id ?? c.account_id ?? ""), name: String(c.category_name ?? c.account_name ?? "") }))
      .filter((c) => c.id && c.name);
  }

  const byName = <T extends { name: string }>(list: T[], name: unknown, what: string) => {
    const n = String(name ?? "").toLowerCase().trim();
    const hit = list.find((x) => x.name.toLowerCase() === n) ?? list.find((x) => x.name.toLowerCase().includes(n) || (n && n.includes(x.name.toLowerCase())));
    if (!hit) throw new Error(`No ${what} called "${name}". Choose one of: ${list.map((x) => x.name).join(", ")}`);
    return hit;
  };

  // ---- Jo's memory and background thinking ----
  // Tables jo_activity (what was asked and done) and jo_suggestions (what Jo noticed); notes in the
  // settings table: jo_facts (what Jo knows about the work), jo_style (how the user writes),
  // jo_think (last run, Gemini calls today). A Supabase cron job calls action "think" every 30 min.
  const table = (name: string) => `${env("SUPABASE_URL")}/rest/v1/${name}`;
  async function rows(name: string, query: string): Promise<Json[]> {
    const r = await http(`${table(name)}?${query}`, { headers: serviceHeaders() });
    if (!r.ok) throw new Error(`Jo's memory isn't set up yet (${name}: ${r.status}). Run the memory SQL in Supabase.`);
    return await r.json();
  }
  async function insert(name: string, data: Json) {
    const r = await http(table(name), { method: "POST", headers: { ...serviceHeaders(), Prefer: "return=minimal" }, body: JSON.stringify(data) });
    if (!r.ok) throw new Error(`Couldn't save to ${name} (${r.status}): ${(await r.text()).slice(0, 200)}`);
  }
  async function patch(name: string, query: string, data: Json) {
    const r = await http(`${table(name)}?${query}`, { method: "PATCH", headers: { ...serviceHeaders(), Prefer: "return=minimal" }, body: JSON.stringify(data) });
    if (!r.ok) throw new Error(`Couldn't update ${name} (${r.status})`);
  }
  const clipText = (s: unknown, n: number) => String(s ?? "").slice(0, n);
  const istDay = (ms = Date.now()) => new Date(ms + 5.5 * HOUR).toISOString().slice(0, 10);
  const istHour = (ms = Date.now()) => new Date(ms + 5.5 * HOUR).getUTCHours();
  const GEMINI_DAILY_LIMIT = 80; // background calls a day: far below the free limit, so Jo itself never runs short

  /** The background's Gemini key: saved from Jo's Settings, or a GEMINI_API_KEY secret. */
  async function geminiKey() { return (env("GEMINI_API_KEY") ?? "").trim() || (await getSetting("gemini_key").catch(() => "")); }

  /** One Gemini answer as JSON; the lighter model if the main one is busy. */
  async function geminiJson(key: string, prompt: string): Promise<Json> {
    let last = "";
    for (const model of ["gemini-flash-latest", "gemini-flash-lite-latest"]) {
      const r = await http(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-goog-api-key": key },
        body: JSON.stringify({ contents: [{ role: "user", parts: [{ text: prompt }] }], generationConfig: { responseMimeType: "application/json", temperature: 0.2 } }),
      });
      const text = await r.text();
      if (r.ok) {
        const out = JSON.parse(text);
        const answer = (out.candidates?.[0]?.content?.parts ?? []).map((p: Json) => p.text ?? "").join("");
        try { return JSON.parse(answer.replace(/^```(?:json)?|```$/g, "").trim()); } catch { last = "unreadable answer"; continue; }
      }
      last = `${r.status} ${text.slice(0, 160)}`;
      if (![429, 500, 503].includes(r.status)) break;
    }
    throw new Error(`Gemini: ${last}`);
  }

  type ThinkState = { lastRun?: number; day?: string; calls?: number; styleDay?: string; lastError?: string };
  async function thinkState(): Promise<ThinkState> { try { return JSON.parse(await getSetting("jo_think") || "{}"); } catch { return {}; } }

  /**
   * Background thinking: looks at what changed (new mail, sent mail, deals, tasks, meetings, what the
   * user asked Jo) and writes suggestions plus updated notes. Runs from cron with the PC off.
   */
  async function think() {
    const key = await geminiKey();
    if (!key) return { skipped: "No Gemini key for background thinking yet: use the button in Jo's Settings." };
    const state = await thinkState();
    const today = istDay();
    if (state.day !== today) { state.day = today; state.calls = 0; }
    if ((state.calls ?? 0) >= GEMINI_DAILY_LIMIT) return { skipped: "Daily limit for background thinking reached." };
    const since = state.lastRun ?? Date.now() - DAY;
    const safe = async <T>(f: () => Promise<T>, fallback: T) => { try { return await f(); } catch (e) { console.error("think:", e instanceof Error ? e.message : e); return fallback; } };
    const me_ = await safe(me, null);
    const [inbox, sent, deals, tasks, events, activity, open, judged, facts, style] = await Promise.all([
      safe(async () => folderMail(await folderOf("Inbox"), 40, false), [] as Json[]),
      safe(async () => folderMail(await folderOf("Sent"), 40, false), [] as Json[]),
      safe(async () => me_ ? (await openDeals(me_.id)).slice(0, 80) : [], [] as Json[]),
      safe(async () => me_ ? await coql(`select id, Subject, Due_Date, Status, Priority from Tasks where (Status != 'Completed' and Owner = '${quote(me_.id)}') order by Due_Date asc limit 50`) : [], [] as Json[]),
      safe(() => calendarEvents(Date.now(), Date.now() + 36 * HOUR), [] as Json[]),
      safe(() => rows("jo_activity", `select=at,kind,text&at=gte.${new Date(Math.min(since, Date.now() - 6 * HOUR)).toISOString()}&order=at.desc&limit=40`), [] as Json[]),
      safe(() => rows("jo_suggestions", "select=id,title,ref&status=eq.new&order=at.desc&limit=30"), [] as Json[]),
      safe(() => rows("jo_suggestions", `select=title,feedback,status&or=(feedback.neq.0,status.eq.dismissed)&order=at.desc&limit=20`), [] as Json[]),
      safe(() => getSetting("jo_facts"), ""),
      safe(() => getSetting("jo_style"), ""),
    ]);
    const day = (ms: number) => ms ? new Date(ms + 5.5 * HOUR).toISOString().slice(0, 16).replace("T", " ") : "";
    const mailLine = (m: Json) => `${day(m.received)} | ${m.unread ? "UNREAD | " : ""}from ${m.from} <${m.email}> | ${m.subject} | ${clipText(m.summary, 160)} | ref=mail:${m.messageId}:${m.folderId}`;
    const sentLine = (m: Json) => `${day(m.received)} | ${m.subject} | ${clipText(m.summary, 120)}`;
    const now = new Date(Date.now() + 5.5 * HOUR).toISOString().slice(0, 16).replace("T", " ");
    const morning = istHour() >= 7 && istHour() < 10;
    const prompt = [
      "You are the background brain of Jo Office, a work assistant for a sales manager at a company in Tamil Nadu, India (Zoho Mail and Zoho CRM).",
      `Now (India time): ${now}.${morning ? " It is morning: also include one suggestion of kind \"brief\" summing up the day ahead in 2 sentences." : ""}`,
      "From the data below, find the few things the user should act on, and update your notes about their work.",
      "Good suggestions: a customer email waiting for a reply for over a day (no matching \"Re:\" in Sent), a deal past its closing date",
      "or closing within 3 days, an overdue task, a meeting in the next 24 hours with an open topic, a promise the user made in an email",
      "or to Jo that is not done. Skip newsletters, promotions, automatic notifications and anything already in OPEN SUGGESTIONS.",
      "Never invent facts: use only the data. Write numbers in digits (₹1,20,000). At most 5 suggestions; none is fine.",
      "Learn from FEEDBACK: don't suggest things like the ones the user dismissed or marked unhelpful.",
      "Each suggestion: kind (reply | deal | task | meeting | followup | brief), title (under 12 words), detail (one or two sentences, why),",
      "say (what the user would ask Jo to act on it, e.g. \"Draft a reply to the email from Vikram Fotedar about RRR.0391 design tree\"),",
      "ref (the ref= value for a mail, deal:<id> or task:<id>; empty for brief).",
      "notes: rewrite NOTES as an up-to-date list of short facts worth remembering (people and their companies, open topics and promises,",
      "customer preferences, what the user usually wants), merging the old notes with what's new; drop what's finished; at most 3000 characters.",
      'Answer JSON only: {"suggestions": [...], "notes": "..."}',
      "",
      `NOTES (your memory so far):\n${clipText(facts, 3500) || "(none yet)"}`,
      `\nWHAT THE USER ASKED JO RECENTLY:\n${activity.map((a: Json) => `${day(Date.parse(a.at))} [${a.kind}] ${clipText(a.text, 300)}`).join("\n") || "(nothing)"}`,
      `\nINBOX (newest first):\n${inbox.map(mailLine).join("\n") || "(empty or unavailable)"}`,
      `\nSENT:\n${sent.map(sentLine).join("\n") || "(empty or unavailable)"}`,
      `\nOPEN DEALS:\n${deals.map((d: Json) => `${d.Deal_Name} | ${d.Stage} | ₹${d.Amount ?? "?"} | closes ${d.Closing_Date ?? "?"} | ${d.Account_Name?.name ?? ""} | ref=deal:${d.id}`).join("\n") || "(none)"}`,
      `\nOPEN TASKS:\n${tasks.map((t: Json) => `${t.Subject} | due ${t.Due_Date ?? "-"} | ${t.Priority ?? ""} | ref=task:${t.id}`).join("\n") || "(none)"}`,
      `\nMEETINGS (next 36 hours):\n${events.map((e: Json) => `${day(e.start)} | ${e.title}${e.location ? ` | ${e.location}` : ""}`).join("\n") || "(none)"}`,
      `\nOPEN SUGGESTIONS (don't repeat):\n${open.map((s: Json) => `${s.title} | ${s.ref}`).join("\n") || "(none)"}`,
      `\nFEEDBACK:\n${judged.map((s: Json) => `${s.title}: ${s.feedback > 0 ? "helpful" : "not helpful"}`).join("\n") || "(none)"}`,
    ].join("\n");
    state.calls = (state.calls ?? 0) + 1;
    let out: Json;
    try {
      out = await geminiJson(key, prompt);
    } catch (e) {
      state.lastError = e instanceof Error ? e.message : String(e);
      await setSetting("jo_think", JSON.stringify(state)).catch(() => {});
      throw e;
    }
    const known = new Set(open.map((s: Json) => String(s.ref || s.title)));
    const fresh = (Array.isArray(out.suggestions) ? out.suggestions : []).slice(0, 5)
      .map((s: Json) => ({ kind: clipText(s.kind, 20) || "followup", title: clipText(s.title, 120), detail: clipText(s.detail, 500), say: clipText(s.say, 300), ref: clipText(s.ref, 200) }))
      .filter((s: Json) => s.title && !known.has(s.ref || s.title));
    if (fresh.length) await insert("jo_suggestions", fresh);
    if (typeof out.notes === "string" && out.notes.trim()) await setSetting("jo_facts", clipText(out.notes.trim(), 4000));

    // Once a day: how the user writes, from their sent mail and the changes they made to Jo's drafts.
    let styled = false;
    if (state.styleDay !== today && sent.length && (state.calls ?? 0) < GEMINI_DAILY_LIMIT) {
      const edits = await safe(() => rows("jo_activity", "select=text&kind=eq.draft_edit&order=at.desc&limit=10"), [] as Json[]);
      const styleOut = await safe(() => geminiJson(key, [
        "Describe how this person writes work emails, as short instructions another writer can follow: greeting and sign-off they use,",
        "tone with customers and with colleagues, typical length, phrases they like, words they avoid, and what they changed in drafts",
        "written for them (learn from those changes most). At most 1200 characters. Answer JSON only: {\"style\": \"...\"}",
        `\nCURRENT NOTES:\n${clipText(style, 1500) || "(none)"}`,
        `\nTHEIR SENT EMAILS (subject | start):\n${sent.slice(0, 25).map(sentLine).join("\n")}`,
        `\nCHANGES THEY MADE TO DRAFTS:\n${edits.map((e: Json) => clipText(e.text, 600)).join("\n---\n") || "(none)"}`,
      ].join("\n")), null as Json);
      state.calls = (state.calls ?? 0) + 1;
      if (styleOut?.style) { await setSetting("jo_style", clipText(styleOut.style, 1500)); styled = true; }
      state.styleDay = today;
    }
    state.lastRun = Date.now();
    delete state.lastError;
    await setSetting("jo_think", JSON.stringify(state));
    return { suggestions: fresh.length, notes: !!out.notes, style: styled, callsToday: state.calls };
  }

  async function run(action: string, body: Json) {
    switch (action) {
      case "ping":
        return { ok: true };
      case "web_search":
        return await webSearch(String(body.query ?? ""));
      case "zoho_scopes":
        return { scopes: ZOHO_SCOPES };
      case "zoho_connect": {
        const json = await zohoToken({ code: String(body.code ?? "").trim(), grant_type: "authorization_code" });
        if (!json.refresh_token) {
          throw new Error(`Zoho said: ${json.error ?? JSON.stringify(json).slice(0, 200)}. Grant codes expire in minutes; make a fresh one.`);
        }
        await setSetting("zoho_refresh_token", json.refresh_token);
        await dropAccess();
        // A new login may be a different account: look everything up again.
        saved = { at: Date.now() };
        folderCache = { at: 0, list: [] };
        usersCache = { at: 0, list: [] };
        await setSetting("zoho_ids", "{}").catch(() => {});
        mailAccount = crmUser = null;
        expenseOrg = calendarUid = crmTabs = "";
        return { connected: true };
      }
      /** Tries each Zoho app once so Settings can show which ones work. */
      case "zoho_status": {
        const check = async (f: () => Promise<unknown>) => { try { await f(); return "ok"; } catch (e) { return e instanceof Error ? e.message : String(e); } };
        return {
          mail: await check(account),
          crm: await check(me),
          expense: await check(expenseOrgId),
          calendar: await check(defaultCalendar),
        };
      }

      // Mail
      case "mail_unread":
        return await unreadInbox(Math.min(Number(body.limit) || 15, 50));
      /** Emails Jo told the user about: marked read in Zoho Mail. */
      case "mail_mark_read": {
        const ids = (Array.isArray(body.messageIds) ? body.messageIds : []).slice(0, 50);
        return { markedRead: await markRead(ids) };
      }

      // Jo's memory
      case "memory_get": {
        const [notes, suggestions, state] = await Promise.all([
          getSettings(["jo_facts", "jo_style", "gemini_key"]),
          rows("jo_suggestions", "select=id,at,kind,title,detail,say,ref&status=eq.new&order=at.desc&limit=12").catch(() => null),
          thinkState(),
        ]);
        return {
          facts: notes.jo_facts ?? "", style: notes.jo_style ?? "", suggestions: suggestions ?? [], ready: suggestions !== null,
          thinking: !!(notes.gemini_key || (env("GEMINI_API_KEY") ?? "").trim()), lastThink: state.lastRun ?? 0, lastError: state.lastError ?? "",
        };
      }
      case "memory_log": {
        const kind = clipText(body.kind, 20) || "ask";
        await insert("jo_activity", { kind, text: clipText(body.text, 2000), meta: body.meta && typeof body.meta === "object" ? body.meta : {} });
        return { ok: true };
      }
      case "suggestion_update": {
        const id = String(body.id ?? "").replace(/\D/g, "");
        if (!id) throw new Error("Which suggestion?");
        const change: Json = {};
        if (["new", "done", "dismissed"].includes(body.status)) change.status = body.status;
        if ([1, -1, 0].includes(Number(body.feedback)) && body.feedback !== undefined) change.feedback = Number(body.feedback);
        await patch("jo_suggestions", `id=eq.${id}`, change);
        return { ok: true };
      }
      /** Saved from the button in Jo's Settings, so background thinking can use Gemini with the PC off. */
      case "memory_set_gemini": {
        const key = String(body.key ?? "").trim();
        // Google's key formats change (older AIza…, newer ones with dots), so the key is checked by asking Gemini.
        if (!/^\S{20,300}$/.test(key)) throw new Error("That doesn't look like a Gemini API key.");
        const test = await http("https://generativelanguage.googleapis.com/v1beta/models?pageSize=1", { headers: { "x-goog-api-key": key } }).catch(() => null);
        if (test && [400, 401, 403].includes(test.status)) throw new Error("Google didn't accept this Gemini key. Check it in Settings.");
        await setSetting("gemini_key", key);
        return { ok: true };
      }
      case "think":
        return await think();
      case "mail_recent": {
        const q = new URLSearchParams({ limit: String(Math.min(Number(body.limit) || 10, 30)), sortorder: "false" });
        return { mails: ((await zoho("Mail", `/accounts/${(await account()).id}/messages/view?${q}`)).data ?? []).map(toMail) };
      }
      /** Searches every folder (Zoho leaves out only Spam and Trash), or one folder by name. */
      case "mail_search": {
        const query = String(body.query ?? "").trim();
        const [acct, folders] = await Promise.all([account(), allFolders().catch(() => [] as Folder[])]);
        const parts = query ? [query.includes(":") ? query : `entire:${query}`] : [];
        const wanted = String(body.folder ?? "").trim().toLowerCase();
        if (wanted) {
          const f = folders.find((x) => x.name.toLowerCase() === wanted) ?? folders.find((x) => x.name.toLowerCase().includes(wanted));
          if (!f) throw new Error(`No mail folder called "${body.folder}". Folders: ${folders.map((x) => x.name).join(", ")}`);
          parts.push(`in:${f.name.includes(" ") ? `"${f.name}"` : f.name}`);
        }
        if (!parts.length) throw new Error("Nothing to search for.");
        const q = new URLSearchParams({ searchKey: parts.join("::"), limit: String(Math.min(Number(body.limit) || 8, 50)) });
        const mails = ((await zoho("Mail", `/accounts/${acct.id}/messages/search?${q}`)).data ?? [])
          .map((m: Json) => ({ ...toMail(m), folder: folderName(folders, String(m.folderId ?? "")) }));
        return { mails };
      }
      case "mail_folders":
        return { folders: (await allFolders()).map((f) => f.name) };
      case "mail_read": {
        const folder = encodeURIComponent(String(body.folderId));
        const message = encodeURIComponent(String(body.messageId));
        const json = await zoho("Mail", `/accounts/${(await account()).id}/folders/${folder}/messages/${message}/content`);
        const content = String(json.data?.content ?? "");
        const markedRead = body.markRead === false ? false : await markRead(String(body.messageId));
        return { content, markedRead };
      }
      case "mail_send":
        return await sendMail(body);

      // CRM
      case "crm_search": {
        const module = crmModule(body.module);
        const q = new URLSearchParams({ word: String(body.word ?? "").slice(0, 100), per_page: String(Math.min(Number(body.limit) || 10, 50)) });
        return await withLinks({ records: (await zoho("CRM", `/${module}/search?${q}`)).data ?? [] });
      }
      case "crm_get": {
        const module = crmModule(body.module);
        return { record: (await zoho("CRM", `/${module}/${encodeURIComponent(String(body.id ?? ""))}`)).data?.[0] ?? null };
      }
      case "crm_query":
        return await withLinks({ records: await coql(String(body.query ?? "")) });
      case "crm_find":
        return await crmFind(body);
      case "crm_agenda": {
        const q = `select id, Event_Title, Start_DateTime, End_DateTime, Venue, All_day from Events where Start_DateTime between '${coqlTime(Number(body.from))}' and '${coqlTime(Number(body.to))}' order by Start_DateTime asc limit 50`;
        return await withLinks({ events: await coql(q) });
      }
      case "crm_my_tasks": {
        const u = await me();
        const q = `select id, Subject, Due_Date, Status, Priority from Tasks where (Status != 'Completed' and Owner = '${quote(u.id)}') order by Due_Date asc limit 50`;
        return await withLinks({ tasks: await coql(q) });
      }
      case "crm_users":
        return { users: await crmUsers() };
      /** Open tasks of one team member (ownerId), or all tasks you gave to others. */
      case "crm_team_tasks": {
        const u = await me();
        const owner = String(body.ownerId ?? "").replace(/\D/g, "");
        // COQL takes at most two conditions per bracket, so three are nested.
        const where = owner
          ? `(Status != 'Completed' and Owner = '${owner}')`
          : `((Status != 'Completed' and Created_By = '${quote(u.id)}') and Owner != '${quote(u.id)}')`;
        const q = `select id, Subject, Due_Date, Status, Priority, Owner from Tasks where ${where} order by Due_Date asc limit 100`;
        const [rows, users] = await Promise.all([coql(q), crmUsers().catch(() => [])]);
        const nameOf = (id: string) => users.find((x) => x.id === id)?.name ?? "";
        return await withLinks({ tasks: rows.map((t: Json) => ({ ...t, Owner_Name: nameOf(String(t.Owner?.id ?? t.Owner ?? "")) })) });
      }
      case "crm_pipeline": {
        const u = await me();
        return await withLinks({ deals: await openDeals(u.id) });
      }
      case "crm_create": {
        const module = crmModule(body.module);
        const fields = cleanFields(body.fields);
        if (!Object.keys(fields).length) throw new Error("Nothing to save.");
        return crmResult(await zoho("CRM", `/${module}`, { method: "POST", body: JSON.stringify({ data: [fields] }) }));
      }
      case "crm_update": {
        const module = crmModule(body.module);
        const fields = { ...cleanFields(body.fields), id: String(body.id ?? "") };
        if (!fields.id) throw new Error("Which record? The id is missing.");
        return crmResult(await zoho("CRM", `/${module}`, { method: "PUT", body: JSON.stringify({ data: [fields] }) }));
      }
      case "crm_add_note": {
        const module = crmModule(body.module);
        const note = { Note_Title: String(body.title ?? "Note").slice(0, 120), Note_Content: String(body.content ?? "").slice(0, 5000) };
        return crmResult(await zoho("CRM", `/${module}/${encodeURIComponent(String(body.id ?? ""))}/Notes`, { method: "POST", body: JSON.stringify({ data: [note] }) }));
      }
      case "crm_log_activity":
        return await logActivity(body);

      // Expense
      case "expense_list": {
        const q = new URLSearchParams({ per_page: String(Math.min(Number(body.limit) || 20, 100)) });
        const keys = ["expense_id", "date", "merchant_name", "category_name", "total", "amount", "currency_code", "status", "report_name", "description", "is_reimbursable"];
        return { expenses: firstArray(await zoho("Expense", `/expenses?${q}`)).map((e: Json) => pick(e, keys)) };
      }
      case "expense_reports": {
        const keys = ["report_id", "report_name", "report_number", "status", "total", "currency_code", "start_date", "end_date", "submitted_date", "approver_name", "due_date"];
        return { reports: firstArray(await zoho("Expense", "/expensereports?per_page=30")).map((r: Json) => pick(r, keys)) };
      }
      case "expense_categories":
        return { categories: await expenseCategories() };
      case "expense_create": {
        const amount = Number(body.amount);
        if (!(amount > 0)) throw new Error("The expense needs an amount above zero.");
        const category = byName(await expenseCategories(), body.category, "expense category");
        const expense: Json = {
          date: String(body.date ?? "").slice(0, 10),
          is_reimbursable: body.reimbursable !== false,
          line_items: [{ category_id: category.id, amount, description: String(body.description ?? "").slice(0, 500) }],
        };
        if (body.merchant) expense.merchant_name = String(body.merchant).slice(0, 100);
        const json = await zoho("Expense", "/expenses", { method: "POST", body: JSON.stringify(expense) });
        return { id: String(json.expense?.expense_id ?? json.expenses?.[0]?.expense_id ?? ""), category: category.name, ok: true };
      }

      // Calendar
      case "calendar_events":
        return { events: await calendarEvents(Number(body.from), Number(body.to)) };
      case "calendar_add": {
        const when = /^\d{8}T\d{6}(Z|[+-]\d{4})$/;
        if (!when.test(String(body.start ?? "")) || !when.test(String(body.end ?? ""))) throw new Error("Bad time; expected 20261024T150000+0530");
        const event: Json = {
          title: String(body.title ?? "").slice(0, 200) || "Meeting",
          dateandtime: { timezone: String(body.timezone || "Asia/Kolkata"), start: String(body.start), end: String(body.end) },
        };
        if (body.location) event.location = String(body.location).slice(0, 200);
        if (body.description) event.description = String(body.description).slice(0, 2000);
        const uid = await defaultCalendar();
        const json = await zoho("Calendar", `/calendars/${uid}/events?eventdata=${encodeURIComponent(JSON.stringify(event))}`, { method: "POST" });
        return { ok: true, id: String(json.events?.[0]?.uid ?? "") };
      }
      default:
        throw new Error(`Unknown action ${action}`);
    }
  }

  return async (req: Request): Promise<Response> => {
    if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
    const json = (data: unknown, status = 200) =>
      new Response(JSON.stringify(data), { status, headers: { ...CORS, "Content-Type": "application/json" } });

    // The scheduled background job has its own key (made by the memory SQL, kept in the settings table),
    // and may only start "think".
    const thinkKey = (req.headers.get("x-think-key") ?? "").trim();
    if (thinkKey) {
      const stored = await getSetting("think_key").catch(() => "");
      if (!stored || thinkKey !== stored) return json({ error: "Wrong think key." }, 401);
      try { return json(await think()); } catch (e) {
        console.error("think error:", e instanceof Error ? e.message : String(e));
        return json({ error: e instanceof Error ? e.message : String(e) }, 400);
      }
    }

    // Trimmed on both sides: a stray space from copy-paste is the most common mismatch.
    const key = (env("OFFICE_BRIDGE_KEY") ?? "").trim();
    if (!key) {
      // Names only (never values) of similar secrets, so a typo or wrong project is easy to spot.
      let seen: string[] = [];
      try { seen = Object.keys(Deno.env.toObject()).filter((n) => /office|zoho|bridge/i.test(n)); } catch { /* not allowed here */ }
      return json({
        error: "OFFICE_BRIDGE_KEY is not set in Supabase (Edge Functions → Secrets). Add it, then try again.",
        similarSecretNames: seen,
      }, 401);
    }
    if ((req.headers.get("x-office-key") ?? "").trim() !== key) {
      return json({ error: "The Bridge key in Jo Office doesn't match OFFICE_BRIDGE_KEY in Supabase. Re-type the same text in both places." }, 401);
    }

    try {
      const body = await req.json().catch(() => ({}));
      return json(await run(String(body.action ?? ""), body));
    } catch (e) {
      console.error("bridge error:", e instanceof Error ? e.message : String(e));
      return json({ error: e instanceof Error ? e.message : String(e) }, 400);
    }
  };
}

if (!Deno.env.get("OFFICE_BRIDGE_TEST")) Deno.serve(createHandler());
