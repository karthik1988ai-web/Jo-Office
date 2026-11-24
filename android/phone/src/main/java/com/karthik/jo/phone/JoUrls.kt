package com.karthik.jo.phone

import java.net.URI

/** Which pages open inside Jo Office. Everything else (links in mail, Google, Supabase) opens in the browser. */
object JoUrls {
    const val HOST = "karthik1988ai-web.github.io"
    const val ORIGIN = "https://$HOST"
    const val HOME = "$ORIGIN/Jo-Office/"

    fun isJo(url: String): Boolean {
        val u = try { URI(url) } catch (e: Exception) { return false }
        val path = u.rawPath ?: return false
        return u.scheme == "https" && u.host == HOST && (u.port == -1 || u.port == 443) &&
            u.rawUserInfo == null && (path == "/Jo-Office" || path.startsWith("/Jo-Office/"))
    }
}
