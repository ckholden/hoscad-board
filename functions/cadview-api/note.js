// Same-origin proxy for the SCMC CAD Viewer's note SAVE action
// (cadview/index.html POSTs here, not directly to Railway).
//
// Why this exists: Cloudflare Access injects the
// Cf-Access-Authenticated-User-Email header into requests that terminate at
// THIS (Access-protected) origin. A direct cross-origin browser fetch to the
// Railway backend would never carry that header -- Access only gates inbound
// requests to its configured hostname, not a page's outbound fetches to other
// origins. So the only way to attribute a save to the person who made it is to
// route it through a same-origin hop first.
//
// Reads don't need this -- the page fetches the note directly from Railway
// (see NOTE_GET_URL in cadview/index.html), since reading doesn't need identity.
const RAILWAY_NOTE_URL = "https://orfireems-scraper-production.up.railway.app/board_note.json";
const NOTE_MAX_LENGTH = 500;

export async function onRequestPost({ request }) {
  const email = request.headers.get("Cf-Access-Authenticated-User-Email");
  if (!email) {
    // Fail closed rather than saving unattributed -- deliberate, not a bug guard.
    // (Access misconfigured for this path, or a non-interactive/service-token
    // request with no user identity.)
    return new Response(
      JSON.stringify({ error: "Missing Cloudflare Access identity; cannot attribute save." }),
      { status: 401, headers: { "Content-Type": "application/json" } }
    );
  }
  const savedBy = email.split("@")[0];

  let text;
  try {
    const body = await request.json();
    text = String(body?.text ?? "").slice(0, NOTE_MAX_LENGTH);
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON body" }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  let upstream;
  try {
    upstream = await fetch(`${RAILWAY_NOTE_URL}?board=scmc`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, saved_by: savedBy }),
    });
  } catch {
    return new Response(JSON.stringify({ error: "Upstream fetch failed" }), {
      status: 502,
      headers: { "Content-Type": "application/json" },
    });
  }

  const bodyText = await upstream.text();
  return new Response(bodyText, {
    status: upstream.status,
    headers: { "Content-Type": "application/json" },
  });
}
