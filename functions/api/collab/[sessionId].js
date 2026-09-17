const DEFAULT_WORKER_URL = "https://generate-mom-collab-worker.alex-marcello08.workers.dev";

function cleanSessionId(value) {
  return `${value || ""}`
    .trim()
    .replace(/[^a-zA-Z0-9_-]/g, "")
    .slice(0, 96);
}

function getCollabWorkerUrl(context, sessionId) {
  const baseUrl = context.env.MOM_COLLAB_WORKER_URL || DEFAULT_WORKER_URL;
  const requestUrl = new URL(context.request.url);
  const workerUrl = new URL(`/api/collab/${encodeURIComponent(sessionId)}`, baseUrl);
  workerUrl.search = requestUrl.search;
  return workerUrl;
}

/**
 * Single proxy for the collaboration Durable Object. It forwards WebSocket upgrades
 * (live editing), POST bodies (the unload beacon that flushes debounced edits), and
 * plain GETs (the deployment health probe) instead of rejecting everything that is
 * not an upgrade.
 */
export async function onRequest(context) {
  const sessionId = cleanSessionId(context.params.sessionId);
  if (!sessionId) {
    return new Response("Missing sessionId", { status: 400 });
  }

  const target = getCollabWorkerUrl(context, sessionId);
  const isUpgrade = context.request.headers.get("Upgrade")?.toLowerCase() === "websocket";

  if (isUpgrade) {
    return fetch(new Request(target, context.request));
  }

  const method = context.request.method.toUpperCase();
  const body = method === "GET" || method === "HEAD" ? undefined : await context.request.arrayBuffer();
  return fetch(target, {
    method,
    headers: {
      "Content-Type": context.request.headers.get("Content-Type") || "application/json",
    },
    body,
  });
}
