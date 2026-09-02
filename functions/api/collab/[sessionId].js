function cleanSessionId(value) {
  return `${value || ""}`.trim().replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96);
}

function getCollabWorkerUrl(context, sessionId) {
  const baseUrl =
    context.env.MOM_COLLAB_WORKER_URL ||
    "https://generate-mom-collab-worker.alex-marcello08.workers.dev";
  const requestUrl = new URL(context.request.url);
  const workerUrl = new URL(`/api/collab/${encodeURIComponent(sessionId)}`, baseUrl);
  workerUrl.search = requestUrl.search;
  return workerUrl;
}

export async function onRequest(context) {
  const sessionId = cleanSessionId(context.params.sessionId);
  if (!sessionId) {
    return new Response("Missing sessionId", { status: 400 });
  }

  if (context.request.headers.get("Upgrade")?.toLowerCase() !== "websocket") {
    return new Response("Expected WebSocket upgrade", { status: 426 });
  }

  return fetch(new Request(getCollabWorkerUrl(context, sessionId), context.request));
}
