function cleanSessionId(value) {
  return `${value || ""}`.trim().replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96);
}

const PRODUCTION_WORKER_URL =
  "https://generate-mom-collab-worker.alex-marcello08.workers.dev";
const ACCESS_TEST_HOSTNAME = "generate-mom-entra-test.pages.dev";
const ACCESS_TEST_WORKER_URL =
  "https://generate-mom-collab-worker-entra-test.alex-marcello08.workers.dev";

export function resolveCollabWorkerUrl(context) {
  if (context.env.MOM_COLLAB_WORKER_URL) {
    return context.env.MOM_COLLAB_WORKER_URL;
  }

  const requestUrl = new URL(context.request.url);
  if (requestUrl.hostname === ACCESS_TEST_HOSTNAME) {
    return ACCESS_TEST_WORKER_URL;
  }

  return PRODUCTION_WORKER_URL;
}

function getCollabWorkerUrl(context, sessionId) {
  const baseUrl = resolveCollabWorkerUrl(context);
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
