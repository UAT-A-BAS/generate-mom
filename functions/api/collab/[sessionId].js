function cleanSessionId(value) {
  return `${value || ""}`.trim().replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96);
}

export async function onRequest(context) {
  const sessionId = cleanSessionId(context.params.sessionId);
  if (!sessionId) {
    return new Response("Missing sessionId", { status: 400 });
  }

  if (context.request.headers.get("Upgrade")?.toLowerCase() !== "websocket") {
    return new Response("Expected WebSocket upgrade", { status: 426 });
  }

  if (!context.env.MOM_COLLAB_SESSIONS) {
    return new Response("Missing MOM_COLLAB_SESSIONS binding", { status: 500 });
  }

  const objectId = context.env.MOM_COLLAB_SESSIONS.idFromName(sessionId);
  const object = context.env.MOM_COLLAB_SESSIONS.get(objectId);
  return object.fetch(context.request);
}
