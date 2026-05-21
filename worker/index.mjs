const STATE_KEY = "mom-collab-latest-state";
const MAX_MESSAGE_BYTES = 900_000;

export function parseCollabSessionId(request) {
  const url = new URL(request.url);
  const match = url.pathname.match(/\/api\/collab\/([^/?#]+)/);
  return match ? decodeURIComponent(match[1]).replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96) : "";
}

function normalizePathPart(part) {
  return /^\d+$/.test(part) ? Number(part) : part;
}

function setPathByParts(root, parts, value) {
  let target = root;
  for (let index = 0; index < parts.length - 1; index += 1) {
    const key = normalizePathPart(parts[index]);
    const nextKey = normalizePathPart(parts[index + 1]);
    if (target[key] === undefined || target[key] === null) {
      target[key] = typeof nextKey === "number" ? [] : {};
    }
    target = target[key];
  }

  target[normalizePathPart(parts[parts.length - 1])] = value;
}

export function setDraftPath(draft, path, value) {
  if (!draft || !path) {
    return draft;
  }

  const parts = `${path}`.split("/").filter(Boolean);
  if (parts[0] === "checklistRows" && parts.length >= 3) {
    const row = draft.checklistRows?.find((entry) => entry.id === parts[1]);
    if (!row) {
      return draft;
    }

    setPathByParts(row, parts.slice(2), value);
    return draft;
  }

  setPathByParts(draft, parts, value);
  return draft;
}

function getClientMeta(request) {
  const url = new URL(request.url);
  return {
    clientId: (url.searchParams.get("clientId") || "").slice(0, 120),
    editorName: (url.searchParams.get("editorName") || "Editor").slice(0, 80),
  };
}

function safeJsonParse(raw) {
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

export class MomCollabSession {
  constructor(state, env) {
    this.state = state;
    this.env = env;
    this.clients = new Map();
    this.latestPayload = null;
    this.latestVersion = 0;
    this.latestUpdatedAt = "";
    this.ready = this.loadState();
  }

  async loadState() {
    const stored = await this.state.storage.get(STATE_KEY);
    if (!stored) {
      return;
    }

    this.latestPayload = stored.payload || null;
    this.latestVersion = Number(stored.version) || 0;
    this.latestUpdatedAt = stored.updatedAt || "";
  }

  async persistState() {
    await this.state.storage.put(STATE_KEY, {
      payload: this.latestPayload,
      version: this.latestVersion,
      updatedAt: this.latestUpdatedAt,
    });
  }

  getPresenceMessage() {
    return {
      type: "presence",
      users: this.clients.size,
      updatedAt: new Date().toISOString(),
    };
  }

  send(socket, message) {
    try {
      socket.send(JSON.stringify(message));
    } catch (error) {
      this.clients.delete(socket);
    }
  }

  broadcast(message, sourceSocket = null) {
    for (const socket of this.clients.keys()) {
      if (socket !== sourceSocket) {
        this.send(socket, message);
      }
    }
  }

  async fetch(request) {
    await this.ready;

    if (request.headers.get("Upgrade")?.toLowerCase() !== "websocket") {
      return new Response("Expected WebSocket upgrade", { status: 426 });
    }

    const pair = new WebSocketPair();
    const [client, server] = Object.values(pair);
    const meta = getClientMeta(request);

    server.accept();
    this.clients.set(server, meta);
    this.send(server, {
      type: "init",
      payload: this.latestPayload,
      version: this.latestVersion,
      updatedAt: this.latestUpdatedAt,
      users: this.clients.size,
      needsPayload: !this.latestPayload && this.clients.size === 1,
    });
    this.broadcast(this.getPresenceMessage());

    server.addEventListener("message", async (event) => {
      const raw = typeof event.data === "string" ? event.data : "";
      if (!raw || raw.length > MAX_MESSAGE_BYTES) {
        return;
      }

      const message = safeJsonParse(raw);
      if (!message || message.clientId === undefined) {
        return;
      }

      message.clientId = `${message.clientId || meta.clientId}`.slice(0, 120);
      message.editorName = `${message.editorName || meta.editorName}`.slice(0, 80);
      message.updatedAt = message.updatedAt || new Date().toISOString();
      message.version = Math.max(Number(message.version) || 0, this.latestVersion + 1);

      if (message.type === "hello") {
        this.send(server, this.getPresenceMessage());
        return;
      }

      if (message.type === "full") {
        this.latestPayload = message.value || null;
      } else if (message.type === "patch" && this.latestPayload && message.path) {
        setDraftPath(this.latestPayload, message.path, message.value);
      } else {
        return;
      }

      this.latestVersion = message.version;
      this.latestUpdatedAt = message.updatedAt;
      await this.persistState();
      this.broadcast(message, server);
    });

    const cleanup = () => {
      if (!this.clients.has(server)) {
        return;
      }
      this.clients.delete(server);
      this.broadcast(this.getPresenceMessage());
    };

    server.addEventListener("close", cleanup);
    server.addEventListener("error", cleanup);

    return new Response(null, { status: 101, webSocket: client });
  }
}

export default {
  async fetch(request, env) {
    const sessionId = parseCollabSessionId(request);
    if (!sessionId) {
      return new Response("Not found", { status: 404 });
    }

    if (!env.MOM_COLLAB_SESSIONS) {
      return new Response("Missing MOM_COLLAB_SESSIONS binding", { status: 500 });
    }

    const objectId = env.MOM_COLLAB_SESSIONS.idFromName(sessionId);
    return env.MOM_COLLAB_SESSIONS.get(objectId).fetch(request);
  },
};
