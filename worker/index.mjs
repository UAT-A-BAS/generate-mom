const STATE_KEY = "mom-collab-latest-state";
const MAX_MESSAGE_BYTES = 900_000;
const MAX_OP_LOG = 4000;
const MAX_OPS_PER_MESSAGE = 400;

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Max-Age": "86400",
};

export function parseCollabSessionId(request) {
  const url = new URL(request.url);
  const match = url.pathname.match(/\/api\/collab\/([^/?#]+)/);
  return match ? decodeURIComponent(match[1]).replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96) : "";
}

function normalizePathPart(part) {
  return /^\d+$/.test(part) ? Number(part) : part;
}

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function readPathByParts(root, parts) {
  let target = root;
  for (const part of parts) {
    if (target === null || target === undefined || typeof target !== "object") {
      return { exists: false, value: undefined };
    }
    const key = normalizePathPart(part);
    if (!(key in target)) {
      return { exists: false, value: undefined };
    }
    target = target[key];
  }
  return { exists: true, value: target };
}

export function readDraftPath(draft, path) {
  if (!draft || !path) {
    return { exists: false, value: undefined };
  }

  const parts = `${path}`.split("/").filter(Boolean);
  if (parts[0] === "checklistRows" && parts.length >= 3) {
    const row = draft.checklistRows?.find((entry) => entry.id === parts[1]);
    if (!row) {
      return { exists: false, value: undefined };
    }
    return readPathByParts(row, parts.slice(2));
  }

  return readPathByParts(draft, parts);
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

export function shouldAcceptFullMessage(latestPayload, latestVersion, message) {
  if (!latestPayload) {
    return true;
  }

  return message?.replace === true && Number(message.baseVersion) === Number(latestVersion);
}

function normalizeOps(message) {
  if (Array.isArray(message?.ops) && message.ops.length) {
    return message.ops
      .filter((op) => op && typeof op.path === "string" && op.path)
      .slice(0, MAX_OPS_PER_MESSAGE)
      // `kind` is preserved so receivers know an array replacement needs a re-render
      // rather than a plain field assignment.
      .map((op) => (op.kind ? { path: op.path, value: op.value, kind: op.kind } : { path: op.path, value: op.value }));
  }

  if (typeof message?.path === "string" && message.path) {
    return [{ path: message.path, value: message.value }];
  }

  return [];
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

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...CORS_HEADERS },
  });
}

export class MomCollabSession {
  constructor(state, env) {
    this.state = state;
    this.env = env;
    // Only used when the platform has no WebSocket hibernation support (unit tests pass a
    // plain state stub). In production the authoritative socket list lives in the
    // platform runtime so the object can be evicted from memory while a room is idle.
    this.clients = new Map();
    this.latestPayload = null;
    this.latestVersion = 0;
    this.latestUpdatedAt = "";
    this.opLog = [];
    this.ready = this.loadState();
  }

  /**
   * Hibernation-aware socket inventory. `state.getWebSockets()` reports the sockets the
   * platform is holding on our behalf, which is what lets the Durable Object be evicted
   * from memory (and stop billing for duration) while editors keep their connection open.
   */
  listSockets() {
    if (typeof this.state.getWebSockets === "function") {
      return this.state.getWebSockets();
    }
    return [...this.clients.keys()];
  }

  socketMeta(socket) {
    try {
      const attachment = socket.deserializeAttachment?.();
      if (attachment) {
        return attachment;
      }
    } catch (error) {
      // Older sockets may carry no attachment; fall back to the in-memory map.
    }
    return this.clients.get(socket) || { clientId: "", editorName: "Editor" };
  }

  presenceCount() {
    return this.listSockets().length;
  }

  async loadState() {
    const stored = await this.state.storage.get(STATE_KEY);
    if (!stored) {
      return;
    }

    this.latestPayload = stored.payload || null;
    this.latestVersion = Number(stored.version) || 0;
    this.latestUpdatedAt = stored.updatedAt || "";
    this.opLog = Array.isArray(stored.opLog) ? stored.opLog.slice(-MAX_OP_LOG) : [];
  }

  async persistState() {
    await this.state.storage.put(STATE_KEY, {
      payload: this.latestPayload,
      version: this.latestVersion,
      updatedAt: this.latestUpdatedAt,
      opLog: this.opLog.slice(-MAX_OP_LOG),
    });
  }

  getPresenceMessage() {
    return {
      type: "presence",
      users: this.presenceCount(),
      seq: this.latestVersion,
      updatedAt: new Date().toISOString(),
    };
  }

  send(socket, message) {
    try {
      socket.send(JSON.stringify(message));
      return true;
    } catch (error) {
      this.clients.delete(socket);
      return false;
    }
  }

  broadcast(message, sourceSocket = null) {
    for (const socket of this.listSockets()) {
      if (socket !== sourceSocket) {
        this.send(socket, message);
      }
    }
  }

  broadcastPresence() {
    this.broadcast(this.getPresenceMessage());
  }

  nextSeq() {
    this.latestVersion += 1;
    return this.latestVersion;
  }

  recordOps(ops, seq, message) {
    const entry = {
      seq,
      clientId: message.clientId,
      editorName: message.editorName,
      updatedAt: message.updatedAt,
      mutationId: message.mutationId || "",
      ops,
    };

    this.opLog.push(entry);
    if (this.opLog.length > MAX_OP_LOG) {
      this.opLog = this.opLog.slice(-MAX_OP_LOG);
    }
  }

  /**
   * Field-level last-write-wins. Every op lands on top of the current server payload,
   * so two editors touching different fields cannot clobber each other. Ops whose
   * parent row was removed concurrently are reported as rejected so the sender
   * resyncs instead of silently diverging.
   */
  applyOps(ops) {
    if (!Array.isArray(ops) || !ops.length) {
      return [];
    }

    if (!this.latestPayload) {
      return ops.map((op) => ({ path: op.path, applied: false, reason: "empty-session" }));
    }

    const draft = cloneJson(this.latestPayload);
    const results = [];

    for (const op of ops) {
      const probe = readDraftPath(draft, op.path);
      if (!probe.exists) {
        results.push({ path: op.path, applied: false, reason: "path-missing" });
        continue;
      }

      setDraftPath(draft, op.path, op.value);
      results.push({ path: op.path, applied: true });
    }

    this.latestPayload = draft;
    return results;
  }

  replayEntriesSince(sinceSeq) {
    const from = Number(sinceSeq) || 0;
    return this.opLog.filter((entry) => entry.seq > from);
  }

  async handleFullMessage(message, server) {
    if (!shouldAcceptFullMessage(this.latestPayload, this.latestVersion, message)) {
      this.send(server, {
        type: "full",
        clientId: "server",
        value: this.latestPayload,
        updatedAt: this.latestUpdatedAt,
        version: this.latestVersion,
        seq: this.latestVersion,
        conflict: true,
      });
      return;
    }

    const seq = this.nextSeq();
    this.latestPayload = message.value || null;
    this.latestUpdatedAt = message.updatedAt;
    this.recordOps([{ path: "draft", value: null }], seq, message);
    await this.persistState();

    this.send(server, {
      type: "ack",
      clientId: message.clientId,
      mutationId: message.mutationId || "",
      path: message.path || "draft",
      ops: [{ path: message.path || "draft", applied: true }],
      updatedAt: message.updatedAt,
      version: seq,
      seq,
    });
    this.broadcast(
      {
        type: "full",
        clientId: message.clientId,
        editorName: message.editorName,
        value: this.latestPayload,
        updatedAt: message.updatedAt,
        version: seq,
        seq,
      },
      server
    );
  }

  async handlePatchMessage(message, server) {
    const ops = normalizeOps(message);
    if (!ops.length) {
      return;
    }

    const results = this.applyOps(ops);
    const appliedPaths = new Set(results.filter((result) => result.applied).map((r) => r.path));
    if (!appliedPaths.size) {
      this.send(server, {
        type: "ack",
        clientId: message.clientId,
        mutationId: message.mutationId || "",
        path: message.path || "",
        ops: results,
        updatedAt: message.updatedAt,
        version: this.latestVersion,
        seq: this.latestVersion,
        rejected: true,
      });
      return;
    }

    const seq = this.nextSeq();
    const appliedOps = ops.filter((op) => appliedPaths.has(op.path));
    this.latestUpdatedAt = message.updatedAt;
    this.recordOps(appliedOps, seq, message);
    await this.persistState();

    this.send(server, {
      type: "ack",
      clientId: message.clientId,
      mutationId: message.mutationId || "",
      path: appliedOps.length === 1 ? appliedOps[0].path : "",
      ops: results,
      updatedAt: message.updatedAt,
      version: seq,
      seq,
    });
    this.broadcast(
      {
        type: "patch",
        clientId: message.clientId,
        editorName: message.editorName,
        path: appliedOps.length === 1 ? appliedOps[0].path : "",
        ops: appliedOps,
        updatedAt: message.updatedAt,
        version: seq,
        seq,
      },
      server
    );
  }

  async handleResync(message, server) {
    const sinceSeq = Number(message.sinceSeq) || 0;
    const entries = this.replayEntriesSince(sinceSeq);
    const canReplay =
      entries.length > 0 &&
      Boolean(this.latestPayload) &&
      sinceSeq > 0 &&
      sinceSeq < this.latestVersion &&
      entries.length <= MAX_OPS_PER_MESSAGE;

    if (canReplay) {
      this.send(server, {
        type: "sync",
        mode: "replay",
        fromSeq: sinceSeq,
        seq: this.latestVersion,
        entries: entries.map((entry) => ({
          seq: entry.seq,
          clientId: entry.clientId,
          editorName: entry.editorName,
          ops: entry.ops.filter((op) => op.path && op.path !== "draft"),
          updatedAt: entry.updatedAt,
        })),
        updatedAt: this.latestUpdatedAt,
      });
      return;
    }

    this.send(server, {
      type: "sync",
      mode: "snapshot",
      value: this.latestPayload,
      seq: this.latestVersion,
      updatedAt: this.latestUpdatedAt,
      users: this.presenceCount(),
    });
  }

  async handleBeacon(request) {
    const body = safeJsonParse(await request.text());
    if (!body) {
      return jsonResponse({ ok: false, error: "invalid-payload" }, 400);
    }

    const ops = normalizeOps(body).filter(
      (op) => op.path && op.path !== "draft" && op.path !== "presence"
    );
    if (!ops.length) {
      return jsonResponse({ ok: true, applied: 0, seq: this.latestVersion });
    }

    const results = this.applyOps(ops);
    const appliedPaths = new Set(results.filter((result) => result.applied).map((r) => r.path));
    if (!appliedPaths.size) {
      return jsonResponse({ ok: true, applied: 0, seq: this.latestVersion, results });
    }

    const seq = this.nextSeq();
    const appliedOps = ops.filter((op) => appliedPaths.has(op.path));
    this.latestUpdatedAt = new Date().toISOString();
    this.recordOps(appliedOps, seq, {
      clientId: body.clientId || "unknown",
      editorName: body.editorName || "Editor",
      updatedAt: this.latestUpdatedAt,
      mutationId: body.mutationId || "",
    });
    await this.persistState();
    this.broadcast({
      type: "patch",
      clientId: body.clientId || "unknown",
      editorName: body.editorName || "Editor",
      ops: appliedOps,
      updatedAt: this.latestUpdatedAt,
      version: seq,
      seq,
    });

    return jsonResponse({ ok: true, applied: appliedOps.length, seq, results });
  }

  async fetch(request) {
    await this.ready;

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (request.method === "POST") {
      return this.handleBeacon(request);
    }

    if (request.headers.get("Upgrade")?.toLowerCase() !== "websocket") {
      return jsonResponse({
        ok: true,
        service: "generate-mom-collab-worker",
        seq: this.latestVersion,
        users: this.presenceCount(),
        hasPayload: Boolean(this.latestPayload),
      });
    }

    const pair = new WebSocketPair();
    const [client, server] = Object.values(pair);
    const meta = getClientMeta(request);

    // Hibernation API: the platform holds the socket, so this object can be evicted from
    // memory while the room sits idle. Editor identity rides on the socket attachment so
    // it survives hibernation, and handlers below are woken on demand.
    const socketInfo = { ...meta, joinedAt: new Date().toISOString() };
    server.serializeAttachment(socketInfo);
    this.state.acceptWebSocket(server);
    this.clients.set(server, socketInfo);
    this.send(server, {
      type: "init",
      payload: this.latestPayload,
      version: this.latestVersion,
      seq: this.latestVersion,
      updatedAt: this.latestUpdatedAt,
      users: this.presenceCount(),
      // Any client joining a session that has no canonical draft yet may seed it. If two
      // editors race, the first accepted `full` wins and the loser receives the winner's
      // snapshot through the normal conflict response, so the room always self-heals.
      needsPayload: !this.latestPayload,
    });
    this.broadcastPresence();

    return new Response(null, { status: 101, webSocket: client });
  }

  /**
   * Woken by the runtime for every inbound frame, including after a hibernation cycle.
   * `await this.ready` is mandatory here: on wake the constructor has just re-run and the
   * canonical draft is still being loaded from storage.
   */
  async webSocketMessage(socket, rawData) {
    await this.ready;

    const raw = typeof rawData === "string" ? rawData : "";
    if (!raw || raw.length > MAX_MESSAGE_BYTES) {
      return;
    }

    const message = safeJsonParse(raw);
    if (!message || message.clientId === undefined) {
      return;
    }

    const meta = this.socketMeta(socket);
    message.clientId = `${message.clientId || meta.clientId}`.slice(0, 120);
    message.editorName = `${message.editorName || meta.editorName}`.slice(0, 80);
    message.updatedAt = message.updatedAt || new Date().toISOString();

    if (message.type === "hello" || message.type === "ping") {
      this.send(socket, {
        type: message.type === "ping" ? "pong" : "presence",
        seq: this.latestVersion,
        users: this.presenceCount(),
        updatedAt: new Date().toISOString(),
      });
      return;
    }

    if (message.type === "resync") {
      await this.handleResync(message, socket);
      return;
    }

    if (message.type === "full") {
      await this.handleFullMessage(message, socket);
      return;
    }

    if (message.type === "patch") {
      await this.handlePatchMessage(message, socket);
    }
  }

  async webSocketClose(socket) {
    await this.ready;
    this.clients.delete(socket);
    this.broadcastPresence();
    try {
      socket.close(1000, "closed");
    } catch (error) {
      // The platform may already have torn the socket down.
    }
  }

  async webSocketError(socket) {
    await this.ready;
    this.clients.delete(socket);
    this.broadcastPresence();
  }
}

export default {
  async fetch(request, env) {
    const sessionId = parseCollabSessionId(request);
    if (!sessionId) {
      return new Response("Not found", { status: 404, headers: CORS_HEADERS });
    }

    if (!env.MOM_COLLAB_SESSIONS) {
      return new Response("Missing MOM_COLLAB_SESSIONS binding", {
        status: 500,
        headers: CORS_HEADERS,
      });
    }

    const objectId = env.MOM_COLLAB_SESSIONS.idFromName(sessionId);
    return env.MOM_COLLAB_SESSIONS.get(objectId).fetch(request);
  },
};
