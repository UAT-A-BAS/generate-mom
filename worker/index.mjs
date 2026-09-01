import { createRemoteJWKSet, jwtVerify } from "jose";

const STATE_KEY = "mom-collab-latest-state";
const AUDIT_HEAD_KEY = "mom-collab-audit-head";
const AUDIT_KEY_PREFIX = "mom-collab-audit/";
const OUTBOX_KEY_PREFIX = "mom-collab-audit-outbox/";
const MAX_MESSAGE_BYTES = 900_000;
const MAX_TOKEN_LENGTH = 20_000;
const MAX_PATH_LENGTH = 1_024;
const MAX_PATH_DEPTH = 16;
const MAX_PATH_PART_LENGTH = 160;
const MAX_ARRAY_INDEX = 10_000;
const FORBIDDEN_PATH_PARTS = new Set(["__proto__", "prototype", "constructor"]);
const INTERNAL_IDENTITY_HEADER = "x-mom-verified-identity";
const INTERNAL_SESSION_HEADER = "x-mom-session-id";
const INTERNAL_AUTH_MODE_HEADER = "x-mom-auth-mode";
const SAFE_WEBSOCKET_PROTOCOL = "mom-collab";
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;
const discoveryCache = new Map();

class IdentityConfigError extends Error {}

export function isEntraRequired(env = {}) {
  return `${env.REQUIRE_ENTRA || ""}`.trim().toLowerCase() === "true";
}

export function parseCollabSessionId(request) {
  const url = new URL(request.url);
  const match = url.pathname.match(/\/api\/collab\/([^/?#]+)/);
  return match ? decodeURIComponent(match[1]).replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 96) : "";
}

function normalizePathPart(part) {
  return /^\d+$/.test(part) ? Number(part) : part;
}

export function validateDraftPath(path) {
  if (typeof path !== "string" || !path || path.length > MAX_PATH_LENGTH) {
    return { valid: false, reason: "INVALID_PATH", parts: [] };
  }

  const rawParts = path.split("/");
  if (
    rawParts.length === 0 ||
    rawParts.length > MAX_PATH_DEPTH ||
    rawParts.some((part) => !part || part.length > MAX_PATH_PART_LENGTH)
  ) {
    return { valid: false, reason: "INVALID_PATH", parts: [] };
  }

  for (const part of rawParts) {
    if (FORBIDDEN_PATH_PARTS.has(part.toLowerCase())) {
      return { valid: false, reason: "UNSAFE_PATH", parts: [] };
    }
    if (/^\d+$/.test(part) && Number(part) > MAX_ARRAY_INDEX) {
      return { valid: false, reason: "ARRAY_INDEX_TOO_LARGE", parts: [] };
    }
  }

  return {
    valid: true,
    reason: "",
    parts: rawParts.map(normalizePathPart),
    normalizedPath: rawParts.join("/"),
  };
}

function isObjectLike(value) {
  return value !== null && (typeof value === "object" || typeof value === "function");
}

function setPathByParts(root, parts, value) {
  if (!isObjectLike(root) || parts.length === 0) {
    return false;
  }

  let target = root;
  for (let index = 0; index < parts.length - 1; index += 1) {
    const key = parts[index];
    const nextKey = parts[index + 1];
    if (!Object.prototype.hasOwnProperty.call(target, key) || target[key] === null) {
      target[key] = typeof nextKey === "number" ? [] : Object.create(null);
    }
    if (!isObjectLike(target[key])) {
      return false;
    }
    target = target[key];
  }

  target[parts[parts.length - 1]] = value;
  return true;
}

export function setDraftPath(draft, path, value) {
  if (!draft || !path) {
    return draft;
  }

  const validation = validateDraftPath(`${path}`);
  if (!validation.valid) {
    return draft;
  }
  const { parts } = validation;

  if (parts[0] === "checklistRows" && parts.length >= 3) {
    const row = draft.checklistRows?.find((entry) => entry.id === `${parts[1]}`);
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

function safeJsonParse(raw) {
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function sanitizeText(value, maxLength) {
  return `${value || ""}`.replace(/[\u0000-\u001f\u007f]/g, "").trim().slice(0, maxLength);
}

function parseWebSocketProtocols(request) {
  return (request.headers.get("Sec-WebSocket-Protocol") || "")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
}

export function extractEntraProtocolToken(request) {
  const matches = parseWebSocketProtocols(request).filter((protocol) => protocol.startsWith("entra."));
  if (matches.length !== 1) {
    return "";
  }

  const token = matches[0].slice("entra.".length);
  if (
    token.length === 0 ||
    token.length > MAX_TOKEN_LENGTH ||
    !/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(token)
  ) {
    return "";
  }
  return token;
}

function extractBearerToken(request) {
  const match = (request.headers.get("Authorization") || "").match(/^Bearer\s+(.+)$/i);
  const token = match?.[1]?.trim() || "";
  if (
    token.length === 0 ||
    token.length > MAX_TOKEN_LENGTH ||
    !/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(token)
  ) {
    return "";
  }
  return token;
}

function requireGuid(value, settingName) {
  const normalized = `${value || ""}`.trim().toLowerCase();
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(normalized)) {
    throw new IdentityConfigError(`${settingName} must be a GUID`);
  }
  return normalized;
}

async function getMicrosoftJwks(tenantId) {
  if (discoveryCache.has(tenantId)) {
    return discoveryCache.get(tenantId);
  }

  const pending = (async () => {
    const issuer = `https://login.microsoftonline.com/${tenantId}/v2.0`;
    const discoveryUrl = `${issuer}/.well-known/openid-configuration`;
    const response = await fetch(discoveryUrl, {
      headers: { Accept: "application/json" },
      redirect: "error",
    });
    if (!response.ok) {
      throw new IdentityConfigError("Microsoft OpenID discovery failed");
    }

    const metadata = await response.json();
    if (metadata?.issuer !== issuer || typeof metadata?.jwks_uri !== "string") {
      throw new IdentityConfigError("Microsoft OpenID discovery metadata is invalid");
    }
    const jwksUrl = new URL(metadata.jwks_uri);
    if (
      jwksUrl.protocol !== "https:" ||
      jwksUrl.hostname !== "login.microsoftonline.com" ||
      jwksUrl.username ||
      jwksUrl.password ||
      jwksUrl.port
    ) {
      throw new IdentityConfigError("Microsoft JWKS URL is invalid");
    }
    return createRemoteJWKSet(jwksUrl);
  })();

  discoveryCache.set(tenantId, pending);
  try {
    return await pending;
  } catch (error) {
    discoveryCache.delete(tenantId);
    throw error;
  }
}

export async function verifyEntraToken(token, env, options = {}) {
  const tenantId = requireGuid(env.ENTRA_TENANT_ID, "ENTRA_TENANT_ID");
  const clientId = requireGuid(env.ENTRA_CLIENT_ID, "ENTRA_CLIENT_ID");
  const issuer = `https://login.microsoftonline.com/${tenantId}/v2.0`;
  const jwks = options.jwks || (await getMicrosoftJwks(tenantId));
  const { payload } = await jwtVerify(token, jwks, {
    algorithms: ["RS256"],
    issuer,
    audience: clientId,
    clockTolerance: 5,
  });

  const tokenTenant = sanitizeText(payload.tid, 64).toLowerCase();
  const subject = sanitizeText(payload.oid || payload.sub, 160);
  const email = sanitizeText(payload.preferred_username || payload.email, 254).toLowerCase();
  const displayName = sanitizeText(payload.name || email.split("@")[0] || "Microsoft user", 120);
  if (tokenTenant !== tenantId || !subject || !email) {
    throw new Error("Required Microsoft identity claims are missing");
  }

  return Object.freeze({
    provider: "entra",
    tenantId,
    subject,
    email,
    displayName,
  });
}

function bytesToBase64Url(bytes) {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlToBytes(value) {
  const base64 = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  const binary = atob(base64);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function encodeInternalIdentity(identity) {
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(identity)));
}

function decodeInternalIdentity(raw) {
  if (!raw || raw.length > 2_048) {
    return null;
  }
  try {
    const parsed = JSON.parse(new TextDecoder().decode(base64UrlToBytes(raw)));
    if (
      parsed?.provider !== "entra" ||
      !parsed.tenantId ||
      !parsed.subject ||
      !parsed.email ||
      !parsed.displayName
    ) {
      return null;
    }
    return Object.freeze({
      provider: "entra",
      tenantId: sanitizeText(parsed.tenantId, 64),
      subject: sanitizeText(parsed.subject, 160),
      email: sanitizeText(parsed.email, 254),
      displayName: sanitizeText(parsed.displayName, 120),
    });
  } catch (error) {
    return null;
  }
}

export function createSanitizedDurableObjectRequest(request, { identity = null, sessionId = "" } = {}) {
  const headers = new Headers(request.headers);
  for (const header of [
    "Authorization",
    "Cookie",
    "Cf-Access-Jwt-Assertion",
    INTERNAL_IDENTITY_HEADER,
    INTERNAL_SESSION_HEADER,
    INTERNAL_AUTH_MODE_HEADER,
    "Sec-WebSocket-Protocol",
  ]) {
    headers.delete(header);
  }

  if (parseWebSocketProtocols(request).includes(SAFE_WEBSOCKET_PROTOCOL)) {
    headers.set("Sec-WebSocket-Protocol", SAFE_WEBSOCKET_PROTOCOL);
  }
  headers.set(INTERNAL_SESSION_HEADER, sanitizeText(sessionId, 96));
  headers.set(INTERNAL_AUTH_MODE_HEADER, identity ? "entra" : "legacy");
  if (identity) {
    headers.set(INTERNAL_IDENTITY_HEADER, encodeInternalIdentity(identity));
  }

  return new Request(request, { headers });
}

function getClientMeta(request) {
  const url = new URL(request.url);
  const identity = decodeInternalIdentity(request.headers.get(INTERNAL_IDENTITY_HEADER));
  if (identity) {
    return {
      clientId: crypto.randomUUID(),
      editorName: identity.displayName,
      identity,
      authMode: "entra",
    };
  }
  return {
    clientId: (url.searchParams.get("clientId") || "").slice(0, 120),
    editorName: (url.searchParams.get("editorName") || "Editor").slice(0, 80),
    identity: null,
    authMode: "legacy",
  };
}

export function applyMessageOwnership(incoming, meta, latestVersion, now = () => new Date().toISOString()) {
  const message = { ...incoming };
  if (meta.authMode === "entra") {
    message.clientId = meta.clientId;
    message.editorName = meta.editorName;
    message.updatedAt = now();
    message.version = Number(latestVersion) + 1;
  } else {
    message.clientId = `${message.clientId || meta.clientId}`.slice(0, 120);
    message.editorName = `${message.editorName || meta.editorName}`.slice(0, 80);
    message.updatedAt = message.updatedAt || now();
    message.version = Math.max(Number(message.version) || 0, Number(latestVersion) + 1);
  }
  return message;
}

export function getVerifiedIdentityInit(meta) {
  if (!meta?.identity) {
    return {};
  }
  return {
    serverClientId: sanitizeText(meta.clientId, 120),
    identity: {
      email: sanitizeText(meta.identity.email, 254),
      displayName: sanitizeText(meta.identity.displayName, 120),
      verified: true,
    },
  };
}

async function sha256(value) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return bytesToBase64Url(new Uint8Array(digest));
}

export async function createAuditEvent({
  sequence,
  previousHash = "",
  sessionId,
  authMode,
  identity,
  clientId,
  editorName,
  action,
  path = "",
  baseVersion = 0,
  version,
  occurredAt,
  auditId = crypto.randomUUID(),
}) {
  const event = {
    schemaVersion: 1,
    auditId: sanitizeText(auditId, 80),
    sequence: Number(sequence),
    sessionId: sanitizeText(sessionId, 96),
    authMode: authMode === "entra" ? "entra" : "legacy",
    actor: {
      subject: identity ? sanitizeText(identity.subject, 160) : "",
      email: identity ? sanitizeText(identity.email, 254) : "",
      displayName: identity ? sanitizeText(identity.displayName, 120) : sanitizeText(editorName, 120),
    },
    clientId: sanitizeText(clientId, 120),
    action: action === "full" ? "full" : "patch",
    path: action === "patch" ? sanitizeText(path, MAX_PATH_LENGTH) : "",
    baseVersion: Number(baseVersion) || 0,
    version: Number(version),
    occurredAt: sanitizeText(occurredAt, 40),
    previousHash: sanitizeText(previousHash, 64),
  };
  return Object.freeze({ ...event, hash: await sha256(JSON.stringify(event)) });
}

function cloneJson(value) {
  return structuredClone(value);
}

function jsonResponse(body, status, headers = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...headers },
  });
}

export class MomCollabSession {
  constructor(state, env) {
    this.state = state;
    this.env = env;
    this.clients = new Map();
    this.latestPayload = null;
    this.latestVersion = 0;
    this.latestUpdatedAt = "";
    this.createdAt = 0;
    this.expiresAt = 0;
    this.auditHead = { sequence: 0, hash: "" };
    this.mutationQueue = Promise.resolve();
    this.ready = this.loadState();
  }

  async loadState() {
    const [stored, auditHead] = await Promise.all([
      this.state.storage.get(STATE_KEY),
      this.state.storage.get(AUDIT_HEAD_KEY),
    ]);
    if (stored) {
      this.latestPayload = stored.payload || null;
      this.latestVersion = Number(stored.version) || 0;
      this.latestUpdatedAt = stored.updatedAt || "";
      this.createdAt = Number(stored.createdAt) || 0;
      this.expiresAt = Number(stored.expiresAt) || 0;
      if (this.latestPayload && !this.expiresAt) {
        const fallback = this.latestUpdatedAt ? Date.parse(this.latestUpdatedAt) : Date.now();
        this.createdAt = Number.isFinite(fallback) ? fallback : Date.now();
        this.expiresAt = this.createdAt + SESSION_TTL_MS;
      }
    }
    if (auditHead) {
      this.auditHead = {
        sequence: Number(auditHead.sequence) || 0,
        hash: sanitizeText(auditHead.hash, 64),
      };
    }
  }

  async alarm() {
    await this.state.storage.deleteAll();
    this.latestPayload = null;
    this.latestVersion = 0;
    this.latestUpdatedAt = "";
    this.createdAt = 0;
    this.expiresAt = 0;
    this.auditHead = { sequence: 0, hash: "" };
  }

  async commitMutation({ payload, version, updatedAt, meta, action, path, baseVersion }) {
    const sequence = this.auditHead.sequence + 1;
    const auditOccurredAt = new Date().toISOString();
    const event = await createAuditEvent({
      sequence,
      previousHash: this.auditHead.hash,
      sessionId: this.sessionId,
      authMode: meta.authMode,
      identity: meta.identity,
      clientId: meta.clientId,
      editorName: meta.editorName,
      action,
      path,
      baseVersion,
      version,
      occurredAt: auditOccurredAt,
    });
    if (!this.expiresAt) {
      this.createdAt = Date.now();
      this.expiresAt = this.createdAt + SESSION_TTL_MS;
    }
    const storedState = { payload, version, updatedAt, createdAt: this.createdAt, expiresAt: this.expiresAt };
    const auditKey = `${AUDIT_KEY_PREFIX}${`${sequence}`.padStart(16, "0")}-${event.auditId}`;
    const outboxKey = `${OUTBOX_KEY_PREFIX}${`${sequence}`.padStart(16, "0")}-${event.auditId}`;

    await this.state.storage.transaction(async (transaction) => {
      await transaction.put(STATE_KEY, storedState);
      await transaction.put(AUDIT_HEAD_KEY, { sequence, hash: event.hash });
      await transaction.put(auditKey, event);
      if (this.env.MOM_COLLAB_AUDIT) {
        await transaction.put(outboxKey, event);
      }
    });

    this.latestPayload = payload;
    this.latestVersion = version;
    this.latestUpdatedAt = updatedAt;
    this.auditHead = { sequence, hash: event.hash };
    try {
      const existingAlarm = await this.state.storage.getAlarm();
      if (existingAlarm == null && this.expiresAt) {
        await this.state.storage.setAlarm(this.expiresAt);
      }
    } catch {}
    return event;
  }

  async flushAuditOutbox() {
    if (!this.env.MOM_COLLAB_AUDIT) {
      return 0;
    }
    const pending = await this.state.storage.list({ prefix: OUTBOX_KEY_PREFIX, limit: 25 });
    let written = 0;
    for (const [storageKey, event] of pending) {
      const objectKey = `${event.sessionId || "unknown"}/${event.occurredAt.replace(/[:.]/g, "-")}-${event.sequence}-${event.auditId}.json`;
      try {
        await this.env.MOM_COLLAB_AUDIT.put(objectKey, JSON.stringify(event), {
          httpMetadata: { contentType: "application/json" },
          customMetadata: {
            sessionId: event.sessionId || "unknown",
            sequence: `${event.sequence}`,
            hash: event.hash,
          },
        });
        await this.state.storage.delete(storageKey);
        written += 1;
      } catch (error) {
        // Leave the transactional outbox record for the next connection or mutation.
      }
    }
    return written;
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

  async handleSocketMessage(server, meta, raw) {
    if (!raw || raw.length > MAX_MESSAGE_BYTES) {
      return;
    }

    const incoming = safeJsonParse(raw);
    if (!incoming || (meta.authMode !== "entra" && incoming.clientId === undefined)) {
      return;
    }

    const message = applyMessageOwnership(incoming, meta, this.latestVersion);

    if (message.type === "hello") {
      this.send(server, this.getPresenceMessage());
      return;
    }

    let nextPayload;
    let normalizedPath = "";
    if (message.type === "full") {
      if (!shouldAcceptFullMessage(this.latestPayload, this.latestVersion, message)) {
        this.send(server, {
          type: "full",
          clientId: "server",
          value: this.latestPayload,
          updatedAt: this.latestUpdatedAt,
          version: this.latestVersion,
          conflict: true,
        });
        return;
      }
      if (meta.authMode === "entra" && (!message.value || typeof message.value !== "object")) {
        this.send(server, { type: "error", code: "INVALID_FULL_PAYLOAD" });
        return;
      }
      nextPayload = cloneJson(message.value || null);
    } else if (message.type === "patch" && this.latestPayload && message.path) {
      const validation = validateDraftPath(`${message.path}`);
      if (!validation.valid) {
        this.send(server, { type: "error", code: validation.reason });
        return;
      }
      normalizedPath = validation.normalizedPath;
      nextPayload = cloneJson(this.latestPayload);
      setDraftPath(nextPayload, normalizedPath, message.value);
      message.path = normalizedPath;
    } else {
      return;
    }

    await this.commitMutation({
      payload: nextPayload,
      version: message.version,
      updatedAt: message.updatedAt,
      meta,
      action: message.type,
      path: normalizedPath,
      baseVersion: message.baseVersion,
    });
    this.send(server, {
      type: "ack",
      clientId: message.clientId,
      path: normalizedPath,
      updatedAt: message.updatedAt,
      version: message.version,
      expiresAt: this.expiresAt || 0,
      createdAt: this.createdAt || 0,
    });
    this.broadcast(message, server);
    await this.flushAuditOutbox();
  }

  async fetch(request) {
    await this.ready;
    this.sessionId = sanitizeText(request.headers.get(INTERNAL_SESSION_HEADER), 96);
    // Fixed 7-day expiry: lazy check + backfill + alarm ensure
    const now = Date.now();
    if (this.expiresAt && now >= this.expiresAt) {
      try { await this.state.storage.deleteAll(); } catch {}
      this.latestPayload = null;
      this.latestVersion = 0;
      this.latestUpdatedAt = "";
      this.createdAt = 0;
      this.expiresAt = 0;
      this.auditHead = { sequence: 0, hash: "" };
    } else if (this.latestPayload && !this.expiresAt) {
      const fallback = this.latestUpdatedAt ? Date.parse(this.latestUpdatedAt) : now;
      this.createdAt = Number.isFinite(fallback) ? fallback : now;
      this.expiresAt = this.createdAt + SESSION_TTL_MS;
      try {
        await this.state.storage.put(STATE_KEY, {
          payload: this.latestPayload,
          version: this.latestVersion,
          updatedAt: this.latestUpdatedAt,
          createdAt: this.createdAt,
          expiresAt: this.expiresAt,
        });
      } catch {}
    }
    if (this.expiresAt) {
      try {
        const existingAlarm = await this.state.storage.getAlarm();
        if (existingAlarm == null) await this.state.storage.setAlarm(this.expiresAt);
      } catch {}
    }

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
      createdAt: this.createdAt || 0,
      expiresAt: this.expiresAt || 0,
      users: this.clients.size,
      needsPayload: !this.latestPayload && this.clients.size === 1,
      ...getVerifiedIdentityInit(meta),
    });
    this.broadcast(this.getPresenceMessage());
    if (this.state.waitUntil) {
      this.state.waitUntil(this.flushAuditOutbox());
    }

    server.addEventListener("message", (event) => {
      const raw = typeof event.data === "string" ? event.data : "";
      this.mutationQueue = this.mutationQueue
        .then(() => this.handleSocketMessage(server, meta, raw))
        .catch(() => this.send(server, { type: "error", code: "SERVER_WRITE_FAILED" }));
      if (this.state.waitUntil) {
        this.state.waitUntil(this.mutationQueue);
      }
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

    const responseHeaders = new Headers();
    if (request.headers.get("Sec-WebSocket-Protocol") === SAFE_WEBSOCKET_PROTOCOL) {
      responseHeaders.set("Sec-WebSocket-Protocol", SAFE_WEBSOCKET_PROTOCOL);
    }
    return new Response(null, { status: 101, webSocket: client, headers: responseHeaders });
  }
}

export async function handleWorkerRequest(request, env, options = {}) {
  const url = new URL(request.url);
  const isHealth = url.pathname === "/health";
  const sessionId = parseCollabSessionId(request);
  if (!isHealth && !sessionId) {
    return new Response("Not found", { status: 404 });
  }

  let identity = null;
  if (isEntraRequired(env)) {
    const token = isHealth ? extractBearerToken(request) : extractEntraProtocolToken(request);
    if (!token) {
      return jsonResponse({ ok: false, error: "Microsoft sign-in required" }, 401);
    }
    try {
      identity = await verifyEntraToken(token, env, options);
    } catch (error) {
      const status = error instanceof IdentityConfigError ? 503 : 401;
      return jsonResponse(
        { ok: false, error: status === 503 ? "Microsoft identity configuration unavailable" : "Invalid Microsoft identity token" },
        status
      );
    }
  }

  if (isHealth) {
    return jsonResponse(
      {
        ok: true,
        mode: identity ? "entra" : "legacy",
        ...(identity ? { identity: { email: identity.email, displayName: identity.displayName } } : {}),
      },
      200,
      { "Cache-Control": "no-store" }
    );
  }

  if (!env.MOM_COLLAB_SESSIONS) {
    return new Response("Missing MOM_COLLAB_SESSIONS binding", { status: 500 });
  }

  const objectId = env.MOM_COLLAB_SESSIONS.idFromName(sessionId);
  const sanitizedRequest = createSanitizedDurableObjectRequest(request, { identity, sessionId });
  return env.MOM_COLLAB_SESSIONS.get(objectId).fetch(sanitizedRequest);
}

export default {
  fetch(request, env) {
    return handleWorkerRequest(request, env);
  },
};
