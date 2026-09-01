import assert from "node:assert/strict";
import test from "node:test";
import {
  SignJWT,
  createLocalJWKSet,
  exportJWK,
  generateKeyPair,
} from "jose";
import {
  MomCollabSession,
  applyMessageOwnership,
  createAuditEvent,
  createSanitizedDurableObjectRequest,
  extractEntraProtocolToken,
  getVerifiedIdentityInit,
  handleWorkerRequest,
  setDraftPath,
  validateDraftPath,
  verifyEntraToken,
} from "../index.mjs";

const tenantId = "11111111-1111-4111-8111-111111111111";
const clientId = "22222222-2222-4222-8222-222222222222";
const issuer = `https://login.microsoftonline.com/${tenantId}/v2.0`;
const envIdentity = { ENTRA_TENANT_ID: tenantId, ENTRA_CLIENT_ID: clientId };
const { publicKey, privateKey } = await generateKeyPair("RS256");
const publicJwk = await exportJWK(publicKey);
publicJwk.kid = "pilot-test-key";
publicJwk.use = "sig";
publicJwk.alg = "RS256";
const jwks = createLocalJWKSet({ keys: [publicJwk] });

async function signToken({ claims = {}, tokenIssuer = issuer, audience = clientId } = {}) {
  return new SignJWT({
    tid: tenantId,
    oid: "33333333-3333-4333-8333-333333333333",
    preferred_username: "u075060@bca.co.id",
    name: "Alex Marcello",
    ...claims,
  })
    .setProtectedHeader({ alg: "RS256", kid: "pilot-test-key", typ: "JWT" })
    .setIssuer(tokenIssuer)
    .setAudience(audience)
    .setIssuedAt()
    .setExpirationTime("5m")
    .sign(privateKey);
}

function collaborationRequest(token, extraHeaders = {}) {
  return new Request("https://pilot.example/api/collab/session-1?clientId=attacker&editorName=Attacker", {
    headers: {
      Upgrade: "websocket",
      "Sec-WebSocket-Protocol": `${SAFE_PROTOCOL}, entra.${token}`,
      ...extraHeaders,
    },
  });
}

const SAFE_PROTOCOL = "mom-collab";

test("validates a fixed-tenant, fixed-audience Microsoft token", async () => {
  const identity = await verifyEntraToken(await signToken(), envIdentity, { jwks });
  assert.deepEqual(identity, {
    provider: "entra",
    tenantId,
    subject: "33333333-3333-4333-8333-333333333333",
    email: "u075060@bca.co.id",
    displayName: "Alex Marcello",
  });
});

test("accepts sub plus email and rejects wrong tenant, issuer, audience, or missing login", async () => {
  const accessIdentity = await verifyEntraToken(
    await signToken({ claims: { oid: undefined, sub: "subject-1", preferred_username: undefined, email: "USER@BCA.CO.ID" } }),
    envIdentity,
    { jwks }
  );
  assert.equal(accessIdentity.subject, "subject-1");
  assert.equal(accessIdentity.email, "user@bca.co.id");

  await assert.rejects(
    verifyEntraToken(await signToken({ claims: { tid: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa" } }), envIdentity, { jwks })
  );
  await assert.rejects(
    verifyEntraToken(await signToken({ tokenIssuer: "https://attacker.example/v2.0" }), envIdentity, { jwks })
  );
  await assert.rejects(
    verifyEntraToken(await signToken({ audience: "33333333-3333-4333-8333-333333333333" }), envIdentity, { jwks })
  );
  await assert.rejects(
    verifyEntraToken(
      await signToken({ claims: { preferred_username: undefined, email: undefined } }),
      envIdentity,
      { jwks }
    )
  );
});

test("extracts only one structurally valid entra WebSocket protocol token", async () => {
  const token = await signToken();
  assert.equal(extractEntraProtocolToken(collaborationRequest(token)), token);
  assert.equal(
    extractEntraProtocolToken(
      new Request("https://pilot.example/api/collab/a", {
        headers: { "Sec-WebSocket-Protocol": `entra.${token}, entra.${token}` },
      })
    ),
    ""
  );
  assert.equal(
    extractEntraProtocolToken(
      new Request("https://pilot.example/api/collab/a", {
        headers: { "Sec-WebSocket-Protocol": "entra.not-a-jwt" },
      })
    ),
    ""
  );
});

test("strips all credentials and spoofed internal headers before the Durable Object", async () => {
  const token = await signToken();
  const identity = await verifyEntraToken(token, envIdentity, { jwks });
  const request = collaborationRequest(token, {
    Authorization: `Bearer ${token}`,
    Cookie: "session=secret",
    "Cf-Access-Jwt-Assertion": "old-access-token",
    "x-mom-verified-identity": "spoofed",
    "x-mom-session-id": "spoofed",
    "x-mom-auth-mode": "spoofed",
  });
  const sanitized = createSanitizedDurableObjectRequest(request, { identity, sessionId: "session-1" });

  assert.equal(sanitized.headers.get("Authorization"), null);
  assert.equal(sanitized.headers.get("Cookie"), null);
  assert.equal(sanitized.headers.get("Cf-Access-Jwt-Assertion"), null);
  assert.equal(sanitized.headers.get("Sec-WebSocket-Protocol"), SAFE_PROTOCOL);
  assert.equal(sanitized.headers.get("x-mom-session-id"), "session-1");
  assert.equal(sanitized.headers.get("x-mom-auth-mode"), "entra");
  assert.ok(sanitized.headers.get("x-mom-verified-identity"));
  assert.doesNotMatch(sanitized.headers.get("x-mom-verified-identity"), /u075060|@bca\.co\.id/);
  assert.doesNotMatch([...sanitized.headers.values()].join("\n"), new RegExp(token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
});

function fakeBinding() {
  const calls = [];
  return {
    calls,
    binding: {
      idFromName(name) {
        return `id:${name}`;
      },
      get(id) {
        return {
          async fetch(request) {
            calls.push({ id, request });
            return new Response("proxied", { status: 200 });
          },
        };
      },
    },
  };
}

test("fails closed before Durable Object and forwards a valid token only as sanitized identity", async () => {
  const missing = fakeBinding();
  const missingResponse = await handleWorkerRequest(
    new Request("https://pilot.example/api/collab/session-1", { headers: { Upgrade: "websocket" } }),
    { ...envIdentity, REQUIRE_ENTRA: "true", MOM_COLLAB_SESSIONS: missing.binding },
    { jwks }
  );
  assert.equal(missingResponse.status, 401);
  assert.equal(missing.calls.length, 0);

  const valid = fakeBinding();
  const token = await signToken();
  const validResponse = await handleWorkerRequest(
    collaborationRequest(token, { Cookie: "must-not-reach-do=true" }),
    { ...envIdentity, REQUIRE_ENTRA: "true", MOM_COLLAB_SESSIONS: valid.binding },
    { jwks }
  );
  assert.equal(validResponse.status, 200);
  assert.equal(valid.calls.length, 1);
  assert.equal(valid.calls[0].id, "id:session-1");
  assert.equal(valid.calls[0].request.headers.get("Cookie"), null);
  assert.equal(valid.calls[0].request.headers.get("Sec-WebSocket-Protocol"), SAFE_PROTOCOL);
});

test("accepts Authorization Bearer only on authenticated health checks", async () => {
  const token = await signToken();
  const health = await handleWorkerRequest(
    new Request("https://pilot.example/health", { headers: { Authorization: `Bearer ${token}` } }),
    { ...envIdentity, REQUIRE_ENTRA: "true" },
    { jwks }
  );
  assert.equal(health.status, 200);
  assert.deepEqual(await health.json(), {
    ok: true,
    mode: "entra",
    identity: { email: "u075060@bca.co.id", displayName: "Alex Marcello" },
  });

  const binding = fakeBinding();
  const collaboration = await handleWorkerRequest(
    new Request("https://pilot.example/api/collab/session-1", {
      headers: { Upgrade: "websocket", Authorization: `Bearer ${token}` },
    }),
    { ...envIdentity, REQUIRE_ENTRA: "true", MOM_COLLAB_SESSIONS: binding.binding },
    { jwks }
  );
  assert.equal(collaboration.status, 401);
  assert.equal(binding.calls.length, 0);
});

test("legacy mode keeps routing without identity while still removing credential headers", async () => {
  const legacy = fakeBinding();
  const token = await signToken();
  const response = await handleWorkerRequest(
    collaborationRequest(token, { Cookie: "legacy-cookie", Authorization: `Bearer ${token}` }),
    { MOM_COLLAB_SESSIONS: legacy.binding }
  );
  assert.equal(response.status, 200);
  assert.equal(legacy.calls.length, 1);
  assert.equal(legacy.calls[0].request.headers.get("x-mom-auth-mode"), "legacy");
  assert.equal(legacy.calls[0].request.headers.get("x-mom-verified-identity"), null);
  assert.equal(legacy.calls[0].request.headers.get("Cookie"), null);
  assert.match(legacy.calls[0].request.url, /clientId=attacker&editorName=Attacker/);
});

test("authenticated mode makes client identity, timestamp, and version server-owned", () => {
  const incoming = {
    type: "patch",
    clientId: "forged-client",
    editorName: "Forged Name",
    updatedAt: "1999-01-01T00:00:00.000Z",
    version: 999_999,
  };
  const authenticated = applyMessageOwnership(
    incoming,
    { authMode: "entra", clientId: "server-client", editorName: "Verified User" },
    7,
    () => "2026-08-13T01:02:03.000Z"
  );
  assert.equal(authenticated.clientId, "server-client");
  assert.equal(authenticated.editorName, "Verified User");
  assert.equal(authenticated.updatedAt, "2026-08-13T01:02:03.000Z");
  assert.equal(authenticated.version, 8);

  const legacy = applyMessageOwnership(incoming, { authMode: "legacy", clientId: "fallback", editorName: "Editor" }, 7);
  assert.equal(legacy.clientId, "forged-client");
  assert.equal(legacy.editorName, "Forged Name");
  assert.equal(legacy.updatedAt, "1999-01-01T00:00:00.000Z");
  assert.equal(legacy.version, 999_999);

  const authenticatedWithoutClientMetadata = applyMessageOwnership(
    { type: "patch" },
    { authMode: "entra", clientId: "server-client", editorName: "Verified User" },
    8,
    () => "2026-08-13T01:02:04.000Z"
  );
  assert.equal(authenticatedWithoutClientMetadata.clientId, "server-client");
  assert.equal(authenticatedWithoutClientMetadata.version, 9);

  assert.deepEqual(
    getVerifiedIdentityInit({
      clientId: "server-client",
      identity: { email: "u075060@bca.co.id", displayName: "Alex Marcello" },
    }),
    {
      serverClientId: "server-client",
      identity: { email: "u075060@bca.co.id", displayName: "Alex Marcello", verified: true },
    }
  );
  assert.deepEqual(getVerifiedIdentityInit({ clientId: "legacy", identity: null }), {});
});

test("blocks prototype pollution, excessive depth, and huge sparse array indices", () => {
  for (const unsafe of [
    "__proto__/polluted",
    "constructor/prototype/polluted",
    "safe/10001/value",
    Array.from({ length: 17 }, (_, index) => `level${index}`).join("/"),
  ]) {
    assert.equal(validateDraftPath(unsafe).valid, false, unsafe);
  }

  const draft = { rows: [{ value: "old" }] };
  setDraftPath(draft, "rows/0/value", "new");
  setDraftPath(draft, "__proto__/polluted", "yes");
  assert.equal(draft.rows[0].value, "new");
  assert.equal({}.polluted, undefined);
});

class FakeStorage {
  constructor() {
    this.values = new Map();
    this.transactionKeys = [];
  }

  async get(key) {
    return this.values.get(key);
  }

  async transaction(callback) {
    const staged = new Map();
    await callback({
      put: async (key, value) => staged.set(key, structuredClone(value)),
    });
    this.transactionKeys.push([...staged.keys()]);
    for (const [key, value] of staged) this.values.set(key, value);
  }

  async list({ prefix, limit }) {
    return new Map([...this.values].filter(([key]) => key.startsWith(prefix)).slice(0, limit));
  }

  async delete(key) {
    this.values.delete(key);
  }
}

test("stores state, audit hash-chain, and R2 outbox transactionally then retries R2", async () => {
  const storage = new FakeStorage();
  const objects = [];
  let failR2 = true;
  const r2 = {
    async put(key, value) {
      if (failR2) throw new Error("temporary R2 failure");
      objects.push({ key, event: JSON.parse(value) });
    },
  };
  const session = new MomCollabSession({ storage }, { MOM_COLLAB_AUDIT: r2 });
  await session.ready;
  session.sessionId = "session-1";
  const meta = {
    authMode: "entra",
    clientId: "server-client",
    editorName: "Alex Marcello",
    identity: {
      subject: "verified-subject",
      email: "u075060@bca.co.id",
      displayName: "Alex Marcello",
    },
  };

  const first = await session.commitMutation({
    payload: { confidentialValue: "never-copy-into-audit" },
    version: 1,
    updatedAt: "2026-08-13T01:00:00.000Z",
    meta,
    action: "full",
    path: "",
    baseVersion: 0,
  });
  assert.equal(storage.transactionKeys.length, 1);
  assert.equal(storage.transactionKeys[0].length, 4);
  assert.ok(storage.transactionKeys[0].includes("mom-collab-latest-state"));
  assert.ok(storage.transactionKeys[0].some((key) => key.startsWith("mom-collab-audit/")));
  assert.ok(storage.transactionKeys[0].some((key) => key.startsWith("mom-collab-audit-outbox/")));
  assert.doesNotMatch(JSON.stringify(first), /never-copy-into-audit|jwt|ipAddress/i);
  assert.equal(await session.flushAuditOutbox(), 0);
  assert.equal([...storage.values.keys()].filter((key) => key.startsWith("mom-collab-audit-outbox/")).length, 1);

  const second = await session.commitMutation({
    payload: { confidentialValue: "changed" },
    version: 2,
    updatedAt: "2026-08-13T01:01:00.000Z",
    meta,
    action: "patch",
    path: "table3State/0/activity",
    baseVersion: 1,
  });
  assert.equal(second.previousHash, first.hash);
  assert.notEqual(second.hash, first.hash);

  failR2 = false;
  assert.equal(await session.flushAuditOutbox(), 2);
  assert.equal(objects.length, 2);
  assert.equal(new Set(objects.map(({ key }) => key)).size, 2);
  assert.equal([...storage.values.keys()].filter((key) => key.startsWith("mom-collab-audit-outbox/")).length, 0);
});

test("audit hashes are deterministic for the same canonical event and contain no raw MOM value", async () => {
  const input = {
    sequence: 1,
    previousHash: "",
    sessionId: "session-1",
    authMode: "entra",
    identity: { subject: "subject", email: "user@example.com", displayName: "User" },
    clientId: "server-client",
    editorName: "User",
    action: "patch",
    path: "table3State/0/activity",
    baseVersion: 0,
    version: 1,
    occurredAt: "2026-08-13T01:00:00.000Z",
    auditId: "fixed-audit-id",
  };
  const first = await createAuditEvent(input);
  const second = await createAuditEvent(input);
  assert.equal(first.hash, second.hash);
  assert.doesNotMatch(JSON.stringify(first), /meeting secret|token|cookie|ipAddress/i);
});
