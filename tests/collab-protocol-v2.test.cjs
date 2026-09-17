const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const workerPath = path.join(__dirname, "..", "worker", "index.mjs");

function createState() {
  const store = {};
  return {
    storage: {
      async get(key) {
        return store[key];
      },
      async put(key, value) {
        store[key] = value;
      },
    },
    snapshot: () => store,
  };
}

function createSocket() {
  const sent = [];
  return {
    sent,
    send(raw) {
      sent.push(JSON.parse(raw));
    },
    last() {
      return sent[sent.length - 1];
    },
  };
}

function makeDraft() {
  return {
    memo: "",
    table1BlueprintLevel: "release",
    table1ProjectsState: [
      { projectName: "Alpha", relationPackages: [{ releaseId: "", documentNeeds: [] }] },
    ],
    table3State: [{ activity: "first row", status: "To Do" }],
    lampiranEnabled: false,
    lampiranState: [],
    checklistRows: [{ id: "c1", status: "To Do" }],
  };
}

(async () => {
  const worker = await import(pathToFileURL(workerPath).href);
  const { MomCollabSession } = worker;

  // --- helper exports -------------------------------------------------------
  assert.equal(typeof worker.MomCollabSession, "function");
  assert.equal(typeof worker.readDraftPath, "function");

  const empty = makeDraft();
  assert.deepEqual(worker.readDraftPath(empty, "table3State/0/activity"), {
    exists: true,
    value: "first row",
  });
  assert.equal(
    worker.readDraftPath(empty, "table3State/9/activity").exists,
    false,
    "a missing row index must be detectable so ghost rows are never created"
  );

  // --- server-authoritative sequencing --------------------------------------
  const state = createState();
  const session = new MomCollabSession(state, {});
  await session.ready;

  const editorA = createSocket();
  session.clients.set(editorA, { clientId: "A" });
  await session.handleFullMessage({
    type: "full",
    clientId: "A",
    path: "draft",
    value: makeDraft(),
    replace: true,
    updatedAt: "t1",
    version: 9999,
  });
  assert.equal(session.latestVersion, 1, "the server assigns the sequence, not the client");

  // A hostile or confused client claiming a huge version must not advance the server.
  const editorB = createSocket();
  session.clients.set(editorB, { clientId: "B" });
  await session.handlePatchMessage(
    {
      type: "patch",
      clientId: "B",
      version: 50000,
      ops: [{ path: "memo", value: "memo from B" }],
    },
    editorB
  );
  assert.equal(session.latestVersion, 2, "client-supplied version numbers are ignored");
  assert.equal(session.latestPayload.memo, "memo from B");

  // --- field-level last-write-wins across editors ---------------------------
  await session.handlePatchMessage(
    {
      type: "patch",
      clientId: "A",
      ops: [{ path: "table3State/0/activity", value: "activity from A" }],
    },
    editorA
  );
  assert.equal(session.latestPayload.table3State[0].activity, "activity from A");
  assert.equal(
    session.latestPayload.memo,
    "memo from B",
    "editing one field must not disturb a field another editor wrote"
  );

  assert.deepEqual(editorA.last().ops, [{ path: "table3State/0/activity", applied: true }]);
  assert.equal(editorA.last().type, "ack", "the sender receives an ack, never its own broadcast");
  const memoBroadcast = editorA.sent.find(
    (message) => message.type === "patch" && (message.ops || []).some((op) => op.path === "memo")
  );
  assert.ok(memoBroadcast, "other editors receive the patch as a broadcast");
  assert.equal(
    memoBroadcast.clientId,
    "B",
    "broadcasts carry the originating editor so senders can ignore echoes"
  );
  assert.equal(
    editorA.sent.some((message) => message.clientId === "A" && message.type === "patch"),
    false,
    "a sender must never receive an echo of its own patch"
  );

  await session.handlePatchMessage(
    {
      type: "patch",
      clientId: "B",
      ops: [{ path: "table3State/0/activity", value: "activity from B" }],
    },
    editorB
  );
  assert.equal(
    session.latestPayload.table3State[0].activity,
    "activity from B",
    "the later writer wins on the same path"
  );

  // --- impossible structural writes are reported, not silently applied ------
  await session.handlePatchMessage(
    {
      type: "patch",
      clientId: "B",
      mutationId: "ghost-1",
      ops: [{ path: "table3State/9/activity", value: "ghost" }],
    },
    editorB
  );
  const ghostAck = editorB.last();
  assert.equal(ghostAck.rejected, true);
  assert.deepEqual(ghostAck.ops, [
    { path: "table3State/9/activity", applied: false, reason: "path-missing" },
  ]);
  assert.equal(session.latestPayload.table3State.length, 1, "no ghost row is created");

  // --- whole-array structural operations ------------------------------------
  const structuralState = createState();
  const structuralSession = new MomCollabSession(structuralState, {});
  await structuralSession.ready;
  const seedSocket = createSocket();
  structuralSession.clients.set(seedSocket, { clientId: "seed" });
  await structuralSession.handleFullMessage({
    type: "full",
    clientId: "seed",
    path: "draft",
    value: makeDraft(),
    replace: true,
    updatedAt: "t1",
  });

  const structuralSocket = createSocket();
  structuralSession.clients.set(structuralSocket, { clientId: "S" });
  await structuralSession.handlePatchMessage(
    {
      type: "patch",
      clientId: "S",
      ops: [
        {
          kind: "array-set",
          path: "table3State",
          value: [{ activity: "row one" }, { activity: "row two" }],
        },
      ],
    },
    structuralSocket
  );
  assert.equal(structuralSession.latestPayload.table3State.length, 2);
  assert.equal(
    structuralSession.latestPayload.memo,
    "",
    "a structural array write must not touch unrelated fields"
  );

  // The nested array is the more precise target: it must travel as-is, and the receiving
  // client needs the array-set marker to know it has to re-render rather than assign a
  // single field.
  await structuralSession.handlePatchMessage(
    {
      type: "patch",
      clientId: "S",
      ops: [
        {
          kind: "array-set",
          path: "table1ProjectsState/0/relationPackages",
          value: [{ releaseId: "R1" }, { releaseId: "R2" }, { releaseId: "R3" }],
        },
      ],
    },
    structuralSocket
  );
  assert.equal(
    structuralSession.latestPayload.table1ProjectsState[0].relationPackages.length,
    3,
    "a nested array replacement applies to the model"
  );
  const broadcast = seedSocket.sent
    .filter((message) => message.type === "patch")
    .flatMap((message) => message.ops || [])
    .find((op) => op.path === "table1ProjectsState/0/relationPackages");
  assert.ok(broadcast, "the nested array change is broadcast to other editors");
  assert.equal(
    broadcast.kind,
    "array-set",
    "the array-set marker must survive the server so receivers re-render the section"
  );

  // --- resync: replay vs snapshot ------------------------------------------
  const replaySocket = createSocket();
  structuralSession.clients.set(replaySocket, { clientId: "R" });
  await structuralSession.handleResync(
    { type: "resync", clientId: "R", sinceSeq: 1 },
    replaySocket
  );
  const replay = replaySocket.last();
  assert.equal(replay.type, "sync");
  assert.equal(replay.mode, "replay", "a small gap is replayed as ops");
  assert.equal(replay.seq, structuralSession.latestVersion);
  assert.deepEqual(
    replay.entries.flatMap((entry) => entry.ops).map((op) => op.path),
    ["table3State", "table1ProjectsState/0/relationPackages"],
    "replay contains every missed operation, in order"
  );
  assert.equal(
    replay.entries
      .flatMap((entry) => entry.ops)
      .find((op) => op.path === "table1ProjectsState/0/relationPackages")?.kind,
    "array-set",
    "replayed array operations keep their array-set marker"
  );

  const snapshotSocket = createSocket();
  structuralSession.clients.set(snapshotSocket, { clientId: "S2" });
  await structuralSession.handleResync({ type: "resync", clientId: "S2", sinceSeq: 0 }, snapshotSocket);
  const snapshot = snapshotSocket.last();
  assert.equal(snapshot.mode, "snapshot", "a resync from scratch gets a snapshot");
  assert.equal(snapshot.value.table3State.length, 2);

  // --- beacon: last-write flush from a closing tab --------------------------
  const beaconState = createState();
  const beaconSession = new MomCollabSession(beaconState, {});
  await beaconSession.ready;
  const beaconSeed = createSocket();
  beaconSession.clients.set(beaconSeed, { clientId: "seed" });
  await beaconSession.handleFullMessage({
    type: "full",
    clientId: "seed",
    path: "draft",
    value: makeDraft(),
    replace: true,
    updatedAt: "t1",
  });

  const beaconResponse = await beaconSession.fetch(
    new Request("https://example.test/api/collab/room-1", {
      method: "POST",
      body: JSON.stringify({
        clientId: "Z",
        ops: [{ path: "memo", value: "flushed on unload" }],
      }),
    })
  );
  assert.equal(beaconResponse.status, 200);
  const beaconBody = await beaconResponse.json();
  assert.equal(beaconBody.ok, true);
  assert.equal(beaconBody.applied, 1);
  assert.equal(beaconSession.latestPayload.memo, "flushed on unload");

  const beaconDraft = await beaconSession.fetch(
    new Request("https://example.test/api/collab/room-1", {
      method: "POST",
      body: JSON.stringify({ clientId: "Z", ops: [{ path: "draft", value: {} }] }),
    })
  );
  const beaconDraftBody = await beaconDraft.json();
  assert.equal(beaconDraftBody.applied, 0, "the beacon must never replace the whole draft");
  assert.equal(beaconSession.latestPayload.memo, "flushed on unload");

  // --- plain HTTP probe used by deployment checks ---------------------------
  const probe = await beaconSession.fetch(new Request("https://example.test/api/collab/room-1"));
  assert.equal(probe.status, 200);
  const probeBody = await probe.json();
  assert.equal(probeBody.ok, true);
  assert.equal(probeBody.seq, beaconSession.latestVersion);

  console.log("collab protocol v2 tests passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
