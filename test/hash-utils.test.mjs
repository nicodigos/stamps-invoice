import assert from "node:assert/strict";
import test from "node:test";
import { assertMatchingSha256, sha256Hex } from "../assets/js/hash-utils.mjs";

test("accepts downloaded bytes that exactly match the uploaded PDF", async () => {
  const uploaded = new TextEncoder().encode("same PDF bytes");
  const downloaded = new Uint8Array(uploaded);

  await assert.doesNotReject(() => assertMatchingSha256(uploaded, downloaded, "Saved PDF"));
  assert.equal(await sha256Hex(uploaded), await sha256Hex(downloaded));
});

test("rejects a downloaded PDF when any byte changed", async () => {
  const uploaded = new TextEncoder().encode("original PDF bytes");
  const downloaded = new TextEncoder().encode("modified PDF bytes");

  await assert.rejects(
    () => assertMatchingSha256(uploaded, downloaded, "Saved PDF"),
    /Saved PDF content mismatch after upload/,
  );
});
