import assert from "node:assert/strict";
import test from "node:test";
import { singleFlight } from "../assets/js/single-flight.mjs";

test("ignores a second process attempt while the first one is running", async () => {
  let calls = 0;
  let finishFirstProcess;
  const firstProcessPending = new Promise((resolve) => {
    finishFirstProcess = resolve;
  });
  const processOnce = singleFlight(async () => {
    calls += 1;
    await firstProcessPending;
  });

  const firstAttempt = processOnce();
  const secondAttempt = processOnce();

  assert.equal(calls, 1);
  assert.equal(await secondAttempt, false);

  finishFirstProcess();
  assert.equal(await firstAttempt, true);
});

test("allows processing again after the active process finishes", async () => {
  let calls = 0;
  const processOnce = singleFlight(async () => {
    calls += 1;
  });

  assert.equal(await processOnce(), true);
  assert.equal(await processOnce(), true);
  assert.equal(calls, 2);
});
