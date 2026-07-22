import assert from "node:assert/strict";
import test from "node:test";
import { describeError, describeHttpError } from "../assets/js/error-utils.mjs";

test("extracts a useful message from a Graph JSON error", () => {
  const message = describeHttpError("SharePoint upload", 403, JSON.stringify({ error: { message: "Access denied to Accounting/July." } }));
  assert.match(message, /SharePoint upload failed \(HTTP 403\)/);
  assert.match(message, /Access denied to Accounting\/July/);
  assert.match(message, /does not have permission/);
});

test("turns an unhelpful fetch failure into actionable guidance", () => {
  assert.equal(describeError(new TypeError("Failed to fetch")), "The network request failed. Check your connection and confirm that Microsoft is still connected, then try again.");
});

test("removes HTML from server errors", () => {
  assert.equal(describeError("<h1>Bad Gateway</h1><p>Try later</p>"), "Bad Gateway Try later");
});
