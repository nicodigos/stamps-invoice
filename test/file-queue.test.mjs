import assert from "node:assert/strict";
import test from "node:test";
import { mergeQueuedFileItems } from "../assets/js/file-queue.mjs";

const item = (name) => ({ id: name, file: { name } });

test("a new EML or MSG replaces the email already in the queue", () => {
  const result = mergeQueuedFileItems(
    [item("invoice.pdf"), item("old.eml")],
    [item("new.MSG")],
  );

  assert.deepEqual(result.map(({ file }) => file.name), ["invoice.pdf", "new.MSG"]);
});

test("only the last email is retained when several are added together", () => {
  const result = mergeQueuedFileItems(
    [item("existing.png")],
    [item("first.eml"), item("attachment.pdf"), item("second.msg")],
  );

  assert.deepEqual(
    result.map(({ file }) => file.name),
    ["existing.png", "attachment.pdf", "second.msg"],
  );
});

test("non-email files continue to accumulate normally", () => {
  const result = mergeQueuedFileItems(
    [item("message.eml"), item("one.pdf")],
    [item("two.pdf"), item("photo.jpg")],
  );

  assert.deepEqual(
    result.map(({ file }) => file.name),
    ["message.eml", "one.pdf", "two.pdf", "photo.jpg"],
  );
});
