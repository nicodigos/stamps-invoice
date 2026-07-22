import assert from "node:assert/strict";
import test from "node:test";
import { filterOptionsContaining } from "../assets/js/combo-utils.mjs";

const options = [
  { value: "office", label: "General Office Expenses" },
  { value: "fees", label: "Professional Fees" },
  { value: "vehicle", label: "Véhicule Expenses" },
];

test("filters category options by a case-insensitive contains match", () => {
  assert.deepEqual(filterOptionsContaining(options, "OFFICE"), [options[0]]);
  assert.deepEqual(filterOptionsContaining(options, "fees"), [options[1]]);
});

test("category filtering ignores accents", () => {
  assert.deepEqual(filterOptionsContaining(options, "vehicule"), [options[2]]);
});
