// Created by Codex on 2026-05-11.
// Smoke tests for the pure strategy scoring engine.

import fs from "node:fs/promises";
import path from "node:path";
import assert from "node:assert/strict";
import { calculateDistrictScores, defaultConstraintSelections } from "./lib/strategy_engine.mjs";

const root = process.cwd();
const strategyModel = JSON.parse(await fs.readFile(path.join(root, "data", "strategy_model.json"), "utf8"));
const districtNames = ["徐汇", "闵行", "浦东"];
const baseWeights = Object.fromEntries(strategyModel.scoreDimensions.map((item) => [item.key, item.defaultWeight]));
const defaultSelections = defaultConstraintSelections(strategyModel.constraintGroups);

const defaultDecision = calculateDistrictScores({
  scoreDimensions: strategyModel.scoreDimensions,
  districtNames,
  baseWeights,
  constraintGroups: strategyModel.constraintGroups,
  selections: defaultSelections,
});

assert.equal(defaultDecision.districtScores.length, 3);
assert.equal(defaultDecision.districtScores[0].districtName, "闵行");
assert.ok(defaultDecision.effects.labels.includes("暂无居住证"));
assert.ok(defaultDecision.effects.labels.includes("100㎡左右会明显考验核心区房源库存。") === false);

const flexibleDecision = calculateDistrictScores({
  scoreDimensions: strategyModel.scoreDimensions,
  districtNames,
  baseWeights,
  constraintGroups: strategyModel.constraintGroups,
  selections: {
    ...defaultSelections,
    commuteTolerance: "flexible",
    schoolPreference: "privateAcceptable",
    elderPickup: false,
    budgetMax: 14000,
  },
});

const pudongDefault = defaultDecision.districtScores.find((item) => item.districtName === "浦东").score;
const pudongFlexible = flexibleDecision.districtScores.find((item) => item.districtName === "浦东").score;
assert.ok(pudongFlexible > pudongDefault, "浦东 should improve when cross-river commute and private options are acceptable");
assert.ok(flexibleDecision.effects.labels.includes("可接受跨江"));
assert.ok(flexibleDecision.effects.labels.includes("预算放宽"));
assert.ok(flexibleDecision.effects.labels.includes("主要父母接送"));

console.log(JSON.stringify({
  ok: true,
  defaultTop: defaultDecision.districtScores[0],
  flexibleTop: flexibleDecision.districtScores[0],
}, null, 2));
