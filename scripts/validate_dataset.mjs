// Created by Codex on 2026-05-11.
// Validate generated kindergarten data before publishing static outputs.

import fs from "node:fs/promises";
import path from "node:path";

const root = process.cwd();
const datasetPath = path.join(root, "data", "kindergartens", "kindergarten_dataset.json");
const byDistrictDir = path.join(root, "data", "kindergartens", "by_district");
const strategyModelPath = path.join(root, "data", "strategy_model.json");
const outputIndexPath = path.join(root, "outputs", "index.html");

const readJson = async (filePath) => JSON.parse(await fs.readFile(filePath, "utf8"));
const fail = (message, details = {}) => {
  console.error(JSON.stringify({ ok: false, message, ...details }, null, 2));
  process.exitCode = 1;
};

const dataset = await readJson(datasetPath);
const strategyModel = await readJson(strategyModelPath);
const rows = Array.isArray(dataset.rows) ? dataset.rows : [];
const requiredTopLevelFields = ["district", "area", "name", "campus", "nature", "category", "level", "address", "phone", "source"];
const allowedDistricts = new Set(["徐汇", "闵行", "浦东"]);
const allowedNature = new Set(["公办", "民办", "中外合作"]);
const allowedConfidence = new Set(["A", "B", "C"]);

const errors = [];
const warnings = [];
const primaryKeys = new Set();
const strategyKeys = new Set(rows.map((row) => [row.nature, row.name, row.campus, row.address].join("|")));

const addError = (message, row) => errors.push(row ? { message, key: [row.district, row.name, row.campus, row.address].join("|") } : { message });
const addWarning = (message, row) => warnings.push(row ? { message, key: [row.district, row.name, row.campus, row.address].join("|") } : { message });

if (dataset.schemaVersion !== "1.0") addError("dataset.schemaVersion must be 1.0");
if (dataset.rowCount !== rows.length) addError("dataset.rowCount does not match rows.length");
for (const [district, slug] of [["徐汇", "xuhui"], ["闵行", "minhang"], ["浦东", "pudong"]]) {
  const districtDataset = await readJson(path.join(byDistrictDir, `${slug}.json`));
  const expectedRows = rows.filter((row) => row.district === district).length;
  if (districtDataset.rowCount !== districtDataset.rows?.length) addError(`${slug} rowCount does not match rows.length`);
  if (districtDataset.rowCount !== expectedRows) addError(`${slug} rowCount does not match full dataset`);
}

for (const row of rows) {
  for (const field of requiredTopLevelFields) {
    if (row[field] === undefined || row[field] === "") addError(`missing required field: ${field}`, row);
  }
  if (!allowedDistricts.has(row.district)) addError(`invalid district: ${row.district}`, row);
  if (!allowedNature.has(row.nature)) addError(`invalid nature: ${row.nature}`, row);
  if (!allowedConfidence.has(row.quality?.confidence)) addError(`invalid confidence: ${row.quality?.confidence}`, row);

  const primaryKey = [row.district, row.name, row.campus, row.address].join("|");
  if (primaryKeys.has(primaryKey)) addError("duplicated primary key", row);
  primaryKeys.add(primaryKey);

  if (!row.admission?.type) addError("missing admission.type", row);
  if (!row.geo?.mapUrl?.startsWith("https://ditu.amap.com/search?query=")) addError("invalid amap search url", row);
  if (row.geo?.amapLocation !== "待高德MCP查询" && !/^-?\d+(\.\d+)?,-?\d+(\.\d+)?$/.test(row.geo?.amapLocation || "")) {
    addError("invalid amap location format", row);
  }
  if (!row.rental?.beikeBoardSlug) addWarning("missing rental beikeBoardSlug", row);
  if (row.phone.includes("待电话确认")) addWarning("phone still pending", row);
}

const scoreTotal = (strategyModel.scoreDimensions || []).reduce((sum, item) => sum + (Number(item.defaultWeight) || 0), 0);
if (scoreTotal !== 100) addWarning(`score dimension default weights total ${scoreTotal}, expected 100`);
for (const item of strategyModel.scoreDimensions || []) {
  for (const district of allowedDistricts) {
    if (!Number.isFinite(Number(item.scores?.[district]))) addError(`score dimension ${item.key} missing ${district} score`);
  }
}
const scoreDimensionKeys = new Set((strategyModel.scoreDimensions || []).map((item) => item.key));
const validateAdjustmentKeys = (owner, effect = {}) => {
  for (const key of Object.keys(effect.weightAdjustments || {})) {
    if (!scoreDimensionKeys.has(key)) addError(`${owner} references unknown score dimension ${key}`);
  }
  for (const district of Object.keys(effect.districtAdjustments || {})) {
    if (!allowedDistricts.has(district)) addError(`${owner} references unknown district ${district}`);
  }
};
const isValidCondition = (condition = {}) => {
  if (Array.isArray(condition.all)) return condition.all.every(isValidCondition);
  if (Array.isArray(condition.any)) return condition.any.every(isValidCondition);
  return Boolean(condition.operator) && condition.value !== undefined;
};
for (const group of strategyModel.constraintGroups || []) {
  if (!group.key) addError("constraint group missing key");
  if (!group.label) addError(`constraint group ${group.key} missing label`);
  if (!["choice", "number", "boolean", undefined].includes(group.type)) addError(`constraint group ${group.key} has invalid type`);
  if (group.type === "number") {
    if (!Number.isFinite(Number(group.defaultValue))) addError(`constraint group ${group.key} missing numeric defaultValue`);
    if (!Array.isArray(group.impacts) || group.impacts.length === 0) addError(`constraint group ${group.key} has no impacts`);
    for (const impact of group.impacts || []) {
      if (!isValidCondition(impact.when)) addError(`constraint group ${group.key} impact missing condition`);
      validateAdjustmentKeys(`constraint impact ${group.key}.${impact.key || "unnamed"}`, impact);
    }
  } else if (group.type === "boolean") {
    if (group.defaultValue !== true && group.defaultValue !== false) addError(`constraint group ${group.key} missing boolean defaultValue`);
    validateAdjustmentKeys(`constraint boolean ${group.key}.whenTrue`, group.whenTrue);
    validateAdjustmentKeys(`constraint boolean ${group.key}.whenFalse`, group.whenFalse);
  } else {
    if (!Array.isArray(group.options) || group.options.length === 0) addError(`constraint group ${group.key} has no options`);
    if (group.defaultOption && !group.options?.some((option) => option.key === group.defaultOption)) {
      addError(`constraint group ${group.key} defaultOption not found`);
    }
    for (const option of group.options || []) {
      if (!option.key) addError(`constraint group ${group.key} option missing key`);
      validateAdjustmentKeys(`constraint option ${group.key}.${option.key}`, option);
    }
  }
}
for (const item of strategyModel.decisionRecommendations || []) {
  const ref = item.itemRef || {};
  const key = [ref.nature, ref.name, ref.campus, ref.address].join("|");
  if (!strategyKeys.has(key)) addError(`strategy itemRef not found: ${key}`);
}

try {
  await fs.access(outputIndexPath);
} catch {
  addError("outputs/index.html is missing; run npm run build");
}

if (errors.length) {
  fail("dataset validation failed", {
    errors: errors.slice(0, 30),
    errorCount: errors.length,
    warningCount: warnings.length,
  });
} else {
  console.log(JSON.stringify({
    ok: true,
    rowCount: rows.length,
    districts: dataset.counts,
    warnings: warnings.slice(0, 20),
    warningCount: warnings.length,
  }, null, 2));
}
