// Created by Codex on 2026-05-11.
// Pure scoring engine for constraint-driven district strategy decisions.

export const clampScore = (value) => Math.max(0, Math.min(100, Math.round(Number(value) || 0)));

export const addAdjustment = (target, source = {}) => {
  for (const [key, value] of Object.entries(source || {})) {
    target[key] = (target[key] || 0) + (Number(value) || 0);
  }
  return target;
};

export const defaultConstraintSelections = (constraintGroups = []) => Object.fromEntries(constraintGroups.map((group) => {
  if (group.type === "number") return [group.key, Number(group.defaultValue ?? group.min ?? 0)];
  if (group.type === "boolean") return [group.key, Boolean(group.defaultValue)];
  return [group.key, group.defaultOption || group.options?.[0]?.key || ""];
}));

export const conditionMatches = (condition = {}, value) => {
  if (Array.isArray(condition.all)) return condition.all.every((item) => conditionMatches(item, value));
  if (Array.isArray(condition.any)) return condition.any.some((item) => conditionMatches(item, value));
  const actual = Number(value);
  const expected = Number(condition.value);
  switch (condition.operator) {
    case "lt": return actual < expected;
    case "lte": return actual <= expected;
    case "gt": return actual > expected;
    case "gte": return actual >= expected;
    case "eq": return actual === expected;
    default: return false;
  }
};

export const resolveConstraintOptionEffects = (group, rawValue) => {
  if (!group) return [];
  if (group.type === "number") {
    const value = Number(rawValue ?? group.defaultValue ?? 0);
    return (group.impacts || [])
      .filter((impact) => conditionMatches(impact.when, value))
      .map((impact) => ({
        key: impact.key || `${group.key}:${value}`,
        label: impact.label || group.label,
        summary: impact.summary || "",
        value,
        weightAdjustments: impact.weightAdjustments || {},
        districtAdjustments: impact.districtAdjustments || {},
      }));
  }
  if (group.type === "boolean") {
    const value = rawValue === true || rawValue === "true" || rawValue === "1";
    const effect = value ? group.whenTrue : group.whenFalse;
    return effect ? [{
      key: value ? "true" : "false",
      label: value ? (group.trueLabel || group.label) : (group.falseLabel || group.label),
      summary: effect.summary || "",
      value,
      weightAdjustments: effect.weightAdjustments || {},
      districtAdjustments: effect.districtAdjustments || {},
    }] : [];
  }
  const selectedKey = rawValue || group.defaultOption;
  const option = (group.options || []).find((item) => item.key === selectedKey) || (group.options || [])[0];
  return option ? [{
    key: option.key,
    label: option.label || option.key,
    summary: option.summary || "",
    value: option.key,
    weightAdjustments: option.weightAdjustments || {},
    districtAdjustments: option.districtAdjustments || {},
  }] : [];
};

export const collectConstraintEffects = (constraintGroups = [], selections = {}) => {
  const defaults = defaultConstraintSelections(constraintGroups);
  const selected = [];
  const weightAdjustments = {};
  const districtAdjustments = {};
  const labels = [];
  const summaries = [];

  for (const group of constraintGroups) {
    const rawValue = selections[group.key] ?? defaults[group.key];
    const effects = resolveConstraintOptionEffects(group, rawValue);
    for (const effect of effects) {
      selected.push({ groupKey: group.key, groupLabel: group.label, ...effect });
      labels.push(effect.label);
      if (effect.summary) summaries.push(effect.summary);
      addAdjustment(weightAdjustments, effect.weightAdjustments);
      addAdjustment(districtAdjustments, effect.districtAdjustments);
    }
  }

  return { selected, weightAdjustments, districtAdjustments, labels, summaries };
};

export const adjustedWeights = (baseWeights = {}, weightAdjustments = {}) => Object.fromEntries(Object.entries(baseWeights).map(([key, value]) => [
  key,
  Math.max(0, (Number(value) || 0) + (Number(weightAdjustments[key]) || 0)),
]));

export const calculateDistrictScores = ({ scoreDimensions = [], districtNames = [], baseWeights = {}, constraintGroups = [], selections = {} }) => {
  const effects = collectConstraintEffects(constraintGroups, selections);
  const weights = adjustedWeights(baseWeights, effects.weightAdjustments);
  const totalWeight = Object.values(weights).reduce((sum, value) => sum + (Number(value) || 0), 0) || 1;
  const districtScores = districtNames.map((districtName) => {
    const raw = scoreDimensions.reduce((sum, item) => {
      return sum + ((Number(item.scores?.[districtName]) || 0) * (Number(weights[item.key]) || 0));
    }, 0);
    const constraintAdjustment = Number(effects.districtAdjustments[districtName]) || 0;
    const score = clampScore((raw / totalWeight) + constraintAdjustment);
    const contributions = scoreDimensions
      .map((item) => ({
        key: item.key,
        label: item.label,
        value: ((Number(item.scores?.[districtName]) || 0) * (Number(weights[item.key]) || 0)) / totalWeight,
      }))
      .sort((a, b) => b.value - a.value);
    return { districtName, score, constraintAdjustment, contributions };
  }).sort((a, b) => b.score - a.score || a.districtName.localeCompare(b.districtName, "zh-CN"));

  return { effects, weights, districtScores };
};
