// Created by Codex on 2026-05-09.
// Batch-enrich kindergarten rows with Amap POI/geocode coordinates and office distance.

import fs from "node:fs/promises";
import path from "node:path";

const apiKey = process.env.AMAP_API_KEY;
if (!apiKey) {
  throw new Error("AMAP_API_KEY is required");
}

const csvPath = path.join(process.cwd(), "outputs", "徐汇区幼儿园园区位置表.csv");
const outputPath = path.join(process.cwd(), "data", "amap_enrichment.json");
const officeName = "网易上海西岸研发中心";
const officeLocation = "121.459000,31.156370";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const url = (endpoint, params) => {
  const query = new URLSearchParams({ key: apiKey, ...params });
  return `https://restapi.amap.com/v3/${endpoint}?${query.toString()}`;
};

const parseCsv = (text) => {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (quoted && char === "\"" && next === "\"") {
      cell += "\"";
      i += 1;
    } else if (char === "\"") {
      quoted = !quoted;
    } else if (!quoted && char === ",") {
      row.push(cell);
      cell = "";
    } else if (!quoted && (char === "\n" || char === "\r")) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  const [header, ...body] = rows.filter((item) => item.length > 1);
  return body.map((cells) => Object.fromEntries(header.map((key, index) => [key, cells[index] || ""])));
};

const itemKey = (row) => [row["性质"], row["幼儿园"], row["园区/分园"], row["地址"]].join("|");
const radians = (degrees) => degrees * Math.PI / 180;
const distanceMeters = (from, to) => {
  const [lng1, lat1] = from.split(",").map(Number);
  const [lng2, lat2] = to.split(",").map(Number);
  const dLat = radians(lat2 - lat1);
  const dLng = radians(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2
    + Math.cos(radians(lat1)) * Math.cos(radians(lat2)) * Math.sin(dLng / 2) ** 2;
  return Math.round(6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
};

const scorePoi = (row, poi) => {
  let score = 0;
  const name = `${poi.name || ""}`;
  const type = `${poi.type || ""}`;
  const district = `${poi.pname || ""}${poi.cityname || ""}${poi.adname || ""}`;
  const address = `${poi.address || ""}`;
  if (type.includes("幼儿园")) score += 5;
  if (name.includes(row["幼儿园"].replace(/^上海市?/, "")) || row["幼儿园"].includes(name.replace(/[()（）].*$/, ""))) score += 5;
  if (row["园区/分园"] && name.includes(row["园区/分园"].replace(/（.*?）|\(.*?\)/g, ""))) score += 3;
  if (district.includes(row["区"])) score += 2;
  const road = row["地址"].match(/([\u4e00-\u9fa5A-Za-z]+路|[\u4e00-\u9fa5A-Za-z]+街|[\u4e00-\u9fa5A-Za-z]+镇|[\u4e00-\u9fa5A-Za-z]+弄)/)?.[1];
  if (road && address.includes(road)) score += 2;
  return score;
};

const fetchJson = async (requestUrl) => {
  const response = await fetch(requestUrl);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return response.json();
};

const buildEntry = (row, fields) => {
  const meters = distanceMeters(fields.location, officeLocation);
  return {
    query: fields.query,
    status: "matched",
    matchScore: fields.matchScore,
    poiId: fields.poiId || "",
    poiName: fields.poiName,
    poiAddress: fields.poiAddress,
    poiDistrict: fields.poiDistrict || "",
    poiType: fields.poiType || "",
    tel: fields.tel || "",
    location: fields.location,
    officeName,
    officeLocation,
    officeDistanceMeters: `${meters}`,
    officeDistanceText: `高德直线约 ${(meters / 1000).toFixed(2)} km`,
    updatedAt: new Date().toISOString(),
  };
};

const text = await fs.readFile(csvPath, "utf8");
const rows = parseCsv(text.replace(/^\uFEFF/, ""));
const enrichment = JSON.parse(await fs.readFile(outputPath, "utf8"));
const pending = rows.filter((row) => !enrichment[itemKey(row)] || enrichment[itemKey(row)].status !== "matched");
let processed = 0;
let poiMatches = 0;
let geocodes = 0;

for (const row of pending) {
  const key = itemKey(row);
  const district = row["区"] === "浦东" ? "浦东新区" : `${row["区"]}区`;
  const query = `上海市${district} ${row["幼儿园"]} ${row["园区/分园"]} ${row["地址"]}`;
  try {
    const poiData = await fetchJson(url("place/text", {
      keywords: query,
      city: "上海",
      citylimit: "true",
      offset: "10",
      extensions: "all",
    }));
    const pois = Array.isArray(poiData.pois) ? poiData.pois : [];
    const best = pois
      .map((poi) => ({ poi, score: scorePoi(row, poi) }))
      .sort((a, b) => b.score - a.score)[0];
    if (best?.poi?.location && best.score >= 8) {
      enrichment[key] = buildEntry(row, {
        query,
        matchScore: best.score,
        poiId: best.poi.id,
        poiName: best.poi.name,
        poiAddress: Array.isArray(best.poi.address) ? best.poi.address.join("") : `${best.poi.address || ""}`,
        poiDistrict: best.poi.adname,
        poiType: best.poi.type,
        tel: Array.isArray(best.poi.tel) ? best.poi.tel.join("/") : `${best.poi.tel || ""}`,
        location: best.poi.location,
      });
      poiMatches += 1;
    } else {
      const geocodeData = await fetchJson(url("geocode/geo", {
        address: `上海市${district}${row["地址"]}`,
        city: "上海",
      }));
      const geocode = Array.isArray(geocodeData.geocodes) ? geocodeData.geocodes[0] : null;
      if (!geocode?.location) throw new Error("no geocode result");
      enrichment[key] = buildEntry(row, {
        query,
        matchScore: best?.score || 0,
        poiName: "高德地址坐标（非POI精确匹配）",
        poiAddress: geocode.formatted_address || row["地址"],
        poiDistrict: geocode.district || district,
        poiType: "地址地理编码",
        location: geocode.location,
      });
      geocodes += 1;
    }
  } catch (error) {
    enrichment[key] = {
      query,
      status: "unmatched",
      error: error instanceof Error ? error.message : `${error}`,
      updatedAt: new Date().toISOString(),
    };
  }
  processed += 1;
  if (processed % 25 === 0) {
    await fs.writeFile(outputPath, `${JSON.stringify(enrichment, null, 2)}\n`);
    console.log(JSON.stringify({ processed, remaining: pending.length - processed, poiMatches, geocodes }));
  }
  await sleep(120);
}

await fs.writeFile(outputPath, `${JSON.stringify(enrichment, null, 2)}\n`);
console.log(JSON.stringify({ processed, poiMatches, geocodes, total: Object.keys(enrichment).length }, null, 2));
