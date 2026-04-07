/**
 * Заменяет @media (prefers-color-scheme: dark) { ... } на html[data-theme="dark"] ... */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const cssPath = path.join(__dirname, "..", "styles.css");

function extractMediaBlocks(css) {
  const marker = "@media (prefers-color-scheme: dark)";
  const out = [];
  let i = 0;
  while (i < css.length) {
    const idx = css.indexOf(marker, i);
    if (idx === -1) break;
    let j = idx + marker.length;
    while (j < css.length && /\s/.test(css[j])) j++;
    if (css[j] !== "{") throw new Error("Expected { after @media at " + idx);
    j++;
    const bodyStart = j;
    let depth = 1;
    while (j < css.length && depth > 0) {
      if (css[j] === "{") depth++;
      else if (css[j] === "}") depth--;
      j++;
    }
    out.push({ start: idx, end: j, body: css.slice(bodyStart, j - 1) });
    i = j;
  }
  return out;
}

/** Разбор правил верхнего уровня: селектор { ... } */
function parseTopLevelRules(body) {
  const rules = [];
  let i = 0;
  const s = body;
  while (i < s.length) {
    while (i < s.length && /\s/.test(s[i])) i++;
    if (i >= s.length) break;
    const selStart = i;
    const openBrace = s.indexOf("{", selStart);
    if (openBrace === -1) break;
    const sel = s.slice(selStart, openBrace).trim();
    let depth = 1;
    i = openBrace + 1;
    const blockStart = openBrace;
    while (i < s.length && depth > 0) {
      if (s[i] === "{") depth++;
      else if (s[i] === "}") depth--;
      i++;
    }
    const fullBlock = s.slice(blockStart, i);
    rules.push({ sel, block: fullBlock });
  }
  return rules;
}

function prefixSelector(sel) {
  const parts = splitSelectors(sel);
  return parts
    .map((p) => {
      const t = p.trim();
      if (t === ":root") return "html[data-theme=\"dark\"]";
      return `html[data-theme="dark"] ${t}`;
    })
    .join(",\n");
}

function splitSelectors(sel) {
  const out = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < sel.length; i++) {
    const c = sel[i];
    if (c === "(") depth++;
    else if (c === ")") depth--;
    else if (c === "," && depth === 0) {
      out.push(sel.slice(start, i));
      start = i + 1;
    }
  }
  out.push(sel.slice(start));
  return out.map((x) => x.trim()).filter(Boolean);
}

function transformBody(body) {
  const rules = parseTopLevelRules(body);
  if (!rules.length) return body.trim();
  return rules.map(({ sel, block }) => `${prefixSelector(sel)}${block}`).join("\n\n");
}

const css = fs.readFileSync(cssPath, "utf8");
const blocks = extractMediaBlocks(css);
if (!blocks.length) {
  console.error("No dark media blocks found");
  process.exit(1);
}

let out = "";
let last = 0;
for (const { start, end, body } of blocks) {
  out += css.slice(last, start);
  out += transformBody(body) + "\n";
  last = end;
}
out += css.slice(last);

fs.writeFileSync(cssPath, out, "utf8");
console.log("OK: flattened", blocks.length, "dark theme blocks");
