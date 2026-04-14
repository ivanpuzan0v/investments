import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");

async function setupApp() {
  const [html, appJs] = await Promise.all([
    readFile(path.join(ROOT, "index.html"), "utf8"),
    readFile(path.join(ROOT, "app.js"), "utf8"),
  ]);

  const dom = new JSDOM(html, {
    url: "http://localhost",
    runScripts: "dangerously",
    pretendToBeVisual: true,
  });
  const { window } = dom;

  window.matchMedia = () => ({
    matches: false,
    addEventListener() {},
    removeEventListener() {},
  });
  window.SVGCircleElement = window.SVGCircleElement || window.SVGElement;
  window.SVGLineElement = window.SVGLineElement || window.SVGElement;
  window.SVGTextElement = window.SVGTextElement || window.SVGElement;
  window.SVGSVGElement = window.SVGSVGElement || window.SVGElement;
  window.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
  window.cancelAnimationFrame = (id) => clearTimeout(id);
  window.alert = () => {};

  window.eval(appJs);
  window.getTodayYMD = () => "2026-01-01";

  return { dom, window, document: window.document };
}

function setTableRows(window, { bonds, holdings, buys = [] }) {
  const { document } = window;
  window.writeRows(document.getElementById("assets-tbody"), document.getElementById("asset-row-template"), bonds);
  window.writeRows(document.getElementById("holdings-tbody"), document.getElementById("holding-row-template"), holdings);
  window.writeRows(document.getElementById("txns-tbody"), document.getElementById("txn-row-template"), buys);
  window.refreshAutoPlanBondPicker();
}

function setSelectedDays(document, dayNums) {
  const chips = Array.from(document.querySelectorAll("#auto-plan-topup-days .planScheduleDayChip"));
  chips.forEach((btn) => {
    btn.classList.remove("planScheduleDayChip--selected");
    btn.setAttribute("aria-pressed", "false");
  });
  for (const day of dayNums) {
    const btn = document.querySelector(`#auto-plan-topup-days .planScheduleDayChip[data-day="${day}"]`);
    if (!btn) continue;
    btn.classList.add("planScheduleDayChip--selected");
    btn.setAttribute("aria-pressed", "true");
  }
}

function setSelectedBonds(document, selectedBondKeys) {
  const selected = new Set((selectedBondKeys || []).map((x) => String(x || "").trim().toUpperCase()));
  const checks = Array.from(document.querySelectorAll('#auto-plan-bonds input[type="checkbox"][data-bond]'));
  checks.forEach((inp) => {
    const key = String(inp.getAttribute("data-bond") || "").trim().toUpperCase();
    inp.checked = selected.has(key);
  });
}

function parseRowItems(rows) {
  return rows.map((r) => ({ date: r.date, items: JSON.parse(r.items) }));
}

function qtyByBond(items) {
  const out = new Map();
  for (const it of items) {
    const k = String(it.bond || "").trim().toUpperCase();
    out.set(k, (out.get(k) || 0) + Number(it.quantity || 0));
  }
  return out;
}

test("reinvest coupons are applied once per purchase month", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "AAA", coupon: "20", bondPrice: "60", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [{ bond: "AAA", quantity: "1" }],
      buys: [],
    });

    setSelectedBonds(document, ["AAA"]);
    setSelectedDays(document, [10, 20]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "100";
    document.getElementById("auto-plan-reinvest").checked = true;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    const result = window.generateAutoPlanBuyRows();
    assert.equal(result.ok, true, result.message || "expected successful generation");

    const rows = parseRowItems(result.rows);
    assert.equal(rows.length, 2, "must generate 2 purchase dates in one month");

    const monthItems = rows.flatMap((r) => r.items);
    const byBond = qtyByBond(monthItems);
    assert.equal(byBond.get("AAA"), 3, "with one monthly reinvest (20) total quantity should be 3, not 4");
    assert.equal(rows[0].items[0].quantity, 2, "first purchase date should consume reinvested budget");
    assert.equal(rows[1].items[0].quantity, 1, "second purchase date should use only regular topup");
  } finally {
    dom.window.close();
  }
});

test("reinvest simulation uses all coupon sources, including non-selected bonds", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "AAA", coupon: "30", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
        { bond: "BBB", coupon: "15", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
        { bond: "CCC", coupon: "5", bondPrice: "50", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [
        { bond: "AAA", quantity: "1" },
        { bond: "BBB", quantity: "2" },
      ],
      buys: [],
    });

    setSelectedBonds(document, ["CCC"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "100";
    document.getElementById("auto-plan-reinvest").checked = true;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    const result = window.generateAutoPlanBuyRows();
    assert.equal(result.ok, true, result.message || "expected successful generation");

    const rows = parseRowItems(result.rows);
    assert.equal(rows.length, 1, "expected one generated date");
    assert.equal(rows[0].items.length, 1, "all budget should go to selected bond CCC");
    assert.equal(rows[0].items[0].bond, "CCC");
    assert.equal(rows[0].items[0].quantity, 3, "topup=100 + all prev-month coupons=60 => 160 => 3 lots at 50");
  } finally {
    dom.window.close();
  }
});

test("horizon prefers bond with more future coupons when nominal yield is equal", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        {
          bond: "SHORT",
          coupon: "10",
          bondPrice: "100",
          payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12",
          startDate: "2025-01-01",
          endDate: "2026-12-31",
        },
        {
          bond: "LONG",
          coupon: "10",
          bondPrice: "100",
          payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12",
          startDate: "2025-01-01",
          endDate: "2030-12-31",
        },
      ],
      holdings: [],
      buys: [],
    });

    setSelectedBonds(document, ["SHORT", "LONG"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "200";
    document.getElementById("auto-plan-reinvest").checked = false;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    const result = window.generateAutoPlanBuyRows();
    assert.equal(result.ok, true, result.message || "expected successful generation");

    const rows = parseRowItems(result.rows);
    assert.equal(rows.length, 1);
    const byBond = qtyByBond(rows[0].items);
    assert.equal(byBond.get("LONG"), 2, "longer horizon should take all lots at equal coupon/price");
    assert.equal(byBond.get("SHORT") || 0, 0);
  } finally {
    dom.window.close();
  }
});

test("budget is allocated to the best-yield bond", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "HIGH", coupon: "20", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
        { bond: "LOW", coupon: "10", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [{ bond: "HIGH", quantity: "5" }],
      buys: [],
    });

    setSelectedBonds(document, ["HIGH", "LOW"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "100";
    document.getElementById("auto-plan-reinvest").checked = true;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    const result = window.generateAutoPlanBuyRows();
    assert.equal(result.ok, true, result.message || "expected successful generation");

    const rows = parseRowItems(result.rows);
    assert.equal(rows.length, 1);
    const byBond = qtyByBond(rows[0].items);
    assert.equal(byBond.get("HIGH"), 2, "100 topup + 100 coupons from HIGH holdings => 2 lots of HIGH");
    assert.equal(byBond.get("LOW") || 0, 0, "LOW should not receive budget when HIGH yield score is better");
  } finally {
    dom.window.close();
  }
});

test("monthly price drift reduces lot count on later purchase dates", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "AAA", coupon: "10", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [],
      buys: [],
    });

    setSelectedBonds(document, ["AAA"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-07-31";
    document.getElementById("auto-plan-topup-amount").value = "500";
    document.getElementById("auto-plan-reinvest").checked = false;
    document.getElementById("auto-plan-diversify").checked = false;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    document.getElementById("auto-plan-price-drift-pct").value = "";
    const noDrift = window.generateAutoPlanBuyRows();
    assert.equal(noDrift.ok, true, noDrift.message || "expected successful generation");
    const rows0 = parseRowItems(noDrift.rows);
    const july0 = rows0.find((r) => String(r.date || "").startsWith("2026-07"));
    assert.ok(july0, "expected a July purchase");
    assert.equal(qtyByBond(july0.items).get("AAA"), 5, "500 / 100 = 5 lots without drift");

    document.getElementById("auto-plan-price-drift-pct").value = "10";
    const withDrift = window.generateAutoPlanBuyRows();
    assert.equal(withDrift.ok, true, withDrift.message || "expected successful generation");
    const rows1 = parseRowItems(withDrift.rows);
    const july1 = rows1.find((r) => String(r.date || "").startsWith("2026-07"));
    assert.ok(july1);
    assert.equal(qtyByBond(july1.items).get("AAA"), 4, "two months from May: factor 1.1^2, price 121, 500/121 -> 4 lots");
    const may1 = rows1.find((r) => String(r.date || "").startsWith("2026-05"));
    assert.ok(may1);
    assert.equal(qtyByBond(may1.items).get("AAA"), 5, "May still at base price");
  } finally {
    dom.window.close();
  }
});

test("trade commission reduces affordable lots", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "AAA", coupon: "10", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [],
      buys: [],
    });

    setSelectedBonds(document, ["AAA"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "205";
    document.getElementById("auto-plan-reinvest").checked = false;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    const commInput = document.getElementById("strategy-sidebar-commission-pct");
    if (commInput) commInput.value = "0";
    const noComm = window.generateAutoPlanBuyRows();
    assert.equal(noComm.ok, true, noComm.message || "expected successful generation");
    const q0 = qtyByBond(JSON.parse(noComm.rows[0].items)).get("AAA");
    assert.equal(q0, 2, "205 without commission -> 2 lots at 100");

    if (commInput) commInput.value = "5";
    const withComm = window.generateAutoPlanBuyRows();
    assert.equal(withComm.ok, true, withComm.message || "expected successful generation");
    const q1 = qtyByBond(JSON.parse(withComm.rows[0].items)).get("AAA");
    assert.equal(q1, 1, "205 with 5% commission -> deal cost 105 per lot -> 1 lot");
  } finally {
    dom.window.close();
  }
});

test("diversification spreads budget across selected bonds", async () => {
  const { dom, window, document } = await setupApp();
  try {
    setTableRows(window, {
      bonds: [
        { bond: "AAA", coupon: "10", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
        { bond: "BBB", coupon: "10", bondPrice: "100", payoutMonths: "1,2,3,4,5,6,7,8,9,10,11,12", startDate: "2025-01-01", endDate: "2030-12-31" },
      ],
      holdings: [],
      buys: [],
    });

    setSelectedBonds(document, ["AAA", "BBB"]);
    setSelectedDays(document, [10]);

    document.getElementById("auto-plan-start").value = "2026-05-01";
    document.getElementById("auto-plan-end").value = "2026-05-31";
    document.getElementById("auto-plan-topup-amount").value = "200";
    document.getElementById("auto-plan-reinvest").checked = false;

    const strategySelect = document.getElementById("auto-plan-strategy");
    if (strategySelect.options.length) strategySelect.value = strategySelect.options[0].value;

    document.getElementById("auto-plan-diversify").checked = false;
    const noDiversify = window.generateAutoPlanBuyRows();
    assert.equal(noDiversify.ok, true, noDiversify.message || "expected successful generation");
    const noDivItems = JSON.parse(noDiversify.rows[0].items);
    const noDivByBond = qtyByBond(noDivItems);
    assert.equal(noDivByBond.get("AAA"), 2, "without diversification, budget should go to first-best bond");
    assert.equal(noDivByBond.get("BBB") || 0, 0);

    document.getElementById("auto-plan-diversify").checked = true;
    const diversify = window.generateAutoPlanBuyRows();
    assert.equal(diversify.ok, true, diversify.message || "expected successful generation");
    const divItems = JSON.parse(diversify.rows[0].items);
    const divByBond = qtyByBond(divItems);
    assert.equal(divByBond.get("AAA"), 1);
    assert.equal(divByBond.get("BBB"), 1);
  } finally {
    dom.window.close();
  }
});
