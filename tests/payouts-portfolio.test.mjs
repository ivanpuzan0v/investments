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
  window.matchMedia = () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
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

function parseRuMoney(text) {
  const cleaned = String(text || "")
    .replace(/[^\d,.-]/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : NaN;
}

test("buildPayoutSeries includes coupon in bond start month for holdings", async () => {
  const { dom, window } = await setupApp();
  try {
    const bonds = [{ bond: "AAA", coupon: "10", payoutMonths: "1,3", startDate: "2026-01-15", endDate: "2026-03-31" }];
    const buys = [{ date: "2026-02-15", items: JSON.stringify([{ bond: "AAA", price: 100, quantity: 3 }]) }];
    const holdings = [{ bond: "AAA", quantity: "2" }];

    const out = window.buildPayoutSeries(bonds, buys, holdings);
    assert.deepEqual(Array.from(out.allDates), ["2026-01", "2026-03"]);
    assert.equal(out.seriesByBond.length, 1);
    const points = Array.from(out.seriesByBond[0].points)
      .map((p) => ({ monthKey: String(p.monthKey), amount: Number(p.amount) }))
      .sort((a, b) => a.monthKey.localeCompare(b.monthKey));
    assert.deepEqual(points, [
      { monthKey: "2026-01", amount: 20 },
      { monthKey: "2026-03", amount: 50 },
    ]);
  } finally {
    dom.window.close();
  }
});

test("coupon and investment helpers return correct totals", async () => {
  const { dom, window } = await setupApp();
  try {
    const chartData = {
      allDates: ["2026-01", "2026-02"],
      seriesByBond: [
        { bond: "AAA", matchBond: "AAA", points: [{ monthKey: "2026-01", amount: 10 }, { monthKey: "2026-02", amount: 30 }] },
        { bond: "BBB", matchBond: "BBB", points: [{ monthKey: "2026-01", amount: 5 }] },
      ],
    };

    assert.equal(window.getGrossCouponTotalForMonth(chartData.seriesByBond, "2026-01"), 15);
    assert.equal(window.getGrossCouponTotalForMonth(chartData.seriesByBond, "2026-02"), 30);

    const buys = [
      { date: "2025-12-31", items: JSON.stringify([{ bond: "AAA", price: 110, quantity: 1 }]) },
      { date: "2026-01-15", items: JSON.stringify([{ bond: "AAA", price: 100, quantity: 2 }, { bond: "BBB", price: 50, quantity: 1 }]) },
      { date: "2026-06-01", bond: "CCC", price: "20", quantity: "3" },
    ];
    assert.equal(window.computeTotalInvestedFromBuys(buys), 420);
    assert.equal(window.computeInvestedFromBuysInCalendarYear(buys, "2026"), 310);
    assert.equal(window.computeInvestedFromBuysInCalendarYear(buys, "2025"), 110);

    const net = window.toNetChartData(chartData, 0.13);
    assert.equal(net.seriesByBond[0].points[0].amount, 8.7);
    assert.equal(net.seriesByBond[1].points[0].amount, 4.35);
  } finally {
    dom.window.close();
  }
});

test("renderSummary calculates gross, tax, net and invested totals", async () => {
  const { dom, window, document } = await setupApp();
  try {
    document.getElementById("tax-rate").value = "10";
    window.localStorage.setItem("invest_planner_summary_pie_year_v1", "2026");

    const chartData = {
      allDates: ["2026-01", "2026-02"],
      seriesByBond: [
        { bond: "AAA", matchBond: "AAA", points: [{ monthKey: "2026-01", amount: 100 }] },
        { bond: "BBB", matchBond: "BBB", points: [{ monthKey: "2026-02", amount: 50 }] },
      ],
    };
    const buys = [
      { date: "2026-01-10", items: JSON.stringify([{ bond: "AAA", price: 100, quantity: 1 }, { bond: "BBB", price: 80, quantity: 1 }]) },
    ];
    const holdings = [{ bond: "AAA", quantity: "2" }];

    window.renderSummary(chartData, buys, holdings);

    assert.equal(parseRuMoney(document.getElementById("total-net").textContent), 135);
    assert.equal(parseRuMoney(document.getElementById("total-gross").textContent), 150);
    assert.equal(parseRuMoney(document.getElementById("total-tax").textContent), 15);
    assert.equal(parseRuMoney(document.getElementById("total-invested").textContent), 180);

    const rows = Array.from(document.querySelectorAll("#summary-tbody tr"));
    assert.equal(rows.length, 2, "summary table should contain one row per bond from chart data");
  } finally {
    dom.window.close();
  }
});

test("coupon payout chart can show all years when chart-year is all time", async () => {
  const { dom, window, document } = await setupApp();
  try {
    document.getElementById("chart-year").value = "__all__";
    window.localStorage.setItem("invest_planner_chart_year_v1", "__all__");

    const chartData = {
      allDates: ["2025-06", "2026-01", "2027-03"],
      seriesByBond: [
        {
          bond: "AAA",
          matchBond: "AAA",
          points: [
            { monthKey: "2025-06", amount: 1 },
            { monthKey: "2026-01", amount: 2 },
            { monthKey: "2027-03", amount: 3 },
          ],
        },
      ],
    };

    window.renderChart(chartData);

    const years = Array.from(
      document.querySelectorAll("#chart-content .chartAxisLabel--interactive[data-axis-month]")
    )
      .map((el) => el.getAttribute("data-axis-month"))
      .filter(Boolean);
    assert.deepEqual(years, ["2025", "2026", "2027"]);

    const chartDataTwoMonthsOneYear = {
      allDates: ["2026-01", "2026-06", "2027-03"],
      seriesByBond: [
        {
          bond: "AAA",
          matchBond: "AAA",
          points: [
            { monthKey: "2026-01", amount: 2 },
            { monthKey: "2026-06", amount: 5 },
            { monthKey: "2027-03", amount: 3 },
          ],
        },
      ],
    };
    window.renderChart(chartDataTwoMonthsOneYear);
    const y2026 = document.querySelector(
      '#chart-content .chartAxisLabel--interactive[data-axis-month="2026"]'
    );
    assert.ok(y2026, "one column per calendar year");
    assert.equal(Number(y2026.getAttribute("data-axis-total")), 7);
  } finally {
    dom.window.close();
  }
});

test("renderPortfolioChart aggregates start, topups, coupons and end value", async () => {
  const { dom, window, document } = await setupApp();
  try {
    document.getElementById("portfolio-start-date").value = "2026-01-05";
    document.getElementById("portfolio-start-value").value = "100";
    window.renderPortfolioTopupPeriodsEditor([
      { startYmd: "2026-01-01", endYmd: "2026-03-31", amount: 50 },
    ]);
    document.getElementById("portfolio-chart-year").value = "__all__";

    const chartData = {
      allDates: ["2026-01", "2026-02", "2026-03"],
      seriesByBond: [
        {
          bond: "AAA",
          matchBond: "AAA",
          points: [
            { monthKey: "2026-01", amount: 10 },
            { monthKey: "2026-02", amount: 20 },
            { monthKey: "2026-03", amount: 30 },
          ],
        },
      ],
    };

    window.renderPortfolioChart(chartData);

    assert.equal(parseRuMoney(document.getElementById("portfolio-start-total").textContent), 100);
    assert.equal(parseRuMoney(document.getElementById("portfolio-coupon-total").textContent), 60);
    assert.equal(parseRuMoney(document.getElementById("portfolio-topup-total").textContent), 150);
    assert.equal(parseRuMoney(document.getElementById("portfolio-end-total").textContent), 310);
  } finally {
    dom.window.close();
  }
});

test("portfolio chart applies different topup amounts by periods", async () => {
  const { dom, window, document } = await setupApp();
  try {
    document.getElementById("portfolio-start-date").value = "2026-01-05";
    document.getElementById("portfolio-start-value").value = "100";
    window.renderPortfolioTopupPeriodsEditor([
      { startYmd: "2026-01-01", endYmd: "2026-01-31", amount: 40 },
      { startYmd: "2026-02-01", endYmd: "2026-03-31", amount: 70 },
    ]);
    document.getElementById("portfolio-chart-year").value = "__all__";

    const chartData = {
      allDates: ["2026-01", "2026-02", "2026-03"],
      seriesByBond: [
        {
          bond: "AAA",
          matchBond: "AAA",
          points: [
            { monthKey: "2026-01", amount: 10 },
            { monthKey: "2026-02", amount: 20 },
            { monthKey: "2026-03", amount: 30 },
          ],
        },
      ],
    };

    window.renderPortfolioChart(chartData);

    assert.equal(parseRuMoney(document.getElementById("portfolio-start-total").textContent), 100);
    assert.equal(parseRuMoney(document.getElementById("portfolio-coupon-total").textContent), 60);
    assert.equal(parseRuMoney(document.getElementById("portfolio-topup-total").textContent), 180);
    assert.equal(parseRuMoney(document.getElementById("portfolio-end-total").textContent), 340);

    const hoverZones = Array.from(document.querySelectorAll("#portfolio-chart-content .portfolioHoverZone"));
    const byMonth = new Map(hoverZones.map((z) => [z.getAttribute("data-portfolio-month"), Number(z.getAttribute("data-portfolio-topup"))]));
    assert.equal(byMonth.get("2026-01"), 40);
    assert.equal(byMonth.get("2026-02"), 70);
    assert.equal(byMonth.get("2026-03"), 70);
  } finally {
    dom.window.close();
  }
});
