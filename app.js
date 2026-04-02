const BONDS_KEY = "invest_planner_bonds_v2";
const BUYS_KEY = "invest_planner_bond_buys_v2";
const HOLDINGS_KEY = "invest_planner_holdings_v1";
const TAX_RATE_KEY = "invest_planner_tax_rate_v1";
const ACTIVE_TAB_KEY = "invest_planner_active_tab_v1";
const CHART_YEAR_KEY = "invest_planner_chart_year_v1";
const PORTFOLIO_CHART_YEAR_KEY = "invest_planner_portfolio_chart_year_v1";
const PORTFOLIO_START_DATE_KEY = "invest_planner_portfolio_start_date_v1";
const PORTFOLIO_START_VALUE_KEY = "invest_planner_portfolio_start_value_v1";
const PORTFOLIO_MONTHLY_TOPUP_KEY = "invest_planner_portfolio_monthly_topup_v1";
const PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY = "invest_planner_portfolio_monthly_topup_end_date_v1";
const ACCRUED_CALC_KEY = "invest_planner_accrued_calc_v1";
const YIELD_CALC_KEY = "invest_planner_yield_calc_v1";

const bondsTbody = document.getElementById("assets-tbody");
const buysTbody = document.getElementById("txns-tbody");
const holdingsTbody = document.getElementById("holdings-tbody");
const yieldTbody = document.getElementById("yield-tbody");

const bondTpl = document.getElementById("asset-row-template");
const buyTpl = document.getElementById("txn-row-template");
const holdingTpl = document.getElementById("holding-row-template");
const yieldTpl = document.getElementById("yield-row-template");

const addBondBtn = document.getElementById("add-asset-row");
const addBuyBtn = document.getElementById("add-txn-row");
const addHoldingBtn = document.getElementById("add-holding-row");
const addYieldRowBtn = document.getElementById("add-yield-row");
const resetAllBtn = document.getElementById("reset-all");
const chartContent = document.getElementById("chart-content");
const returnsChartSvg = document.getElementById("returns-chart");
const chartLegend = document.getElementById("chart-legend");
const chartYearSelect = document.getElementById("chart-year");
const portfolioChartYearSelect = document.getElementById("portfolio-chart-year");
const portfolioChartSvg = document.getElementById("portfolio-chart");
const portfolioChartContent = document.getElementById("portfolio-chart-content");
const portfolioStartDateInput = document.getElementById("portfolio-start-date");
const portfolioStartValueInput = document.getElementById("portfolio-start-value");
const portfolioMonthlyTopupInput = document.getElementById("portfolio-monthly-topup");
const portfolioStartTotalEl = document.getElementById("portfolio-start-total");
const portfolioCouponTotalEl = document.getElementById("portfolio-coupon-total");
const portfolioEndTotalEl = document.getElementById("portfolio-end-total");
const portfolioMonthlyTopupEndDateInput = document.getElementById("portfolio-monthly-topup-end-date");
const summaryTbody = document.getElementById("summary-tbody");
const totalGrossEl = document.getElementById("total-gross");
const totalTaxEl = document.getElementById("total-tax");
const totalNetEl = document.getElementById("total-net");
const totalInvestedEl = document.getElementById("total-invested");
const yieldAnnualEl = document.getElementById("yield-annual");
const yieldAnnualLabelEl = document.getElementById("yield-annual-label");
const yieldMonthlyEl = document.getElementById("yield-monthly");
const taxRateInput = document.getElementById("tax-rate");
const yearPiesEl = document.getElementById("year-pies");
const monthModalOverlay = document.getElementById("month-modal-overlay");
const monthModalClose = document.getElementById("month-modal-close");
const monthGrid = document.getElementById("month-grid");
const monthClearBtn = document.getElementById("month-clear");
const monthDoneBtn = document.getElementById("month-done");
const monthStartDateInput = document.getElementById("month-start-date");
const monthEndDateInput = document.getElementById("month-end-date");
const monthModeAllBtn = document.getElementById("month-mode-all");
const monthModeCustomBtn = document.getElementById("month-mode-custom");
const buyModalOverlay = document.getElementById("buy-modal-overlay");
const buyModalClose = document.getElementById("buy-modal-close");
const buyModalTitle = document.getElementById("buy-modal-title");
const buyDateInput = document.getElementById("buy-date");
const buyItemsWrap = document.getElementById("buy-items");
const buyAddItemBtn = document.getElementById("buy-add-item");
const buySaveBtn = document.getElementById("buy-save");
const tabPortfolioBtn = document.getElementById("tab-portfolio");
const tabPlanningBtn = document.getElementById("tab-planning");
const tabChartsBtn = document.getElementById("tab-charts");
const tabCalculatorsBtn = document.getElementById("tab-calculators");
const portfolioPanel = document.getElementById("tab-panel-portfolio");
const planningPanel = document.getElementById("tab-panel-planning");
const chartsPanel = document.getElementById("tab-panel-charts");
const calculatorsPanel = document.getElementById("tab-panel-calculators");
const accruedNominalInput = document.getElementById("accrued-nominal");
const accruedCouponRateInput = document.getElementById("accrued-coupon-rate");
const accruedLastCouponDateInput = document.getElementById("accrued-last-coupon-date");
const accruedNextCouponDateInput = document.getElementById("accrued-next-coupon-date");
const accruedSettlementDateInput = document.getElementById("accrued-settlement-date");
const accruedDaysElapsedEl = document.getElementById("accrued-days-elapsed");
const accruedDaysPeriodEl = document.getElementById("accrued-days-period");
const accruedIncomeResultEl = document.getElementById("accrued-income-result");
const accruedFormulaEl = document.getElementById("accrued-formula");
const yieldFormulaEl = document.getElementById("yield-formula");
const yieldModalOverlay = document.getElementById("yield-modal-overlay");
const yieldModalClose = document.getElementById("yield-modal-close");
const yieldModalTitle = document.getElementById("yield-modal-title");
const yieldModalBondNameInput = document.getElementById("yield-modal-bond-name");
const yieldModalPurchaseDateInput = document.getElementById("yield-modal-purchase-date");
const yieldModalMaturityDateInput = document.getElementById("yield-modal-maturity-date");
const yieldModalPaymentsPerYearInput = document.getElementById("yield-modal-payments-per-year");
const yieldModalPaidCouponsCountInput = document.getElementById("yield-modal-paid-coupons-count");
const yieldModalNominalInput = document.getElementById("yield-modal-nominal");
const yieldModalPurchasePriceInput = document.getElementById("yield-modal-purchase-price");
const yieldModalCouponRateInput = document.getElementById("yield-modal-coupon-rate");
const yieldModalAccruedIncomeInput = document.getElementById("yield-modal-accrued-income");
const yieldModalRedemptionPriceInput = document.getElementById("yield-modal-redemption-price");
const yieldSaveBtn = document.getElementById("yield-save");

const DATE_YMD_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
const MONTHS_RU = [
  "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
];

const BUY_MODAL_TITLE_NEW = "Новая покупка облигаций";
const BUY_MODAL_TITLE_EDIT = "Редактирование покупки";
const YIELD_MODAL_TITLE_NEW = "Новая облигация для сравнения";
const YIELD_MODAL_TITLE_EDIT = "Редактирование облигации";
const tableSortState = {
  assets: { field: "", dir: "asc" },
  holdings: { field: "", dir: "asc" },
  summary: { field: "", dir: "asc" },
};

function readRows(tbody) {
  return Array.from(tbody.querySelectorAll("tr")).map((tr) => {
    const row = {};
    tr.querySelectorAll("input[data-field], select[data-field], textarea[data-field]").forEach((el) => {
      const key = el.getAttribute("data-field");
      if (el instanceof HTMLSelectElement && el.multiple) {
        row[key] = Array.from(el.selectedOptions).map((o) => o.value).join(",");
      } else {
        row[key] = (el.value || "").trim();
      }
    });
    return row;
  });
}

function setActiveTab(tab) {
  const nextTab = ["portfolio", "planning", "charts", "calculators"].includes(tab) ? tab : "planning";
  if (portfolioPanel) portfolioPanel.hidden = nextTab !== "portfolio";
  if (planningPanel) planningPanel.hidden = nextTab !== "planning";
  if (chartsPanel) chartsPanel.hidden = nextTab !== "charts";
  if (calculatorsPanel) calculatorsPanel.hidden = nextTab !== "calculators";
  if (tabPortfolioBtn) tabPortfolioBtn.classList.toggle("is-active", nextTab === "portfolio");
  if (tabPlanningBtn) tabPlanningBtn.classList.toggle("is-active", nextTab === "planning");
  if (tabChartsBtn) tabChartsBtn.classList.toggle("is-active", nextTab === "charts");
  if (tabCalculatorsBtn) tabCalculatorsBtn.classList.toggle("is-active", nextTab === "calculators");
  localStorage.setItem(ACTIVE_TAB_KEY, nextTab);
}

function writeRows(tbody, template, rows) {
  tbody.innerHTML = "";
  rows.forEach((row) => {
    const node = template.content.cloneNode(true);
    const tr = node.querySelector("tr");
    tr.querySelectorAll("input[data-field], select[data-field], textarea[data-field]").forEach((el) => {
      const key = el.getAttribute("data-field");
      if (row[key] === undefined) return;
      if (el instanceof HTMLSelectElement && el.multiple) {
        const selected = String(row[key] || "")
          .split(",")
          .map((v) => v.trim())
          .filter(Boolean);
        Array.from(el.options).forEach((opt) => {
          opt.selected = selected.includes(opt.value);
        });
      } else {
        el.value = row[key];
      }
    });
    tbody.appendChild(node);
  });
}

function renderTableEmptyState(tbody, colSpan, title, hint) {
  if (!tbody) return;
  tbody.innerHTML = `<tr data-empty-state="1"><td colspan="${colSpan}" class="tableEmptyCell"><div class="emptyState emptyState--table"><div class="emptyState__title">${escapeHtml(
    title
  )}</div><div class="emptyState__hint">${escapeHtml(hint)}</div></div></td></tr>`;
}

function clearTableEmptyState(tbody) {
  if (!tbody) return;
  const emptyRow = tbody.querySelector('tr[data-empty-state="1"]');
  if (emptyRow) emptyRow.remove();
}

function syncStaticTableEmptyStates() {
  if (bondsTbody && !bondsTbody.querySelector("tr")) {
    renderTableEmptyState(
      bondsTbody,
      4,
      "Список облигаций пока пуст",
      "Добавьте облигацию, укажите купон и выберите даты выплат, чтобы строить прогноз выплат."
    );
  }
  if (holdingsTbody && !holdingsTbody.querySelector("tr")) {
    renderTableEmptyState(
      holdingsTbody,
      3,
      "Текущий портфель не заполнен",
      "Добавьте уже купленные облигации и их количество, чтобы учитывать их в графиках и сводке."
    );
  }
  if (buysTbody && !buysTbody.querySelector("tr")) {
    renderTableEmptyState(
      buysTbody,
      4,
      "План покупок пока пуст",
      "Добавьте будущие покупки облигаций, чтобы увидеть, как они повлияют на выплаты и стоимость портфеля."
    );
  }
  if (yieldTbody && !yieldTbody.querySelector("tr")) {
    renderTableEmptyState(
      yieldTbody,
      9,
      "Сравнение доходности пока пусто",
      "Добавьте облигацию через модальное окно, чтобы сравнить купонную доходность и доходность к погашению."
    );
  }
}

function buildSvgEmptyState(title, hint) {
  return `<text x="450" y="160" text-anchor="middle" fill="currentColor" opacity="0.78" font-size="17" font-weight="600">${escapeHtml(
    title
  )}</text><text x="450" y="188" text-anchor="middle" fill="currentColor" opacity="0.6" font-size="12.5">${escapeHtml(hint)}</text>`;
}

function updateSortableHeaders() {
  document.querySelectorAll("th[data-sort-table][data-sort-field]").forEach((th) => {
    const tableName = th.getAttribute("data-sort-table") || "";
    const field = th.getAttribute("data-sort-field") || "";
    const state = tableSortState[tableName];
    if (!state || state.field !== field) {
      th.removeAttribute("data-sort-dir");
      return;
    }
    th.setAttribute("data-sort-dir", state.dir === "desc" ? "desc" : "asc");
  });
}

function getBondPayoutSortValue(row) {
  const start = normalizeYMD(row.startDate) || "9999-99-99";
  const end = normalizeYMD(row.endDate) || "9999-99-99";
  const months = parseMonthList(row.payoutMonths).join(",").padStart(24, "0");
  if (!row.startDate && !row.endDate && !row.payoutMonths) return "";
  return `${start}|${end}|${months}`;
}

function getSortValue(row, field) {
  switch (field) {
    case "bond":
      return String(row.bond || "").trim().toUpperCase();
    case "coupon":
    case "quantity": {
      const n = parseNumber(row[field]);
      return Number.isFinite(n) ? n : NaN;
    }
    case "qty":
    case "invested":
    case "tax":
    case "net": {
      const n = Number(row[field]);
      return Number.isFinite(n) ? n : NaN;
    }
    case "payoutPeriod":
      return getBondPayoutSortValue(row);
    default:
      return String(row[field] || "").trim();
  }
}

function compareSortValues(a, b, dir) {
  const aEmpty = typeof a === "number" ? !Number.isFinite(a) : !String(a || "").trim();
  const bEmpty = typeof b === "number" ? !Number.isFinite(b) : !String(b || "").trim();
  if (aEmpty && bEmpty) return 0;
  if (aEmpty) return 1;
  if (bEmpty) return -1;

  let cmp = 0;
  if (typeof a === "number" && typeof b === "number") cmp = a - b;
  else cmp = String(a).localeCompare(String(b), "ru", { numeric: true, sensitivity: "base" });
  return dir === "desc" ? -cmp : cmp;
}

function sortRowsForTable(rows, field, dir) {
  return [...rows].sort((a, b) => compareSortValues(getSortValue(a, field), getSortValue(b, field), dir));
}

function sortPlanningTable(tableName) {
  const state = tableSortState[tableName];
  if (!state?.field) return;
  if (tableName === "assets") {
    const rows = sortRowsForTable(sanitizeBondRows(readRows(bondsTbody)), state.field, state.dir);
    writeRows(bondsTbody, bondTpl, rows);
    syncDateSummaries();
    persistAndRender();
    return;
  }
  if (tableName === "holdings") {
    const rows = sortRowsForTable(readRows(holdingsTbody), state.field, state.dir);
    writeRows(holdingsTbody, holdingTpl, rows);
    persistAndRender();
    return;
  }
  if (tableName === "summary") {
    renderAll();
  }
}

function defaultBondRows() {
  return [];
}

function sanitizeBondRows(rows) {
  return rows.filter((row) => {
    const bond = String(row.bond || "").trim();
    const coupon = String(row.coupon || "").trim();
    const payoutMonths = String(row.payoutMonths || "").trim();
    const startDate = String(row.startDate || "").trim();
    const endDate = String(row.endDate || "").trim();
    return Boolean(bond || coupon || payoutMonths || startDate || endDate);
  });
}

function defaultBuyRows() {
  return [];
}

function defaultHoldingRows() {
  return [];
}

function normalizeBuyRowsForUI(rows) {
  return rows
    .map((row) => {
      const date = normalizeYMD(row.date) || "";
      const items = parseBuyItems(row.items);
      if (items.length) return { date, items: JSON.stringify(items) };

      // migration from legacy single-item shape
      const legacyBond = String(row.bond || "").trim().toUpperCase();
      const legacyPrice = parseNumber(row.price);
      const legacyQty = parseNumber(row.quantity);
      if (legacyBond && Number.isFinite(legacyPrice) && Number.isFinite(legacyQty) && legacyQty > 0) {
        return {
          date,
          items: JSON.stringify([{ bond: legacyBond, price: legacyPrice, quantity: legacyQty }]),
        };
      }

      return { date, items: "[]" };
    })
    .filter(isBuyRowComplete);
}

function parseBuyItems(raw) {
  try {
    const parsed = JSON.parse(String(raw || "[]"));
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((x) => {
        const qRaw = parseNumber(x.quantity);
        const q = Number.isFinite(qRaw) ? Math.round(qRaw) : NaN;
        return {
          bond: String(x.bond || "").trim().toUpperCase(),
          price: parseNumber(x.price),
          quantity: q,
        };
      })
      .filter((x) => x.bond && Number.isFinite(x.price) && Number.isFinite(x.quantity) && x.quantity > 0);
  } catch {
    return [];
  }
}

function isBuyRowComplete(row) {
  const dateYmd = normalizeYMD(row.date);
  const items = parseBuyItems(row.items);
  if (dateYmd && items.length > 0) return true;
  const legacyBond = String(row.bond || "").trim();
  const legacyPrice = parseNumber(row.price);
  const legacyQty = Math.round(parseNumber(row.quantity));
  return Boolean(dateYmd && legacyBond && Number.isFinite(legacyPrice) && Number.isFinite(legacyQty) && legacyQty > 0);
}

function sortBuyRowsByDate(rows) {
  const completeRows = rows.filter(isBuyRowComplete);
  if (completeRows.length <= 1) {
    return { rows: [...rows], didSort: false };
  }

  const sortedComplete = [...completeRows].sort((a, b) => {
    const aYmd = normalizeYMD(a.date);
    const bYmd = normalizeYMD(b.date);
    const aTs = aYmd ? ymdToUTCms(aYmd) : Number.MAX_SAFE_INTEGER;
    const bTs = bYmd ? ymdToUTCms(bYmd) : Number.MAX_SAFE_INTEGER;
    if (aTs !== bTs) return aTs - bTs;
    const aTotal = parseBuyItems(a.items).reduce((s, x) => s + x.price * x.quantity, 0);
    const bTotal = parseBuyItems(b.items).reduce((s, x) => s + x.price * x.quantity, 0);
    return aTotal - bTotal;
  });

  let ptr = 0;
  const rebuilt = rows.map((row) => (isBuyRowComplete(row) ? sortedComplete[ptr++] : row));
  const before = JSON.stringify(rows);
  const after = JSON.stringify(rebuilt);
  return { rows: rebuilt, didSort: before !== after };
}

function parseNumber(value) {
  if (!value) return NaN;
  const normalized = String(value)
    .replace(/\s+/g, "")
    .replace(/,/g, ".");
  return Number(normalized);
}

function normalizeYMD(raw) {
  const v = String(raw || "").trim();
  const m = DATE_YMD_RE.exec(v);
  if (!m) return null;

  const y = Number(m[1]);
  const mo = Number(m[2]); // 1..12
  const d = Number(m[3]); // 1..31

  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;

  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;

  // Возвращаем в исходном y-m-d формате с ведущими нулями
  return `${String(y).padStart(4, "0")}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

function ymdToUTCms(ymd) {
  const m = DATE_YMD_RE.exec(ymd);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  return Date.UTC(y, mo - 1, d);
}

function diffDaysInclusiveStartExclusiveEnd(startYmd, endYmd) {
  const startTs = ymdToUTCms(startYmd);
  const endTs = ymdToUTCms(endYmd);
  if (!Number.isFinite(startTs) || !Number.isFinite(endTs) || endTs < startTs) return NaN;
  return Math.round((endTs - startTs) / 86400000);
}

function toMonthKeyFromYMD(ymd) {
  const m = DATE_YMD_RE.exec(ymd);
  if (!m) return "";
  return `${m[1]}-${m[2]}`;
}

function formatMonthKey(monthKey) {
  const m = /^(\d{4})-(\d{2})$/.exec(monthKey);
  if (!m) return monthKey;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const dt = new Date(Date.UTC(y, mo, 1));
  return {
    month: new Intl.DateTimeFormat("ru-RU", { month: "long" }).format(dt),
    year: String(y),
  };
}

function getSeriesPalette() {
  return [
    "#007AFF",
    "#34C759",
    "#FF9500",
    "#AF52DE",
    "#FF2D55",
    "#30B0C7",
    "#8E8E93",
  ];
}

function getYearFromMonthKey(monthKey) {
  const m = /^(\d{4})-(\d{2})$/.exec(monthKey);
  return m ? m[1] : "";
}

function getAvailableYears(chartData) {
  const years = new Set((chartData?.allDates || []).map(getYearFromMonthKey).filter(Boolean));
  return Array.from(years).sort();
}

function ensureChartYearOptions(chartData) {
  if (!chartYearSelect) return null;
  const years = getAvailableYears(chartData);
  if (!years.length) {
    chartYearSelect.innerHTML = "";
    return null;
  }

  const saved = localStorage.getItem(CHART_YEAR_KEY);
  const current = chartYearSelect.value;
  let selected = saved && years.includes(saved) ? saved : years[0];
  if (current && years.includes(current)) selected = current;

  chartYearSelect.innerHTML = years.map((y) => `<option value="${y}">${y}</option>`).join("");
  chartYearSelect.value = selected;
  localStorage.setItem(CHART_YEAR_KEY, selected);
  return selected;
}

function ensurePortfolioChartYearOptions(points) {
  if (!portfolioChartYearSelect) return null;
  const years = Array.from(new Set((points || []).map((p) => getYearFromMonthKey(p.monthKey)).filter(Boolean))).sort();
  if (!years.length) {
    portfolioChartYearSelect.innerHTML = "";
    return null;
  }

  const saved = localStorage.getItem(PORTFOLIO_CHART_YEAR_KEY);
  const current = portfolioChartYearSelect.value;
  let selected = saved && years.includes(saved) ? saved : years[0];
  if (current && years.includes(current)) selected = current;

  portfolioChartYearSelect.innerHTML = years.map((y) => `<option value="${y}">${y}</option>`).join("");
  portfolioChartYearSelect.value = selected;
  localStorage.setItem(PORTFOLIO_CHART_YEAR_KEY, selected);
  return selected;
}

function polarToCartesian(cx, cy, r, angleDeg) {
  const rad = ((angleDeg - 90) * Math.PI) / 180.0;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function arcPath(cx, cy, r, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, r, endAngle);
  const end = polarToCartesian(cx, cy, r, startAngle);
  const largeArc = endAngle - startAngle <= 180 ? 0 : 1;
  return `M ${cx} ${cy} L ${start.x} ${start.y} A ${r} ${r} 0 ${largeArc} 0 ${end.x} ${end.y} Z`;
}

function renderYearPies(chartData) {
  if (!yearPiesEl) return;
  hideChartTooltip();
  clearAllChartHover();
  const seriesByBond = chartData?.seriesByBond || [];
  const palette = getSeriesPalette();

  // year -> map(bond -> net amount)
  const byYear = new Map();
  seriesByBond.forEach((series) => {
    series.points.forEach((p) => {
      const year = getYearFromMonthKey(p.monthKey);
      if (!year) return;
      if (!byYear.has(year)) byYear.set(year, new Map());
      const m = byYear.get(year);
      m.set(series.bond, (m.get(series.bond) || 0) + p.amount);
    });
  });

  const years = Array.from(byYear.keys()).sort();
  if (!years.length) {
    yearPiesEl.innerHTML = `<div class="emptyState emptyState--card"><div class="emptyState__title">Пока нет выплат по годам</div><div class="emptyState__hint">Добавьте облигации, текущие позиции или план покупок, чтобы увидеть распределение купонов по годам.</div></div>`;
    return;
  }

  const cards = years.map((year) => {
    const bondMap = byYear.get(year);
    const entries = Array.from(bondMap.entries())
      .filter(([, value]) => value > 0)
      .sort((a, b) => b[1] - a[1]);
    const total = entries.reduce((sum, [, value]) => sum + value, 0);
    if (total <= 0) {
      return `<div class="yearPieCard"><div class="yearPieCard__title">${year}</div><div class="yearPieCard__sum">0.00</div></div>`;
    }

    let currentAngle = 0;
    const slices = entries.length === 1
      ? `<circle class="yearPieSlice" cx="50" cy="50" r="44" fill="${palette[0]}" data-slice-bond="${escapeHtml(
          entries[0][0]
        )}" data-slice-year="${escapeHtml(String(year))}" data-slice-amount="${entries[0][1]}" data-slice-pct="100"></circle>`
      : entries
          .map(([bond, value], idx) => {
            const share = value / total;
            const angle = share * 360;
            const start = currentAngle;
            const end = currentAngle + angle;
            currentAngle = end;
            const color = palette[idx % palette.length];
            const pct = (value / total) * 100;
            return { bond, value, color, path: arcPath(50, 50, 44, start, end), pct };
          })
          .map(
            (s) =>
              `<path class="yearPieSlice" d="${s.path}" fill="${s.color}" data-slice-bond="${escapeHtml(
                s.bond
              )}" data-slice-year="${escapeHtml(String(year))}" data-slice-amount="${s.value}" data-slice-pct="${s.pct}"></path>`
          )
          .join("");

    const legend = entries
      .map(([bond, value], idx) => {
        const color = palette[idx % palette.length];
        return `<div class="yearPieLegend__item">
          <span class="yearPieLegend__dot" style="background:${color}"></span>
          <span class="yearPieLegend__bond">${escapeHtml(bond)}</span>
          <span class="yearPieLegend__sum">${formatMoney(value)}</span>
        </div>`;
      })
      .join("");

    return `<div class="yearPieCard">
      <div class="yearPieCard__title">${year}</div>
      <div class="yearPieCard__sum">${formatMoney(total)}</div>
      <svg class="yearPie" viewBox="0 0 100 100" aria-label="Выплаты за ${year}">
        ${slices}
      </svg>
      <div class="yearPieLegend">${legend}</div>
    </div>`;
  });

  yearPiesEl.innerHTML = cards.join("");
}

function renderChartLegend(seriesByBond, palette) {
  if (!chartLegend) return;
  const visibleSeries = seriesByBond.filter((series) => (series.points || []).some((point) => Number(point.amount) > 0));
  if (!visibleSeries.length) {
    chartLegend.innerHTML = "";
    return;
  }

  chartLegend.innerHTML = visibleSeries
    .map((series, idx) => {
      const color = palette[idx % palette.length];
      const safeBond = String(series.bond || `Bond ${idx + 1}`)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;");
      const safeMatchBond = String(series.matchBond || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
      const safeColor = String(color || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
      const total = (series.points || []).reduce((sum, point) => sum + (Number(point.amount) || 0), 0);
      return `<span class="chartLegend__item" data-legend-bond="${safeMatchBond}" data-legend-color="${safeColor}">
        <span class="chartLegend__dot" style="background:${color}"></span>
        <span class="chartLegend__bond">${safeBond}</span>
        <span class="chartLegend__sep" aria-hidden="true">·</span>
        <span class="chartLegend__sum">${formatMoney(total)}</span>
      </span>`;
    })
    .join("");
}

function formatAmount(amount) {
  return new Intl.NumberFormat("ru-RU", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount);
}

function formatMoney(amount) {
  return `${formatAmount(amount)} ₽`;
}

/** Целое количество для полей ввода модалки (без десятичных) */
function formatQuantityInt(n) {
  if (!Number.isFinite(n)) return "";
  return String(Math.round(Number(n)));
}

/** Отображение количества бумаг (целое, ru-RU) */
function formatQuantityRu(n) {
  if (!Number.isFinite(n)) return "0";
  return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(Math.round(Number(n)));
}

function formatPercentValue(value) {
  if (!Number.isFinite(value)) return "—";
  return `${formatAmount(value)}%`;
}

function toNetChartData(chartData, taxRate) {
  const rate = Number.isFinite(taxRate) ? taxRate : 0;
  const factor = 1 - rate;
  return {
    allDates: chartData?.allDates || [],
    seriesByBond: (chartData?.seriesByBond || []).map((series) => ({
      ...series,
      points: (series.points || []).map((point) => ({
        ...point,
        amount: (Number(point.amount) || 0) * factor,
      })),
    })),
  };
}

function setInputNumericValue(input, value, fallback = "0") {
  if (!input) return;
  const numeric = parseNumber(value);
  input.value = Number.isFinite(numeric) ? String(numeric) : fallback;
}

function getDefaultAccruedCalcState() {
  return {
    nominal: 1000,
    couponRate: 12,
    lastCouponDate: "",
    nextCouponDate: "",
    settlementDate: getTodayYMD(),
  };
}

function getDefaultYieldCalcState() {
  return [];
}

function getDefaultYieldCalcRow() {
  return {
    bondName: "",
    purchaseDate: "",
    maturityDate: "",
    paymentsPerYear: 4,
    paidCouponsCount: 0,
    nominal: 1000,
    purchasePrice: 950,
    couponRate: 12,
    accruedIncome: 0,
    redemptionPrice: 1000,
  };
}

function readCalcState(storageKey, fallback) {
  try {
    const raw = localStorage.getItem(storageKey);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    if (Array.isArray(fallback)) return Array.isArray(parsed) ? parsed : fallback;
    return parsed && typeof parsed === "object" ? { ...fallback, ...parsed } : fallback;
  } catch {
    return fallback;
  }
}

function saveAccruedCalcState() {
  const payload = {
    nominal: parseNumber(accruedNominalInput?.value || ""),
    couponRate: parseNumber(accruedCouponRateInput?.value || ""),
    lastCouponDate: normalizeYMD(accruedLastCouponDateInput?.value || "") || "",
    nextCouponDate: normalizeYMD(accruedNextCouponDateInput?.value || "") || "",
    settlementDate: normalizeYMD(accruedSettlementDateInput?.value || "") || "",
  };
  localStorage.setItem(ACCRUED_CALC_KEY, JSON.stringify(payload));
}

function saveYieldCalcState() {
  const payload = readRows(yieldTbody);
  localStorage.setItem(YIELD_CALC_KEY, JSON.stringify(payload));
}

function loadCalculatorStates() {
  const accrued = readCalcState(ACCRUED_CALC_KEY, getDefaultAccruedCalcState());
  setInputNumericValue(accruedNominalInput, accrued.nominal, "1000");
  setInputNumericValue(accruedCouponRateInput, accrued.couponRate, "12");
  if (accruedLastCouponDateInput) accruedLastCouponDateInput.value = normalizeYMD(accrued.lastCouponDate || "") || "";
  if (accruedNextCouponDateInput) accruedNextCouponDateInput.value = normalizeYMD(accrued.nextCouponDate || "") || "";
  if (accruedSettlementDateInput) accruedSettlementDateInput.value = normalizeYMD(accrued.settlementDate || "") || getTodayYMD();

  const yieldCalc = readCalcState(YIELD_CALC_KEY, getDefaultYieldCalcState());
  const normalizedRows = Array.isArray(yieldCalc) ? yieldCalc : getDefaultYieldCalcState();
  writeRows(yieldTbody, yieldTpl, normalizedRows);
  renderYieldCalculator();
}

function formatDateRu(ymd) {
  const normalized = normalizeYMD(ymd || "");
  if (!normalized) return "—";
  const [y, m, d] = normalized.split("-");
  return `${d}.${m}.${y}`;
}

function renderAccruedIncomeCalculator() {
  if (!accruedDaysElapsedEl || !accruedDaysPeriodEl || !accruedIncomeResultEl) return;
  const nominal = parseNumber(accruedNominalInput?.value || "");
  const couponRate = parseNumber(accruedCouponRateInput?.value || "");
  const lastCouponDate = normalizeYMD(accruedLastCouponDateInput?.value || "");
  const nextCouponDate = normalizeYMD(accruedNextCouponDateInput?.value || "");
  const settlementDate = normalizeYMD(accruedSettlementDateInput?.value || "");
  const daysPeriod = diffDaysInclusiveStartExclusiveEnd(lastCouponDate, nextCouponDate);
  const daysElapsed = diffDaysInclusiveStartExclusiveEnd(lastCouponDate, settlementDate);
  const isValidRange = Number.isFinite(daysPeriod) && Number.isFinite(daysElapsed) && daysPeriod > 0 && daysElapsed >= 0 && daysElapsed <= daysPeriod;
  const annualCouponIncome = Number.isFinite(nominal) && Number.isFinite(couponRate) ? nominal * (couponRate / 100) : NaN;
  const accruedIncome = isValidRange && Number.isFinite(annualCouponIncome) ? annualCouponIncome * (daysElapsed / 365) : NaN;
  const canShowFormula = isValidRange && Number.isFinite(nominal) && Number.isFinite(couponRate) && Number.isFinite(accruedIncome);

  accruedDaysElapsedEl.textContent = Number.isFinite(daysElapsed) && daysElapsed >= 0 ? String(daysElapsed) : "0";
  accruedDaysPeriodEl.textContent = Number.isFinite(daysPeriod) && daysPeriod >= 0 ? String(daysPeriod) : "0";
  accruedIncomeResultEl.textContent = Number.isFinite(accruedIncome) ? formatMoney(accruedIncome) : "—";
  if (accruedFormulaEl) {
    accruedFormulaEl.textContent = canShowFormula
      ? `Формула: ${formatAmount(nominal)} × ${formatAmount(couponRate)}% / 100 × ${daysElapsed} / 365 = ${formatMoney(accruedIncome)}.`
      : "Укажите номинал, ставку купона и корректные даты внутри одного купонного периода.";
  }
}

function computeYieldMetrics(row) {
  const purchaseDate = normalizeYMD(row.purchaseDate);
  const maturityDate = normalizeYMD(row.maturityDate);
  const paymentsPerYearRaw = parseNumber(row.paymentsPerYear);
  const paidCouponsCountRaw = parseNumber(row.paidCouponsCount);
  const nominal = parseNumber(row.nominal);
  const purchasePrice = parseNumber(row.purchasePrice);
  const couponRate = parseNumber(row.couponRate);
  const accruedIncome = parseNumber(row.accruedIncome);
  const redemptionPrice = parseNumber(row.redemptionPrice);
  const purchaseTs = ymdToUTCms(purchaseDate || "");
  const maturityTs = ymdToUTCms(maturityDate || "");
  const holdingYears = Number.isFinite(purchaseTs) && Number.isFinite(maturityTs) && maturityTs > purchaseTs
    ? (maturityTs - purchaseTs) / (365 * 86400000)
    : NaN;
  const paymentsPerYear = Number.isFinite(paymentsPerYearRaw) && paymentsPerYearRaw > 0
    ? Math.max(1, Math.round(paymentsPerYearRaw))
    : NaN;
  const paidCouponsCount = Number.isFinite(paidCouponsCountRaw) && paidCouponsCountRaw >= 0
    ? Math.max(0, Math.round(paidCouponsCountRaw))
    : NaN;

  const investedAmount = Number.isFinite(purchasePrice) && Number.isFinite(accruedIncome)
    ? purchasePrice + accruedIncome
    : NaN;
  const annualCouponIncome = Number.isFinite(nominal) && Number.isFinite(couponRate)
    ? nominal * (couponRate / 100)
    : NaN;
  const couponPerPayment = Number.isFinite(annualCouponIncome) && Number.isFinite(paymentsPerYear) && paymentsPerYear > 0
    ? annualCouponIncome / paymentsPerYear
    : NaN;
  const plannedCouponPayments = Number.isFinite(holdingYears) && Number.isFinite(paymentsPerYear)
    ? Math.max(0, Math.ceil(holdingYears * paymentsPerYear))
    : NaN;
  const remainingCouponPayments = Number.isFinite(plannedCouponPayments) && Number.isFinite(paidCouponsCount)
    ? Math.max(0, plannedCouponPayments - paidCouponsCount)
    : NaN;
  const couponIncome = Number.isFinite(couponPerPayment) && Number.isFinite(remainingCouponPayments)
    ? couponPerPayment * remainingCouponPayments
    : NaN;
  const redemptionIncome = Number.isFinite(redemptionPrice) && Number.isFinite(investedAmount)
    ? redemptionPrice - investedAmount
    : NaN;
  const totalIncome = Number.isFinite(couponIncome) && Number.isFinite(redemptionIncome)
    ? couponIncome + redemptionIncome
    : NaN;
  const investmentReturn = Number.isFinite(totalIncome) && Number.isFinite(investedAmount) && investedAmount > 0
    ? (totalIncome / investedAmount) * 100
    : NaN;
  const currentYield = Number.isFinite(annualCouponIncome) && Number.isFinite(investedAmount) && investedAmount > 0
    ? (annualCouponIncome / investedAmount) * 100
    : NaN;
  const ytmApprox = Number.isFinite(totalIncome)
    && Number.isFinite(investedAmount)
    && Number.isFinite(holdingYears)
    && holdingYears > 0
    && investedAmount > 0
    ? (totalIncome / investedAmount / holdingYears) * 100
    : NaN;

  return {
    annualCouponIncome,
    investedAmount,
    currentYield,
    ytmApprox,
    paymentsPerYear,
    paidCouponsCount,
    holdingYears,
    remainingCouponPayments,
    couponIncome,
    redemptionIncome,
    totalIncome,
    investmentReturn,
  };
}

function renderYieldCalculator() {
  if (!yieldTbody) return;
  if (!yieldTbody.querySelector("tr")) {
    syncStaticTableEmptyStates();
  }
  const rows = Array.from(yieldTbody.querySelectorAll("tr"));
  let hasCalculatedRow = false;
  rows.forEach((tr) => {
    const row = {};
    tr.querySelectorAll("input[data-field]").forEach((el) => {
      const key = el.getAttribute("data-field");
      if (key) row[key] = el.value;
    });
    const metrics = computeYieldMetrics(row);
    const bondNameEl = tr.querySelector('[data-display-for="bondName"]');
    const paymentsPerYearEl = tr.querySelector('[data-display-for="paymentsPerYear"]');
    const couponIncomeEl = tr.querySelector('[data-display-for="couponIncome"]');
    const redemptionIncomeEl = tr.querySelector('[data-display-for="redemptionIncome"]');
    const totalIncomeEl = tr.querySelector('[data-display-for="totalIncome"]');
    const investmentReturnEl = tr.querySelector('[data-display-for="investmentReturn"]');
    const holdingYearsEl = tr.querySelector('[data-display-for="holdingYears"]');
    const currentYieldEl = tr.querySelector('[data-display-for="currentYield"]');
    const ytmEl = tr.querySelector('[data-display-for="ytm"]');
    if (bondNameEl) bondNameEl.textContent = String(row.bondName || "").trim() || "—";
    if (paymentsPerYearEl) paymentsPerYearEl.textContent = Number.isFinite(metrics.paymentsPerYear) ? String(metrics.paymentsPerYear) : "—";
    if (couponIncomeEl) couponIncomeEl.textContent = Number.isFinite(metrics.couponIncome) ? formatMoney(metrics.couponIncome) : "—";
    if (redemptionIncomeEl) redemptionIncomeEl.textContent = Number.isFinite(metrics.redemptionIncome) ? formatMoney(metrics.redemptionIncome) : "—";
    if (totalIncomeEl) totalIncomeEl.textContent = Number.isFinite(metrics.totalIncome) ? formatMoney(metrics.totalIncome) : "—";
    if (investmentReturnEl) investmentReturnEl.textContent = formatPercentValue(metrics.investmentReturn);
    if (holdingYearsEl) holdingYearsEl.textContent = Number.isFinite(metrics.holdingYears) ? `${formatAmount(metrics.holdingYears)} г.` : "—";
    if (currentYieldEl) currentYieldEl.textContent = formatPercentValue(metrics.currentYield);
    if (ytmEl) ytmEl.textContent = formatPercentValue(metrics.ytmApprox);
    if (Number.isFinite(metrics.ytmApprox)) hasCalculatedRow = true;
  });
  if (yieldFormulaEl) {
    yieldFormulaEl.textContent = hasCalculatedRow
      ? "Текущая доходность = купон за год / грязная цена. Доходность к погашению рассчитывается по дате покупки и погашения с учетом количества выплат в год, уже выплаченных купонов и оставшегося купонного потока."
      : "Добавьте облигацию через модальное окно: укажите даты, частоту выплат, уже выплаченные купоны, цену, купон и НКД.";
  }
}

function renderCalculators() {
  renderAccruedIncomeCalculator();
  renderYieldCalculator();
  syncStaticTableEmptyStates();
}

function computeTotalInvestedFromBuys(buys) {
  return buys.reduce((sum, row) => {
    const items = parseBuyItems(row.items);
    if (items.length) {
      return sum + items.reduce((s, i) => s + i.price * i.quantity, 0);
    }
    const bond = String(row.bond || "").trim();
    const price = parseNumber(row.price);
    const qty = Math.round(parseNumber(row.quantity));
    if (bond && Number.isFinite(price) && Number.isFinite(qty)) return sum + price * qty;
    return sum;
  }, 0);
}

/** @type {HTMLDivElement | null} */
let chartTooltipHost = null;

/** @type {SVGElement | null} */
let lastHoveredBar = null;
/** @type {SVGElement | null} */
let lastHoveredSlice = null;
/** @type {HTMLElement | null} */
let lastHoveredLegendItem = null;
/** @type {SVGElement | null} */
let lastHoveredPortfolioZone = null;

function clearBarHover() {
  if (lastHoveredBar) {
    lastHoveredBar.classList.remove("chartHover--active");
    lastHoveredBar = null;
  }
}

function clearSliceHover() {
  if (lastHoveredSlice) {
    lastHoveredSlice.classList.remove("chartHover--active");
    lastHoveredSlice = null;
  }
}

function clearLegendHover() {
  if (lastHoveredLegendItem) {
    lastHoveredLegendItem.classList.remove("chartLegend__item--active");
    lastHoveredLegendItem = null;
  }
}

function clearPortfolioHover() {
  if (lastHoveredPortfolioZone) {
    lastHoveredPortfolioZone.classList.remove("portfolioHoverZone--active");
    lastHoveredPortfolioZone = null;
  }
}

function clearChartTextHighlight() {
  document.querySelectorAll(".chartValueLabel, .chartAxisLabel").forEach((el) => {
    if (!(el instanceof SVGTextElement)) return;
    el.classList.remove("chartText--active");
    el.classList.remove("chartValueLabel--active");
    el.setAttribute("fill", el.getAttribute("data-default-fill") || "currentColor");
    el.setAttribute("opacity", el.getAttribute("data-default-opacity") || "1");
  });
}

function applyChartTextHighlight(matchBond, color, monthKey = "", valueMode = "series") {
  const bond = String(matchBond || "").trim().toUpperCase();
  const activeColor = String(color || "").trim();
  clearChartTextHighlight();
  if (!bond || !activeColor) return;

  const activeMonths = new Set();
  document.querySelectorAll(".chartValueLabel").forEach((el) => {
    if (!(el instanceof SVGTextElement)) return;
    const labelBond = String(el.getAttribute("data-label-match") || "").trim().toUpperCase();
    const labelMonth = String(el.getAttribute("data-label-month") || "").trim();
    const shouldActivate = valueMode === "single"
      ? labelBond === bond && labelMonth === monthKey
      : labelBond === bond;
    if (!shouldActivate) return;
    activeMonths.add(labelMonth);
    el.classList.add("chartText--active");
    el.classList.add("chartValueLabel--active");
    el.setAttribute("fill", activeColor);
    el.setAttribute("opacity", "0.98");
  });

  document.querySelectorAll(".chartAxisLabel").forEach((el) => {
    if (!(el instanceof SVGTextElement)) return;
    const labelMonth = String(el.getAttribute("data-axis-month") || "").trim();
    const shouldActivate = valueMode === "single" ? labelMonth === monthKey : activeMonths.has(labelMonth);
    if (!shouldActivate) return;
    el.classList.add("chartText--active");
    el.setAttribute("fill", activeColor);
    el.setAttribute("opacity", "0.98");
  });
}

function setLegendFilterBond(matchBond, color = "", legendItem = null) {
  const bond = String(matchBond || "").trim().toUpperCase();
  clearBarHover();
  clearLegendHover();
  clearChartTextHighlight();
  if (!chartContent || !bond) {
    if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
    return;
  }
  chartContent.classList.add("chartContent--dimOthers");
  document.querySelectorAll(".chartBar").forEach((el) => {
    if (!(el instanceof SVGElement)) return;
    const barBond = String(el.getAttribute("data-bar-match") || "").trim().toUpperCase();
    el.classList.toggle("chartHover--active", barBond === bond);
  });
  applyChartTextHighlight(bond, color);
  if (legendItem) {
    legendItem.classList.add("chartLegend__item--active");
    lastHoveredLegendItem = legendItem;
  }
}

function clearAllChartHover() {
  clearBarHover();
  clearSliceHover();
  clearLegendHover();
  clearPortfolioHover();
  clearChartTextHighlight();
  if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
  document.querySelectorAll(".yearPie.yearPie--dimOthers").forEach((svg) => svg.classList.remove("yearPie--dimOthers"));
}

function monthLabelFromKey(monthKey) {
  const ml = formatMonthKey(monthKey);
  if (typeof ml === "string") return ml;
  return `${ml.month} ${ml.year}`;
}

function ensureChartTooltip() {
  if (chartTooltipHost) return chartTooltipHost;
  const el = document.createElement("div");
  el.className = "chartTooltip";
  el.setAttribute("role", "tooltip");
  el.hidden = true;
  document.body.appendChild(el);
  chartTooltipHost = el;
  return el;
}

function hideChartTooltip() {
  if (chartTooltipHost) {
    chartTooltipHost.hidden = true;
    chartTooltipHost.classList.remove("chartTooltip--visible");
    chartTooltipHost.innerHTML = "";
  }
}

/**
 * @param {string} titleHtml
 * @param {string} metaHtml
 * @param {number} clientX
 * @param {number} clientY
 */
function showChartTooltip(titleHtml, metaHtml, clientX, clientY) {
  const el = ensureChartTooltip();
  el.innerHTML = `<div class="chartTooltip__title">${titleHtml}</div><div class="chartTooltip__meta">${metaHtml}</div>`;
  el.hidden = false;
  el.classList.add("chartTooltip--visible");
  const margin = 14;
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  requestAnimationFrame(() => {
    const tw = el.offsetWidth;
    const th = el.offsetHeight;
    let left = clientX + margin;
    let top = clientY + margin;
    if (left + tw > vw - 8) left = clientX - tw - margin;
    if (top + th > vh - 8) top = clientY - th - margin;
    left = Math.max(8, Math.min(left, vw - tw - 8));
    top = Math.max(8, Math.min(top, vh - th - 8));
    el.style.left = `${left}px`;
    el.style.top = `${top}px`;
  });
}

/** @param {PointerEvent} e */
function onBarChartPointerMove(e) {
  const t = /** @type {EventTarget | null} */ (e.target);
  if (!(t instanceof Element)) return;
  const bar = t.closest(".chartBar");
  if (!bar || !(bar instanceof SVGElement)) {
    const axisLabel = t.closest(".chartAxisLabel[data-axis-total]");
    if (axisLabel instanceof SVGElement) {
      hideChartTooltip();
      clearBarHover();
      if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
      const monthKey = axisLabel.getAttribute("data-axis-month") || "";
      const total = parseNumber(axisLabel.getAttribute("data-axis-total"));
      const period = monthLabelFromKey(monthKey);
      showChartTooltip("Сумма выплат", `${escapeHtml(period)} · ${escapeHtml(formatMoney(total))}`, e.clientX, e.clientY);
      return;
    }
    hideChartTooltip();
    clearBarHover();
    if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
    return;
  }
  if (chartContent) chartContent.classList.add("chartContent--dimOthers");
  if (lastHoveredBar !== bar) {
    clearBarHover();
    lastHoveredBar = bar;
    bar.classList.add("chartHover--active");
  }
  const bond = bar.getAttribute("data-bar-bond") || "";
  const matchBond = bar.getAttribute("data-bar-match") || "";
  const monthKey = bar.getAttribute("data-bar-month") || "";
  const amount = parseNumber(bar.getAttribute("data-bar-amount"));
  const color = bar.getAttribute("fill") || "";
  const period = monthLabelFromKey(monthKey);
  applyChartTextHighlight(matchBond, color, monthKey, "single");
  showChartTooltip(escapeHtml(bond), `${escapeHtml(period)} · ${escapeHtml(formatMoney(amount))}`, e.clientX, e.clientY);
}

/** @param {PointerEvent} e */
function onPieChartPointerMove(e) {
  const t = /** @type {EventTarget | null} */ (e.target);
  if (!(t instanceof Element)) return;
  const slice = t.closest(".yearPieSlice");
  if (!slice || !(slice instanceof SVGElement)) {
    hideChartTooltip();
    clearSliceHover();
    document.querySelectorAll(".yearPie.yearPie--dimOthers").forEach((svg) => svg.classList.remove("yearPie--dimOthers"));
    return;
  }
  const pieSvg = slice.closest("svg.yearPie");
  if (pieSvg instanceof SVGSVGElement) {
    document.querySelectorAll(".yearPie.yearPie--dimOthers").forEach((svg) => {
      if (svg !== pieSvg) svg.classList.remove("yearPie--dimOthers");
    });
    pieSvg.classList.add("yearPie--dimOthers");
  }
  if (lastHoveredSlice !== slice) {
    clearSliceHover();
    lastHoveredSlice = slice;
    slice.classList.add("chartHover--active");
  }
  const bond = slice.getAttribute("data-slice-bond") || "";
  const year = slice.getAttribute("data-slice-year") || "";
  const amount = parseNumber(slice.getAttribute("data-slice-amount"));
  const share = parseNumber(slice.getAttribute("data-slice-pct"));
  const shareStr = Number.isFinite(share) ? `${formatAmount(share)}%` : "—";
  showChartTooltip(
    escapeHtml(bond),
    `${escapeHtml(year)} · ${escapeHtml(shareStr)} · ${escapeHtml(formatMoney(amount))}`,
    e.clientX,
    e.clientY
  );
}

function onBarChartPointerLeave() {
  hideChartTooltip();
  clearBarHover();
  if (lastHoveredLegendItem) {
    const bond = lastHoveredLegendItem.getAttribute("data-legend-bond") || "";
    const color = lastHoveredLegendItem.getAttribute("data-legend-color") || "";
    setLegendFilterBond(bond, color, lastHoveredLegendItem);
    return;
  }
  clearChartTextHighlight();
  if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
}

function onPieChartPointerLeave() {
  hideChartTooltip();
  clearSliceHover();
  document.querySelectorAll(".yearPie.yearPie--dimOthers").forEach((svg) => svg.classList.remove("yearPie--dimOthers"));
}

/** @param {PointerEvent} e */
function onPortfolioChartPointerMove(e) {
  const t = /** @type {EventTarget | null} */ (e.target);
  if (!(t instanceof Element)) return;
  const zone = t.closest(".portfolioHoverZone");
  if (!zone || !(zone instanceof SVGElement)) {
    hideChartTooltip();
    clearPortfolioHover();
    return;
  }
  if (lastHoveredPortfolioZone !== zone) {
    clearPortfolioHover();
    lastHoveredPortfolioZone = zone;
    zone.classList.add("portfolioHoverZone--active");
  }
  const monthKey = zone.getAttribute("data-portfolio-month") || "";
  const value = parseNumber(zone.getAttribute("data-portfolio-value"));
  const x = parseNumber(zone.getAttribute("data-portfolio-x"));
  const y = parseNumber(zone.getAttribute("data-portfolio-y"));
  if (portfolioChartContent) {
    const hoverLine = portfolioChartContent.querySelector(".portfolioHoverLine");
    if (hoverLine instanceof SVGLineElement && Number.isFinite(x) && Number.isFinite(y)) {
      hoverLine.setAttribute("x1", String(x));
      hoverLine.setAttribute("x2", String(x));
      hoverLine.setAttribute("y1", String(y));
      hoverLine.setAttribute("opacity", "1");
    }
  }
  const period = monthLabelFromKey(monthKey);
  showChartTooltip("Стоимость портфеля", `${escapeHtml(period)} · ${escapeHtml(formatMoney(value))}`, e.clientX, e.clientY);
}

function onPortfolioChartPointerLeave() {
  hideChartTooltip();
  clearPortfolioHover();
  if (portfolioChartContent) {
    const hoverLine = portfolioChartContent.querySelector(".portfolioHoverLine");
    if (hoverLine instanceof SVGLineElement) hoverLine.setAttribute("opacity", "0");
  }
}

function bindChartTooltips() {
  if (returnsChartSvg && !returnsChartSvg.dataset.tooltipBound) {
    returnsChartSvg.dataset.tooltipBound = "1";
    returnsChartSvg.addEventListener("pointermove", onBarChartPointerMove);
    returnsChartSvg.addEventListener("pointerleave", onBarChartPointerLeave);
  }
  if (portfolioChartSvg && !portfolioChartSvg.dataset.tooltipBound) {
    portfolioChartSvg.dataset.tooltipBound = "1";
    portfolioChartSvg.addEventListener("pointermove", onPortfolioChartPointerMove);
    portfolioChartSvg.addEventListener("pointerleave", onPortfolioChartPointerLeave);
  }
  if (chartLegend && !chartLegend.dataset.tooltipBound) {
    chartLegend.dataset.tooltipBound = "1";
    chartLegend.addEventListener("pointermove", (e) => {
      const t = /** @type {EventTarget | null} */ (e.target);
      if (!(t instanceof Element)) return;
      const item = t.closest(".chartLegend__item");
      if (!item || !(item instanceof HTMLElement)) {
        clearLegendHover();
        clearBarHover();
        clearChartTextHighlight();
        if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
        return;
      }
      const bond = item.getAttribute("data-legend-bond") || "";
      const color = item.getAttribute("data-legend-color") || "";
      setLegendFilterBond(bond, color, item);
    });
    chartLegend.addEventListener("pointerleave", () => {
      clearLegendHover();
      clearBarHover();
      clearChartTextHighlight();
      if (chartContent) chartContent.classList.remove("chartContent--dimOthers");
    });
  }
  if (yearPiesEl && !yearPiesEl.dataset.tooltipBound) {
    yearPiesEl.dataset.tooltipBound = "1";
    yearPiesEl.addEventListener("pointermove", onPieChartPointerMove);
    yearPiesEl.addEventListener("pointerleave", onPieChartPointerLeave);
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getTaxRateDecimal() {
  const raw = taxRateInput ? parseNumber(taxRateInput.value) : 13;
  if (!Number.isFinite(raw)) return 0.13;
  const clamped = Math.max(0, Math.min(100, raw));
  return clamped / 100;
}

function parseMonthList(raw) {
  return String(raw || "")
    .split(",")
    .map((v) => Number(v.trim()))
    .filter((n) => Number.isInteger(n) && n >= 1 && n <= 12);
}

function summarizeMonths(months) {
  if (!months.length) return "—";
  if (months.length === 12) return "Все месяцы";
  return months.map((m) => MONTHS_RU[m - 1]).join(", ");
}

function formatMonthYearFromYMD(ymd) {
  const m = DATE_YMD_RE.exec(String(ymd || ""));
  if (!m) return "";
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const dt = new Date(Date.UTC(y, mo, 1));
  const month = new Intl.DateTimeFormat("ru-RU", { month: "long" }).format(dt);
  return `${month} ${y}`;
}

function monthKeyToUTCms(monthKey) {
  return Date.parse(`${monthKey}-01T00:00:00Z`) || 0;
}

function monthKeyFromDateInput(ymd) {
  const normalized = normalizeYMD(ymd || "");
  return normalized ? toMonthKeyFromYMD(normalized) : "";
}

function isYearLabelMonth(monthKey, idx, total) {
  const m = /^(\d{4})-(\d{2})$/.exec(monthKey);
  if (!m) return false;
  if (idx === 0 || idx === total - 1) return true;
  return Number(m[2]) === 1;
}

function addMonthsToMonthKey(monthKey, delta) {
  const m = /^(\d{4})-(\d{2})$/.exec(monthKey);
  if (!m) return "";
  const dt = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1 + delta, 1));
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, "0")}`;
}

function buildMonthRange(startMonthKey, endMonthKey) {
  if (!startMonthKey || !endMonthKey) return [];
  if (monthKeyToUTCms(startMonthKey) > monthKeyToUTCms(endMonthKey)) return [];
  const result = [];
  let current = startMonthKey;
  while (current && monthKeyToUTCms(current) <= monthKeyToUTCms(endMonthKey)) {
    result.push(current);
    current = addMonthsToMonthKey(current, 1);
  }
  return result;
}

function getTodayYMD() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function normalizeHoldingRowsForUI(rows) {
  return rows
    .map((row) => ({
      bond: String(row.bond || "").trim().toUpperCase(),
      quantity: formatQuantityInt(parseNumber(row.quantity)),
    }))
    .filter((row) => row.bond && Number.isFinite(parseNumber(row.quantity)) && Math.round(parseNumber(row.quantity)) > 0);
}

function parseHoldingItems(rows) {
  return rows
    .map((row) => {
      const qRaw = parseNumber(row.quantity);
      const quantity = Number.isFinite(qRaw) ? Math.round(qRaw) : NaN;
      return {
        bond: String(row.bond || "").trim().toUpperCase(),
        quantity,
      };
    })
    .filter((row) => row.bond && Number.isFinite(row.quantity) && row.quantity > 0);
}

function sanitizeHoldingRows(rows) {
  return rows
    .map((row) => ({
      bond: String(row.bond || "").trim().toUpperCase(),
      quantity: formatQuantityInt(parseNumber(row.quantity)),
    }))
    .filter((row) => row.bond && Number.isFinite(parseNumber(row.quantity)) && Math.round(parseNumber(row.quantity)) > 0);
}

function buildPayoutSeries(bonds, buys, holdings) {
  const normalizedBonds = bonds
    .map((row, idx) => ({
      bond: String(row.bond || "").toUpperCase(),
      matchBond: String(row.bond || "").toUpperCase(),
      legendBond: String(row.bond || "").toUpperCase() || `BOND ${idx + 1}`,
      coupon: parseNumber(row.coupon),
      payoutMonths: Array.from(new Set(parseMonthList(row.payoutMonths))).sort((a, b) => a - b),
      startDate: normalizeYMD(row.startDate),
      endDate: normalizeYMD(row.endDate),
    }))
    .filter((row) => row.startDate && row.endDate && ymdToUTCms(row.startDate) <= ymdToUTCms(row.endDate));

  const normalizedBuys = buys
    .flatMap((row) => {
      const dateYMD = normalizeYMD(row.date);
      const dateUTCms = ymdToUTCms(dateYMD || "");
      if (!dateYMD) return [];

      // New format: grouped items in one row
      const groupedItems = parseBuyItems(row.items);
      if (groupedItems.length) {
        return groupedItems.map((item) => ({
          dateYMD,
          dateUTCms,
          bond: item.bond,
          quantity: item.quantity,
        }));
      }

      // Backward compatibility: old single-item rows
      const bond = String(row.bond || "").toUpperCase();
      const quantity = parseNumber(row.quantity);
      if (!bond || !Number.isFinite(quantity) || quantity <= 0) return [];
      return [{ dateYMD, dateUTCms, bond, quantity }];
    });

  const holdingsAsBuys = parseHoldingItems(holdings).map((item) => {
    const dateYMD = getTodayYMD();
    return {
      dateYMD,
      dateUTCms: ymdToUTCms(dateYMD),
      bond: item.bond,
      quantity: item.quantity,
    };
  });

  const allDateSet = new Set();
  const seriesByBond = normalizedBonds.map((bondRow) => {
    const amountByMonth = new Map();
    if (!bondRow.payoutMonths.length) {
      return { bond: bondRow.legendBond, matchBond: bondRow.matchBond, points: [] };
    }

    const startTs = ymdToUTCms(bondRow.startDate);
    const endTs = ymdToUTCms(bondRow.endDate);
    const start = new Date(startTs);
    const end = new Date(endTs);

    for (let y = start.getUTCFullYear(); y <= end.getUTCFullYear(); y += 1) {
      bondRow.payoutMonths.forEach((monthNum) => {
        const monthIndex = monthNum - 1;
        const paymentDate = new Date(Date.UTC(y, monthIndex, 1));
        const paymentTs = paymentDate.getTime();
        if (paymentTs < startTs || paymentTs > endTs) return;

        const monthKey = `${String(y).padStart(4, "0")}-${String(monthNum).padStart(2, "0")}`;
        allDateSet.add(monthKey);

        const quantityAtDate = normalizedBuys
          .concat(holdingsAsBuys)
          .filter((buyRow) => buyRow.bond === bondRow.matchBond && buyRow.dateUTCms <= paymentTs)
          .reduce((sum, buyRow) => sum + buyRow.quantity, 0);

        const couponAmount = Number.isFinite(bondRow.coupon) && bondRow.coupon > 0 ? bondRow.coupon : 0;
        const amount = Math.max(0, couponAmount * quantityAtDate);
        amountByMonth.set(monthKey, (amountByMonth.get(monthKey) || 0) + amount);
      });
    }

    const points = Array.from(amountByMonth.entries()).map(([monthKey, amount]) => {
      return { monthKey, amount };
    });

    return { bond: bondRow.legendBond, matchBond: bondRow.matchBond, points };
  });

  const allDates = Array.from(allDateSet).sort((a, b) => {
    const aTs = Date.parse(`${a}-01T00:00:00Z`) || 0;
    const bTs = Date.parse(`${b}-01T00:00:00Z`) || 0;
    return aTs - bTs;
  });
  return { allDates, seriesByBond };
}

function renderChart(chartData) {
  hideChartTooltip();
  clearAllChartHover();
  const allDatesRaw = chartData?.allDates || [];
  const seriesByBondRaw = chartData?.seriesByBond || [];
  const selectedYear = ensureChartYearOptions(chartData);
  const allDates = selectedYear ? allDatesRaw.filter((mk) => getYearFromMonthKey(mk) === selectedYear) : allDatesRaw;
  const seriesByBond = seriesByBondRaw.map((s) => ({
    ...s,
    points: s.points.filter((p) => (selectedYear ? getYearFromMonthKey(p.monthKey) === selectedYear : true)),
  }));
  const visibleDates = allDates.filter((monthKey) =>
    seriesByBond.some((series) => series.points.some((point) => point.monthKey === monthKey && point.amount > 0))
  );
  const W = 800;
  const H = 360;
  const left = 20;
  const right = 12;
  const top = 24;
  const bottom = 34;
  const innerW = W - left - right;
  const innerH = H - top - bottom;

  const lines = [];

  for (let i = 0; i <= 4; i += 1) {
    const y = top + (innerH / 4) * i;
    lines.push(`<path d="M ${left} ${y} H ${W - right}" stroke="rgba(60,60,67,0.2)" stroke-width="1"></path>`);
  }

  if (!visibleDates.length || !seriesByBond.length) {
    if (chartLegend) chartLegend.innerHTML = "";
    lines.push(buildSvgEmptyState("Нет выплат за выбранный год", "Добавьте облигации и даты купонов, чтобы построить график выплат."));
    chartContent.innerHTML = lines.join("");
    return;
  }

  const allAmounts = seriesByBond.flatMap((series) => series.points.map((point) => point.amount));
  const maxY = Math.max(...allAmounts, 0);
  const safeMaxY = maxY <= 0 ? 1 : maxY * 1.15;

  const groupWidth = innerW / visibleDates.length;
  const groupInnerWidth = Math.max(16, groupWidth * 0.78);

  const palette = getSeriesPalette();
  renderChartLegend(seriesByBond, palette);

  const amountByBondAndDate = seriesByBond.map((series) => {
    const map = new Map();
    series.points.forEach((p) => map.set(p.monthKey, p.amount));
    return map;
  });

  visibleDates.forEach((monthKey, dateIdx) => {
    const activeBars = seriesByBond
      .map((bondRow, sIdx) => ({
        bondRow,
        sIdx,
        amount: amountByBondAndDate[sIdx].get(monthKey) || 0,
      }))
      .filter((item) => item.amount > 0);
    const barsCount = activeBars.length;
    if (!barsCount) return;
    const barGap = 2;
    const maxBarWidth = 22;
    const computedWidth = (groupInnerWidth - (barsCount - 1) * barGap) / barsCount;
    const barWidth = Math.max(4, Math.min(maxBarWidth, computedWidth));
    const usedWidth = barWidth * barsCount + barGap * (barsCount - 1);
    const groupStartX = left + dateIdx * groupWidth + (groupWidth - usedWidth) / 2;

    activeBars.forEach((item, visibleIdx) => {
      const { bondRow, sIdx, amount } = item;
      const color = palette[sIdx % palette.length];
      const barHeight = (amount / safeMaxY) * innerH;
      const x = groupStartX + visibleIdx * (barWidth + barGap);
      const visibleHeight = barHeight;
      const y = top + innerH - visibleHeight;
      lines.push(
        `<rect class="chartBar" data-bar-bond="${escapeHtml(bondRow.bond)}" data-bar-match="${escapeHtml(
          bondRow.matchBond
        )}" data-bar-month="${monthKey}" data-bar-amount="${amount}" x="${x}" y="${y}" width="${barWidth}" height="${visibleHeight}" rx="3" fill="${color}" stroke="none" opacity="0.86"></rect>`
      );
      lines.push(
        `<text class="chartValueLabel" data-label-match="${escapeHtml(bondRow.matchBond)}" data-label-month="${monthKey}" data-default-fill="currentColor" data-default-opacity="0.66" pointer-events="none" x="${x + barWidth / 2}" y="${y - 4}" text-anchor="middle" fill="currentColor" opacity="0.66" font-size="7.5">${formatMoney(amount)}</text>`
      );
    });

    const labelX = left + dateIdx * groupWidth + groupWidth / 2;
    const monthLabel = formatMonthKey(monthKey);
    const monthTotal = activeBars.reduce((sum, item) => sum + item.amount, 0);
    lines.push(
      `<text class="chartAxisLabel chartAxisLabel--interactive" data-axis-month="${monthKey}" data-axis-total="${monthTotal}" data-default-fill="currentColor" data-default-opacity="0.68" x="${labelX}" y="${H - 22}" text-anchor="middle" fill="currentColor" opacity="0.68" font-size="8.5">${monthLabel.month}</text>`
    );
    lines.push(
      `<text class="chartAxisLabel" data-axis-month="${monthKey}" data-default-fill="currentColor" data-default-opacity="0.62" pointer-events="none" x="${labelX}" y="${H - 11}" text-anchor="middle" fill="currentColor" opacity="0.62" font-size="8">${monthLabel.year}</text>`
    );
  });

  chartContent.innerHTML = lines.join("");
}

function renderSummary(chartData, buys, holdings) {
  if (!summaryTbody || !totalGrossEl || !totalTaxEl || !totalNetEl) return;

  const taxRate = getTaxRateDecimal();
  const netChartData = toNetChartData(chartData, taxRate);
  const seriesByBond = netChartData?.seriesByBond || [];
  const totalInvested = computeTotalInvestedFromBuys(buys);
  renderYearPies(chartData);

  const quantityMap = new Map();
  const investedMap = new Map();
  buys
    .flatMap((row) => {
      const groupedItems = parseBuyItems(row.items);
      if (groupedItems.length) {
        return groupedItems.map((item) => ({
          bond: item.bond,
          quantity: item.quantity,
          invested: item.price * item.quantity,
        }));
      }
      const legacyBond = String(row.bond || "").toUpperCase();
      const legacyPrice = parseNumber(row.price);
      const legacyQty = parseNumber(row.quantity);
      if (!legacyBond || !Number.isFinite(legacyQty) || !Number.isFinite(legacyPrice)) return [];
      return [{
        bond: legacyBond,
        quantity: Math.round(legacyQty),
        invested: legacyPrice * Math.round(legacyQty),
      }];
    })
    .forEach((item) => {
      quantityMap.set(item.bond, (quantityMap.get(item.bond) || 0) + item.quantity);
      investedMap.set(item.bond, (investedMap.get(item.bond) || 0) + item.invested);
    });
  parseHoldingItems(holdings).forEach((item) => {
    quantityMap.set(item.bond, (quantityMap.get(item.bond) || 0) + item.quantity);
  });

  const rows = seriesByBond.map((series) => {
    const net = series.points.reduce((sum, point) => sum + point.amount, 0);
    const grossEstimated = taxRate < 1 ? net / (1 - taxRate) : net;
    const tax = grossEstimated - net;
    const qty = quantityMap.get(series.matchBond) || 0;
    const invested = investedMap.get(series.matchBond) || 0;
    return { bond: series.bond, qty, invested, tax, net };
  });
  const summarySort = tableSortState.summary;
  const visibleRows = summarySort.field ? sortRowsForTable(rows, summarySort.field, summarySort.dir) : rows;

  const totalNet = rows.reduce((sum, r) => sum + r.net, 0);
  const totalGrossEstimated = taxRate < 1 ? totalNet / (1 - taxRate) : totalNet;
  const totalTax = totalGrossEstimated - totalNet;

  totalGrossEl.textContent = formatMoney(totalGrossEstimated);
  totalTaxEl.textContent = formatMoney(totalTax);
  totalNetEl.textContent = formatMoney(totalNet);

  if (totalInvestedEl) totalInvestedEl.textContent = formatMoney(totalInvested);

  let monthlyYieldPct = NaN;
  let annualYieldPct = NaN;

  const selectedYearForYield = chartYearSelect?.value || "";
  if (totalInvested > 0) {
    if (selectedYearForYield) {
      // `totalNet` в этой сводке суммируется по всему горизонту.
      // Для "среднегодовой" доходности используем выплаты только в выбранном календарном году,
      // чтобы не "пересчитывать" их ещё раз коэффициентом 12/число_месяцев_с_выплатами.
      const netForYear = seriesByBond.reduce((sum, series) => {
        const pts = series.points || [];
        return sum + pts.reduce((s, point) => {
          const y = getYearFromMonthKey(point.monthKey);
          return y && y === selectedYearForYield ? s + (Number(point.amount) || 0) : s;
        }, 0);
      }, 0);

      annualYieldPct = (netForYear / totalInvested) * 100;
      monthlyYieldPct = annualYieldPct / 12;
    } else {
      // fallback: annualize over time span of monthKeys
      const allDates = Array.isArray(chartData?.allDates) ? chartData.allDates : [];
      const first = allDates[0];
      const last = allDates[allDates.length - 1];
      const m1 = first ? /^(\d{4})-(\d{2})$/.exec(first) : null;
      const m2 = last ? /^(\d{4})-(\d{2})$/.exec(last) : null;
      const monthsSpan = m1 && m2
        ? (Number(m2[1]) - Number(m1[1])) * 12 + (Number(m2[2]) - Number(m1[2])) + 1
        : 12;

      annualYieldPct = (totalNet / totalInvested) * (12 / Math.max(1, monthsSpan)) * 100;
      monthlyYieldPct = annualYieldPct / 12;
    }
  }
  if (yieldAnnualLabelEl) {
    yieldAnnualLabelEl.textContent = selectedYearForYield
      ? `Средняя доходность за ${selectedYearForYield} год`
      : "Средняя доходность за год";
  }
  if (yieldAnnualEl) yieldAnnualEl.textContent = formatPercentValue(annualYieldPct);
  if (yieldMonthlyEl) yieldMonthlyEl.textContent = formatPercentValue(monthlyYieldPct);

  if (!rows.length) {
    renderTableEmptyState(
      summaryTbody,
      5,
      "Сводка пока не рассчитана",
      "Здесь будут показаны выплаты по облигациям, запланированным к покупкам, а также налог и итоговая доходность."
    );
    return;
  }

  summaryTbody.innerHTML = visibleRows
    .map(
      (r) => `<tr>
        <td>${r.bond}</td>
        <td>${formatQuantityRu(r.qty)}</td>
        <td>${formatMoney(r.invested)}</td>
        <td>${formatMoney(r.tax)}</td>
        <td>${formatMoney(r.net)}</td>
      </tr>`
    )
    .join("");
}

function renderPortfolioChart(chartData) {
  if (!portfolioChartContent || !portfolioStartTotalEl || !portfolioCouponTotalEl || !portfolioEndTotalEl) return;

  const startDate = normalizeYMD(portfolioStartDateInput?.value || "") || getTodayYMD();
  const startMonthKey = toMonthKeyFromYMD(startDate);
  const startValueRaw = parseNumber(portfolioStartValueInput?.value || "");
  const startValue = Number.isFinite(startValueRaw) && startValueRaw >= 0 ? startValueRaw : 0;
  const monthlyTopupRaw = parseNumber(portfolioMonthlyTopupInput?.value || "");
  const monthlyTopup = Number.isFinite(monthlyTopupRaw) && monthlyTopupRaw >= 0 ? monthlyTopupRaw : 0;
  const monthlyTopupEndYMD = normalizeYMD(portfolioMonthlyTopupEndDateInput?.value || "") || "";
  const monthlyTopupEndTs = monthlyTopupEndYMD ? ymdToUTCms(monthlyTopupEndYMD) : null;
  const allDates = chartData?.allDates || [];
  const seriesByBond = chartData?.seriesByBond || [];
  const defaultEndMonthKey = addMonthsToMonthKey(startMonthKey, 11);
  const lastCouponMonthKey = allDates.length ? allDates[allDates.length - 1] : "";
  const lastMonthKey = lastCouponMonthKey && monthKeyToUTCms(lastCouponMonthKey) >= monthKeyToUTCms(startMonthKey)
    ? lastCouponMonthKey
    : defaultEndMonthKey;

  portfolioStartTotalEl.textContent = formatMoney(startValue);

  const monthRange = buildMonthRange(startMonthKey, lastMonthKey);
  const couponByMonth = new Map();
  monthRange.forEach((monthKey) => couponByMonth.set(monthKey, 0));
  seriesByBond.forEach((series) => {
    series.points.forEach((point) => {
      if (!couponByMonth.has(point.monthKey)) return;
      couponByMonth.set(point.monthKey, (couponByMonth.get(point.monthKey) || 0) + point.amount);
    });
  });

  let cumulative = startValue;
  const allPoints = monthRange.map((monthKey) => {
    const topupForMonth = !monthlyTopupEndTs || monthKeyToUTCms(monthKey) <= monthlyTopupEndTs ? monthlyTopup : 0;
    cumulative += topupForMonth + (couponByMonth.get(monthKey) || 0);
    return { monthKey, value: cumulative };
  });
  const selectedYear = ensurePortfolioChartYearOptions(allPoints);
  const points = selectedYear ? allPoints.filter((p) => getYearFromMonthKey(p.monthKey) === selectedYear) : allPoints;

  if (!points.length) {
    portfolioStartTotalEl.textContent = formatMoney(startValue);
    portfolioCouponTotalEl.textContent = formatMoney(0);
    portfolioEndTotalEl.textContent = formatMoney(startValue);
    portfolioChartContent.innerHTML = buildSvgEmptyState(
      "За выбранный год нет движений",
      "Смените год или добавьте облигации с купонными выплатами в этот период."
    );
    return;
  }

  const firstPoint = points[0];
  const firstPointMonthCoupon = couponByMonth.get(firstPoint.monthKey) || 0;
  const periodStartValue = firstPoint.value - monthlyTopup - firstPointMonthCoupon;
  const couponOnlyAdded = points.reduce((sum, point) => sum + (couponByMonth.get(point.monthKey) || 0), 0);
  const endValue = points[points.length - 1].value;
  portfolioStartTotalEl.textContent = formatMoney(periodStartValue);
  portfolioCouponTotalEl.textContent = formatMoney(couponOnlyAdded);
  portfolioEndTotalEl.textContent = formatMoney(endValue);

  const W = 900;
  const H = 360;
  const left = 78;
  const right = 18;
  const top = 24;
  const bottom = 42;
  const innerW = W - left - right;
  const innerH = H - top - bottom;
  const minY = Math.min(periodStartValue, ...points.map((p) => p.value));
  const maxY = Math.max(periodStartValue, ...points.map((p) => p.value));
  const safeMinY = minY <= 0 ? 0 : minY * 0.98;
  const safeMaxY = maxY <= safeMinY ? safeMinY + 1 : maxY * 1.02;
  const xStep = points.length === 1 ? innerW / 2 : innerW / Math.max(1, points.length - 1);
  const xAt = (idx) => left + (points.length === 1 ? innerW / 2 : idx * xStep);
  const yAt = (value) => top + innerH - ((value - safeMinY) / (safeMaxY - safeMinY)) * innerH;

  const lines = [];
  for (let i = 0; i <= 4; i += 1) {
    const y = top + (innerH / 4) * i;
    const value = safeMaxY - ((safeMaxY - safeMinY) / 4) * i;
    lines.push(`<path d="M ${left} ${y} H ${W - right}" stroke="rgba(60,60,67,0.2)" stroke-width="1"></path>`);
    lines.push(`<text x="${left - 12}" y="${y + 4}" text-anchor="end" fill="currentColor" opacity="0.58" font-size="9">${formatAmount(value)}</text>`);
  }
  lines.push(`<line class="portfolioHoverLine" x1="${left}" y1="${top + innerH}" x2="${left}" y2="${top + innerH}" stroke="#007AFF" stroke-width="1" opacity="0"></line>`);

  const areaPath = points
    .map((p, idx) => `${idx === 0 ? "M" : "L"} ${xAt(idx)} ${yAt(p.value)}`)
    .join(" ");
  const areaClose = `L ${xAt(points.length - 1)} ${top + innerH} L ${xAt(0)} ${top + innerH} Z`;
  const linePath = points
    .map((p, idx) => `${idx === 0 ? "M" : "L"} ${xAt(idx)} ${yAt(p.value)}`)
    .join(" ");
  lines.push(`<path d="${areaPath} ${areaClose}" fill="rgba(0,122,255,0.12)" stroke="none"></path>`);
  lines.push(`<path d="${linePath}" fill="none" stroke="#007AFF" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"></path>`);

  points.forEach((p, idx) => {
    const x = xAt(idx);
    const y = yAt(p.value);
    const monthLabel = formatMonthKey(p.monthKey);
    lines.push(`<circle cx="${x}" cy="${y}" r="3.5" fill="#007AFF"></circle>`);
    lines.push(`<text x="${x}" y="${H - 22}" text-anchor="middle" fill="currentColor" opacity="0.68" font-size="8.5">${monthLabel.month}</text>`);
    lines.push(`<text x="${x}" y="${H - 11}" text-anchor="middle" fill="currentColor" opacity="0.62" font-size="8">${monthLabel.year}</text>`);
    const zoneWidth = points.length === 1 ? innerW : Math.max(12, xStep);
    const zoneX = points.length === 1 ? left : x - zoneWidth / 2;
    lines.push(
      `<rect class="portfolioHoverZone" data-portfolio-month="${p.monthKey}" data-portfolio-value="${p.value}" data-portfolio-x="${x}" data-portfolio-y="${y}" x="${zoneX}" y="${top}" width="${zoneWidth}" height="${innerH}" fill="transparent"></rect>`
    );
  });

  portfolioChartContent.innerHTML = lines.join("");
}

function setPortfolioMoneyInputValue(input, value) {
  if (!input) return;
  const numeric = parseNumber(value);
  input.value = Number.isFinite(numeric) ? String(numeric) : "0";
}

function bindPortfolioMoneyInput(input, storageKey) {
  if (!input) return;
  input.addEventListener("input", () => {
    localStorage.setItem(storageKey, String(parseNumber(input.value) || 0));
    renderAll();
  });
}

function renderAll() {
  const bonds = sanitizeBondRows(readRows(bondsTbody));
  const buys = sortBuyRowsByDate(readRows(buysTbody)).rows.filter(isBuyRowComplete);
  const holdings = sanitizeHoldingRows(readRows(holdingsTbody));
  const payoutSeries = buildPayoutSeries(bonds, buys, holdings);
  renderPortfolioChart(payoutSeries);
  renderChart(payoutSeries);
  renderSummary(payoutSeries, buys, holdings);
  renderCalculators();
}

function persistAndRender() {
  // Важно: отменяем отложенный save, чтобы старое состояние не перезаписало новое.
  clearTimeout(saveTimer);
  syncDateSummaries();
  const bonds = sanitizeBondRows(readRows(bondsTbody));
  const sortResult = sortBuyRowsByDate(readRows(buysTbody).filter(isBuyRowComplete));
  const buysSorted = sortResult.rows;
  const holdings = sanitizeHoldingRows(readRows(holdingsTbody));
  localStorage.setItem(BONDS_KEY, JSON.stringify(bonds));
  localStorage.setItem(BUYS_KEY, JSON.stringify(buysSorted));
  localStorage.setItem(HOLDINGS_KEY, JSON.stringify(holdings));
  // Перерисовываем строки только если реально поменялся порядок.
  if (sortResult.didSort) {
    writeRows(buysTbody, buyTpl, buysSorted);
  }
  syncBuySummaries();
  syncStaticTableEmptyStates();
  if (taxRateInput) localStorage.setItem(TAX_RATE_KEY, String(parseNumber(taxRateInput.value) || 0));
  renderAll();
}

function loadAll() {
  try {
    const bondsRaw = localStorage.getItem(BONDS_KEY);
    const buysRaw = localStorage.getItem(BUYS_KEY);
    const holdingsRaw = localStorage.getItem(HOLDINGS_KEY);
    const portfolioStartDateRaw = localStorage.getItem(PORTFOLIO_START_DATE_KEY);
    const portfolioStartValueRaw = localStorage.getItem(PORTFOLIO_START_VALUE_KEY);
    const portfolioMonthlyTopupRaw = localStorage.getItem(PORTFOLIO_MONTHLY_TOPUP_KEY);
    const portfolioMonthlyTopupEndDateRaw = localStorage.getItem(PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY);
    const bonds = bondsRaw ? sanitizeBondRows(JSON.parse(bondsRaw)) : defaultBondRows();
    const buys = buysRaw ? sortBuyRowsByDate(normalizeBuyRowsForUI(JSON.parse(buysRaw))).rows : defaultBuyRows();
    const holdings = holdingsRaw ? normalizeHoldingRowsForUI(JSON.parse(holdingsRaw)) : defaultHoldingRows();
    const taxRaw = localStorage.getItem(TAX_RATE_KEY);

    writeRows(bondsTbody, bondTpl, Array.isArray(bonds) && bonds.length ? bonds : defaultBondRows());
    writeRows(buysTbody, buyTpl, Array.isArray(buys) && buys.length ? buys : defaultBuyRows());
    writeRows(holdingsTbody, holdingTpl, Array.isArray(holdings) && holdings.length ? holdings : defaultHoldingRows());
    if (taxRateInput && taxRaw !== null) taxRateInput.value = String(parseNumber(taxRaw) || 13);
    if (portfolioStartDateInput) portfolioStartDateInput.value = normalizeYMD(portfolioStartDateRaw || "") || getTodayYMD();
    setPortfolioMoneyInputValue(portfolioStartValueInput, String(parseNumber(portfolioStartValueRaw) || 0));
    setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, String(parseNumber(portfolioMonthlyTopupRaw) || 0));
    if (portfolioMonthlyTopupEndDateInput) {
      portfolioMonthlyTopupEndDateInput.value = normalizeYMD(portfolioMonthlyTopupEndDateRaw || "") || "";
    }
    loadCalculatorStates();
  } catch {
    writeRows(bondsTbody, bondTpl, defaultBondRows());
    writeRows(buysTbody, buyTpl, defaultBuyRows());
    writeRows(holdingsTbody, holdingTpl, defaultHoldingRows());
    if (taxRateInput) taxRateInput.value = "13";
    if (portfolioStartDateInput) portfolioStartDateInput.value = getTodayYMD();
    if (portfolioMonthlyTopupEndDateInput) portfolioMonthlyTopupEndDateInput.value = "";
    setPortfolioMoneyInputValue(portfolioStartValueInput, "0");
    setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, "0");
    loadCalculatorStates();
  }
  syncStaticTableEmptyStates();
  syncDateSummaries();
  syncBuySummaries();
  renderAll();
}

let saveTimer = null;
function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    persistAndRender();
  }, 250);
}

function bindRowDelete(tbody) {
  tbody.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-action='remove']");
    if (!btn) return;
    const tr = btn.closest("tr");
    if (tr) tr.remove();
    scheduleSave();
  });

  tbody.addEventListener("input", (e) => {
    if (e.target && e.target.matches("[data-field]")) scheduleSave();
  });
}

bindRowDelete(bondsTbody);
bindRowDelete(buysTbody);
bindRowDelete(holdingsTbody);

function copyBuyRow(tr) {
  if (!tr || tr.parentElement !== buysTbody) return;
  if (tr.hasAttribute("data-empty-state")) return;

  const dateInput = tr.querySelector('[data-field="date"]');
  const itemsInput = tr.querySelector('[data-field="items"]');
  if (!dateInput || !itemsInput) return;

  clearTableEmptyState(buysTbody);

  const node = buyTpl.content.cloneNode(true);
  const newTr = node.querySelector("tr");
  const newDateInput = newTr?.querySelector('[data-field="date"]');
  const newItemsInput = newTr?.querySelector('[data-field="items"]');

  if (newDateInput) newDateInput.value = dateInput.value || "";
  if (newItemsInput) newItemsInput.value = itemsInput.value || "[]";

  buysTbody.appendChild(node);
  syncBuySummaries();

}

if (yieldTbody) {
  yieldTbody.addEventListener("click", (e) => {
    const btn = e.target.closest('[data-action="remove"]');
    if (btn) {
      const tr = btn.closest("tr");
      if (tr) tr.remove();
      saveYieldCalcState();
      renderYieldCalculator();
      return;
    }
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement !== yieldTbody) return;
    if (tr.hasAttribute("data-empty-state")) return;
    openYieldModalForEdit(tr);
  });
}

addBondBtn.addEventListener("click", () => {
  clearTableEmptyState(bondsTbody);
  bondsTbody.appendChild(bondTpl.content.cloneNode(true));
  syncDateSummaries();
  scheduleSave();
});

addBuyBtn.addEventListener("click", () => {
  openBuyModal();
});

if (addHoldingBtn) {
  addHoldingBtn.addEventListener("click", () => {
    clearTableEmptyState(holdingsTbody);
    holdingsTbody.appendChild(holdingTpl.content.cloneNode(true));
    scheduleSave();
  });
}
if (addYieldRowBtn) {
  addYieldRowBtn.addEventListener("click", () => {
    openYieldModal();
  });
}

document.querySelectorAll("th[data-sort-table][data-sort-field]").forEach((th) => {
  th.addEventListener("click", () => {
    const tableName = th.getAttribute("data-sort-table") || "";
    const field = th.getAttribute("data-sort-field") || "";
    const state = tableSortState[tableName];
    if (!state || !field) return;
    if (state.field === field) state.dir = state.dir === "asc" ? "desc" : "asc";
    else {
      state.field = field;
      state.dir = "asc";
    }
    updateSortableHeaders();
    sortPlanningTable(tableName);
  });
});

buysTbody.addEventListener("click", (e) => {
  if (e.target.closest('[data-action="remove"]')) return;
  const copyBtn = e.target.closest('[data-action="copy"]');
  if (copyBtn) {
    const tr = copyBtn.closest("tr");
    if (!tr || tr.parentElement !== buysTbody) return;
    copyBuyRow(tr);
    return;
  }
  if (e.target.closest("input[data-field]")) return;
  const tr = e.target.closest("tr");
  if (!tr || tr.parentElement !== buysTbody) return;
  if (tr.hasAttribute("data-empty-state")) return;
  openBuyModalForEdit(tr);
});

resetAllBtn.addEventListener("click", () => {
  localStorage.removeItem(BONDS_KEY);
  localStorage.removeItem(BUYS_KEY);
  localStorage.removeItem(HOLDINGS_KEY);
  localStorage.removeItem(TAX_RATE_KEY);
  localStorage.removeItem(ACTIVE_TAB_KEY);
  localStorage.removeItem(CHART_YEAR_KEY);
  localStorage.removeItem(PORTFOLIO_CHART_YEAR_KEY);
  localStorage.removeItem(PORTFOLIO_START_DATE_KEY);
  localStorage.removeItem(PORTFOLIO_START_VALUE_KEY);
  localStorage.removeItem(PORTFOLIO_MONTHLY_TOPUP_KEY);
  localStorage.removeItem(PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY);
  localStorage.removeItem(ACCRUED_CALC_KEY);
  localStorage.removeItem(YIELD_CALC_KEY);
  if (taxRateInput) taxRateInput.value = "13";
  if (portfolioStartDateInput) portfolioStartDateInput.value = getTodayYMD();
  setPortfolioMoneyInputValue(portfolioStartValueInput, "0");
  setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, "0");
  loadAll();
});

if (taxRateInput) {
  taxRateInput.addEventListener("input", () => {
    persistAndRender();
  });
}

if (chartYearSelect) {
  chartYearSelect.addEventListener("change", () => {
    localStorage.setItem(CHART_YEAR_KEY, chartYearSelect.value || "");
    renderAll();
  });
}

if (portfolioChartYearSelect) {
  portfolioChartYearSelect.addEventListener("change", () => {
    localStorage.setItem(PORTFOLIO_CHART_YEAR_KEY, portfolioChartYearSelect.value || "");
    renderAll();
  });
}

if (portfolioStartDateInput) {
  portfolioStartDateInput.addEventListener("input", () => {
    localStorage.setItem(PORTFOLIO_START_DATE_KEY, normalizeYMD(portfolioStartDateInput.value || "") || "");
    renderAll();
  });
}

if (portfolioMonthlyTopupEndDateInput) {
  portfolioMonthlyTopupEndDateInput.addEventListener("input", () => {
    localStorage.setItem(
      PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY,
      normalizeYMD(portfolioMonthlyTopupEndDateInput.value || "") || ""
    );
    renderAll();
  });
}

bindPortfolioMoneyInput(portfolioStartValueInput, PORTFOLIO_START_VALUE_KEY);
bindPortfolioMoneyInput(portfolioMonthlyTopupInput, PORTFOLIO_MONTHLY_TOPUP_KEY);

if (tabPortfolioBtn) {
  tabPortfolioBtn.addEventListener("click", () => setActiveTab("portfolio"));
}
if (tabPlanningBtn) {
  tabPlanningBtn.addEventListener("click", () => setActiveTab("planning"));
}
if (tabChartsBtn) {
  tabChartsBtn.addEventListener("click", () => setActiveTab("charts"));
}
if (tabCalculatorsBtn) {
  tabCalculatorsBtn.addEventListener("click", () => setActiveTab("calculators"));
}

[
  accruedNominalInput,
  accruedCouponRateInput,
  accruedLastCouponDateInput,
  accruedNextCouponDateInput,
  accruedSettlementDateInput,
].forEach((input) => {
  if (!input) return;
  input.addEventListener("input", () => {
    saveAccruedCalcState();
    renderAccruedIncomeCalculator();
  });
  input.addEventListener("change", () => {
    saveAccruedCalcState();
    renderAccruedIncomeCalculator();
  });
});

function closeSummaryInfoPopovers(exceptBtn = null) {
  document.querySelectorAll(".summaryInfoPopover").forEach((el) => el.remove());
  document.querySelectorAll(".summaryInfoBtn[aria-expanded='true']").forEach((btn) => {
    if (btn !== exceptBtn) btn.setAttribute("aria-expanded", "false");
  });
}

function positionSummaryInfoPopover(pop, stat) {
  if (!(pop instanceof HTMLElement) || !(stat instanceof HTMLElement)) return;
  pop.style.left = "";
  pop.style.right = "8px";
  pop.style.top = "";
  pop.style.bottom = "calc(100% - 8px)";
  pop.style.width = "";
  const margin = 8;
  const rect = pop.getBoundingClientRect();
  const statRect = stat.getBoundingClientRect();
  const viewportWidth = window.innerWidth;

  if (rect.right > viewportWidth - margin) {
    const overflow = rect.right - (viewportWidth - margin);
    const desiredLeft = Math.max(margin - statRect.left, stat.clientWidth - rect.width - 8 - overflow);
    pop.style.left = `${Math.max(8 - statRect.left, desiredLeft)}px`;
    pop.style.right = "auto";
  }

  const nextRect = pop.getBoundingClientRect();
  if (nextRect.left < margin) {
    pop.style.left = `${margin - statRect.left}px`;
    pop.style.right = "auto";
  }

  const finalRect = pop.getBoundingClientRect();
  if (finalRect.top < margin) {
    pop.style.bottom = "auto";
    pop.style.top = "34px";
  }
}

document.addEventListener("click", (e) => {
  const target = e.target;
  if (!(target instanceof Element)) return;
  const btn = target.closest(".summaryInfoBtn");
  if (!btn) {
    closeSummaryInfoPopovers();
    return;
  }
  const stat = btn.closest(".summaryStat");
  if (!stat) return;
  const isOpen = btn.getAttribute("aria-expanded") === "true";
  closeSummaryInfoPopovers(btn);
  if (isOpen) {
    btn.setAttribute("aria-expanded", "false");
    return;
  }
  const text = btn.getAttribute("data-summary-help") || "";
  const pop = document.createElement("div");
  pop.className = "summaryInfoPopover";
  pop.textContent = text;
  stat.appendChild(pop);
  positionSummaryInfoPopover(pop, stat);
  btn.setAttribute("aria-expanded", "true");
});

setActiveTab(localStorage.getItem(ACTIVE_TAB_KEY) || "planning");

bindChartTooltips();
updateSortableHeaders();
loadAll();

function syncDateSummaries() {
  Array.from(bondsTbody.querySelectorAll("tr")).forEach((tr) => {
    const monthsHidden = tr.querySelector('[data-field="payoutMonths"]');
    const startHidden = tr.querySelector('[data-field="startDate"]');
    const endHidden = tr.querySelector('[data-field="endDate"]');
    const monthsBtn = tr.querySelector('[data-action="open-month-picker"]');
    if (!monthsHidden || !startHidden || !endHidden) return;

    const selected = parseMonthList(monthsHidden.value);
    const normalized = Array.from(new Set(selected)).sort((a, b) => a - b);
    monthsHidden.value = normalized.join(",");
    startHidden.value = normalizeYMD(startHidden.value) || "";
    endHidden.value = normalizeYMD(endHidden.value) || "";

    if (monthsBtn) {
      if (normalized.length && startHidden.value && endHidden.value) {
        monthsBtn.textContent = `${formatMonthYearFromYMD(startHidden.value)} - ${formatMonthYearFromYMD(endHidden.value)}`;
      } else {
        monthsBtn.textContent = "Выбрать даты";
      }
    }
  });
}

function syncBuySummaries() {
  Array.from(buysTbody.querySelectorAll("tr")).forEach((tr) => {
    const dateHidden = tr.querySelector('[data-field="date"]');
    const dateDisplay = tr.querySelector('[data-display-for="date"]');
    if (dateHidden && dateDisplay) {
      dateDisplay.textContent = formatDateRuMonthWords(dateHidden.value);
    }
    const itemsHidden = tr.querySelector('[data-field="items"]');
    const itemsDisplay = tr.querySelector('[data-display-for="items"]');
    const totalDisplay = tr.querySelector('[data-display-for="total"]');
    if (!itemsHidden || !itemsDisplay || !totalDisplay) return;

    const items = parseBuyItems(itemsHidden.value);
    if (!items.length) {
      itemsDisplay.textContent = "—";
      totalDisplay.textContent = "0.00 ₽";
      return;
    }

    itemsDisplay.innerHTML = `<div class="buyChips">${items
      .map((i) => {
        const detail = `${formatQuantityRu(i.quantity)} × ${formatMoney(i.price)}`;
        const tip = `${i.bond} — ${detail}`;
        return `<span class="buyChip" title="${escapeHtml(tip)}">
            <span class="buyChip__name">${escapeHtml(i.bond)}</span>
            <span class="buyChip__sep" aria-hidden="true">·</span>
            <span class="buyChip__detail">${detail}</span>
          </span>`;
      })
      .join("")}</div>`;
    const total = items.reduce((sum, i) => sum + i.price * i.quantity, 0);
    totalDisplay.textContent = formatMoney(total);
  });
}

function formatDateRuMonthWords(ymd) {
  const normalized = normalizeYMD(ymd || "");
  if (!normalized) return "—";
  const [y, m, d] = normalized.split("-").map((s) => Number(s));
  // Форматируем через UTC, чтобы не зависеть от локального таймзоны при парсинге.
  const dt = new Date(Date.UTC(y, m - 1, d));
  return new Intl.DateTimeFormat("ru-RU", {
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  }).format(dt);
}

let monthPickerState = null; // { tr, months:number[], startDate:string, endDate:string }

function isAllMonthsSelected(months) {
  return Array.isArray(months) && months.length === 12;
}

function updateMonthModeUI() {
  if (!monthPickerState) return;
  const all = isAllMonthsSelected(monthPickerState.months);
  if (monthModeAllBtn) monthModeAllBtn.classList.toggle("is-active", all);
  if (monthModeCustomBtn) monthModeCustomBtn.classList.toggle("is-active", !all);
}

function setMonthMonths(months) {
  if (!monthPickerState) return;
  monthPickerState.months = Array.from(new Set(months)).filter((m) => Number.isInteger(m) && m >= 1 && m <= 12).sort((a, b) => a - b);
  renderMonthGrid(monthPickerState.months);
  updateMonthModeUI();
}

function openModalOverlay(overlay) {
  if (!overlay) return;
  if (!overlay.hasAttribute("hidden") && overlay.classList.contains("is-open")) return;
  overlay.removeAttribute("hidden");
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      overlay.classList.add("is-open");
    });
  });
}

function closeModalOverlay(overlay) {
  if (!overlay || overlay.hasAttribute("hidden")) return;
  overlay.classList.remove("is-open");
  let finished = false;
  const complete = () => {
    if (finished) return;
    finished = true;
    overlay.setAttribute("hidden", "");
  };
  const onEnd = (e) => {
    if (e.target !== overlay || e.propertyName !== "opacity") return;
    overlay.removeEventListener("transitionend", onEnd);
    complete();
  };
  overlay.addEventListener("transitionend", onEnd);
  window.setTimeout(complete, 320);
}

function setMonthModalOpen(open) {
  if (!monthModalOverlay) return;
  if (open) openModalOverlay(monthModalOverlay);
  else closeModalOverlay(monthModalOverlay);
}

function renderMonthGrid(months) {
  if (!monthGrid) return;
  monthGrid.innerHTML = "";
  for (let i = 1; i <= 12; i += 1) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `monthBtn${months.includes(i) ? " is-active" : ""}`;
    btn.textContent = MONTHS_RU[i - 1];
    btn.addEventListener("click", () => {
      if (!monthPickerState) return;
      const current = monthPickerState.months;
      if (current.includes(i)) {
        monthPickerState.months = current.filter((m) => m !== i);
      } else {
        monthPickerState.months = [...current, i].sort((a, b) => a - b);
      }
      renderMonthGrid(monthPickerState.months);
      updateMonthModeUI();
    });
    monthGrid.appendChild(btn);
  }
}

function applyMonthsToRow() {
  if (!monthPickerState) return;
  const hidden = monthPickerState.tr.querySelector('[data-field="payoutMonths"]');
  const startHidden = monthPickerState.tr.querySelector('[data-field="startDate"]');
  const endHidden = monthPickerState.tr.querySelector('[data-field="endDate"]');
  if (!hidden || !startHidden || !endHidden) return;
  hidden.value = Array.from(new Set(monthPickerState.months)).sort((a, b) => a - b).join(",");
  startHidden.value = normalizeYMD(monthPickerState.startDate) || "";
  endHidden.value = normalizeYMD(monthPickerState.endDate) || "";
}

bondsTbody.addEventListener("click", (e) => {
  const btn = e.target.closest('[data-action="open-month-picker"]');
  if (!btn) return;
  const tr = btn.closest("tr");
  if (!tr) return;
  const hidden = tr.querySelector('[data-field="payoutMonths"]');
  const startHidden = tr.querySelector('[data-field="startDate"]');
  const endHidden = tr.querySelector('[data-field="endDate"]');
  if (!hidden || !startHidden || !endHidden) return;

  monthPickerState = {
    tr,
    months: parseMonthList(hidden.value),
    startDate: startHidden.value || "",
    endDate: endHidden.value || "",
  };
  if (monthStartDateInput) monthStartDateInput.value = monthPickerState.startDate;
  if (monthEndDateInput) monthEndDateInput.value = monthPickerState.endDate;
  renderMonthGrid(monthPickerState.months);
  updateMonthModeUI();
  setMonthModalOpen(true);
});

if (monthClearBtn) {
  monthClearBtn.addEventListener("click", () => {
    if (!monthPickerState) return;
    setMonthMonths([]);
  });
}

if (monthModeAllBtn) {
  monthModeAllBtn.addEventListener("click", () => {
    if (!monthPickerState) return;
    setMonthMonths([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
  });
}

if (monthModeCustomBtn) {
  monthModeCustomBtn.addEventListener("click", () => {
    if (!monthPickerState) return;
    if (isAllMonthsSelected(monthPickerState.months)) {
      setMonthMonths([]);
    } else {
      updateMonthModeUI();
    }
  });
}

if (monthDoneBtn) {
  monthDoneBtn.addEventListener("click", () => {
    if (monthPickerState) {
      monthPickerState.startDate = monthStartDateInput?.value || "";
      monthPickerState.endDate = monthEndDateInput?.value || "";
      applyMonthsToRow();
      syncDateSummaries();
      persistAndRender();
    }
    monthPickerState = null;
    setMonthModalOpen(false);
  });
}

if (monthModalClose) {
  monthModalClose.addEventListener("click", () => {
    monthPickerState = null;
    setMonthModalOpen(false);
  });
}

if (monthModalOverlay) {
  monthModalOverlay.addEventListener("click", (e) => {
    if (e.target === monthModalOverlay) {
      monthPickerState = null;
      setMonthModalOpen(false);
    }
  });
}

/** @type {HTMLTableRowElement | null} */
let yieldModalEditingTr = null;

function setYieldModalOpen(open) {
  if (!yieldModalOverlay) return;
  if (open) openModalOverlay(yieldModalOverlay);
  else closeModalOverlay(yieldModalOverlay);
}

function fillYieldModalForm(row = {}) {
  if (yieldModalBondNameInput) yieldModalBondNameInput.value = String(row.bondName || "").trim();
  if (yieldModalPurchaseDateInput) yieldModalPurchaseDateInput.value = normalizeYMD(row.purchaseDate || "") || "";
  if (yieldModalMaturityDateInput) yieldModalMaturityDateInput.value = normalizeYMD(row.maturityDate || "") || "";
  setInputNumericValue(yieldModalPaymentsPerYearInput, row.paymentsPerYear, "4");
  setInputNumericValue(yieldModalPaidCouponsCountInput, row.paidCouponsCount, "0");
  setInputNumericValue(yieldModalNominalInput, row.nominal, "1000");
  setInputNumericValue(yieldModalPurchasePriceInput, row.purchasePrice, "950");
  setInputNumericValue(yieldModalCouponRateInput, row.couponRate, "12");
  setInputNumericValue(yieldModalAccruedIncomeInput, row.accruedIncome, "0");
  setInputNumericValue(yieldModalRedemptionPriceInput, row.redemptionPrice, "1000");
}

function getYieldModalRowData() {
  return {
    bondName: String(yieldModalBondNameInput?.value || "").trim(),
    purchaseDate: normalizeYMD(yieldModalPurchaseDateInput?.value || "") || "",
    maturityDate: normalizeYMD(yieldModalMaturityDateInput?.value || "") || "",
    paymentsPerYear: String(Math.max(1, Math.round(parseNumber(yieldModalPaymentsPerYearInput?.value || "") || 0))),
    paidCouponsCount: String(Math.max(0, Math.round(parseNumber(yieldModalPaidCouponsCountInput?.value || "") || 0))),
    nominal: String(parseNumber(yieldModalNominalInput?.value || "") || 0),
    purchasePrice: String(parseNumber(yieldModalPurchasePriceInput?.value || "") || 0),
    couponRate: String(parseNumber(yieldModalCouponRateInput?.value || "") || 0),
    accruedIncome: String(parseNumber(yieldModalAccruedIncomeInput?.value || "") || 0),
    redemptionPrice: String(parseNumber(yieldModalRedemptionPriceInput?.value || "") || 0),
  };
}

function openYieldModal() {
  yieldModalEditingTr = null;
  if (yieldModalTitle) yieldModalTitle.textContent = YIELD_MODAL_TITLE_NEW;
  fillYieldModalForm(getDefaultYieldCalcRow());
  if (yieldModalPurchaseDateInput) yieldModalPurchaseDateInput.value = getTodayYMD();
  setYieldModalOpen(true);
}

function openYieldModalForEdit(tr) {
  if (!tr || tr.closest("tbody") !== yieldTbody) return;
  yieldModalEditingTr = tr;
  if (yieldModalTitle) yieldModalTitle.textContent = YIELD_MODAL_TITLE_EDIT;
  const row = {};
  tr.querySelectorAll("input[data-field]").forEach((el) => {
    const key = el.getAttribute("data-field");
    if (key) row[key] = el.value;
  });
  fillYieldModalForm(row);
  setYieldModalOpen(true);
}

function closeYieldModal() {
  yieldModalEditingTr = null;
  if (yieldModalTitle) yieldModalTitle.textContent = YIELD_MODAL_TITLE_NEW;
  setYieldModalOpen(false);
}

let buyModalItems = [];
/** @type {HTMLTableRowElement | null} */
let buyModalEditingTr = null;
/** Блокирует обработчики input/change во время перестроения модалки (срыв DOM даёт ложные события) */
let buyModalRenderLock = false;

function syncBuyModalItemsFromDom() {
  if (!buyItemsWrap) return;
  const fields = buyItemsWrap.querySelectorAll("input[data-buy-idx], select[data-buy-idx]");
  let maxIdx = -1;
  fields.forEach((el) => {
    const idx = Number(el.getAttribute("data-buy-idx"));
    if (Number.isInteger(idx)) maxIdx = Math.max(maxIdx, idx);
  });
  if (maxIdx < 0) return;

  const next = Array.from({ length: maxIdx + 1 }, (_, i) => {
    const prev = buyModalItems[i];
    return {
      bond: prev && typeof prev.bond === "string" ? prev.bond : "",
      price: prev && typeof prev.price === "string" ? prev.price : "",
      quantity: prev && typeof prev.quantity === "string" ? prev.quantity : "",
    };
  });

  fields.forEach((el) => {
    const idx = Number(el.getAttribute("data-buy-idx"));
    const field = el.getAttribute("data-buy-field");
    if (!Number.isInteger(idx) || idx < 0 || idx > maxIdx || !field) return;
    if (!(el instanceof HTMLInputElement || el instanceof HTMLSelectElement)) return;
    if (field !== "bond" && field !== "price" && field !== "quantity") return;
    next[idx][field] = el.value;
  });

  buyModalItems.length = 0;
  buyModalItems.push(...next);
}

function buyRowItemsToModalState(tr) {
  const itemsInput = tr.querySelector('[data-field="items"]');
  const parsed = parseBuyItems(itemsInput?.value);
  if (!parsed.length) return [{ bond: "", price: "", quantity: "" }];
  return parsed.map((i) => ({
    bond: i.bond,
    price: formatAmount(i.price),
    quantity: formatQuantityInt(i.quantity),
  }));
}

function getAvailableBondNames() {
  const seen = new Set();
  return readRows(bondsTbody)
    .map((row) => String(row.bond || "").trim().toUpperCase())
    .filter((bond) => {
      if (!bond) return false;
      if (seen.has(bond)) return false;
      seen.add(bond);
      return true;
    })
    .sort((a, b) => a.localeCompare(b, "ru"));
}

function setBuyModalOpen(open) {
  if (!buyModalOverlay) return;
  if (open) openModalOverlay(buyModalOverlay);
  else closeModalOverlay(buyModalOverlay);
}

function renderBuyModalItems() {
  if (!buyItemsWrap) return;
  buyModalRenderLock = true;
  try {
    buyItemsWrap.replaceChildren();
    const availableBonds = getAvailableBondNames();
    const optionsHtml = availableBonds
      .map((bond) => `<option value="${escapeHtml(bond)}">${escapeHtml(bond)}</option>`)
      .join("");

    buyModalItems.forEach((item, idx) => {
      const selectedBond = String(item.bond || "").trim();
      const bondUpper = selectedBond.toUpperCase();
      const hasSelectedInList = bondUpper && availableBonds.includes(bondUpper);
      const fallbackOption = bondUpper && !hasSelectedInList
        ? `<option value="${escapeHtml(bondUpper)}">${escapeHtml(bondUpper)}</option>`
        : "";

      const row = document.createElement("div");
      row.className = "buyItemRow";
      row.innerHTML = `
      <select class="input buyBondSelect" data-buy-field="bond" data-buy-idx="${idx}">
        <option value="">Выберите облигацию</option>
        ${fallbackOption}
        ${optionsHtml}
      </select>
      <label class="floatingField">
        <input class="input input--num floatingField__input" data-buy-field="price" data-buy-idx="${idx}" inputmode="decimal" placeholder=" " value="${item.price || ""}" />
        <span class="floatingField__label">Цена</span>
      </label>
      <label class="floatingField">
        <input class="input input--num floatingField__input" data-buy-field="quantity" data-buy-idx="${idx}" inputmode="numeric" pattern="[0-9]*" placeholder=" " value="${item.quantity ?? ""}" />
        <span class="floatingField__label">Количество</span>
      </label>
      <button type="button" class="iconBtn" data-buy-action="remove" data-buy-idx="${idx}" aria-label="Удалить">✕</button>
    `;
      const bondSelect = row.querySelector('[data-buy-field="bond"]');
      if (bondSelect) {
        const want = String(item.bond || "").trim().toUpperCase();
        const opt = Array.from(bondSelect.options).find((o) => o.value.toUpperCase() === want);
        if (opt) bondSelect.value = opt.value;
        else if (want) bondSelect.value = want;
      }
      buyItemsWrap.appendChild(row);
    });
  } finally {
    buyModalRenderLock = false;
  }
}

function openBuyModal() {
  buyModalEditingTr = null;
  if (buyModalTitle) buyModalTitle.textContent = BUY_MODAL_TITLE_NEW;
  buyModalItems = [{ bond: "", price: "", quantity: "" }];
  if (buyDateInput) buyDateInput.value = "";
  renderBuyModalItems();
  setBuyModalOpen(true);
  if (buyDateInput) {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        buyDateInput.focus();
      });
    });
  }
}

function openBuyModalForEdit(tr) {
  if (!tr || tr.closest("tbody") !== buysTbody) return;
  buyModalEditingTr = tr;
  if (buyModalTitle) buyModalTitle.textContent = BUY_MODAL_TITLE_EDIT;
  const dateInput = tr.querySelector('[data-field="date"]');
  if (buyDateInput) buyDateInput.value = normalizeYMD(dateInput?.value || "") || "";
  buyModalItems = buyRowItemsToModalState(tr);
  renderBuyModalItems();
  setBuyModalOpen(true);
  if (buyDateInput) {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        buyDateInput.focus();
      });
    });
  }
}

function closeBuyModal() {
  buyModalEditingTr = null;
  if (buyModalTitle) buyModalTitle.textContent = BUY_MODAL_TITLE_NEW;
  if (buyItemsWrap) {
    buyModalRenderLock = true;
    try {
      buyItemsWrap.replaceChildren();
    } finally {
      buyModalRenderLock = false;
    }
  }
  setBuyModalOpen(false);
}

if (buyItemsWrap) {
  const onBuyItemFieldChanged = (e) => {
    if (buyModalRenderLock) return;
    const el = e.target;
    if (!(el instanceof HTMLInputElement || el instanceof HTMLSelectElement)) return;
    const idx = Number(el.getAttribute("data-buy-idx"));
    const field = el.getAttribute("data-buy-field");
    if (!Number.isInteger(idx) || idx < 0 || idx >= buyModalItems.length || !field) return;
    buyModalItems[idx][field] = el.value;
  };
  buyItemsWrap.addEventListener("input", onBuyItemFieldChanged);
  buyItemsWrap.addEventListener("change", onBuyItemFieldChanged);

  buyItemsWrap.addEventListener("click", (e) => {
    const btn = e.target.closest('[data-buy-action="remove"]');
    if (!btn) return;
    syncBuyModalItemsFromDom();
    const idx = Number(btn.getAttribute("data-buy-idx"));
    if (!Number.isInteger(idx) || idx < 0 || idx >= buyModalItems.length) return;
    buyModalItems.splice(idx, 1);
    if (!buyModalItems.length) buyModalItems.push({ bond: "", price: "", quantity: "" });
    renderBuyModalItems();
  });
}

if (buyAddItemBtn) {
  buyAddItemBtn.addEventListener("click", () => {
    syncBuyModalItemsFromDom();
    buyModalItems.push({ bond: "", price: "", quantity: "" });
    renderBuyModalItems();
  });
}

if (buySaveBtn) {
  buySaveBtn.addEventListener("click", () => {
    syncBuyModalItemsFromDom();
    const date = normalizeYMD(buyDateInput?.value || "");
    const items = buyModalItems
      .map((i) => {
        const qRaw = parseNumber(i.quantity);
        const qty = Number.isFinite(qRaw) ? Math.round(qRaw) : NaN;
        return {
          bond: String(i.bond || "").trim().toUpperCase(),
          price: parseNumber(i.price),
          quantity: qty,
        };
      })
      .filter((i) => i.bond && Number.isFinite(i.price) && Number.isFinite(i.quantity) && i.quantity > 0);

    if (!date || !items.length) return;

    if (buyModalEditingTr) {
      const tr = buyModalEditingTr;
      buyModalEditingTr = null;
      const dateInput = tr.querySelector('[data-field="date"]');
      const itemsInput = tr.querySelector('[data-field="items"]');
      if (dateInput) dateInput.value = date;
      if (itemsInput) itemsInput.value = JSON.stringify(items);
      syncBuySummaries();
      persistAndRender();
      closeBuyModal();
      return;
    }

    const node = buyTpl.content.cloneNode(true);
    const tr = node.querySelector("tr");
    const dateInput = tr.querySelector('[data-field="date"]');
    const itemsInput = tr.querySelector('[data-field="items"]');
    if (dateInput) dateInput.value = date;
    if (itemsInput) itemsInput.value = JSON.stringify(items);
    clearTableEmptyState(buysTbody);
    buysTbody.appendChild(node);
    syncBuySummaries();
    persistAndRender();
    closeBuyModal();
  });
}

if (buyModalClose) {
  buyModalClose.addEventListener("click", () => closeBuyModal());
}

if (buyModalOverlay) {
  buyModalOverlay.addEventListener("click", (e) => {
    if (e.target === buyModalOverlay) closeBuyModal();
  });
}

if (yieldSaveBtn) {
  yieldSaveBtn.addEventListener("click", () => {
    const row = getYieldModalRowData();
    const purchaseTs = ymdToUTCms(row.purchaseDate);
    const maturityTs = ymdToUTCms(row.maturityDate);
    if (!row.bondName || !Number.isFinite(purchaseTs) || !Number.isFinite(maturityTs) || maturityTs <= purchaseTs) return;

    if (yieldModalEditingTr) {
      const tr = yieldModalEditingTr;
      yieldModalEditingTr = null;
      tr.querySelectorAll("input[data-field]").forEach((el) => {
        const key = el.getAttribute("data-field");
        if (key && row[key] !== undefined) el.value = row[key];
      });
      saveYieldCalcState();
      renderYieldCalculator();
      closeYieldModal();
      return;
    }

    const node = yieldTpl.content.cloneNode(true);
    const tr = node.querySelector("tr");
    tr.querySelectorAll("input[data-field]").forEach((el) => {
      const key = el.getAttribute("data-field");
      if (key && row[key] !== undefined) el.value = row[key];
    });
    clearTableEmptyState(yieldTbody);
    if (yieldTbody) yieldTbody.appendChild(node);
    saveYieldCalcState();
    renderYieldCalculator();
    closeYieldModal();
  });
}

if (yieldModalClose) {
  yieldModalClose.addEventListener("click", () => closeYieldModal());
}

if (yieldModalOverlay) {
  yieldModalOverlay.addEventListener("click", (e) => {
    if (e.target === yieldModalOverlay) closeYieldModal();
  });
}

document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (yieldModalOverlay && !yieldModalOverlay.hidden) {
    closeYieldModal();
    e.preventDefault();
    return;
  }
  if (buyModalOverlay && !buyModalOverlay.hidden) {
    closeBuyModal();
    e.preventDefault();
    return;
  }
  if (monthModalOverlay && !monthModalOverlay.hidden) {
    monthPickerState = null;
    setMonthModalOpen(false);
    e.preventDefault();
  }
});

