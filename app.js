const BONDS_KEY = "invest_planner_bonds_v2";
const BUYS_KEY = "invest_planner_bond_buys_v2";
const HOLDINGS_KEY = "invest_planner_holdings_v1";
const TAX_RATE_KEY = "invest_planner_tax_rate_v1";
const ACTIVE_TAB_KEY = "invest_planner_active_tab_v1";
const CHART_YEAR_KEY = "invest_planner_chart_year_v1";
const SUMMARY_PIE_YEAR_KEY = "invest_planner_summary_pie_year_v1";
const PORTFOLIO_CHART_YEAR_KEY = "invest_planner_portfolio_chart_year_v1";
const PORTFOLIO_CHART_STRATEGY_KEY = "invest_planner_portfolio_chart_strategy_v1";
/** Значение option: весь горизонт от даты старта до последнего месяца с выплатами. */
const PORTFOLIO_CHART_YEAR_ALL = "__all__";
const PORTFOLIO_START_DATE_KEY = "invest_planner_portfolio_start_date_v1";
const PORTFOLIO_START_VALUE_KEY = "invest_planner_portfolio_start_value_v1";
const PORTFOLIO_MONTHLY_TOPUP_KEY = "invest_planner_portfolio_monthly_topup_v1";
const PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY = "invest_planner_portfolio_monthly_topup_end_date_v1";
const BUY_STRATEGIES_KEY = "invest_planner_buy_strategies_v1";
const ACTIVE_BUY_STRATEGY_KEY = "invest_planner_active_buy_strategy_v1";
/** Центр вкладки «Стратегия»: buys | charts (localStorage). */
const STRATEGY_CENTER_VIEW_KEY = "invest_planner_strategy_center_view_v1";
const ACCRUED_CALC_KEY = "invest_planner_accrued_calc_v1";
const YIELD_CALC_KEY = "invest_planner_yield_calc_v1";
/** localStorage: «light» | «dark» | «system» (по умолчанию — как в ОС). */
const THEME_PREF_KEY = "invest_planner_theme_v1";
/** Ожидаемое удорожание чистой цены в автоплане, % в месяц (составная модель). */
const AUTO_PLAN_MONTHLY_PRICE_DRIFT_PCT_KEY = "invest_planner_auto_plan_price_drift_pct_v1";

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
const buyStrategySelect = document.getElementById("buy-strategy-select");
const buyTableYearFilterSelect = document.getElementById("buy-table-year-filter");
const buyStrategyHeaderSelect = document.getElementById("buy-strategy-header-select");
const summaryStrategySelect = document.getElementById("summary-strategy-select");
const autoPlanStrategySelect = document.getElementById("auto-plan-strategy");
const autoPlanStartInput = document.getElementById("auto-plan-start");
const autoPlanEndInput = document.getElementById("auto-plan-end");
const autoPlanTopupAmountInput = document.getElementById("auto-plan-topup-amount");
const autoPlanReinvestCheckbox = document.getElementById("auto-plan-reinvest");
const autoPlanDiversifyCheckbox = document.getElementById("auto-plan-diversify");
const autoPlanPriceDriftPctInput = document.getElementById("auto-plan-price-drift-pct");
const autoPlanSaveStrategyBtn = document.getElementById("auto-plan-save-strategy-btn");
const buyStrategyNewBtn = document.getElementById("buy-strategy-new-btn");
const buyStrategyEditBtn = document.getElementById("buy-strategy-edit-btn");
const buyStrategyDuplicateBtn = document.getElementById("buy-strategy-duplicate-btn");
const buyStrategyDeleteBtn = document.getElementById("buy-strategy-delete-btn");
const resetAllBtn = document.getElementById("reset-all");
const chartContent = document.getElementById("chart-content");
const returnsChartSvg = document.getElementById("returns-chart");
const chartLegend = document.getElementById("chart-legend");
const chartYearSelect = document.getElementById("chart-year");
const portfolioChartYearSelect = document.getElementById("portfolio-chart-year");
const portfolioChartStrategySelect = document.getElementById("portfolio-chart-strategy-select");
const portfolioChartSvg = document.getElementById("portfolio-chart");
const portfolioChartContent = document.getElementById("portfolio-chart-content");
const portfolioStartDateInput = document.getElementById("portfolio-start-date");
const portfolioStartValueInput = document.getElementById("portfolio-start-value");
const portfolioMonthlyTopupInput = document.getElementById("portfolio-monthly-topup");
const portfolioStartTotalEl = document.getElementById("portfolio-start-total");
const portfolioCouponTotalEl = document.getElementById("portfolio-coupon-total");
const portfolioTopupTotalEl = document.getElementById("portfolio-topup-total");
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
const yieldMonthlyLabelEl = document.getElementById("yield-monthly-label");
const yieldMonthlyHelpBtn = document.getElementById("yield-monthly-help");
const totalIncomeLabelEl = document.getElementById("total-income-label");
const totalIncomeHelpBtn = document.getElementById("total-income-help");
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

const bondModalOverlay = document.getElementById("bond-modal-overlay");
const bondModalClose = document.getElementById("bond-modal-close");
const bondModalTitle = document.getElementById("bond-modal-title");
const bondModalNameInput = document.getElementById("bond-modal-name");
const bondModalCouponInput = document.getElementById("bond-modal-coupon");
const bondModalBondPriceInput = document.getElementById("bond-modal-bond-price");
const bondModalOpenMonthPickerBtn = document.getElementById("bond-modal-open-month-picker");
const bondModalPayoutMonthsInput = document.getElementById("bond-modal-payoutMonths");
const bondModalPayoutMonthsSelect = document.getElementById("bond-modal-payoutMonths-select");
const bondModalStartDateVisibleInput = document.getElementById("bond-modal-start-date");
const bondModalEndDateVisibleInput = document.getElementById("bond-modal-end-date");
const bondModalStartDateInput = document.getElementById("bond-modal-startDate");
const bondModalEndDateInput = document.getElementById("bond-modal-endDate");
const bondMonthPickerPanel = document.getElementById("bond-month-picker-panel");
const bondModalSaveBtn = document.getElementById("bond-modal-save");
const bondModalDeleteBtn = document.getElementById("bond-modal-delete");
const strategyActionsOverlay = document.getElementById("strategy-actions-overlay");
const strategyActionsClose = document.getElementById("strategy-actions-close");
const strategyActionsRenameBtn = document.getElementById("strategy-actions-rename");
const strategyActionsDeleteBtn = document.getElementById("strategy-actions-delete");
const strategyActionsTitle = document.getElementById("strategy-actions-title");
const buyModalOverlay = document.getElementById("buy-modal-overlay");
const buyModalClose = document.getElementById("buy-modal-close");
const buyModalTitle = document.getElementById("buy-modal-title");
const buyDateInput = document.getElementById("buy-date");
const buyItemsWrap = document.getElementById("buy-items");
const buyAddItemBtn = document.getElementById("buy-add-item");
const buySaveBtn = document.getElementById("buy-save");
const strategyModalOverlay = document.getElementById("strategy-modal-overlay");
const strategyModalClose = document.getElementById("strategy-modal-close");
const strategyModalTitle = document.getElementById("strategy-modal-title");
const strategyNameInput = document.getElementById("strategy-name-input");
const strategyDeleteBtn = document.getElementById("strategy-delete");
const strategySaveBtn = document.getElementById("strategy-save");
const tabPortfolioBtn = document.getElementById("tab-portfolio");
const tabPlanningBtn = document.getElementById("tab-planning");
const tabChartsBtn = document.getElementById("tab-charts");
const tabCalculatorsBtn = document.getElementById("tab-calculators");
const tabStrategyBtn = document.getElementById("tab-strategy");
const portfolioPanel = document.getElementById("tab-panel-portfolio");
const planningPanel = document.getElementById("tab-panel-planning");
const chartsPanel = document.getElementById("tab-panel-charts");
const calculatorsPanel = document.getElementById("tab-panel-calculators");
const strategyPanel = document.getElementById("tab-panel-strategy");
const strategyTabStrategyList = document.getElementById("strategy-tab-strategy-list");
const strategyTabBuysTbody = document.getElementById("strategy-tab-buys-tbody");
const strategyTabBondsList = document.getElementById("strategy-tab-bonds-list");
const strategyTabNewStrategyBtn = document.getElementById("strategy-tab-new-strategy");
const strategyTabNewBondBtn = document.getElementById("strategy-tab-new-bond");
const strategyTabAddBuyBtn = document.getElementById("strategy-tab-add-buy");
const strategyPaneBuys = document.getElementById("strategy-pane-buys");
const strategyPaneCharts = document.getElementById("strategy-pane-charts");
const strategyPanePortfolio = document.getElementById("strategy-pane-portfolio");
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
const yieldModalCouponPerPaymentInput = document.getElementById("yield-modal-coupon-per-payment");
const yieldModalAccruedIncomeInput = document.getElementById("yield-modal-accrued-income");
const yieldModalRedemptionPriceInput = document.getElementById("yield-modal-redemption-price");
const yieldSaveBtn = document.getElementById("yield-save");
const themeSelect = document.getElementById("theme-select");

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

const DEFAULT_BUY_STRATEGY_ID = "strategy-default";
/** Ограничение числа дат покупок за один прогон (защита от зависания). */
const AUTO_PLAN_MAX_SCHEDULE_DATES = 800;
/** Бонус к скору за «раннюю» фазу купонного периода (после выплаты до следующей): доля периода × вес → до +30%. */
const AUTO_PLAN_COUPON_DISTANCE_WEIGHT = 0.3;
/** До какого числа будущих купонов растёт множитель горизонта (дальше — насыщение). */
const AUTO_PLAN_HORIZON_PAYMENTS_REF = 32;
/**
 * Вес горизонта: скор умножается на (1 + вес × min(1, будущиеКупоны / ref)).
 * Приоритет бумагам с большим числом будущих выплат при равной текущей доходности.
 */
const AUTO_PLAN_HORIZON_WEIGHT = 0.11;
/** @type {{id:string,name:string}[]} */
let buyStrategies = [];
/** @type {Map<string, any[]>} */
let buyRowsByStrategyId = new Map();
let activeBuyStrategyId = DEFAULT_BUY_STRATEGY_ID;
let strategyModalEditingId = null;
/** Источник для «Дублировать»: id стратегии, строки которой копируются. */
let strategyModalDuplicatingFromId = null;
let autoPlanBondPickerSig = "";
/** Кэш данных сводки для переключения года без повторного расчёта из DOM. */
let lastSummaryPieChartData = null;

function getThemePreference() {
  try {
    const v = localStorage.getItem(THEME_PREF_KEY);
    if (v === "light" || v === "dark" || v === "system") return v;
  } catch {
    /* ignore */
  }
  return "system";
}

function systemPrefersDark() {
  return Boolean(window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
}

/** @returns {"light"|"dark"} */
function effectiveThemeFromPreference(pref) {
  if (pref === "dark") return "dark";
  if (pref === "light") return "light";
  return systemPrefersDark() ? "dark" : "light";
}

function applyDocumentTheme(effective) {
  document.documentElement.setAttribute("data-theme", effective);
  document.documentElement.style.colorScheme = effective === "dark" ? "dark" : "light";
}

function initThemeUi() {
  const pref = getThemePreference();
  applyDocumentTheme(effectiveThemeFromPreference(pref));
  if (themeSelect) {
    themeSelect.value = pref;
    themeSelect.addEventListener("change", () => {
      const next = String(themeSelect.value || "").trim();
      if (next !== "light" && next !== "dark" && next !== "system") return;
      try {
        localStorage.setItem(THEME_PREF_KEY, next);
      } catch {
        /* ignore */
      }
      applyDocumentTheme(effectiveThemeFromPreference(next));
    });
  }
  if (window.matchMedia) {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    const onSchemeChange = () => {
      if (getThemePreference() !== "system") return;
      applyDocumentTheme(effectiveThemeFromPreference("system"));
    };
    mq.addEventListener("change", onSchemeChange);
  }
}

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

function refreshBuyYearFilterUi() {
  if (!buyTableYearFilterSelect || !buysTbody) return;
  if (buysTbody.querySelector('tr[data-empty-state="1"]')) {
    buyTableYearFilterSelect.innerHTML = `<option value="">${escapeHtml("Все годы")}</option>`;
    buyTableYearFilterSelect.value = "";
    return;
  }

  const rows = Array.from(buysTbody.querySelectorAll("tr.buyTable__dataRow"));
  const yearsSet = new Set();
  for (const tr of rows) {
    const ymd = normalizeYMD(tr.querySelector('[data-field="date"]')?.value || "");
    if (ymd.length >= 4) yearsSet.add(ymd.slice(0, 4));
  }
  const sorted = Array.from(yearsSet).sort((a, b) => a.localeCompare(b));
  const prev = String(buyTableYearFilterSelect.value || "").trim();

  buyTableYearFilterSelect.innerHTML =
    `<option value="">${escapeHtml("Все годы")}</option>` +
    sorted.map((yy) => `<option value="${escapeHtml(yy)}">${escapeHtml(yy)}</option>`).join("");

  if (prev && sorted.includes(prev)) buyTableYearFilterSelect.value = prev;
  else buyTableYearFilterSelect.value = "";

  applyBuyYearFilter();
}

function applyBuyYearFilter() {
  if (!buysTbody) return;
  const y = buyTableYearFilterSelect ? String(buyTableYearFilterSelect.value || "").trim() : "";
  Array.from(buysTbody.querySelectorAll("tr.buyTable__dataRow")).forEach((tr) => {
    if (!y) {
      tr.hidden = false;
      return;
    }
    const ymd = normalizeYMD(tr.querySelector('[data-field="date"]')?.value || "");
    const rowY = ymd.length >= 4 ? ymd.slice(0, 4) : "";
    tr.hidden = rowY !== y;
  });
}

function ensureChartsWidgetsInChartsTab() {
  if (!strategyPaneCharts || !chartsPanel) return;
  while (strategyPaneCharts.firstChild) {
    chartsPanel.appendChild(strategyPaneCharts.firstChild);
  }
}

function ensureChartsWidgetsInStrategyPane() {
  if (!strategyPaneCharts || !chartsPanel) return;
  const returnsWidget = chartsPanel.querySelector('[data-widget="returns"]');
  if (returnsWidget && returnsWidget.parentNode === chartsPanel) {
    strategyPaneCharts.appendChild(returnsWidget);
  }
}

function ensurePortfolioWidgetsInPortfolioTab() {
  if (!strategyPanePortfolio || !portfolioPanel) return;
  while (strategyPanePortfolio.firstChild) {
    portfolioPanel.appendChild(strategyPanePortfolio.firstChild);
  }
}

function ensurePortfolioWidgetsInStrategyPane() {
  if (!strategyPanePortfolio || !portfolioPanel) return;
  while (portfolioPanel.firstChild) {
    strategyPanePortfolio.appendChild(portfolioPanel.firstChild);
  }
}

function updateStrategyDisplayNavUi(mode) {
  const list = document.getElementById("strategy-tab-display-list");
  if (!list) return;
  list.querySelectorAll("[data-strategy-display]").forEach((b) => {
    const m = b.getAttribute("data-strategy-display");
    const active = m === mode;
    b.classList.toggle("is-active", active);
    b.closest(".strategyTabShell__displayNavRow")?.classList.toggle("is-active", active);
  });
}

function setStrategyCenterView(mode) {
  if (mode !== "buys" && mode !== "charts" && mode !== "portfolio") return;
  localStorage.setItem(STRATEGY_CENTER_VIEW_KEY, mode);
  if (!strategyPaneBuys || !strategyPaneCharts || !strategyPanePortfolio) return;

  if (mode === "charts") {
    ensurePortfolioWidgetsInPortfolioTab();
    ensureChartsWidgetsInStrategyPane();
    strategyPaneBuys.hidden = true;
    strategyPaneCharts.hidden = false;
    strategyPanePortfolio.hidden = true;
  } else if (mode === "portfolio") {
    ensureChartsWidgetsInChartsTab();
    ensurePortfolioWidgetsInStrategyPane();
    strategyPaneBuys.hidden = true;
    strategyPaneCharts.hidden = true;
    strategyPanePortfolio.hidden = false;
  } else {
    ensureChartsWidgetsInChartsTab();
    ensurePortfolioWidgetsInPortfolioTab();
    strategyPaneBuys.hidden = false;
    strategyPaneCharts.hidden = true;
    strategyPanePortfolio.hidden = true;
  }
  updateStrategyDisplayNavUi(mode);
  requestAnimationFrame(() => {
    try {
      renderAll();
    } catch {
      /* ignore */
    }
  });
}

function applyStrategyCenterViewFromStorage() {
  const raw = localStorage.getItem(STRATEGY_CENTER_VIEW_KEY);
  const mode = raw === "charts" || raw === "portfolio" ? raw : "buys";
  setStrategyCenterView(mode);
}

function setActiveTab(tab) {
  const nextTab = ["portfolio", "planning", "charts", "calculators", "strategy"].includes(tab) ? tab : "planning";
  if (nextTab !== "strategy") {
    ensureChartsWidgetsInChartsTab();
    ensurePortfolioWidgetsInPortfolioTab();
  }
  if (portfolioPanel) portfolioPanel.hidden = nextTab !== "portfolio";
  if (planningPanel) planningPanel.hidden = nextTab !== "planning";
  if (chartsPanel) chartsPanel.hidden = nextTab !== "charts";
  if (calculatorsPanel) calculatorsPanel.hidden = nextTab !== "calculators";
  if (strategyPanel) strategyPanel.hidden = nextTab !== "strategy";
  if (tabPortfolioBtn) tabPortfolioBtn.classList.toggle("is-active", nextTab === "portfolio");
  if (tabPlanningBtn) tabPlanningBtn.classList.toggle("is-active", nextTab === "planning");
  if (tabChartsBtn) tabChartsBtn.classList.toggle("is-active", nextTab === "charts");
  if (tabCalculatorsBtn) tabCalculatorsBtn.classList.toggle("is-active", nextTab === "calculators");
  if (tabStrategyBtn) tabStrategyBtn.classList.toggle("is-active", nextTab === "strategy");
  localStorage.setItem(ACTIVE_TAB_KEY, nextTab);
  if (nextTab === "strategy") {
    applyStrategyCenterViewFromStorage();
  }
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
        const nextValue = String(row[key] ?? "");
        if (el instanceof HTMLSelectElement && !el.multiple) {
          // Если опция отсутствует, создаём её, чтобы значение не сбрасывалось в пустую строку.
          if (nextValue) {
            const hasOption = Array.from(el.options).some((o) => o.value === nextValue);
            if (!hasOption) {
              const opt = document.createElement("option");
              opt.value = nextValue;
              opt.textContent = nextValue;
              el.appendChild(opt);
            }
          }
        }
        el.value = nextValue;
      }
    });
    tbody.appendChild(node);
  });
  if (tbody === buysTbody) {
    refreshBuyYearFilterUi();
  }
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
      5,
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
      10,
      "Сравнение доходности пока пусто",
      "Добавьте облигацию: цена покупки, НКД, купон в ₽ за выплату, выплаты в год и даты."
    );
  }
}

function buildSvgEmptyState(title, hint) {
  return `<text x="450" y="160" text-anchor="middle" fill="currentColor" opacity="0.78" font-size="20" font-weight="600">${escapeHtml(
    title
  )}</text><text x="450" y="190" text-anchor="middle" fill="currentColor" opacity="0.6" font-size="15">${escapeHtml(hint)}</text>`;
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
    case "bondPrice":
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

function getDefaultBuyStrategy() {
  return { id: DEFAULT_BUY_STRATEGY_ID, name: "Основная" };
}

function normalizeStrategyName(name) {
  return String(name || "").trim();
}

function createBuyStrategyId() {
  return `strategy-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
}

function cloneStrategyBuyRows(rows) {
  if (!Array.isArray(rows)) return [];
  try {
    return JSON.parse(JSON.stringify(rows));
  } catch {
    return [];
  }
}

function getActiveStrategyBuysRows() {
  const rows = buyRowsByStrategyId.get(activeBuyStrategyId);
  return Array.isArray(rows) ? rows : [];
}

function readSanitizedBuysFromTable() {
  return sortBuyRowsByDate(readRows(buysTbody).filter(isBuyRowComplete)).rows;
}

function saveCurrentBuysToActiveStrategy() {
  buyRowsByStrategyId.set(activeBuyStrategyId, readSanitizedBuysFromTable());
}

/** Строки покупок стратегии для симуляции купонов: с экрана, если стратегия активна, иначе из кэша. */
function getPlannedBuyRowsForAutoPlan(strategyId) {
  const sid = String(strategyId || "").trim();
  if (!sid) return [];
  if (sid === activeBuyStrategyId) {
    return readSanitizedBuysFromTable();
  }
  return (buyRowsByStrategyId.get(sid) || []).filter(isBuyRowComplete);
}

function buildBuyEventsFromPlannedRows(rows) {
  const events = [];
  for (const row of rows || []) {
    if (!isBuyRowComplete(row)) continue;
    const dateYmd = normalizeYMD(row.date);
    if (!dateYmd) continue;
    const ts = ymdToUTCms(dateYmd) || 0;
    const groupedItems = parseBuyItems(row.items);
    if (groupedItems.length) {
      for (const it of groupedItems) {
        const q = Math.round(it.quantity);
        const b = normalizeBondKey(it.bond);
        if (!b || !Number.isFinite(q) || q <= 0) continue;
        events.push({ dateYMD: dateYmd, dateUTCms: ts, bond: b, quantity: q });
      }
      continue;
    }
    const legacyBond = String(row.bond || "").trim();
    const legacyQty = Math.round(parseNumber(row.quantity));
    if (!legacyBond || !Number.isFinite(legacyQty) || legacyQty <= 0) continue;
    events.push({
      dateYMD: dateYmd,
      dateUTCms: ts,
      bond: normalizeBondKey(legacyBond),
      quantity: legacyQty,
    });
  }
  return events;
}

function persistBuyStrategiesState() {
  const strategiesPayload = Array.isArray(buyStrategies) && buyStrategies.length
    ? buyStrategies.map((s) => ({ id: String(s.id || ""), name: normalizeStrategyName(s.name || "") })).filter((s) => s.id && s.name)
    : [getDefaultBuyStrategy()];

  const buysByStrategy = {};
  strategiesPayload.forEach((s) => {
    const rows = buyRowsByStrategyId.get(s.id);
    buysByStrategy[s.id] = Array.isArray(rows) ? rows : [];
  });

  localStorage.setItem(
    BUY_STRATEGIES_KEY,
    JSON.stringify({
      strategies: strategiesPayload,
      buysByStrategy,
    })
  );
  localStorage.setItem(ACTIVE_BUY_STRATEGY_KEY, activeBuyStrategyId);
  // Совместимость со старой схемой: дублируем активную стратегию в старый ключ покупок.
  localStorage.setItem(BUYS_KEY, JSON.stringify(getActiveStrategyBuysRows()));
}

function renderBuyStrategySelect() {
  const optionsHtml = buyStrategies
    .map((s) => `<option value="${escapeHtml(String(s.id))}">${escapeHtml(String(s.name))}</option>`)
    .join("");
  const hasActive = buyStrategies.some((s) => s.id === activeBuyStrategyId);
  if (!hasActive && buyStrategies.length) activeBuyStrategyId = buyStrategies[0].id;

  if (buyStrategySelect) {
    buyStrategySelect.innerHTML = optionsHtml;
    buyStrategySelect.value = activeBuyStrategyId;
  }
  if (buyStrategyHeaderSelect) {
    buyStrategyHeaderSelect.innerHTML = optionsHtml;
    buyStrategyHeaderSelect.value = activeBuyStrategyId;
  }
  if (summaryStrategySelect) {
    summaryStrategySelect.innerHTML = optionsHtml;
    summaryStrategySelect.value = activeBuyStrategyId;
  }
  if (autoPlanStrategySelect) {
    const prev = String(autoPlanStrategySelect.value || "").trim();
    const head = `<option value="">${escapeHtml("Выберите стратегию")}</option>`;
    autoPlanStrategySelect.innerHTML = head + optionsHtml;
    const validPrev = buyStrategies.some((s) => s.id === prev);
    autoPlanStrategySelect.value = validPrev ? prev : "";
  }
  if (buyStrategyEditBtn) {
    const isDefault = activeBuyStrategyId === DEFAULT_BUY_STRATEGY_ID;
    buyStrategyEditBtn.hidden = isDefault;
  }
  if (buyStrategyDeleteBtn) {
    buyStrategyDeleteBtn.hidden = activeBuyStrategyId === DEFAULT_BUY_STRATEGY_ID;
  }
  if (portfolioChartStrategySelect) {
    const prevPortfolio = String(
      portfolioChartStrategySelect.value || localStorage.getItem(PORTFOLIO_CHART_STRATEGY_KEY) || ""
    ).trim();
    portfolioChartStrategySelect.innerHTML = optionsHtml;
    const validPrev = buyStrategies.some((s) => s.id === prevPortfolio);
    const nextVal = validPrev ? prevPortfolio : activeBuyStrategyId;
    portfolioChartStrategySelect.value = nextVal;
    if (!portfolioChartStrategySelect.value && buyStrategies.length) {
      portfolioChartStrategySelect.value = activeBuyStrategyId;
    }
    localStorage.setItem(PORTFOLIO_CHART_STRATEGY_KEY, portfolioChartStrategySelect.value || "");
    if (isStrategyTabPortfolioView()) {
      portfolioChartStrategySelect.value = activeBuyStrategyId;
      localStorage.setItem(PORTFOLIO_CHART_STRATEGY_KEY, activeBuyStrategyId);
    }
  }
}

function isStrategyTabPortfolioView() {
  return Boolean(
    strategyPanel && !strategyPanel.hidden && strategyPanePortfolio && !strategyPanePortfolio.hidden
  );
}

/** Покупки выбранной в виджете «Стоимость портфеля» стратегии (для графика и итогов виджета). */
function getBuysForPortfolioChart() {
  let id = "";
  if (isStrategyTabPortfolioView()) {
    id = activeBuyStrategyId;
  } else {
    id = String(portfolioChartStrategySelect?.value || "").trim();
    if (!buyStrategies.some((s) => s.id === id)) {
      id = String(localStorage.getItem(PORTFOLIO_CHART_STRATEGY_KEY) || "").trim();
    }
  }
  if (!buyStrategies.some((s) => s.id === id)) id = activeBuyStrategyId;
  const raw = (buyRowsByStrategyId.get(id) || []).filter(isBuyRowComplete);
  return sortBuyRowsByDate(raw).rows;
}

function switchActiveBuyStrategy(nextIdRaw) {
  const nextId = String(nextIdRaw || "").trim();
  if (!nextId || nextId === activeBuyStrategyId) return;
  saveCurrentBuysToActiveStrategy();
  const hasStrategy = buyStrategies.some((s) => s.id === nextId);
  if (!hasStrategy) return;
  activeBuyStrategyId = nextId;
  const rows = getActiveStrategyBuysRows();
  writeRows(buysTbody, buyTpl, Array.isArray(rows) && rows.length ? rows : defaultBuyRows());
  syncBuySummaries();
  syncStaticTableEmptyStates();
  renderBuyStrategySelect();
  persistBuyStrategiesState();
  renderAll();
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

/** Сдвиг календарной даты YMD на целое число суток (UTC). */
function addDaysToYmd(ymdRaw, deltaDays) {
  const base = normalizeYMD(ymdRaw) || String(ymdRaw || "").trim();
  const ts = ymdToUTCms(base);
  const dDays = Number(deltaDays);
  if (ts == null || !Number.isFinite(ts) || !Number.isFinite(dDays)) return "";
  const d = new Date(ts + dDays * 86400000);
  const y = d.getUTCFullYear();
  const mo = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  return `${String(y).padStart(4, "0")}-${String(mo).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
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
  const allOptionHtml = `<option value="${PORTFOLIO_CHART_YEAR_ALL}">За все время</option>`;

  if (!years.length) {
    chartYearSelect.innerHTML = allOptionHtml;
    chartYearSelect.value = PORTFOLIO_CHART_YEAR_ALL;
    localStorage.setItem(CHART_YEAR_KEY, PORTFOLIO_CHART_YEAR_ALL);
    return PORTFOLIO_CHART_YEAR_ALL;
  }

  const saved = localStorage.getItem(CHART_YEAR_KEY);
  const current = chartYearSelect.value;

  let selected = years[0];
  if (saved === PORTFOLIO_CHART_YEAR_ALL) selected = PORTFOLIO_CHART_YEAR_ALL;
  else if (saved && years.includes(saved)) selected = saved;

  if (current === PORTFOLIO_CHART_YEAR_ALL) selected = PORTFOLIO_CHART_YEAR_ALL;
  else if (current && years.includes(current)) selected = current;

  chartYearSelect.innerHTML =
    allOptionHtml + years.map((y) => `<option value="${escapeHtml(y)}">${escapeHtml(y)}</option>`).join("");
  chartYearSelect.value = selected;
  if (!chartYearSelect.value) {
    chartYearSelect.value = selected === PORTFOLIO_CHART_YEAR_ALL ? PORTFOLIO_CHART_YEAR_ALL : years[0];
  }
  const persisted = chartYearSelect.value;
  localStorage.setItem(CHART_YEAR_KEY, persisted);
  return persisted;
}

function ensurePortfolioChartYearOptions(points) {
  if (!portfolioChartYearSelect) return null;
  const years = Array.from(new Set((points || []).map((p) => getYearFromMonthKey(p.monthKey)).filter(Boolean))).sort();
  const allOptionHtml = `<option value="${PORTFOLIO_CHART_YEAR_ALL}">За все время</option>`;

  if (!years.length) {
    portfolioChartYearSelect.innerHTML = allOptionHtml;
    portfolioChartYearSelect.value = PORTFOLIO_CHART_YEAR_ALL;
    localStorage.setItem(PORTFOLIO_CHART_YEAR_KEY, PORTFOLIO_CHART_YEAR_ALL);
    return PORTFOLIO_CHART_YEAR_ALL;
  }

  const saved = localStorage.getItem(PORTFOLIO_CHART_YEAR_KEY);
  const current = portfolioChartYearSelect.value;

  let selected = years[0];
  if (saved === PORTFOLIO_CHART_YEAR_ALL) selected = PORTFOLIO_CHART_YEAR_ALL;
  else if (saved && years.includes(saved)) selected = saved;

  if (current === PORTFOLIO_CHART_YEAR_ALL) selected = PORTFOLIO_CHART_YEAR_ALL;
  else if (current && years.includes(current)) selected = current;

  portfolioChartYearSelect.innerHTML =
    allOptionHtml + years.map((y) => `<option value="${escapeHtml(y)}">${escapeHtml(y)}</option>`).join("");
  portfolioChartYearSelect.value = selected;
  if (!portfolioChartYearSelect.value) {
    portfolioChartYearSelect.value = selected === PORTFOLIO_CHART_YEAR_ALL ? PORTFOLIO_CHART_YEAR_ALL : years[0];
  }
  const persisted = portfolioChartYearSelect.value;
  localStorage.setItem(PORTFOLIO_CHART_YEAR_KEY, persisted);
  return persisted;
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
  lastSummaryPieChartData = chartData;

  const seriesByBond = chartData?.seriesByBond || [];
  const palette = getSeriesPalette();

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

  const selectedYear = resolveSummaryPieYearFromYearList(years);

  const segButtons = years
    .map((y) => {
      const active = y === selectedYear;
      return `<button type="button" class="yearPieYearSeg__btn${active ? " is-active" : ""}" role="tab" aria-selected="${active ? "true" : "false"}" data-year="${escapeHtml(y)}">${escapeHtml(y)}</button>`;
    })
    .join("");

  const bondMap = byYear.get(selectedYear);
  const entries = Array.from(bondMap.entries())
    .filter(([, value]) => value > 0)
    .sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, value]) => sum + value, 0);

  let chartBlock = "";
  if (total <= 0) {
    chartBlock = `<div class="yearPiePanel__empty">Нет выплат за ${escapeHtml(selectedYear)}</div>`;
  } else {
    let currentAngle = 0;
    const slices =
      entries.length === 1
        ? `<circle class="yearPieSlice" cx="50" cy="50" r="44" fill="${palette[0]}" data-slice-bond="${escapeHtml(
            entries[0][0]
          )}" data-slice-year="${escapeHtml(String(selectedYear))}" data-slice-amount="${entries[0][1]}" data-slice-pct="100"></circle>`
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
                )}" data-slice-year="${escapeHtml(String(selectedYear))}" data-slice-amount="${s.value}" data-slice-pct="${s.pct}"></path>`
            )
            .join("");

    chartBlock = `<svg class="yearPie" viewBox="0 0 100 100" aria-label="Выплаты net за ${escapeHtml(selectedYear)}">${slices}</svg>`;
  }

  const legend =
    total > 0
      ? entries
          .map(([bond, value], idx) => {
            const color = palette[idx % palette.length];
            return `<div class="yearPieLegend__item">
          <span class="yearPieLegend__dot" style="background:${color}"></span>
          <span class="yearPieLegend__bond">${escapeHtml(bond)}</span>
          <span class="yearPieLegend__sum">${formatMoney(value)}</span>
        </div>`;
          })
          .join("")
      : "";

  yearPiesEl.innerHTML = `<div class="yearPiePanel">
    <div class="yearPiePanel__chart">
      <div class="yearPiePanel__head">
        <span class="yearPiePanel__headYear">${escapeHtml(selectedYear)}</span>
        <span class="yearPiePanel__headSum">${formatMoney(total)}</span>
      </div>
      ${chartBlock}
      ${legend ? `<div class="yearPieLegend">${legend}</div>` : ""}
    </div>
    <div class="yearPieYearSeg" role="tablist" aria-label="Год выплат">${segButtons}</div>
  </div>`;
}

function onYearPieYearSegmentClick(e) {
  const t = e.target;
  if (!(t instanceof Element)) return;
  const btn = t.closest(".yearPieYearSeg__btn");
  if (!(btn instanceof HTMLButtonElement)) return;
  const y = String(btn.getAttribute("data-year") || "").trim();
  if (!y) return;
  localStorage.setItem(SUMMARY_PIE_YEAR_KEY, y);
  renderAll();
}

function renderChartLegend(seriesByBond, palette) {
  if (!chartLegend) return;
  // Тот же sIdx, что у столбцов в renderChart (palette[sIdx]), а не порядковый номер среди «видимых» серий.
  const chunks = seriesByBond
    .map((series, sIdx) => {
      const visible = (series.points || []).some((point) => Number(point.amount) > 0);
      if (!visible) return "";
      const color = palette[sIdx % palette.length];
      const safeBond = String(series.bond || `Bond ${sIdx + 1}`)
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
    .filter((html) => html.length > 0);
  if (!chunks.length) {
    chartLegend.innerHTML = "";
    return;
  }
  chartLegend.innerHTML = chunks.join("");
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
    couponPerPayment: 30,
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
  writeRows(yieldTbody, yieldTpl, normalizeYieldCalcRowsForStorage(normalizedRows));
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

/**
 * Купон за одну выплату в ₽. Новые строки: поле couponPerPayment.
 * Старые данные: годовая ставка % в couponRate → пересчёт через номинал и выплат в год.
 */
function resolveYieldCouponPerPaymentRub(row) {
  const cppDirect = parseNumber(row.couponPerPayment);
  if (Number.isFinite(cppDirect) && cppDirect >= 0) return cppDirect;

  const nominal = parseNumber(row.nominal);
  const couponRatePct = parseNumber(row.couponRate);
  const ppyRaw = parseNumber(row.paymentsPerYear);
  const ppy = Number.isFinite(ppyRaw) && ppyRaw > 0 ? Math.max(1, Math.round(ppyRaw)) : NaN;
  if (Number.isFinite(nominal) && Number.isFinite(couponRatePct) && Number.isFinite(ppy)) {
    return (nominal * (couponRatePct / 100)) / ppy;
  }
  return NaN;
}

/** Строки калькулятора доходности из localStorage: перенос couponRate → couponPerPayment для нового шаблона. */
function normalizeYieldCalcRowsForStorage(rows) {
  if (!Array.isArray(rows)) return [];
  return rows.map((row) => {
    const out = { ...row };
    const cpp = resolveYieldCouponPerPaymentRub(out);
    if (Number.isFinite(cpp) && cpp >= 0) {
      out.couponPerPayment = String(roundRub2(cpp));
    } else if (out.couponPerPayment === undefined || out.couponPerPayment === null) {
      out.couponPerPayment = "";
    }
    delete out.couponRate;
    return out;
  });
}

function computeYieldMetrics(row) {
  const purchaseDate = normalizeYMD(row.purchaseDate);
  const maturityDate = normalizeYMD(row.maturityDate);
  const paymentsPerYearRaw = parseNumber(row.paymentsPerYear);
  const paidCouponsCountRaw = parseNumber(row.paidCouponsCount);
  const purchasePrice = parseNumber(row.purchasePrice);
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
  const couponPerPayment = resolveYieldCouponPerPaymentRub(row);
  const annualCouponIncome = Number.isFinite(couponPerPayment) && Number.isFinite(paymentsPerYear)
    ? couponPerPayment * paymentsPerYear
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
      ? "Грязная цена = цена покупки + НКД. Купоны к получению = купон за выплату (₽) × оставшиеся выплаты. Текущая доходность = (купон × выплат в год) / грязная цена. Доходность к погашению ≈ (купоны + погашение − грязная цена) / грязная цена / срок в годах."
      : "Добавьте облигацию: цена покупки, НКД, номинал, купон в ₽ за одну выплату, выплат в год, уже выплаченных купонов, даты покупки и погашения, цена погашения.";
  }
}

function renderCalculators() {
  renderAccruedIncomeCalculator();
  renderYieldCalculator();
  syncStaticTableEmptyStates();
}

/**
 * Мягкая анимация изменения layout (в т.ч. высоты) при перерисовке виджетов.
 * Реализация в стиле FLIP: измеряем до/после и делаем короткую инверсию transform.
 * @returns {Map<HTMLElement, {top:number,height:number}>}
 */
function captureVisibleWidgetLayout() {
  const widgets = Array.from(document.querySelectorAll('.tabPanel:not([hidden]) .widget')).filter((el) => el instanceof HTMLElement);
  const snap = new Map();
  widgets.forEach((el) => {
    const r = el.getBoundingClientRect();
    snap.set(el, { top: r.top, height: r.height });
  });
  return snap;
}

function animateVisibleWidgetLayout(snap) {
  if (!snap || !(snap instanceof Map) || snap.size === 0) return;
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      snap.forEach((old, el) => {
        if (!el || !el.isConnected) return;
        const r = el.getBoundingClientRect();
        const newTop = r.top;
        const newHeight = r.height;
        if (!Number.isFinite(old.top) || !Number.isFinite(old.height) || newHeight <= 0) return;

        const deltaY = old.top - newTop;
        const scaleY = old.height / newHeight;

        // Слишком мелкие изменения не анимируем.
        if (Math.abs(deltaY) < 1 && Math.abs(scaleY - 1) < 0.01) return;

        // Мягкое "перелистывание" без изменения смысла.
        el.animate(
          [
            { transformOrigin: 'top', transform: `translateY(${deltaY}px) scaleY(${scaleY})`, opacity: 0.985 },
            { transformOrigin: 'top', transform: 'translateY(0px) scaleY(1)', opacity: 1 },
          ],
          { duration: 220, easing: 'cubic-bezier(0.2, 0.8, 0.2, 1)', fill: 'forwards' }
        );
      });
    });
  });
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

/** Календарные годы, в которых есть выплаты по `chartData`. */
function getPayoutYearsFromChartData(chartData) {
  const seriesByBond = chartData?.seriesByBond || [];
  const ys = new Set();
  seriesByBond.forEach((series) => {
    (series.points || []).forEach((p) => {
      const y = getYearFromMonthKey(p.monthKey);
      if (y) ys.add(y);
    });
  });
  return Array.from(ys).sort();
}

/** Год, выбранный переключателем под круговой диаграммой в сводке (`SUMMARY_PIE_YEAR_KEY`). */
function resolveSummaryPieYearFromYearList(years) {
  if (!Array.isArray(years) || !years.length) return "";
  const saved = String(localStorage.getItem(SUMMARY_PIE_YEAR_KEY) || "").trim();
  if (saved && years.includes(saved)) return saved;
  return years[years.length - 1];
}

/** Сумма вложений по покупкам (цена × количество), дата сделки в календарном году `yearStr` (YYYY). Без купонов. */
function computeInvestedFromBuysInCalendarYear(buys, yearStr) {
  const want = String(yearStr || "").trim();
  if (!want) return 0;
  return buys.reduce((sum, row) => {
    const dateYmd = normalizeYMD(row.date);
    if (!dateYmd) return sum;
    const rowYear = getYearFromMonthKey(toMonthKeyFromYMD(dateYmd));
    if (rowYear !== want) return sum;
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
  // Сбрасываем подсветку сразу у всех столбцов (для сценариев: выделение месяца/нескольких столбцов).
  document.querySelectorAll(".chartBar.chartHover--active").forEach((el) => {
    if (!(el instanceof SVGElement)) return;
    el.classList.remove("chartHover--active");
  });
  lastHoveredBar = null;
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

function hidePortfolioHoverDot() {
  if (!portfolioChartContent) return;
  const dot = portfolioChartContent.querySelector(".portfolioHoverDot");
  if (dot instanceof SVGCircleElement) dot.setAttribute("opacity", "0");
}

function clearPortfolioHover() {
  if (lastHoveredPortfolioZone) {
    lastHoveredPortfolioZone.classList.remove("portfolioHoverZone--active");
    lastHoveredPortfolioZone = null;
  }
  hidePortfolioHoverDot();
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

/**
 * Hover по подписи месяца на оси X: подсветка всех сегментов месяца; цвет подписи итога и оси — у верхнего сегмента столбца.
 */
function applyMonthAxisHoverHighlight(monthKey) {
  if (!chartContent || !monthKey) return;
  chartContent.classList.add("chartContent--dimOthers");

  const bars = Array.from(chartContent.querySelectorAll(`.chartBar[data-bar-month="${monthKey}"]`)).filter(
    (el) => el instanceof SVGElement
  );
  bars.forEach((el) => el.classList.add("chartHover--active"));

  const barByMatch = new Map();
  let visuallyTopBar = /** @type {SVGElement | null} */ (null);
  let minY = Infinity;
  for (const barEl of bars) {
    const m = normalizeBondKey(barEl.getAttribute("data-bar-match") || "");
    if (m) barByMatch.set(m, barEl);
    const y = chartBarTopEdgeY(barEl);
    if (Number.isFinite(y) && y < minY) {
      minY = y;
      visuallyTopBar = barEl;
    }
  }

  const stackTopFill =
    visuallyTopBar instanceof SVGElement ? String(visuallyTopBar.getAttribute("fill") || "").trim() : "";
  chartContent.querySelectorAll(`.chartValueLabel[data-label-month="${monthKey}"]`).forEach((el) => {
    if (!(el instanceof SVGTextElement)) return;
    if (el.classList.contains("chartBondMonthHint")) return;
    el.classList.add("chartValueLabel--active");
    const lm = normalizeBondKey(el.getAttribute("data-label-match") || "");
    const barEl = lm ? barByMatch.get(lm) : null;
    const fill =
      barEl instanceof SVGElement
        ? String(barEl.getAttribute("fill") || "").trim()
        : stackTopFill ||
          (bars[bars.length - 1] instanceof SVGElement
            ? String(bars[bars.length - 1].getAttribute("fill") || "").trim()
            : "");
    if (fill) {
      el.setAttribute("fill", fill);
      el.setAttribute("opacity", "0.98");
    }
  });

  const axisFill =
    visuallyTopBar instanceof SVGElement
      ? String(visuallyTopBar.getAttribute("fill") || "").trim()
      : bars[0] instanceof SVGElement
        ? String(bars[0].getAttribute("fill") || "").trim()
        : "";
  chartContent.querySelectorAll(`.chartAxisLabel[data-axis-month="${monthKey}"]`).forEach((el) => {
    if (!(el instanceof SVGTextElement)) return;
    el.classList.add("chartText--active");
    if (axisFill) {
      el.setAttribute("fill", axisFill);
      el.setAttribute("opacity", "0.98");
    }
  });
}

function monthLabelFromKey(monthKey) {
  const raw = String(monthKey || "").trim();
  if (/^\d{4}$/.test(raw)) return raw;
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
    const axisLabelTotal = t.closest(".chartAxisLabel[data-axis-total]");
    const axisMonthLabel = t.closest(".chartAxisLabel[data-axis-month]");
    if (axisMonthLabel instanceof SVGElement) {
      hideChartTooltip();
      clearBarHover();
      clearChartTextHighlight();
      if (chartContent) chartContent.classList.remove("chartContent--dimOthers");

      const monthKey = axisMonthLabel.getAttribute("data-axis-month") || "";
      const totalFromAxis = axisLabelTotal ? parseNumber(axisLabelTotal.getAttribute("data-axis-total")) : NaN;
      const total = Number.isFinite(totalFromAxis) ? totalFromAxis : Array.from(document.querySelectorAll(`.chartBar[data-bar-month="${monthKey}"]`)).reduce((sum, barEl) => {
        if (!(barEl instanceof SVGElement)) return sum;
        return sum + (Number(parseNumber(barEl.getAttribute("data-bar-amount"))) || 0);
      }, 0);

      if (chartContent && monthKey) applyMonthAxisHoverHighlight(monthKey);

      const period = monthLabelFromKey(monthKey);

      // Сводка по каждой облигации для этого месяца.
      const bondAmounts = new Map();
      document.querySelectorAll(`.chartBar[data-bar-month="${monthKey}"]`).forEach((barEl) => {
        if (!(barEl instanceof SVGElement)) return;
        const bond = barEl.getAttribute("data-bar-bond") || "";
        const amt = parseNumber(barEl.getAttribute("data-bar-amount"));
        if (!bond || !Number.isFinite(amt)) return;
        bondAmounts.set(bond, (bondAmounts.get(bond) || 0) + amt);
      });
      const bondLines = Array.from(bondAmounts.entries())
        .sort((a, b) => b[1] - a[1])
        .map(([bond, amt]) => `${escapeHtml(bond)}: ${escapeHtml(formatMoney(amt))}`);
      const title = `${escapeHtml(period)} • ${escapeHtml(formatMoney(total))}`;
      const meta = bondLines.length ? `${bondLines.join("<br/>")}` : "—";

      showChartTooltip(title, meta, e.clientX, e.clientY);
      return;
    }
    hideChartTooltip();
    clearBarHover();
    clearChartTextHighlight();
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
  showChartTooltip(
    escapeHtml(bond),
    `Выплата за ${escapeHtml(period)}<br/>${escapeHtml(formatMoney(amount))}`,
    e.clientX,
    e.clientY
  );
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
    `${escapeHtml(year)} · net · ${escapeHtml(shareStr)} · ${escapeHtml(formatMoney(amount))}`,
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
  const topup = parseNumber(zone.getAttribute("data-portfolio-topup"));
  const coupon = parseNumber(zone.getAttribute("data-portfolio-coupon"));
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
    const hoverDot = portfolioChartContent.querySelector(".portfolioHoverDot");
    if (hoverDot instanceof SVGCircleElement && Number.isFinite(x) && Number.isFinite(y)) {
      hoverDot.setAttribute("cx", String(x));
      hoverDot.setAttribute("cy", String(y));
      hoverDot.setAttribute("opacity", "1");
    }
  }
  const period = monthLabelFromKey(monthKey);
  const topupStr = Number.isFinite(topup) ? formatMoney(topup) : "0.00 ₽";
  const couponStr = Number.isFinite(coupon) ? formatMoney(coupon) : "0.00 ₽";
  showChartTooltip(
    "Стоимость портфеля",
    `${escapeHtml(period)}<br/>Ежемесячные взносы: ${escapeHtml(topupStr)}<br/>Купоны: ${escapeHtml(couponStr)}<br/>Итого: ${escapeHtml(formatMoney(value))}`,
    e.clientX,
    e.clientY
  );
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
  if (yearPiesEl && !yearPiesEl.dataset.yearSegClickBound) {
    yearPiesEl.dataset.yearSegClickBound = "1";
    yearPiesEl.addEventListener("click", onYearPieYearSegmentClick);
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

/** Склонение «N выплат/выплаты/выплата в год» для подписей в UI. */
function formatPayoutsPerYearPhrase(n) {
  const k = Math.abs(Number(n)) | 0;
  if (k === 1) return "1 выплата в год";
  if (k >= 2 && k <= 4) return `${k} выплаты в год`;
  return `${k} выплат в год`;
}

function buildStrategyTabBondMetaHtml(price, coupon) {
  const priceOk = Number.isFinite(price) && price > 0;
  const couponOk = Number.isFinite(coupon) && coupon > 0;
  const line1 = priceOk ? `Стоимость ${formatMoney(price)}` : "Стоимость —";
  const line2 = couponOk ? `Купон ${formatMoney(coupon)}` : "Купон —";
  return `<div class="strategyTabShell__bondMetaLine">${escapeHtml(line1)}</div>
            <div class="strategyTabShell__bondMetaLine">${escapeHtml(line2)}</div>`;
}

/** Справа в списке облигаций: частота и (если не ежемесячно) месяцы выплат. */
function formatStrategyTabBondPayoutBlock(monthsSortedUnique) {
  const months = monthsSortedUnique || [];
  const n = months.length;
  if (!n) {
    return { freq: "—", monthsHtml: "" };
  }
  const freq = formatPayoutsPerYearPhrase(n);
  if (n === 12) {
    return { freq, monthsHtml: "" };
  }
  const monthsText = summarizeMonths(months);
  const monthsHtml =
    monthsText && monthsText !== "—"
      ? `<div class="strategyTabShell__bondPayoutMonths">${escapeHtml(monthsText)}</div>`
      : "";
  return { freq, monthsHtml };
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

/** Сумма gross-купонов за календарный месяц по всем сериям — тот же расчёт, что у столбцов на графике выплат. */
function getGrossCouponTotalForMonth(seriesByBond, monthKey) {
  if (!monthKey) return 0;
  let sum = 0;
  for (const series of seriesByBond || []) {
    for (const p of series.points || []) {
      if (p.monthKey === monthKey) sum += Number(p.amount) || 0;
    }
  }
  return sum;
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
      bond: normalizeBondKey(row.bond),
      matchBond: normalizeBondKey(row.bond),
      legendBond: normalizeBondKey(row.bond) || `BOND ${idx + 1}`,
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
          bond: normalizeBondKey(item.bond),
          quantity: item.quantity,
        }));
      }

      // Backward compatibility: old single-item rows
      const bond = normalizeBondKey(row.bond);
      const quantity = parseNumber(row.quantity);
      if (!bond || !Number.isFinite(quantity) || quantity <= 0) return [];
      return [{ dateYMD, dateUTCms, bond, quantity }];
    });

  /** Портфель: для помесячной модели считаем, что владение начинается с 1-го числа месяца начала обращения. */
  const holdingsAsBuys = parseHoldingItems(holdings).map((item) => {
    const bondKey = normalizeBondKey(item.bond);
    const bondMeta = normalizedBonds.find((b) => b.matchBond === bondKey);
    let dateYMD = bondMeta?.startDate ? normalizeYMD(bondMeta.startDate) || bondMeta.startDate : "";
    if (dateYMD) {
      const mk = toMonthKeyFromYMD(dateYMD);
      if (mk) dateYMD = `${mk}-01`;
    }
    if (!dateYMD) dateYMD = getTodayYMD();
    let dateUTCms = ymdToUTCms(dateYMD);
    if (dateUTCms == null || !Number.isFinite(dateUTCms)) {
      dateYMD = getTodayYMD();
      dateUTCms = ymdToUTCms(dateYMD) ?? 0;
    }
    return {
      dateYMD,
      dateUTCms,
      bond: bondKey,
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
    const startMonthKey = toMonthKeyFromYMD(bondRow.startDate);
    const endMonthKey = toMonthKeyFromYMD(bondRow.endDate);
    const start = new Date(startTs);
    const end = new Date(endTs);

    for (let y = start.getUTCFullYear(); y <= end.getUTCFullYear(); y += 1) {
      bondRow.payoutMonths.forEach((monthNum) => {
        const monthIndex = monthNum - 1;
        const monthKey = `${String(y).padStart(4, "0")}-${String(monthNum).padStart(2, "0")}`;
        if (startMonthKey && monthKeyToUTCms(monthKey) < monthKeyToUTCms(startMonthKey)) return;
        if (endMonthKey && monthKeyToUTCms(monthKey) > monthKeyToUTCms(endMonthKey)) return;
        const paymentDate = new Date(Date.UTC(y, monthIndex, 1));
        const paymentTs = paymentDate.getTime();
        allDateSet.add(monthKey);

        const quantityAtDate = normalizedBuys
          .concat(holdingsAsBuys)
          .filter(
            (buyRow) =>
              normalizeBondKey(buyRow.bond) === bondRow.matchBond && buyRow.dateUTCms <= paymentTs
          )
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

/** Минимальная высота каждого ненулевого сегмента составного столбца (px). */
const CHART_STACK_MIN_SEG_PX = 10;

/**
 * Высоты сегментов: у каждого ≥ minPx (если высоты столбца хватает: barTotalH ≥ n·minPx),
 * остаток barTotalH распределяется пропорционально суммам — без масштабирования «вниз», которое ломало минимум.
 * @param {number[]} amounts
 * @param {number} barTotalH
 * @param {number} minPx
 */
function stackSegmentHeightsPx(amounts, barTotalH, minPx) {
  const n = amounts.length;
  if (!n || !(barTotalH > 0)) return amounts.map(() => 0);
  const sumAmt = amounts.reduce((a, b) => a + (Number(b) || 0), 0);
  if (!(sumAmt > 0)) return amounts.map(() => 0);

  const minSum = n * minPx;
  if (barTotalH >= minSum) {
    const remaining = barTotalH - minSum;
    return amounts.map((a) => minPx + ((Number(a) || 0) / sumAmt) * remaining);
  }

  return amounts.map(() => minPx);
}

const CHART_STACK_TOP_CORNER_R = 2;

function chartBarTopEdgeY(barEl) {
  if (!(barEl instanceof SVGElement)) return NaN;
  const fromData = barEl.getAttribute("data-bar-top-y");
  if (fromData != null && String(fromData).trim() !== "") {
    const n = parseNumber(fromData);
    if (Number.isFinite(n)) return n;
  }
  return parseNumber(barEl.getAttribute("y"));
}

/** Верхние углы скруглены (дуги SVG), низ прямой. */
function chartStackTopSegmentPathD(x, yTop, w, h, rRaw) {
  if (!(w > 0 && h > 0)) return "";
  const r = Math.min(Math.max(0, rRaw), w / 2, h / 2);
  if (r < 0.25) {
    return `M ${x} ${yTop} L ${x + w} ${yTop} L ${x + w} ${yTop + h} L ${x} ${yTop + h} Z`;
  }
  return `M ${x} ${yTop + r} A ${r} ${r} 0 0 1 ${x + r} ${yTop} L ${x + w - r} ${yTop} A ${r} ${r} 0 0 1 ${x + w} ${yTop + r} L ${x + w} ${yTop + h} L ${x} ${yTop + h} Z`;
}

function renderChart(chartData) {
  hideChartTooltip();
  clearAllChartHover();
  const allDatesRaw = chartData?.allDates || [];
  const seriesByBondRaw = chartData?.seriesByBond || [];
  const selectedYear = ensureChartYearOptions(chartData);
  const isCouponChartAllTime = selectedYear === PORTFOLIO_CHART_YEAR_ALL;
  const allDates = isCouponChartAllTime
    ? allDatesRaw
    : selectedYear
      ? allDatesRaw.filter((mk) => getYearFromMonthKey(mk) === selectedYear)
      : allDatesRaw;
  const seriesByBond = seriesByBondRaw.map((s) => ({
    ...s,
    points: s.points.filter((p) =>
      isCouponChartAllTime || !selectedYear ? true : getYearFromMonthKey(p.monthKey) === selectedYear
    ),
  }));

  /** Для «За все время» — столбцы по календарному году; иначе по месяцу YYYY-MM. */
  /** @type {string[]} */
  let visibleDates;
  /** @type {Map<string, number>[]} */
  let amountByBondAndDate;

  if (isCouponChartAllTime) {
    amountByBondAndDate = seriesByBond.map(() => new Map());
    const yearTotals = new Map();
    seriesByBond.forEach((series, sIdx) => {
      const m = amountByBondAndDate[sIdx];
      for (const p of series.points) {
        const y = getYearFromMonthKey(p.monthKey);
        if (!y) continue;
        const amt = Number(p.amount) || 0;
        if (!(amt > 0)) continue;
        m.set(y, (m.get(y) || 0) + amt);
        yearTotals.set(y, (yearTotals.get(y) || 0) + amt);
      }
    });
    visibleDates = Array.from(yearTotals.keys())
      .filter((y) => (yearTotals.get(y) || 0) > 0)
      .sort();
  } else {
    visibleDates = allDates.filter((monthKey) =>
      seriesByBond.some((series) => series.points.some((point) => point.monthKey === monthKey && point.amount > 0))
    );
    amountByBondAndDate = seriesByBond.map((series) => {
      const map = new Map();
      series.points.forEach((p) => map.set(p.monthKey, p.amount));
      return map;
    });
  }
  const W = 800;
  const H = 360;
  const left = 20;
  const right = 12;
  const top = 24;
  const bottom = 42;
  const innerW = W - left - right;
  const innerH = H - top - bottom;

  const lines = [];

  for (let i = 0; i <= 4; i += 1) {
    const y = top + (innerH / 4) * i;
    lines.push(`<path d="M ${left} ${y} H ${W - right}" stroke="rgba(60,60,67,0.2)" stroke-width="1"></path>`);
  }

  if (!visibleDates.length || !seriesByBond.length) {
    if (chartLegend) chartLegend.innerHTML = "";
    const emptyTitle = isCouponChartAllTime
      ? "Нет выплат за весь выбранный период"
      : "Нет выплат за выбранный год";
    lines.push(buildSvgEmptyState(emptyTitle, "Добавьте облигации и даты купонов, чтобы построить график выплат."));
    chartContent.innerHTML = lines.join("");
    return;
  }

  const palette = getSeriesPalette();
  renderChartLegend(seriesByBond, palette);

  const monthTotals = visibleDates.map((monthKey) =>
    seriesByBond.reduce((sum, _series, sIdx) => sum + (amountByBondAndDate[sIdx].get(monthKey) || 0), 0)
  );
  const maxY = Math.max(...monthTotals, 0);
  const safeMaxY = maxY <= 0 ? 1 : maxY * 1.15;

  const groupWidth = innerW / visibleDates.length;
  const barWidth = Math.max(14, Math.min(36, groupWidth * 0.52));

  visibleDates.forEach((bucketKey, dateIdx) => {
    const activeBars = seriesByBond
      .map((bondRow, sIdx) => ({
        bondRow,
        sIdx,
        amount: amountByBondAndDate[sIdx].get(bucketKey) || 0,
      }))
      .filter((item) => item.amount > 0);
    if (!activeBars.length) return;

    const monthTotal = activeBars.reduce((sum, item) => sum + item.amount, 0);
    const barTotalH = (monthTotal / safeMaxY) * innerH;
    const groupCenterX = left + dateIdx * groupWidth + groupWidth / 2;
    const x = groupCenterX - barWidth / 2;
    let yStackBottom = top + innerH;

    const ordered = [...activeBars].sort((a, b) => a.sIdx - b.sIdx);
    const segHeights = stackSegmentHeightsPx(
      ordered.map((it) => it.amount),
      barTotalH,
      CHART_STACK_MIN_SEG_PX
    );
    ordered.forEach((item, i) => {
      const { bondRow, sIdx, amount } = item;
      const color = palette[sIdx % palette.length];
      const segH = segHeights[i] || 0;
      yStackBottom -= segH;
      const yTop = yStackBottom;
      const isTopSeg = i === ordered.length - 1;
      const commonData = `data-bar-bond="${escapeHtml(bondRow.bond)}" data-bar-match="${escapeHtml(
        bondRow.matchBond
      )}" data-bar-month="${bucketKey}" data-bar-amount="${amount}" data-bar-top-y="${yTop}"`;
      if (isTopSeg) {
        const r = Math.min(CHART_STACK_TOP_CORNER_R, barWidth / 2, segH / 2);
        const d = chartStackTopSegmentPathD(x, yTop, barWidth, segH, r);
        lines.push(
          `<path class="chartBar" ${commonData} d="${d}" fill="${color}" stroke="none" opacity="0.9"></path>`
        );
      } else {
        lines.push(
          `<rect class="chartBar" ${commonData} x="${x}" y="${yTop}" width="${barWidth}" height="${segH}" fill="${color}" stroke="none" opacity="0.9"></rect>`
        );
      }
    });

    const stackTopY = yStackBottom;
    seriesByBond.forEach((series, sIdx) => {
      const amt = amountByBondAndDate[sIdx].get(bucketKey) || 0;
      if (!(amt > 0)) return;
      const color = palette[sIdx % palette.length];
      const safeColor = escapeHtml(color);
      const safeMatch = escapeHtml(String(series.matchBond || ""));
      lines.push(
        `<text class="chartValueLabel chartBondMonthHint" data-label-match="${safeMatch}" data-label-month="${bucketKey}" data-default-fill="${safeColor}" data-default-opacity="0" pointer-events="none" x="${groupCenterX}" y="${stackTopY - 13}" text-anchor="middle" fill="${safeColor}" opacity="0" font-size="9.5">${formatMoney(amt)}</text>`
      );
    });
    lines.push(
      `<text class="chartValueLabel" data-label-month="${bucketKey}" data-label-match="" data-default-fill="currentColor" data-default-opacity="0.72" pointer-events="none" x="${groupCenterX}" y="${stackTopY - 6}" text-anchor="middle" fill="currentColor" opacity="0.72" font-size="10">${formatMoney(monthTotal)}</text>`
    );

    const labelX = groupCenterX;
    if (isCouponChartAllTime) {
      const yStr = escapeHtml(bucketKey);
      lines.push(
        `<text class="chartAxisLabel chartAxisLabel--interactive" data-axis-month="${yStr}" data-axis-total="${monthTotal}" data-default-fill="currentColor" data-default-opacity="0.68" x="${labelX}" y="${H - 22}" text-anchor="middle" fill="currentColor" opacity="0.68" font-size="11.5">${yStr}</text>`
      );
    } else {
      const monthLabel = formatMonthKey(bucketKey);
      lines.push(
        `<text class="chartAxisLabel chartAxisLabel--interactive" data-axis-month="${bucketKey}" data-axis-total="${monthTotal}" data-default-fill="currentColor" data-default-opacity="0.68" x="${labelX}" y="${H - 28}" text-anchor="middle" fill="currentColor" opacity="0.68" font-size="11">${monthLabel.month}</text>`
      );
      lines.push(
        `<text class="chartAxisLabel" data-axis-month="${bucketKey}" data-default-fill="currentColor" data-default-opacity="0.62" x="${labelX}" y="${H - 12}" text-anchor="middle" fill="currentColor" opacity="0.62" font-size="10">${monthLabel.year}</text>`
      );
    }
  });

  chartContent.innerHTML = lines.join("");
}

function renderSummary(chartData, buys, holdings) {
  if (!summaryTbody || !totalGrossEl || !totalTaxEl || !totalNetEl) return;

  const taxRate = getTaxRateDecimal();
  const netChartData = toNetChartData(chartData, taxRate);
  const seriesByBond = netChartData?.seriesByBond || [];
  const totalInvested = computeTotalInvestedFromBuys(buys);
  renderYearPies(netChartData);

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
  if (totalIncomeLabelEl) totalIncomeLabelEl.textContent = "К выплате";
  if (totalIncomeHelpBtn) {
    totalIncomeHelpBtn.setAttribute(
      "data-summary-help",
      "К выплате: итоговая сумма купонов после налога. Формула: сумма всех net-выплат."
    );
  }

  if (totalInvestedEl) totalInvestedEl.textContent = formatMoney(totalInvested);

  let monthlyYieldPct = NaN;
  let annualYieldPct = NaN;

  const payoutYearsForYield = getPayoutYearsFromChartData(chartData);
  const selectedYearForYield = resolveSummaryPieYearFromYearList(payoutYearsForYield);
  const investedInSelectedYear = selectedYearForYield
    ? computeInvestedFromBuysInCalendarYear(buys, selectedYearForYield)
    : 0;

  if (selectedYearForYield) {
    const netForYear = seriesByBond.reduce((sum, series) => {
      const pts = series.points || [];
      return sum + pts.reduce((s, point) => {
        const y = getYearFromMonthKey(point.monthKey);
        return y && y === selectedYearForYield ? s + (Number(point.amount) || 0) : s;
      }, 0);
    }, 0);

    if (investedInSelectedYear > 0) {
      annualYieldPct = (netForYear / investedInSelectedYear) * 100;
      // Доля от вложений за этот же год в среднем за календарный месяц (годовая ставка / 12).
      monthlyYieldPct = (netForYear / investedInSelectedYear) * (100 / 12);
    }
  } else if (totalInvested > 0) {
    const allDates = Array.isArray(chartData?.allDates) ? chartData.allDates : [];
    const first = allDates[0];
    const last = allDates[allDates.length - 1];
    const m1 = first ? /^(\d{4})-(\d{2})$/.exec(first) : null;
    const m2 = last ? /^(\d{4})-(\d{2})$/.exec(last) : null;
    const monthsSpan = m1 && m2
      ? (Number(m2[1]) - Number(m1[1])) * 12 + (Number(m2[2]) - Number(m1[2])) + 1
      : 12;

    annualYieldPct = (totalNet / totalInvested) * (12 / Math.max(1, monthsSpan)) * 100;
    monthlyYieldPct = (totalNet / totalInvested) * (100 / Math.max(1, monthsSpan));
  }
  if (yieldAnnualLabelEl) {
    yieldAnnualLabelEl.textContent = selectedYearForYield
      ? `Средняя доходность за ${selectedYearForYield} год`
      : "Средняя доходность за год";
  }
  if (yieldMonthlyLabelEl) {
    yieldMonthlyLabelEl.textContent = selectedYearForYield
      ? `Среднемесячная доходность в ${selectedYearForYield} году`
      : "Среднемесячная доходность за период";
  }
  if (yieldMonthlyHelpBtn) {
    yieldMonthlyHelpBtn.setAttribute(
      "data-summary-help",
      "За год из переключателя под круговой диаграммой: сколько процентов от вложений за этот же год в среднем приходится на один календарный месяц. Формула: (выплаты net за год ÷ вложения за год) ÷ 12 × 100%. Если годов нет — по всему горизонту: (все net-выплаты ÷ все вложения) ÷ число месяцев горизонта × 100%."
    );
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
  if (!portfolioChartContent || !portfolioStartTotalEl || !portfolioCouponTotalEl || !portfolioTopupTotalEl || !portfolioEndTotalEl)
    return;

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

  let cumulative = startValue;
  const allPoints = monthRange.map((monthKey) => {
    const topupForMonth = !monthlyTopupEndTs || monthKeyToUTCms(monthKey) <= monthlyTopupEndTs ? monthlyTopup : 0;
    const couponForMonth = getGrossCouponTotalForMonth(seriesByBond, monthKey);
    cumulative += topupForMonth + couponForMonth;
    return { monthKey, value: cumulative, topup: topupForMonth, coupon: couponForMonth };
  });
  const selectedYear = ensurePortfolioChartYearOptions(allPoints);
  const isAllTime = selectedYear === PORTFOLIO_CHART_YEAR_ALL;
  const points = isAllTime ? allPoints : allPoints.filter((p) => getYearFromMonthKey(p.monthKey) === selectedYear);

  if (!points.length) {
    portfolioStartTotalEl.textContent = formatMoney(startValue);
    portfolioCouponTotalEl.textContent = formatMoney(0);
    portfolioTopupTotalEl.textContent = formatMoney(0);
    portfolioEndTotalEl.textContent = formatMoney(startValue);
    portfolioChartContent.innerHTML = buildSvgEmptyState(
      isAllTime ? "Нет данных для графика" : "За выбранный год нет движений",
      isAllTime
        ? "Проверьте дату старта и облигации с выплатами."
        : "Смените год или добавьте облигации с купонными выплатами в этот период."
    );
    return;
  }

  const firstPoint = points[0];
  const firstPointTopup = Number.isFinite(firstPoint.topup) ? firstPoint.topup : 0;
  const firstPointMonthCoupon = Number.isFinite(firstPoint.coupon)
    ? firstPoint.coupon
    : getGrossCouponTotalForMonth(seriesByBond, firstPoint.monthKey);
  const periodStartValue = firstPoint.value - firstPointTopup - firstPointMonthCoupon;
  const couponOnlyAdded = points.reduce((sum, point) => sum + (Number(point.coupon) || 0), 0);
  const topupOnlyAdded = points.reduce((sum, point) => sum + (Number(point.topup) || 0), 0);
  const endValue = points[points.length - 1].value;
  portfolioStartTotalEl.textContent = formatMoney(periodStartValue);
  portfolioCouponTotalEl.textContent = formatMoney(couponOnlyAdded);
  portfolioTopupTotalEl.textContent = formatMoney(topupOnlyAdded);
  portfolioEndTotalEl.textContent = formatMoney(endValue);

  const W = 900;
  const H = 360;
  const left = 78;
  const right = 18;
  const top = 24;
  const bottom = 50;
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
    lines.push(`<text x="${left - 12}" y="${y + 5}" text-anchor="end" fill="currentColor" opacity="0.58" font-size="11">${formatAmount(value)}</text>`);
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
    if (!isAllTime) {
      lines.push(`<circle cx="${x}" cy="${y}" r="4" fill="#007AFF"></circle>`);
      const monthLabel = formatMonthKey(p.monthKey);
      lines.push(`<text x="${x}" y="${H - 30}" text-anchor="middle" fill="currentColor" opacity="0.68" font-size="11">${monthLabel.month}</text>`);
      lines.push(`<text x="${x}" y="${H - 13}" text-anchor="middle" fill="currentColor" opacity="0.62" font-size="10">${monthLabel.year}</text>`);
    }
    const zoneWidth = points.length === 1 ? innerW : Math.max(12, xStep);
    const zoneX = points.length === 1 ? left : x - zoneWidth / 2;
    lines.push(
      `<rect class="portfolioHoverZone" data-portfolio-month="${p.monthKey}" data-portfolio-value="${p.value}" data-portfolio-topup="${p.topup}" data-portfolio-coupon="${p.coupon}" data-portfolio-x="${x}" data-portfolio-y="${y}" x="${zoneX}" y="${top}" width="${zoneWidth}" height="${innerH}" fill="transparent"></rect>`
    );
  });

  if (isAllTime && points.length) {
    let segStart = 0;
    for (let i = 1; i <= points.length; i += 1) {
      const yStart = getYearFromMonthKey(points[segStart].monthKey);
      const yNext = i < points.length ? getYearFromMonthKey(points[i].monthKey) : null;
      if (i === points.length || yNext !== yStart) {
        const cx = (xAt(segStart) + xAt(i - 1)) / 2;
        lines.push(
          `<text x="${cx}" y="${H - 22}" text-anchor="middle" fill="currentColor" opacity="0.72" font-size="11.5">${escapeHtml(yStart)}</text>`
        );
        segStart = i;
      }
    }
  }

  if (isAllTime && points.length) {
    const hx = xAt(0);
    const hy = yAt(points[0].value);
    lines.push(
      `<circle class="portfolioHoverDot" cx="${hx}" cy="${hy}" r="3.5" fill="#007AFF" opacity="0" pointer-events="none" aria-hidden="true"></circle>`
    );
  }

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

function buildBuyChipsHtmlFromItems(items) {
  if (!Array.isArray(items) || !items.length) return "";
  const itemsSorted = [...items].sort((a, b) =>
    String(a.bond || "").localeCompare(String(b.bond || ""), "ru", { numeric: true, sensitivity: "base" })
  );
  return `<div class="buyChips">${itemsSorted
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
}

function findBondTableRowByBondKey(bondKeyRaw) {
  const want = normalizeBondKey(bondKeyRaw);
  if (!want || !bondsTbody) return null;
  for (const tr of bondsTbody.querySelectorAll("tr")) {
    if (tr.hasAttribute("data-empty-state")) continue;
    const hid = tr.querySelector('input[data-field="bond"]');
    if (!hid) continue;
    if (normalizeBondKey(hid.value) === want) return tr;
  }
  return null;
}

/** Стабильный ключ строки покупки для сопоставления зеркала на вкладке «Стратегия» с таблицей планирования. */
function stableBuyRowKeyFromRowData(row) {
  const dateYmd = normalizeYMD(row.date) || "";
  const itemsStr = String(row.items ?? "").trim();
  return `${dateYmd}\t${itemsStr}`;
}

function stableBuyRowKeyFromTr(tr) {
  if (!tr) return "";
  const date = tr.querySelector('[data-field="date"]')?.value || "";
  const items = tr.querySelector('[data-field="items"]')?.value || "";
  return stableBuyRowKeyFromRowData({ date, items });
}

function findBuyTableRowByStableKey(wantKey) {
  if (!wantKey || !buysTbody) return null;
  for (const tr of buysTbody.querySelectorAll("tr.buyTable__dataRow")) {
    if (tr.hasAttribute("data-empty-state")) continue;
    if (stableBuyRowKeyFromTr(tr) === wantKey) return tr;
  }
  return null;
}

function removeBondRowByBondKey(bondKeyRaw) {
  const tr = findBondTableRowByBondKey(bondKeyRaw);
  if (!tr) return false;
  tr.remove();
  syncStaticTableEmptyStates();
  persistAndRender();
  return true;
}

function renderStrategyTab() {
  if (!strategyTabStrategyList || !strategyTabBuysTbody || !strategyTabBondsList) return;

  strategyTabStrategyList.innerHTML = buyStrategies
    .map((s) => {
      const isActive = s.id === activeBuyStrategyId;
      const isDefault = s.id === DEFAULT_BUY_STRATEGY_ID;
      const idEsc = escapeHtml(String(s.id));
      const actions = isDefault
        ? ""
        : `<button type="button" class="strategyTabShell__menuBtn" data-strategy-action="open-menu" data-strategy-id="${idEsc}" title="Действия" aria-label="Действия со стратегией" aria-haspopup="dialog">⋮</button>`;
      return `<li class="strategyTabShell__strategyRow${isActive ? " is-active" : ""}" data-strategy-id="${idEsc}">
        <button type="button" class="strategyTabShell__strategySelect">
          <span class="strategyTabShell__strategyName">${escapeHtml(String(s.name))}</span>
        </button>
        ${actions}
      </li>`;
    })
    .join("");

  const buyRowsComplete = getActiveStrategyBuysRows().filter(isBuyRowComplete);
  const { rows: sortedBuys } = sortBuyRowsByDate(buyRowsComplete);
  if (!sortedBuys.length) {
    strategyTabBuysTbody.innerHTML = `<tr><td class="strategyTabShell__empty" colspan="3">Нет запланированных покупок. Нажмите «Добавить покупку» или заполните план во вкладке «Планирование».</td></tr>`;
  } else {
    strategyTabBuysTbody.innerHTML = sortedBuys
      .map((row) => {
        const d = normalizeYMD(row.date);
        const items = parseBuyItems(row.items);
        const total = items.reduce((sum, i) => sum + i.price * i.quantity, 0);
        const itemsHtml = !items.length ? "—" : `<div class="buySummary">${buildBuyChipsHtmlFromItems(items)}</div>`;
        const dateCell = `<div class="buyTableDate">${escapeHtml(d ? formatDateRuMonthWords(d) : "—")}</div>`;
        const keyAttr = encodeURIComponent(stableBuyRowKeyFromRowData(row));
        return `<tr class="strategyTabShell__buyRowClick" data-strategy-buy-key="${keyAttr}" tabindex="0" role="button">
        <td>${dateCell}</td>
        <td>${itemsHtml}</td>
        <td><div class="buySummary">${escapeHtml(formatMoney(total))}</div></td>
      </tr>`;
      })
      .join("");
  }

  const bonds = sanitizeBondRows(readRows(bondsTbody));
  if (!bonds.length) {
    strategyTabBondsList.innerHTML = `<li class="strategyTabShell__bondEmpty strategyTabShell__empty">Нет облигаций. Нажмите «+» выше или добавьте строку во вкладке «Планирование».</li>`;
  } else {
    strategyTabBondsList.innerHTML = bonds
      .map((row) => {
        const nameRaw = String(row.bond || "").trim();
        const name = normalizeBondKey(nameRaw) || nameRaw || "—";
        const coupon = parseNumber(row.coupon);
        const price = parseNumber(row.bondPrice);
        const monthsNorm = Array.from(new Set(parseMonthList(row.payoutMonths))).sort((a, b) => a - b);
        const { freq, monthsHtml } = formatStrategyTabBondPayoutBlock(monthsNorm);
        const metaHtml = buildStrategyTabBondMetaHtml(price, coupon);
        const keyEsc = escapeHtml(normalizeBondKey(nameRaw));
        return `<li class="strategyTabShell__bondRow">
          <button type="button" class="strategyTabShell__bondOpenBtn" data-strategy-bond="${keyEsc}">
            <div class="strategyTabShell__bondMain">
              <div class="strategyTabShell__bondName">${escapeHtml(name)}</div>
              <div class="strategyTabShell__bondMeta">${metaHtml}</div>
            </div>
            <div class="strategyTabShell__bondPayout">
              <div class="strategyTabShell__bondPayoutFreq">${escapeHtml(freq)}</div>
              ${monthsHtml}
            </div>
          </button>
        </li>`;
      })
      .join("");
  }
}

function initStrategyTabPanel() {
  if (!strategyTabStrategyList || strategyTabStrategyList.dataset.strategyTabInit) return;
  strategyTabStrategyList.dataset.strategyTabInit = "1";
  strategyTabStrategyList.addEventListener("click", (e) => {
    const menuBtn = e.target.closest('[data-strategy-action="open-menu"]');
    if (menuBtn) {
      e.stopPropagation();
      openStrategyActionsModal(menuBtn.getAttribute("data-strategy-id") || "");
      return;
    }
    const sel = e.target.closest(".strategyTabShell__strategySelect");
    if (!sel) return;
    const li = sel.closest("[data-strategy-id]");
    if (!li) return;
    switchActiveBuyStrategy(li.getAttribute("data-strategy-id") || "");
  });
  strategyTabStrategyList.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    const menuBtn = e.target.closest(".strategyTabShell__menuBtn");
    if (menuBtn) {
      e.preventDefault();
      openStrategyActionsModal(menuBtn.getAttribute("data-strategy-id") || "");
      return;
    }
    const sel = e.target.closest(".strategyTabShell__strategySelect");
    if (!sel) return;
    const li = sel.closest("[data-strategy-id]");
    if (!li) return;
    e.preventDefault();
    switchActiveBuyStrategy(li.getAttribute("data-strategy-id") || "");
  });
  if (strategyTabNewStrategyBtn) {
    strategyTabNewStrategyBtn.addEventListener("click", () => openStrategyModalCreate());
  }
  if (strategyTabNewBondBtn) {
    strategyTabNewBondBtn.addEventListener("click", () => openBondModalNew());
  }
  if (strategyTabAddBuyBtn) {
    strategyTabAddBuyBtn.addEventListener("click", () => openBuyModal());
  }
  const strategyDisplayList = document.getElementById("strategy-tab-display-list");
  if (strategyDisplayList && !strategyDisplayList.dataset.strategyDisplayInit) {
    strategyDisplayList.dataset.strategyDisplayInit = "1";
    strategyDisplayList.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-strategy-display]");
      if (!btn) return;
      const mode = btn.getAttribute("data-strategy-display");
      if (mode !== "buys" && mode !== "charts" && mode !== "portfolio") return;
      setStrategyCenterView(mode);
    });
  }
  if (strategyTabBuysTbody && !strategyTabBuysTbody.dataset.strategyBuysInit) {
    strategyTabBuysTbody.dataset.strategyBuysInit = "1";
    strategyTabBuysTbody.addEventListener("click", (e) => {
      const tr = e.target.closest("tr[data-strategy-buy-key]");
      if (!tr) return;
      const enc = tr.getAttribute("data-strategy-buy-key") || "";
      let key = "";
      try {
        key = decodeURIComponent(enc);
      } catch {
        return;
      }
      const mainTr = findBuyTableRowByStableKey(key);
      if (mainTr) openBuyModalForEdit(mainTr);
    });
    strategyTabBuysTbody.addEventListener("keydown", (e) => {
      if (e.key !== "Enter" && e.key !== " ") return;
      const tr = e.target.closest("tr[data-strategy-buy-key]");
      if (!tr) return;
      e.preventDefault();
      const enc = tr.getAttribute("data-strategy-buy-key") || "";
      let key = "";
      try {
        key = decodeURIComponent(enc);
      } catch {
        return;
      }
      const mainTr = findBuyTableRowByStableKey(key);
      if (mainTr) openBuyModalForEdit(mainTr);
    });
  }
  if (strategyTabBondsList) {
    strategyTabBondsList.addEventListener("click", (e) => {
      const openBtn = e.target.closest(".strategyTabShell__bondOpenBtn");
      if (!openBtn) return;
      const mainTr = findBondTableRowByBondKey(openBtn.getAttribute("data-strategy-bond") || "");
      if (mainTr) openBondModalForEdit(mainTr);
    });
    strategyTabBondsList.addEventListener("keydown", (e) => {
      if (e.key !== "Enter" && e.key !== " ") return;
      const openBtn = e.target.closest(".strategyTabShell__bondOpenBtn");
      if (!openBtn) return;
      e.preventDefault();
      const mainTr = findBondTableRowByBondKey(openBtn.getAttribute("data-strategy-bond") || "");
      if (mainTr) openBondModalForEdit(mainTr);
    });
  }
}

function renderAll() {
  syncHoldingsBondSelects();
  const layoutSnap = captureVisibleWidgetLayout();
  const bonds = sanitizeBondRows(readRows(bondsTbody));
  const buys = sortBuyRowsByDate(readRows(buysTbody)).rows.filter(isBuyRowComplete);
  const portfolioBuys = getBuysForPortfolioChart();
  const holdings = sanitizeHoldingRows(readRows(holdingsTbody));
  const payoutSeries = buildPayoutSeries(bonds, buys, holdings);
  const payoutSeriesPortfolio = buildPayoutSeries(bonds, portfolioBuys, holdings);
  renderPortfolioChart(payoutSeriesPortfolio);
  renderChart(payoutSeries);
  renderSummary(payoutSeries, buys, holdings);
  renderCalculators();
  refreshAutoPlanBondPicker();
  renderStrategyTab();
  animateVisibleWidgetLayout(layoutSnap);
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
  buyRowsByStrategyId.set(activeBuyStrategyId, buysSorted);
  localStorage.setItem(BUYS_KEY, JSON.stringify(buysSorted));
  localStorage.setItem(HOLDINGS_KEY, JSON.stringify(holdings));
  persistBuyStrategiesState();
  // Перерисовываем строки только если реально поменялся порядок.
  if (sortResult.didSort) {
    writeRows(buysTbody, buyTpl, buysSorted);
  } else {
    refreshBuyYearFilterUi();
  }
  syncBuySummaries();
  syncStaticTableEmptyStates();
  if (taxRateInput) localStorage.setItem(TAX_RATE_KEY, String(parseNumber(taxRateInput.value) || 0));
  renderAll();
}

function compareYmd(a, b) {
  const ax = ymdToUTCms(normalizeYMD(a) || "");
  const ay = ymdToUTCms(normalizeYMD(b) || "");
  if (!Number.isFinite(ax) || !Number.isFinite(ay)) return 0;
  if (ax < ay) return -1;
  if (ax > ay) return 1;
  return 0;
}

function daysInCalendarMonthUTC(year, month1to12) {
  return new Date(Date.UTC(year, month1to12, 0)).getUTCDate();
}

function roundRub2(n) {
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

/**
 * Все даты выплат купона (1-е число месяцев из графика) в пределах [startDate, endDate].
 * @param {{ startDate: string, endDate: string, payoutMonths: number[] }} bondRow
 */
function listBondPaymentYmdsSorted(bondRow) {
  const start = bondRow.startDate;
  const end = bondRow.endDate;
  const startTs = ymdToUTCms(start);
  const endTs = ymdToUTCms(end);
  const months = bondRow.payoutMonths || [];
  if (!months.length || !Number.isFinite(startTs) || !Number.isFinite(endTs) || startTs > endTs) return [];

  const out = [];
  const y0 = new Date(startTs).getUTCFullYear();
  const y1 = new Date(endTs).getUTCFullYear();
  for (let y = y0; y <= y1; y += 1) {
    for (const monthNum of months) {
      const ymd = `${y}-${String(monthNum).padStart(2, "0")}-01`;
      const ts = ymdToUTCms(ymd);
      if (ts >= startTs && ts <= endTs) out.push(ymd);
    }
  }
  return out.sort((x, y) => ymdToUTCms(x) - ymdToUTCms(y));
}

/** Есть ли после даты сделки хотя бы одна запланированная выплата купона (строго позже даты покупки). */
function bondHasFutureCouponPaymentAfter(bondRow, purchaseYmd) {
  const purchaseTs = ymdToUTCms(normalizeYMD(purchaseYmd) || "");
  const payDates = listBondPaymentYmdsSorted(bondRow);
  if (!payDates.length || !Number.isFinite(purchaseTs)) return false;
  return payDates.some((d) => ymdToUTCms(d) > purchaseTs);
}

/** Число запланированных купонов строго после даты покупки (для приоритизации горизонта). */
function countFutureCouponPaymentsAfterPurchase(bondRow, purchaseYmd) {
  const purchaseTs = ymdToUTCms(normalizeYMD(purchaseYmd) || "");
  const payDates = listBondPaymentYmdsSorted(bondRow);
  if (!payDates.length || !Number.isFinite(purchaseTs)) return 0;
  let n = 0;
  for (const d of payDates) {
    if (ymdToUTCms(d) > purchaseTs) n += 1;
  }
  return n;
}

/**
 * НКД на одну облигацию на дату сделки (руб.), линейная модель между соседними выплатами.
 */
function accruedCouponPerBondRub(bondRow, purchaseYmd) {
  const purchaseTs = ymdToUTCms(normalizeYMD(purchaseYmd) || "");
  const coupon = bondRow.coupon;
  if (!Number.isFinite(purchaseTs) || !Number.isFinite(coupon) || coupon <= 0) return 0;

  const payDates = listBondPaymentYmdsSorted(bondRow);
  const startTsBond = ymdToUTCms(bondRow.startDate);
  if (!payDates.length || !Number.isFinite(startTsBond)) return 0;

  let prevTs = null;
  let nextTs = null;
  for (const d of payDates) {
    const ts = ymdToUTCms(d);
    if (ts <= purchaseTs) prevTs = ts;
    if (ts > purchaseTs) {
      nextTs = ts;
      break;
    }
  }

  if (nextTs === null) return 0;

  if (prevTs === null) prevTs = startTsBond;

  const periodMs = nextTs - prevTs;
  const elapsedMs = purchaseTs - prevTs;
  if (periodMs <= 0 || elapsedMs <= 0) return 0;

  return coupon * (elapsedMs / periodMs);
}

function dirtyPricePerBondRub(bondRow, purchaseYmd) {
  const clean = bondRow.cleanPrice;
  if (!Number.isFinite(clean) || clean <= 0) return NaN;
  if (!bondHasFutureCouponPaymentAfter(bondRow, purchaseYmd)) return NaN;
  return roundRub2(clean + accruedCouponPerBondRub(bondRow, purchaseYmd));
}

function getSelectedAutoPlanBondKeys() {
  const el = document.getElementById("auto-plan-bonds");
  if (!el) return [];
  return Array.from(el.querySelectorAll('input[type="checkbox"][data-bond]:checked'))
    .map((inp) => normalizeBondKey(inp.getAttribute("data-bond")))
    .filter(Boolean);
}

function getSelectedAutoPlanDayNums() {
  const wrap = document.getElementById("auto-plan-topup-days");
  if (!wrap) return [];
  return Array.from(wrap.querySelectorAll(".planScheduleDayChip--selected"))
    .map((btn) => Number(btn.getAttribute("data-day")))
    .filter((n) => Number.isInteger(n) && n >= 1 && n <= 31)
    .sort((a, b) => a - b);
}

/**
 * @param {{ onlyKeys: Set<string>|null, requireCleanPrice: boolean }} options
 * onlyKeys: null — все строки таблицы, прошедшие проверку; иначе только из множества.
 */
function readBondConfigsForAutoPlanFromTable(options) {
  const { onlyKeys, requireCleanPrice } = options;
  if (!bondsTbody) return [];

  const out = [];
  for (const row of readRows(bondsTbody)) {
    const matchBond = normalizeBondKey(row.bond);
    if (!matchBond) continue;
    if (onlyKeys && !onlyKeys.has(matchBond)) continue;

    const startDate = normalizeYMD(row.startDate);
    const endDate = normalizeYMD(row.endDate);
    const payoutMonths = Array.from(new Set(parseMonthList(row.payoutMonths))).sort((a, b) => a - b);
    const coupon = parseNumber(row.coupon);
    const cleanPrice = parseNumber(row.bondPrice);

    if (!startDate || !endDate || compareYmd(startDate, endDate) > 0) continue;
    if (!payoutMonths.length) continue;
    if (!Number.isFinite(coupon) || coupon <= 0) continue;
    if (requireCleanPrice && (!Number.isFinite(cleanPrice) || cleanPrice <= 0)) continue;

    out.push({
      matchBond,
      coupon,
      payoutMonths,
      startDate,
      endDate,
      cleanPrice,
    });
  }
  return out;
}

function readAutoPlanBondConfigs(selectedBondKeys) {
  const sel = new Set((selectedBondKeys || []).map((k) => normalizeBondKey(k)));
  if (!sel.size) return [];
  return readBondConfigsForAutoPlanFromTable({ onlyKeys: sel, requireCleanPrice: true });
}

/** Все облигации с валидным графиком купонов — для оценки реинвеста (в т.ч. вне списка отмеченных для покупки). */
function readBondConfigsForCouponSimulationFromTable() {
  return readBondConfigsForAutoPlanFromTable({ onlyKeys: null, requireCleanPrice: false });
}

/** Минимальные поля строки облигации для графика выплат (без НКД). */
function autoPlanBondScheduleShape(bondConfig) {
  return {
    startDate: bondConfig.startDate,
    endDate: bondConfig.endDate,
    payoutMonths: bondConfig.payoutMonths,
  };
}

/** Ежемесячные выплаты в модели графика: ровно 12 разных месяцев в году. Для таких бумаг эвристика «дальше до купона» не применяется. */
function isAutoPlanMonthlyCouponBond(bondConfig) {
  const months = bondConfig.payoutMonths;
  if (!Array.isArray(months) || months.length !== 12) return false;
  const set = new Set(months);
  if (set.size !== 12) return false;
  for (const m of months) {
    if (!Number.isInteger(m) || m < 1 || m > 12) return false;
  }
  return true;
}

/**
 * Доля купонного периода от последней опорной даты до следующей выплаты, ещё не прошедшая к дате покупки: 0…1.
 * Опорная дата — последняя выплата ≤ даты покупки, либо дата начала обращения, если выплат ещё не было.
 * Значение близко к 1 сразу после выплаты купона и в течение следующих месяцев (до следующей выплаты далеко);
 * падает к дате купона. Так можно докупать бумагу месяцами после выплаты, пока она по скору выгодна.
 */
function couponCycleFavorableFraction(bondConfig, purchaseYmd) {
  const row = autoPlanBondScheduleShape(bondConfig);
  const payDates = listBondPaymentYmdsSorted(row);
  const purchaseTs = ymdToUTCms(normalizeYMD(purchaseYmd) || "");
  const startTs = ymdToUTCms(row.startDate);
  if (!payDates.length || !Number.isFinite(purchaseTs) || !Number.isFinite(startTs)) return null;

  let nextTs = null;
  for (const d of payDates) {
    const ts = ymdToUTCms(d);
    if (ts > purchaseTs) {
      nextTs = ts;
      break;
    }
  }
  if (nextTs === null) return null;

  let prevTs = startTs;
  for (const d of payDates) {
    const ts = ymdToUTCms(d);
    if (ts <= purchaseTs) prevTs = ts;
  }

  const periodMs = nextTs - prevTs;
  const remainingMs = nextTs - purchaseTs;
  if (periodMs <= 0 || remainingMs <= 0) return null;
  const f = remainingMs / periodMs;
  if (!Number.isFinite(f)) return null;
  return Math.min(1, Math.max(0, f));
}

function buildAutoPlanScheduleDates(periodStartYmd, periodEndYmd, dayNums, minScheduleYmd) {
  const a = normalizeYMD(periodStartYmd);
  const b = normalizeYMD(periodEndYmd);
  const minY = normalizeYMD(minScheduleYmd) || getTodayYMD();
  if (!a || !b || compareYmd(a, b) > 0) return [];

  const start = compareYmd(a, minY) >= 0 ? a : minY;
  if (compareYmd(start, b) > 0) return [];

  const monthRange = buildMonthRange(toMonthKeyFromYMD(start), toMonthKeyFromYMD(b));
  const uniqueDays = Array.from(new Set(dayNums))
    .filter((d) => Number.isInteger(d) && d >= 1 && d <= 31)
    .sort((a, c) => a - c);

  const dates = [];
  for (const mk of monthRange) {
    const m = /^(\d{4})-(\d{2})$/.exec(mk);
    if (!m) continue;
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const dim = daysInCalendarMonthUTC(y, mo);
    for (const d of uniqueDays) {
      const dd = Math.min(d, dim);
      const ymd = `${y}-${String(mo).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
      if (compareYmd(ymd, start) < 0) continue;
      if (compareYmd(ymd, b) > 0) continue;
      dates.push(ymd);
    }
  }

  /** По возрастанию даты: симуляция и реинвест идут вперёд во времени, ранние покупки попадают в стакан событий раньше. */
  const sorted = [...new Set(dates)].sort((x, y) => ymdToUTCms(x) - ymdToUTCms(y));
  if (sorted.length > AUTO_PLAN_MAX_SCHEDULE_DATES) {
    return sorted.slice(0, AUTO_PLAN_MAX_SCHEDULE_DATES);
  }
  return sorted;
}

function buildInitialBuyEventsFromHoldings() {
  const holdings = parseHoldingItems(readRows(holdingsTbody));
  const bonds = sanitizeBondRows(readRows(bondsTbody));
  const startByKey = new Map();
  for (const r of bonds) {
    const k = normalizeBondKey(r.bond);
    const sd = normalizeYMD(r.startDate) || String(r.startDate || "").trim();
    if (k && sd) startByKey.set(k, sd);
  }
  return holdings.map((h) => {
    const k = normalizeBondKey(h.bond);
    let dateYMD = startByKey.get(k) || getTodayYMD();
    dateYMD = normalizeYMD(dateYMD) || dateYMD;
    let dateUTCms = ymdToUTCms(dateYMD);
    if (dateUTCms == null || !Number.isFinite(dateUTCms)) {
      dateYMD = getTodayYMD();
      dateUTCms = ymdToUTCms(dateYMD) ?? 0;
    }
    return { dateYMD, dateUTCms, bond: k, quantity: h.quantity };
  });
}

function quantityHeldAtPaymentTs(bondKey, paymentTs, buyEvents) {
  const k = normalizeBondKey(bondKey);
  return buyEvents
    .filter((ev) => normalizeBondKey(ev.bond) === k && ev.dateUTCms <= paymentTs)
    .reduce((s, ev) => s + ev.quantity, 0);
}

/** Валовые купоны за календарный месяц (выплата 1-го), в руб. */
function grossCouponsForCalendarMonth(monthKey, bondConfigs, buyEvents) {
  if (!monthKey) return 0;
  const m = /^(\d{4})-(\d{2})$/.exec(monthKey);
  if (!m) return 0;
  const monthNum = Number(m[2]);
  const paymentYmd = `${monthKey}-01`;
  const paymentTs = ymdToUTCms(paymentYmd);
  if (!Number.isFinite(paymentTs)) return 0;

  let sum = 0;
  for (const br of bondConfigs) {
    if (!br.payoutMonths.includes(monthNum)) continue;
    const startMonthKey = toMonthKeyFromYMD(br.startDate);
    const endMonthKey = toMonthKeyFromYMD(br.endDate);
    if (startMonthKey && monthKeyToUTCms(monthKey) < monthKeyToUTCms(startMonthKey)) continue;
    if (endMonthKey && monthKeyToUTCms(monthKey) > monthKeyToUTCms(endMonthKey)) continue;
    const q = quantityHeldAtPaymentTs(br.matchBond, paymentTs, buyEvents);
    sum += q * br.coupon;
  }
  return sum;
}

/**
 * Распределение бюджета по чистой цене из таблицы (НКД не оценивается).
 * На дату покупки жадно выбираются наиболее выгодные бумаги; не требуется покупать каждую отмеченную.
 * В режиме diversification=true используется убывающий скор по уже набранному количеству,
 * чтобы распределять бюджет между несколькими бумагами (но не обязательно всеми).
 * Скор: (купон×число выплат в год) / чистая цена; для не-ежемесячных — небольшой бонус в «ранней» фазе периода
 * (после выплаты купона можно докупать ещё долго, если базовая доходность лучше других).
 * Для 12 выплат в год бонус по фазе периода не применяется.
 * Учитывается горизонт: бумаги с большим числом будущих купонов получают до AUTO_PLAN_HORIZON_WEIGHT доп. к скору
 * (равная номинальная доходность, но дольше получать купоны — выгоднее удерживать в портфеле).
 * НКД в жадном шаге не вычитается из бюджета (цена сделки в таблице — чистая); оценка эффективности по купонному потоку.
 */
function allocateAutoPlanBudget(budgetRub, bondConfigs, purchaseYmd, diversification = false) {
  const prelim = bondConfigs
    .map((b) => {
      const row = autoPlanBondScheduleShape(b);
      if (!bondHasFutureCouponPaymentAfter(row, purchaseYmd)) return null;
      const clean = b.cleanPrice;
      if (!Number.isFinite(clean) || clean <= 0) return null;
      const ppy = b.payoutMonths.length;
      const baseYield = (b.coupon * ppy) / clean;
      if (!(baseYield > 0)) return null;
      const monthly = isAutoPlanMonthlyCouponBond(b);
      const cycleFrac = monthly ? null : couponCycleFavorableFraction(b, purchaseYmd);
      const futurePays = countFutureCouponPaymentsAfterPurchase(row, purchaseYmd);
      return { bond: b.matchBond, clean, baseYield, cycleFrac, monthly, futurePays };
    })
    .filter(Boolean);

  if (!prelim.length || !Number.isFinite(budgetRub) || budgetRub < 0.005) return [];

  const entries = prelim.map((p) => {
    let yieldScore = p.baseYield;
    if (!p.monthly && p.cycleFrac !== null && Number.isFinite(p.cycleFrac)) {
      yieldScore = p.baseYield * (1 + AUTO_PLAN_COUPON_DISTANCE_WEIGHT * p.cycleFrac);
    }
    const h = Math.min(1, p.futurePays / Math.max(1, AUTO_PLAN_HORIZON_PAYMENTS_REF));
    yieldScore *= 1 + AUTO_PLAN_HORIZON_WEIGHT * h;
    return { bond: p.bond, clean: p.clean, yieldScore };
  });

  const qtyByBond = Object.fromEntries(entries.map((e) => [e.bond, 0]));
  let B = budgetRub;
  const minClean = Math.min(...entries.map((e) => e.clean));
  let guard = 0;
  const scoreEps = 1e-12;
  while (B + 1e-9 >= minClean && guard < 200000) {
    guard += 1;
    let best = null;
    let bestEff = -1;
    let bestClean = Infinity;
    let bestBondKey = "";
    for (const e of entries) {
      if (e.clean > B + 1e-9) continue;
      const pickedQty = qtyByBond[e.bond] || 0;
      const effectiveScore = diversification ? e.yieldScore / (pickedQty + 1) : e.yieldScore;
      if (effectiveScore > bestEff + scoreEps) {
        bestEff = effectiveScore;
        best = e;
        bestClean = e.clean;
        bestBondKey = e.bond;
      } else if (best && !diversification && Math.abs(effectiveScore - bestEff) <= scoreEps) {
        if (e.clean < bestClean - 1e-9 || (Math.abs(e.clean - bestClean) <= 1e-9 && e.bond < bestBondKey)) {
          best = e;
          bestClean = e.clean;
          bestBondKey = e.bond;
        }
      }
    }
    if (!best) break;
    qtyByBond[best.bond] += 1;
    B -= best.clean;
  }

  return entries
    .map((e) => ({
      bond: e.bond,
      price: roundRub2(e.clean),
      quantity: qtyByBond[e.bond] || 0,
    }))
    .filter((x) => x.quantity > 0);
}

function mergeAutoPlanBuyItems(a, b) {
  const map = new Map();
  const key = (it) => `${normalizeBondKey(it.bond)}\t${roundRub2(it.price)}`;
  for (const it of [...a, ...b]) {
    const k = key(it);
    map.set(k, (map.get(k) || 0) + it.quantity);
  }
  return [...map.entries()].map(([k, q]) => {
    const tab = k.indexOf("\t");
    const bond = tab >= 0 ? k.slice(0, tab) : k;
    const price = parseNumber(k.slice(tab + 1));
    return { bond: normalizeBondKey(bond), price, quantity: q };
  });
}

/** Число полных календарных месяцев от месяца начала плана до месяца даты покупки (включительно одна точка: тот же месяц → 0). */
function countPlanMonthsFromStartToPurchase(periodStartYmd, purchaseYmd) {
  const startMk = toMonthKeyFromYMD(normalizeYMD(periodStartYmd) || "");
  const purchaseMk = toMonthKeyFromYMD(normalizeYMD(purchaseYmd) || "");
  if (!startMk || !purchaseMk) return 0;
  if (monthKeyToUTCms(purchaseMk) < monthKeyToUTCms(startMk)) return 0;
  let n = 0;
  let cur = startMk;
  let guard = 0;
  while (cur !== purchaseMk && guard < 2400) {
    guard += 1;
    n += 1;
    cur = addMonthsToMonthKey(cur, 1);
  }
  return cur === purchaseMk ? n : 0;
}

/**
 * Клоны конфигов с ценой из таблицы, умноженной на (1 + pct/100)^n, n — месяцы от начала плана.
 * @param {number} monthlyDriftPct процент удорожания за месяц; ≤0 — без изменений
 */
function applyAutoPlanMonthlyPriceDriftToConfigs(bondConfigs, periodStartYmd, purchaseYmd, monthlyDriftPct) {
  if (!Array.isArray(bondConfigs) || !bondConfigs.length) return bondConfigs;
  if (!Number.isFinite(monthlyDriftPct) || monthlyDriftPct <= 0) return bondConfigs;
  const n = countPlanMonthsFromStartToPurchase(periodStartYmd, purchaseYmd);
  const factor = Math.pow(1 + monthlyDriftPct / 100, n);
  if (!Number.isFinite(factor) || factor <= 0) return bondConfigs;
  return bondConfigs.map((b) => ({ ...b, cleanPrice: roundRub2(b.cleanPrice * factor) }));
}

function parseAutoPlanMonthlyPriceDriftPct() {
  if (!autoPlanPriceDriftPctInput) return 0;
  const raw = String(autoPlanPriceDriftPctInput.value ?? "").trim();
  if (!raw) return 0;
  const v = parseNumber(raw);
  if (!Number.isFinite(v) || v < 0) return 0;
  return Math.min(99, v);
}

function persistAutoPlanMonthlyPriceDriftPct() {
  try {
    const v = String(autoPlanPriceDriftPctInput?.value ?? "").trim();
    if (v) localStorage.setItem(AUTO_PLAN_MONTHLY_PRICE_DRIFT_PCT_KEY, v);
    else localStorage.removeItem(AUTO_PLAN_MONTHLY_PRICE_DRIFT_PCT_KEY);
  } catch {
    /* ignore */
  }
}

function mergeBuyRowsByDateForAutoPlan(existingRows, newRows) {
  const byDate = new Map();
  for (const row of existingRows) {
    if (!isBuyRowComplete(row)) continue;
    const d = normalizeYMD(row.date);
    if (!d) continue;
    byDate.set(d, { date: d, items: row.items });
  }
  for (const row of newRows) {
    const d = normalizeYMD(row.date);
    if (!d) continue;
    const prev = byDate.get(d);
    if (!prev) {
      byDate.set(d, { date: d, items: row.items });
    } else {
      const merged = mergeAutoPlanBuyItems(parseBuyItems(prev.items), parseBuyItems(row.items));
      byDate.set(d, { date: d, items: JSON.stringify(merged) });
    }
  }
  return sortBuyRowsByDate(Array.from(byDate.values())).rows;
}

function generateAutoPlanBuyRows() {
  const periodStart = autoPlanStartInput?.value || "";
  const periodEnd = autoPlanEndInput?.value || "";
  const topup = parseNumber(autoPlanTopupAmountInput?.value ?? "");
  const dayNums = getSelectedAutoPlanDayNums();
  const reinvest = Boolean(autoPlanReinvestCheckbox?.checked);
  const diversification = Boolean(autoPlanDiversifyCheckbox?.checked);
  const selectedKeys = getSelectedAutoPlanBondKeys();
  const bondConfigs = readAutoPlanBondConfigs(selectedKeys);

  if (!normalizeYMD(periodStart) || !normalizeYMD(periodEnd)) {
    return { ok: false, message: "Укажите даты начала и окончания инвестиций." };
  }
  if (compareYmd(periodStart, periodEnd) > 0) {
    return { ok: false, message: "Дата начала не может быть позже даты окончания." };
  }
  if (!dayNums.length) {
    return { ok: false, message: "Выберите хотя бы один день месяца для пополнения." };
  }
  if (!Number.isFinite(topup) || topup <= 0) {
    return { ok: false, message: "Укажите положительную сумму пополнения на дату." };
  }
  if (!selectedKeys.length) {
    return { ok: false, message: "Отметьте хотя бы одну облигацию для автопланирования." };
  }
  if (!bondConfigs.length) {
    return {
      ok: false,
      message:
        "Нет пригодных облигаций: проверьте тикер, цену, купон, месяцы выплат и период обращения в таблице «Облигации».",
    };
  }
  if (selectedKeys.length !== bondConfigs.length) {
    const found = new Set(bondConfigs.map((b) => b.matchBond));
    const missing = selectedKeys.filter((k) => !found.has(normalizeBondKey(k)));
    return {
      ok: false,
      message: `Не удалось учесть облигации: ${missing.join(", ")}. Проверьте цену, купон и график выплат.`,
    };
  }

  const schedule = buildAutoPlanScheduleDates(periodStart, periodEnd, dayNums, getTodayYMD());
  if (!schedule.length) {
    return { ok: false, message: "В выбранном периоде нет дат покупок (учитывается сегодняшняя дата и дни месяца)." };
  }

  const strategyIdForSim = String(autoPlanStrategySelect?.value || "").trim();
  const plannedRowsForSim = strategyIdForSim ? getPlannedBuyRowsForAutoPlan(strategyIdForSim) : [];
  const bondConfigsForCoupons = readBondConfigsForCouponSimulationFromTable();

  const simulatedEvents = [
    ...buildInitialBuyEventsFromHoldings(),
    ...buildBuyEventsFromPlannedRows(plannedRowsForSim),
  ];
  const generatedRows = [];
  /** Купоны за прошлый месяц — один раз на календарный месяц покупок, не на каждую дату пополнения. */
  const reinvestFromPrevMonthUsedForPurchaseMonth = new Set();

  for (const dateYmd of schedule) {
    const purchaseMonthKey = toMonthKeyFromYMD(dateYmd);
    const prevMonthKey = addMonthsToMonthKey(purchaseMonthKey, -1);
    let budget = topup;
    const applyReinvest =
      reinvest &&
      prevMonthKey &&
      !reinvestFromPrevMonthUsedForPurchaseMonth.has(purchaseMonthKey);
    if (applyReinvest) {
      budget += grossCouponsForCalendarMonth(prevMonthKey, bondConfigsForCoupons, simulatedEvents);
    }
    budget = roundRub2(budget);

    const driftPct = parseAutoPlanMonthlyPriceDriftPct();
    const bondConfigsForDate = applyAutoPlanMonthlyPriceDriftToConfigs(bondConfigs, periodStart, dateYmd, driftPct);
    const items = allocateAutoPlanBudget(budget, bondConfigsForDate, dateYmd, diversification);
    if (!items.length) continue;

    if (applyReinvest) reinvestFromPrevMonthUsedForPurchaseMonth.add(purchaseMonthKey);

    generatedRows.push({ date: dateYmd, items: JSON.stringify(items) });

    const ts = ymdToUTCms(dateYmd) || 0;
    for (const it of items) {
      simulatedEvents.push({
        dateYMD: dateYmd,
        dateUTCms: ts,
        bond: it.bond,
        quantity: it.quantity,
      });
    }
  }

  if (!generatedRows.length) {
    return { ok: false, message: "На выбранных датах не удалось купить ни одного лота (сумма или цены слишком малы)." };
  }

  return { ok: true, rows: generatedRows };
}

function onAutoPlanSaveToStrategy() {
  const strategyId = String(autoPlanStrategySelect?.value || "").trim();
  if (!strategyId) {
    window.alert("Выберите стратегию для сохранения результата.");
    return;
  }

  const gen = generateAutoPlanBuyRows();
  if (!gen.ok) {
    window.alert(gen.message);
    return;
  }

  saveCurrentBuysToActiveStrategy();
  const existing = (buyRowsByStrategyId.get(strategyId) || []).filter(isBuyRowComplete);
  const merged = mergeBuyRowsByDateForAutoPlan(existing, gen.rows);

  buyRowsByStrategyId.set(strategyId, merged);
  activeBuyStrategyId = strategyId;

  writeRows(buysTbody, buyTpl, merged.length ? merged : defaultBuyRows());
  renderBuyStrategySelect();
  if (buyStrategySelect) buyStrategySelect.value = strategyId;
  if (summaryStrategySelect) summaryStrategySelect.value = strategyId;
  if (autoPlanStrategySelect) autoPlanStrategySelect.value = strategyId;

  syncBuySummaries();
  syncStaticTableEmptyStates();
  persistBuyStrategiesState();
  renderAll();
}

function initAutoPlanWidgetUi() {
  const wrap = document.getElementById("auto-plan-topup-days");
  if (wrap && wrap.getAttribute("data-auto-plan-days-init") !== "1") {
    wrap.setAttribute("data-auto-plan-days-init", "1");
    wrap.querySelectorAll(".planScheduleDayChip").forEach((btn) => {
      btn.addEventListener("click", () => {
        const on = btn.classList.toggle("planScheduleDayChip--selected");
        btn.setAttribute("aria-pressed", on ? "true" : "false");
      });
    });
  }

  if (autoPlanSaveStrategyBtn && autoPlanSaveStrategyBtn.dataset.autoPlanBound !== "1") {
    autoPlanSaveStrategyBtn.dataset.autoPlanBound = "1";
    autoPlanSaveStrategyBtn.addEventListener("click", onAutoPlanSaveToStrategy);
  }
  if (autoPlanPriceDriftPctInput && autoPlanPriceDriftPctInput.dataset.autoPlanPersistBound !== "1") {
    autoPlanPriceDriftPctInput.dataset.autoPlanPersistBound = "1";
    autoPlanPriceDriftPctInput.addEventListener("input", persistAutoPlanMonthlyPriceDriftPct);
    autoPlanPriceDriftPctInput.addEventListener("change", persistAutoPlanMonthlyPriceDriftPct);
  }
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
    const buyStrategiesRaw = localStorage.getItem(BUY_STRATEGIES_KEY);
    const activeBuyStrategyRaw = localStorage.getItem(ACTIVE_BUY_STRATEGY_KEY);
    const bonds = bondsRaw ? sanitizeBondRows(JSON.parse(bondsRaw)) : defaultBondRows();
    const buysLegacy = buysRaw ? sortBuyRowsByDate(normalizeBuyRowsForUI(JSON.parse(buysRaw))).rows : defaultBuyRows();
    const holdings = holdingsRaw ? normalizeHoldingRowsForUI(JSON.parse(holdingsRaw)) : defaultHoldingRows();
    const taxRaw = localStorage.getItem(TAX_RATE_KEY);

    let nextStrategies = [];
    let nextRowsByStrategy = new Map();
    if (buyStrategiesRaw) {
      try {
        const parsed = JSON.parse(buyStrategiesRaw);
        const parsedStrategies = Array.isArray(parsed?.strategies) ? parsed.strategies : [];
        parsedStrategies.forEach((s) => {
          const id = String(s?.id || "").trim();
          const name = normalizeStrategyName(s?.name || "");
          if (!id || !name) return;
          nextStrategies.push({ id, name });
        });
        const parsedByStrategy = parsed?.buysByStrategy && typeof parsed.buysByStrategy === "object" ? parsed.buysByStrategy : {};
        nextStrategies.forEach((s) => {
          const rowsRaw = parsedByStrategy[s.id];
          const rows = Array.isArray(rowsRaw) ? sortBuyRowsByDate(normalizeBuyRowsForUI(rowsRaw)).rows : [];
          nextRowsByStrategy.set(s.id, rows);
        });
      } catch {
        nextStrategies = [];
        nextRowsByStrategy = new Map();
      }
    }

    if (!nextStrategies.length) {
      const base = getDefaultBuyStrategy();
      nextStrategies = [base];
      nextRowsByStrategy.set(base.id, buysLegacy);
    }

    buyStrategies = nextStrategies;
    buyRowsByStrategyId = nextRowsByStrategy;
    const hasStoredActive = buyStrategies.some((s) => s.id === String(activeBuyStrategyRaw || "").trim());
    activeBuyStrategyId = hasStoredActive ? String(activeBuyStrategyRaw || "").trim() : buyStrategies[0].id;
    const activeBuys = getActiveStrategyBuysRows();

    writeRows(bondsTbody, bondTpl, Array.isArray(bonds) && bonds.length ? bonds : defaultBondRows());
    writeRows(buysTbody, buyTpl, Array.isArray(activeBuys) && activeBuys.length ? activeBuys : defaultBuyRows());
    writeRows(holdingsTbody, holdingTpl, Array.isArray(holdings) && holdings.length ? holdings : defaultHoldingRows());
    renderBuyStrategySelect();
    if (taxRateInput && taxRaw !== null) taxRateInput.value = String(parseNumber(taxRaw) || 13);
    if (portfolioStartDateInput) portfolioStartDateInput.value = normalizeYMD(portfolioStartDateRaw || "") || getTodayYMD();
    setPortfolioMoneyInputValue(portfolioStartValueInput, String(parseNumber(portfolioStartValueRaw) || 0));
    setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, String(parseNumber(portfolioMonthlyTopupRaw) || 0));
    if (portfolioMonthlyTopupEndDateInput) {
      portfolioMonthlyTopupEndDateInput.value = normalizeYMD(portfolioMonthlyTopupEndDateRaw || "") || "";
    }
    loadCalculatorStates();
    try {
      const driftRaw = localStorage.getItem(AUTO_PLAN_MONTHLY_PRICE_DRIFT_PCT_KEY);
      if (autoPlanPriceDriftPctInput && driftRaw !== null && driftRaw !== undefined) {
        autoPlanPriceDriftPctInput.value = String(driftRaw);
      }
    } catch {
      /* ignore */
    }
  } catch {
    writeRows(bondsTbody, bondTpl, defaultBondRows());
    writeRows(buysTbody, buyTpl, defaultBuyRows());
    writeRows(holdingsTbody, holdingTpl, defaultHoldingRows());
    buyStrategies = [getDefaultBuyStrategy()];
    buyRowsByStrategyId = new Map([[DEFAULT_BUY_STRATEGY_ID, []]]);
    activeBuyStrategyId = DEFAULT_BUY_STRATEGY_ID;
    renderBuyStrategySelect();
    if (taxRateInput) taxRateInput.value = "13";
    if (portfolioStartDateInput) portfolioStartDateInput.value = getTodayYMD();
    if (portfolioMonthlyTopupEndDateInput) portfolioMonthlyTopupEndDateInput.value = "";
    setPortfolioMoneyInputValue(portfolioStartValueInput, "0");
    setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, "0");
    loadCalculatorStates();
    if (autoPlanPriceDriftPctInput) autoPlanPriceDriftPctInput.value = "";
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

  const onFieldChange = (e) => {
    if (e.target && e.target.matches("[data-field]")) scheduleSave();
  };
  tbody.addEventListener("input", onFieldChange);
  // Для `select` событие `input` может не срабатывать, поэтому добавляем `change`.
  tbody.addEventListener("change", onFieldChange);
}

bindRowDelete(bondsTbody);
bindRowDelete(buysTbody);
bindRowDelete(holdingsTbody);

function normalizeBondKey(bond) {
  return String(bond || "").trim().toUpperCase();
}

function propagateBondRename(oldBondKey, newBondKey) {
  const oldKey = normalizeBondKey(oldBondKey);
  const nextKey = normalizeBondKey(newBondKey);
  if (!oldKey || !nextKey) return;
  if (oldKey === nextKey) return;

  // Обновляем текущий портфель
  if (holdingsTbody) {
    Array.from(holdingsTbody.querySelectorAll('[data-field="bond"]')).forEach((el) => {
      if (!(el instanceof HTMLInputElement || el instanceof HTMLSelectElement)) return;
      if (normalizeBondKey(el.value) !== oldKey) return;
      if (el instanceof HTMLSelectElement) {
        const hasOption = Array.from(el.options).some((o) => o.value === nextKey);
        if (!hasOption) {
          const opt = document.createElement("option");
          opt.value = nextKey;
          opt.textContent = nextKey;
          el.appendChild(opt);
        }
      }
      el.value = nextKey;
    });
  }

  // Обновляем покупки (bond внутри items JSON)
  if (buysTbody) {
    Array.from(buysTbody.querySelectorAll('input[data-field="items"]')).forEach((itemsInput) => {
      if (!(itemsInput instanceof HTMLInputElement)) return;
      const tr = itemsInput.closest("tr");
      if (!tr) return;

      const items = parseBuyItems(itemsInput.value);
      if (!items.length) return;

      let changed = false;
      const nextItems = items.map((i) => {
        if (normalizeBondKey(i.bond) !== oldKey) return i;
        changed = true;
        return { ...i, bond: nextKey };
      });

      if (!changed) return;
      itemsInput.value = JSON.stringify(nextItems);
    });

    // Обновляем отображение chips/totals в таблице покупок
    syncBuySummaries();
  }
}

// Чтобы переименование облигации в виджете "Облигации" отображалось везде,
// на лету обновляем ссылки в покупках и текущем портфеле.
bondsTbody?.addEventListener("focusin", (e) => {
  const el = e.target;
  if (!(el instanceof HTMLInputElement)) return;
  if (!el.matches('input[data-field="bond"]')) return;
  el.dataset.prevBondKey = normalizeBondKey(el.value);
});

bondsTbody?.addEventListener("input", (e) => {
  const el = e.target;
  if (!(el instanceof HTMLInputElement)) return;
  if (!el.matches('input[data-field="bond"]')) return;

  const prevKey = normalizeBondKey(el.dataset.prevBondKey || "");
  const nextKey = normalizeBondKey(el.value);

  // Не распространяем пустое имя, чтобы не потерять связь при частичном вводе.
  if (!nextKey) return;
  if (!prevKey) {
    el.dataset.prevBondKey = nextKey;
    return;
  }
  if (prevKey === nextKey) return;

  propagateBondRename(prevKey, nextKey);
  el.dataset.prevBondKey = nextKey;
});

function copyBuyRow(tr) {
  if (!tr || tr.parentElement !== buysTbody) return;
  if (!tr.classList.contains("buyTable__dataRow")) return;
  if (tr.hasAttribute("data-empty-state")) return;

  const dateInput = tr.querySelector('[data-field="date"]');
  if (!dateInput) return;

  // Открываем модалку в режиме "создать новую покупку",
  // но с данными из выбранной строки (чтобы пользователь мог изменить только дату).
  buyModalEditingTr = null;
  if (buyModalTitle) buyModalTitle.textContent = BUY_MODAL_TITLE_NEW;

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
  openBondModalNew();
});

if (bondsTbody) {
  bondsTbody.addEventListener("click", (e) => {
    const btnRemove = e.target.closest('[data-action="remove"]');
    if (btnRemove) return;
    if (e.target.closest('[data-action="open-month-picker"]')) return;
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement !== bondsTbody) return;
    if (tr.hasAttribute("data-empty-state")) return;
    openBondModalForEdit(tr);
  });
}

addBuyBtn.addEventListener("click", () => {
  openBuyModal();
});

if (buyStrategySelect) {
  buyStrategySelect.addEventListener("change", () => {
    switchActiveBuyStrategy(buyStrategySelect.value);
  });
}

if (buyTableYearFilterSelect) {
  buyTableYearFilterSelect.addEventListener("change", () => {
    applyBuyYearFilter();
  });
}

if (buyStrategyHeaderSelect) {
  buyStrategyHeaderSelect.addEventListener("change", () => {
    switchActiveBuyStrategy(buyStrategyHeaderSelect.value);
  });
}

if (summaryStrategySelect) {
  summaryStrategySelect.addEventListener("change", () => {
    switchActiveBuyStrategy(summaryStrategySelect.value);
  });
}

if (addHoldingBtn) {
  addHoldingBtn.addEventListener("click", () => {
    const availableBonds = getAvailableBondNames();
    if (!availableBonds.length) return;
    clearTableEmptyState(holdingsTbody);
    holdingsTbody.appendChild(holdingTpl.content.cloneNode(true));
    syncHoldingsBondSelects();
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
  localStorage.removeItem(BUY_STRATEGIES_KEY);
  localStorage.removeItem(ACTIVE_BUY_STRATEGY_KEY);
  localStorage.removeItem(HOLDINGS_KEY);
  localStorage.removeItem(TAX_RATE_KEY);
  localStorage.removeItem(ACTIVE_TAB_KEY);
  localStorage.removeItem(CHART_YEAR_KEY);
  localStorage.removeItem(SUMMARY_PIE_YEAR_KEY);
  localStorage.removeItem(PORTFOLIO_CHART_YEAR_KEY);
  localStorage.removeItem(PORTFOLIO_CHART_STRATEGY_KEY);
  localStorage.removeItem(PORTFOLIO_START_DATE_KEY);
  localStorage.removeItem(PORTFOLIO_START_VALUE_KEY);
  localStorage.removeItem(PORTFOLIO_MONTHLY_TOPUP_KEY);
  localStorage.removeItem(PORTFOLIO_MONTHLY_TOPUP_END_DATE_KEY);
  localStorage.removeItem(ACCRUED_CALC_KEY);
  localStorage.removeItem(YIELD_CALC_KEY);
  localStorage.removeItem(AUTO_PLAN_MONTHLY_PRICE_DRIFT_PCT_KEY);
  localStorage.removeItem(STRATEGY_CENTER_VIEW_KEY);
  if (taxRateInput) taxRateInput.value = "13";
  if (portfolioStartDateInput) portfolioStartDateInput.value = getTodayYMD();
  setPortfolioMoneyInputValue(portfolioStartValueInput, "0");
  setPortfolioMoneyInputValue(portfolioMonthlyTopupInput, "0");
  if (autoPlanPriceDriftPctInput) autoPlanPriceDriftPctInput.value = "";
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

if (portfolioChartStrategySelect) {
  portfolioChartStrategySelect.addEventListener("change", () => {
    localStorage.setItem(PORTFOLIO_CHART_STRATEGY_KEY, portfolioChartStrategySelect.value || "");
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
if (tabStrategyBtn) {
  tabStrategyBtn.addEventListener("click", () => setActiveTab("strategy"));
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

{
  const savedTab = localStorage.getItem(ACTIVE_TAB_KEY);
  setActiveTab(savedTab === "profile" ? "strategy" : savedTab || "planning");
}

initThemeUi();
bindChartTooltips();
updateSortableHeaders();
loadAll();
initAutoPlanWidgetUi();
initStrategyTabPanel();

function syncDateSummaries() {
  Array.from(bondsTbody.querySelectorAll("tr")).forEach((tr) => {
    const monthsHidden = tr.querySelector('[data-field="payoutMonths"]');
    const bondHidden = tr.querySelector('[data-field="bond"]');
    const couponHidden = tr.querySelector('[data-field="coupon"]');
    const startHidden = tr.querySelector('[data-field="startDate"]');
    const endHidden = tr.querySelector('[data-field="endDate"]');
    const monthsBtn = tr.querySelector('[data-action="open-month-picker"]');
    const bondDisplay = tr.querySelector('[data-display-for="bond"]');
    const couponDisplay = tr.querySelector('[data-display-for="coupon"]');
    const payoutCountDisplay = tr.querySelector('[data-display-for="payoutCount"]');
    const payoutRangeDisplay = tr.querySelector('[data-display-for="payoutRange"]');
    const bondPriceHidden = tr.querySelector('[data-field="bondPrice"]');
    const bondPriceDisplay = tr.querySelector('[data-display-for="bondPrice"]');
    if (!monthsHidden || !startHidden || !endHidden) return;

    const selected = parseMonthList(monthsHidden.value);
    const normalized = Array.from(new Set(selected)).sort((a, b) => a - b);
    monthsHidden.value = normalized.join(",");
    startHidden.value = normalizeYMD(startHidden.value) || "";
    endHidden.value = normalizeYMD(endHidden.value) || "";

    if (bondDisplay && bondHidden) {
      const v = String(bondHidden.value || "").trim();
      bondDisplay.textContent = v ? normalizeBondKey(v) : "—";
    }

    if (couponDisplay && couponHidden) {
      const coupon = parseNumber(couponHidden.value);
      couponDisplay.textContent = Number.isFinite(coupon) ? formatMoney(coupon) : "—";
    }

    if (bondPriceDisplay && bondPriceHidden) {
      const bp = parseNumber(bondPriceHidden.value);
      bondPriceDisplay.textContent = Number.isFinite(bp) && bp > 0 ? formatMoney(bp) : "—";
    }

    if (payoutCountDisplay) {
      const n = normalized.length;
      if (!n) {
        payoutCountDisplay.textContent = "—";
      } else {
        const text = n === 1 ? "выплата" : n >= 2 && n <= 4 ? "выплаты" : "выплат";
        payoutCountDisplay.textContent = `${n} ${text} в год`;
      }
    }

    if (payoutRangeDisplay) {
      if (startHidden.value && endHidden.value) {
        payoutRangeDisplay.textContent = `${formatMonthYearFromYMD(startHidden.value)} - ${formatMonthYearFromYMD(endHidden.value)}`;
      } else {
        payoutRangeDisplay.textContent = "—";
      }
    }

    if (monthsBtn) {
      if (normalized.length && startHidden.value && endHidden.value) {
        monthsBtn.textContent = `${formatMonthYearFromYMD(startHidden.value)} - ${formatMonthYearFromYMD(endHidden.value)}`;
      } else {
        monthsBtn.textContent = "Выбрать даты";
      }
    }
  });
}

function syncHoldingsBondSelects() {
  if (!holdingsTbody) return;
  const availableBonds = getAvailableBondNames();
  const availableSet = new Set(availableBonds);

  Array.from(holdingsTbody.querySelectorAll('select[data-field="bond"]')).forEach((select) => {
    const current = normalizeBondKey(select.value);

    select.innerHTML = "";
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "—";
    select.appendChild(placeholder);

    for (const bond of availableBonds) {
      const opt = document.createElement("option");
      opt.value = bond;
      opt.textContent = bond;
      select.appendChild(opt);
    }

    // Если у пользователя уже сохранена облигация, которой сейчас нет в таблице,
    // добавляем её как fallback, чтобы не потерять сохранённое значение.
    if (current && !availableSet.has(current)) {
      const opt = document.createElement("option");
      opt.value = current;
      opt.textContent = current;
      select.appendChild(opt);
    }

    select.value = current && (availableSet.has(current) || Array.from(select.options).some((o) => o.value === current)) ? current : "";
  });
}

function syncBuySummaries() {
  Array.from(buysTbody.querySelectorAll("tr.buyTable__dataRow")).forEach((tr) => {
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

    itemsDisplay.innerHTML = buildBuyChipsHtmlFromItems(items);
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

// Поддерживает выбор месяцев и для строк таблицы, и для модальных форм
// monthPickerState = {
//   payoutMonthsInput: HTMLInputElement,
//   startDateInput: HTMLInputElement,
//   endDateInput: HTMLInputElement,
//   months:number[],
//   startDate:string,
//   endDate:string,
//   persistOnDone:boolean,
//   onApply: (() => void) | null
// }
let monthPickerState = null;

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
  // Сразу сохраняем выбранные месяцы в hidden-поле модалки облигации.
  applyMonthsToRow();
  if (typeof monthPickerState.onApply === "function") monthPickerState.onApply();
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
    const m = i;
    const isActive = Array.isArray(months) && months.includes(m);

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `monthBtn${isActive ? " is-active" : ""}`;
    btn.textContent = MONTHS_RU[m - 1];

    btn.addEventListener("click", () => {
      if (!monthPickerState) return;
      const current = monthPickerState.months;
      if (current.includes(m)) {
        monthPickerState.months = current.filter((x) => x !== m);
      } else {
        monthPickerState.months = [...current, m].sort((a, b) => a - b);
      }
      renderMonthGrid(monthPickerState.months);
      updateMonthModeUI();
      applyMonthsToRow();
      if (typeof monthPickerState.onApply === "function") monthPickerState.onApply();
    });

    monthGrid.appendChild(btn);
  }
}

function applyMonthsToRow() {
  if (!monthPickerState) return;
  const { payoutMonthsInput } = monthPickerState;
  if (!payoutMonthsInput) return;
  payoutMonthsInput.value = Array.from(new Set(monthPickerState.months)).sort((a, b) => a - b).join(",");
}

bondsTbody.addEventListener("click", (e) => {
  const btn = e.target.closest('[data-action="open-month-picker"]');
  if (!btn) return;
  const tr = btn.closest("tr");
  if (!tr) return;
  const hidden = tr.querySelector('[data-field="payoutMonths"]');
  if (!hidden) return;

  monthPickerState = {
    payoutMonthsInput: hidden,
    months: parseMonthList(hidden.value),
    persistOnDone: true,
    onApply: () => {
      syncDateSummaries();
    },
  };
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
      applyMonthsToRow();
      if (typeof monthPickerState.onApply === "function") monthPickerState.onApply();
      if (monthPickerState.persistOnDone) persistAndRender();
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

let bondModalEditingTr = null;
let bondModalPrevBondKey = "";
/** @type {string | null} */
let strategyActionsTargetId = null;

function setBondModalOpen(open) {
  if (!bondModalOverlay) return;
  if (open) openModalOverlay(bondModalOverlay);
  else closeModalOverlay(bondModalOverlay);
}

function syncBondModalMonthSummary() {
  // Не используется: месяц-выбор теперь сделан через multi-select.
  return;
}

function initBondMonthPickerState() {
  if (!bondModalPayoutMonthsInput) return;

  let months = parseMonthList(bondModalPayoutMonthsInput.value);

  monthPickerState = {
    payoutMonthsInput: bondModalPayoutMonthsInput,
    months,
    persistOnDone: false,
    onApply: null,
  };

  if (bondMonthPickerPanel) bondMonthPickerPanel.removeAttribute("hidden");
  renderMonthGrid(monthPickerState.months);
  updateMonthModeUI();
  applyMonthsToRow(); // фиксируем нормализованное значение в hidden
}

function syncBondModalPayoutMonthsInputFromSelect() {
  if (!bondModalPayoutMonthsInput || !bondModalPayoutMonthsSelect) return;

  const months = Array.from(bondModalPayoutMonthsSelect.selectedOptions)
    .map((o) => Number(o.value))
    .filter((n) => Number.isInteger(n) && n >= 1 && n <= 12);

  const uniqueSorted = Array.from(new Set(months)).sort((a, b) => a - b);
  bondModalPayoutMonthsInput.value = uniqueSorted.join(",");
}

function openBondMonthPicker() {
  if (!bondModalPayoutMonthsInput || !bondMonthPickerPanel) return;

  if (!monthPickerState) initBondMonthPickerState();

  // Показываем/прячем панель выбора внутри bond-модалки.
  const nextOpen = bondMonthPickerPanel.hasAttribute("hidden");
  if (nextOpen) bondMonthPickerPanel.removeAttribute("hidden");
  else bondMonthPickerPanel.setAttribute("hidden", "");

  // На всякий случай обновляем UI по текущим выбранным месяцам.
  renderMonthGrid(monthPickerState.months);
  updateMonthModeUI();
  syncBondModalMonthSummary();
}

function fillBondModalFormFromRow(tr) {
  if (!tr) return;
  const bondInput = tr.querySelector('input[data-field="bond"]');
  const couponInput = tr.querySelector('input[data-field="coupon"]');
  const payoutMonthsInput = tr.querySelector('input[data-field="payoutMonths"]');
  const startHidden = tr.querySelector('input[data-field="startDate"]');
  const endHidden = tr.querySelector('input[data-field="endDate"]');
  if (!bondInput || !couponInput || !payoutMonthsInput || !startHidden || !endHidden) return;

  bondModalNameInput.value = String(bondInput.value || "").trim();
  // Нормализуем купон на случай значений вида "36,5" из старых сохранений.
  setInputNumericValue(bondModalCouponInput, parseNumber(couponInput.value), "");
  const priceHidden = tr.querySelector('input[data-field="bondPrice"]');
  if (bondModalBondPriceInput) setInputNumericValue(bondModalBondPriceInput, parseNumber(priceHidden?.value || ""), "");
  bondModalPayoutMonthsInput.value = String(payoutMonthsInput.value || "").trim();
  if (bondModalStartDateVisibleInput) bondModalStartDateVisibleInput.value = String(startHidden.value || "").trim();
  if (bondModalEndDateVisibleInput) bondModalEndDateVisibleInput.value = String(endHidden.value || "").trim();
  if (bondModalStartDateInput) bondModalStartDateInput.value = String(startHidden.value || "").trim();
  if (bondModalEndDateInput) bondModalEndDateInput.value = String(endHidden.value || "").trim();

  syncBondModalMonthSummary();
}

function getBondModalRowData() {
  const bond = normalizeBondKey(bondModalNameInput?.value || "");
  const coupon = parseNumber(bondModalCouponInput?.value || "");
  const payoutMonths = String(bondModalPayoutMonthsInput?.value || "").trim();
  const startDate = normalizeYMD(bondModalStartDateVisibleInput?.value || "") || "";
  const endDate = normalizeYMD(bondModalEndDateVisibleInput?.value || "") || "";
  const bondPriceRaw = parseNumber(bondModalBondPriceInput?.value || "");
  const bondPrice = Number.isFinite(bondPriceRaw) && bondPriceRaw > 0 ? String(bondPriceRaw) : "";

  return { bond, coupon, payoutMonths, startDate, endDate, bondPrice };
}

function openBondModalNew() {
  bondModalEditingTr = null;
  bondModalPrevBondKey = "";
  if (bondModalDeleteBtn) bondModalDeleteBtn.setAttribute("hidden", "");
  if (bondModalTitle) bondModalTitle.textContent = "Новая облигация";
  if (bondModalNameInput) bondModalNameInput.value = "";
  if (bondModalCouponInput) bondModalCouponInput.value = "";
  if (bondModalBondPriceInput) bondModalBondPriceInput.value = "";
  if (bondModalPayoutMonthsInput) bondModalPayoutMonthsInput.value = "";
  if (bondModalStartDateVisibleInput) bondModalStartDateVisibleInput.value = "";
  if (bondModalEndDateVisibleInput) bondModalEndDateVisibleInput.value = "";
  if (bondModalStartDateInput) bondModalStartDateInput.value = "";
  if (bondModalEndDateInput) bondModalEndDateInput.value = "";
  initBondMonthPickerState();
  setBondModalOpen(true);
  if (bondModalNameInput) {
    requestAnimationFrame(() => requestAnimationFrame(() => bondModalNameInput.focus()));
  }
}

function openBondModalForEdit(tr) {
  if (!tr || tr.closest("tbody") !== bondsTbody) return;
  bondModalEditingTr = tr;
  const bondInput = tr.querySelector('input[data-field="bond"]');
  bondModalPrevBondKey = normalizeBondKey(bondInput?.value || "");
  if (bondModalDeleteBtn) bondModalDeleteBtn.removeAttribute("hidden");
  if (bondModalTitle) bondModalTitle.textContent = "Редактирование облигации";
  fillBondModalFormFromRow(tr);
  initBondMonthPickerState();
  setBondModalOpen(true);
  if (bondModalNameInput) {
    requestAnimationFrame(() => requestAnimationFrame(() => bondModalNameInput.focus()));
  }
}

function closeBondModal() {
  bondModalEditingTr = null;
  bondModalPrevBondKey = "";
  monthPickerState = null;
  if (bondMonthPickerPanel) bondMonthPickerPanel.setAttribute("hidden", "");
  if (bondModalDeleteBtn) bondModalDeleteBtn.setAttribute("hidden", "");
  setBondModalOpen(false);
}

if (bondModalClose) {
  bondModalClose.addEventListener("click", () => closeBondModal());
}

if (bondModalOverlay) {
  bondModalOverlay.addEventListener("click", (e) => {
    if (e.target === bondModalOverlay) closeBondModal();
  });
}

if (bondModalOpenMonthPickerBtn) {
  bondModalOpenMonthPickerBtn.addEventListener("click", () => {
    // старый UI не используется после перехода на multi-select
    initBondMonthPickerState();
  });
}

// select[multiple] больше не используется; синхронизация идёт через клики по чипам.

if (bondModalSaveBtn) {
  bondModalSaveBtn.addEventListener("click", () => {
    if (!bondModalNameInput || !bondModalCouponInput) return;
    const row = getBondModalRowData();
    if (!row.bond || !Number.isFinite(row.coupon) || row.coupon <= 0) return;

    const months = parseMonthList(row.payoutMonths);
    if (!months.length) return;
    if (!row.startDate || !row.endDate) return;
    const startTs = ymdToUTCms(row.startDate);
    const endTs = ymdToUTCms(row.endDate);
    if (!Number.isFinite(startTs) || !Number.isFinite(endTs) || endTs < startTs) return;

    if (bondModalEditingTr) {
      const tr = bondModalEditingTr;
      const bondInput = tr.querySelector('input[data-field="bond"]');
      const couponInput = tr.querySelector('input[data-field="coupon"]');
      const payoutMonthsHidden = tr.querySelector('input[data-field="payoutMonths"]');
      const startHidden = tr.querySelector('input[data-field="startDate"]');
      const endHidden = tr.querySelector('input[data-field="endDate"]');
      const bondPriceHidden = tr.querySelector('input[data-field="bondPrice"]');
      if (!bondInput || !couponInput || !payoutMonthsHidden || !startHidden || !endHidden || !bondPriceHidden) return;

      const newKey = normalizeBondKey(row.bond);
      const oldKey = bondModalPrevBondKey;

      bondInput.value = row.bond;
      couponInput.value = String(row.coupon);
      payoutMonthsHidden.value = months.join(",");
      startHidden.value = row.startDate;
      endHidden.value = row.endDate;
      bondPriceHidden.value = row.bondPrice || "";

      syncDateSummaries();

      if (oldKey && newKey && oldKey !== newKey) {
        propagateBondRename(oldKey, newKey);
      }

      persistAndRender();
      closeBondModal();
      return;
    }

    // create new row
    clearTableEmptyState(bondsTbody);
    const node = bondTpl.content.cloneNode(true);
    const tr = node.querySelector("tr");
    if (!tr) return;

    const bondInput = tr.querySelector('input[data-field="bond"]');
    const couponInput = tr.querySelector('input[data-field="coupon"]');
    const payoutMonthsHidden = tr.querySelector('input[data-field="payoutMonths"]');
    const startHidden = tr.querySelector('input[data-field="startDate"]');
    const endHidden = tr.querySelector('input[data-field="endDate"]');
    const bondPriceHidden = tr.querySelector('input[data-field="bondPrice"]');
    if (!bondInput || !couponInput || !payoutMonthsHidden || !startHidden || !endHidden || !bondPriceHidden) return;

    bondInput.value = row.bond;
    couponInput.value = String(row.coupon);
    payoutMonthsHidden.value = months.join(",");
    startHidden.value = row.startDate;
    endHidden.value = row.endDate;
    bondPriceHidden.value = row.bondPrice || "";

    bondsTbody.appendChild(node);
    syncDateSummaries();
    persistAndRender();
    closeBondModal();
  });
}

if (bondModalDeleteBtn) {
  bondModalDeleteBtn.addEventListener("click", () => {
    if (!bondModalEditingTr) return;
    const bondInput = bondModalEditingTr.querySelector('input[data-field="bond"]');
    const name = bondInput ? String(bondInput.value || "").trim() : "";
    const label = name ? `«${name}»` : "эту облигацию";
    const ok = window.confirm(`Удалить облигацию ${label} из справочника? Покупки и позиции с этим тикером могут потребовать правки.`);
    if (!ok) return;
    const key = bondModalPrevBondKey || normalizeBondKey(name);
    removeBondRowByBondKey(key);
    closeBondModal();
  });
}

if (strategyActionsClose) {
  strategyActionsClose.addEventListener("click", () => closeStrategyActionsModal());
}
if (strategyActionsOverlay) {
  strategyActionsOverlay.addEventListener("click", (e) => {
    if (e.target === strategyActionsOverlay) closeStrategyActionsModal();
  });
}
if (strategyActionsRenameBtn) {
  strategyActionsRenameBtn.addEventListener("click", () => {
    const id = strategyActionsTargetId;
    closeStrategyActionsModal();
    if (id) openStrategyModalEditForId(id);
  });
}
if (strategyActionsDeleteBtn) {
  strategyActionsDeleteBtn.addEventListener("click", () => {
    const id = strategyActionsTargetId;
    if (!id) return;
    const s = buyStrategies.find((x) => x.id === id);
    const label = s?.name ? `«${String(s.name)}»` : "эту стратегию";
    const ok = window.confirm(`Удалить стратегию ${label} и все запланированные в ней покупки?`);
    if (!ok) return;
    deleteBuyStrategyById(id);
    closeStrategyActionsModal();
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
  const cppResolved = resolveYieldCouponPerPaymentRub(row);
  setInputNumericValue(
    yieldModalCouponPerPaymentInput,
    Number.isFinite(cppResolved) ? cppResolved : parseNumber(row.couponPerPayment),
    "30"
  );
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
    couponPerPayment: String(parseNumber(yieldModalCouponPerPaymentInput?.value || "") || 0),
    accruedIncome: String(parseNumber(yieldModalAccruedIncomeInput?.value || "") || 0),
    redemptionPrice: String(parseNumber(yieldModalRedemptionPriceInput?.value || "") || 0),
  };
}

function openYieldModal() {
  yieldModalEditingTr = null;
  if (yieldModalTitle) yieldModalTitle.textContent = YIELD_MODAL_TITLE_NEW;
  fillYieldModalForm(getDefaultYieldCalcRow());
  const today = getTodayYMD();
  if (yieldModalPurchaseDateInput) yieldModalPurchaseDateInput.value = today;
  if (yieldModalMaturityDateInput) {
    yieldModalMaturityDateInput.value = addDaysToYmd(today, 365) || "";
  }
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

function newBuyModalItem() {
  return { bond: "", price: "", quantity: "", priceUserEdited: false };
}

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
      priceUserEdited: Boolean(prev && prev.priceUserEdited),
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

  for (let i = 0; i < next.length; i++) {
    if (!String(next[i].price || "").trim()) next[i].priceUserEdited = false;
  }

  buyModalItems.length = 0;
  buyModalItems.push(...next);
}

function buyRowItemsToModalState(tr) {
  const itemsInput = tr.querySelector('[data-field="items"]');
  const parsed = parseBuyItems(itemsInput?.value);
  if (!parsed.length) return [newBuyModalItem()];
  return parsed.map((i) => ({
    bond: i.bond,
    price: formatAmount(i.price),
    quantity: formatQuantityInt(i.quantity),
    // Сохранённые покупки не перезаписываем ценой из справочника при смене облигации.
    priceUserEdited: true,
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

function refreshAutoPlanBondPicker() {
  const el = document.getElementById("auto-plan-bonds");
  if (!el) return;
  const bonds = getAvailableBondNames();
  const sig = bonds.join("\0");
  if (sig === autoPlanBondPickerSig && el.querySelector('input[type="checkbox"][data-bond], .autoPlanBonds__empty')) {
    return;
  }

  const prev = new Map();
  el.querySelectorAll('input[type="checkbox"][data-bond]').forEach((inp) => {
    prev.set(String(inp.getAttribute("data-bond") || "").trim().toUpperCase(), inp.checked);
  });

  if (!bonds.length) {
    autoPlanBondPickerSig = "";
    el.innerHTML =
      '<span class="autoPlanBonds__empty">Добавьте облигации в таблицу выше — они появятся в этом списке.</span>';
    return;
  }

  el.innerHTML = bonds
    .map((bond, idx) => {
      const checked = prev.has(bond) ? prev.get(bond) : true;
      return `<label class="autoPlanBondPick"><input type="checkbox" id="auto-plan-bond-cb-${idx}" data-bond="${escapeHtml(
        bond
      )}" ${checked ? "checked" : ""} /><span>${escapeHtml(bond)}</span></label>`;
    })
    .join("");
  autoPlanBondPickerSig = sig;
}

/** Плановая цена из карточки облигации (для автоподстановки в покупках). */
function getBondReferencePrice(bondRaw) {
  const want = normalizeBondKey(bondRaw);
  if (!want || !bondsTbody) return NaN;
  for (const tr of bondsTbody.querySelectorAll("tr")) {
    const bondH = tr.querySelector('[data-field="bond"]');
    if (!bondH || normalizeBondKey(bondH.value) !== want) continue;
    const ph = tr.querySelector('[data-field="bondPrice"]');
    const p = parseNumber(ph?.value ?? "");
    return Number.isFinite(p) && p > 0 ? p : NaN;
  }
  return NaN;
}

function setBuyModalOpen(open) {
  if (!buyModalOverlay) return;
  if (open) openModalOverlay(buyModalOverlay);
  else closeModalOverlay(buyModalOverlay);
}

function setStrategyModalOpen(open) {
  if (!strategyModalOverlay) return;
  if (open) openModalOverlay(strategyModalOverlay);
  else closeModalOverlay(strategyModalOverlay);
}

function openStrategyModalCreate() {
  strategyModalEditingId = null;
  strategyModalDuplicatingFromId = null;
  if (strategyModalTitle) strategyModalTitle.textContent = "Новая стратегия";
  if (strategyNameInput) strategyNameInput.value = "";
  if (strategyDeleteBtn) strategyDeleteBtn.setAttribute("hidden", "");
  setStrategyModalOpen(true);
  if (strategyNameInput) {
    requestAnimationFrame(() => requestAnimationFrame(() => strategyNameInput.focus()));
  }
}

function openStrategyModalEditForId(strategyIdRaw) {
  const id = String(strategyIdRaw || "").trim();
  const current = buyStrategies.find((s) => s.id === id);
  if (!current || current.id === DEFAULT_BUY_STRATEGY_ID) return;

  strategyModalDuplicatingFromId = null;
  strategyModalEditingId = current.id;
  if (strategyModalTitle) strategyModalTitle.textContent = "Редактирование стратегии";
  if (strategyNameInput) strategyNameInput.value = current.name;
  if (strategyDeleteBtn) strategyDeleteBtn.removeAttribute("hidden");
  setStrategyModalOpen(true);
  if (strategyNameInput) {
    requestAnimationFrame(() => requestAnimationFrame(() => strategyNameInput.focus()));
  }
}

function openStrategyModalEdit() {
  openStrategyModalEditForId(activeBuyStrategyId);
}

function openStrategyActionsModal(strategyIdRaw) {
  const id = String(strategyIdRaw || "").trim();
  if (!id || id === DEFAULT_BUY_STRATEGY_ID || !strategyActionsOverlay) return;
  const s = buyStrategies.find((x) => x.id === id);
  if (!s) return;
  strategyActionsTargetId = id;
  if (strategyActionsTitle) strategyActionsTitle.textContent = s.name;
  openModalOverlay(strategyActionsOverlay);
}

function closeStrategyActionsModal() {
  strategyActionsTargetId = null;
  if (strategyActionsOverlay) closeModalOverlay(strategyActionsOverlay);
}

function openStrategyModalDuplicate() {
  const current = buyStrategies.find((s) => s.id === activeBuyStrategyId);
  if (!current) return;

  saveCurrentBuysToActiveStrategy();
  strategyModalEditingId = null;
  strategyModalDuplicatingFromId = current.id;
  if (strategyModalTitle) strategyModalTitle.textContent = "Дублировать стратегию";
  const baseName = String(current.name || "").trim() || "Стратегия";
  if (strategyNameInput) strategyNameInput.value = `Копия — ${baseName}`;
  if (strategyDeleteBtn) strategyDeleteBtn.setAttribute("hidden", "");
  setStrategyModalOpen(true);
  if (strategyNameInput) {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      strategyNameInput.focus();
      strategyNameInput.select();
    }));
  }
}

function closeStrategyModal() {
  strategyModalEditingId = null;
  strategyModalDuplicatingFromId = null;
  if (strategyModalTitle) strategyModalTitle.textContent = "Новая стратегия";
  if (strategyNameInput) strategyNameInput.value = "";
  if (strategyDeleteBtn) strategyDeleteBtn.setAttribute("hidden", "");
  setStrategyModalOpen(false);
}

function createBuyStrategy(nameRaw) {
  const name = normalizeStrategyName(nameRaw);
  if (!name) return false;
  const exists = buyStrategies.some((s) => s.name.toLocaleLowerCase("ru") === name.toLocaleLowerCase("ru"));
  if (exists) return false;

  saveCurrentBuysToActiveStrategy();
  const id = createBuyStrategyId();
  buyStrategies.push({ id, name });
  buyRowsByStrategyId.set(id, []);
  activeBuyStrategyId = id;
  renderBuyStrategySelect();
  writeRows(buysTbody, buyTpl, defaultBuyRows());
  syncBuySummaries();
  syncStaticTableEmptyStates();
  persistBuyStrategiesState();
  renderAll();
  return true;
}

function duplicateBuyStrategy(sourceIdRaw, nameRaw) {
  const sourceId = String(sourceIdRaw || "").trim();
  if (!sourceId || !buyStrategies.some((s) => s.id === sourceId)) return false;

  const name = normalizeStrategyName(nameRaw);
  if (!name) return false;
  const exists = buyStrategies.some((s) => s.name.toLocaleLowerCase("ru") === name.toLocaleLowerCase("ru"));
  if (exists) return false;

  if (sourceId === activeBuyStrategyId) {
    saveCurrentBuysToActiveStrategy();
  }
  const cloned = cloneStrategyBuyRows(buyRowsByStrategyId.get(sourceId));

  const id = createBuyStrategyId();
  buyStrategies.push({ id, name });
  buyRowsByStrategyId.set(id, cloned);
  activeBuyStrategyId = id;
  renderBuyStrategySelect();
  writeRows(buysTbody, buyTpl, cloned.length ? cloned : defaultBuyRows());
  syncBuySummaries();
  syncStaticTableEmptyStates();
  persistBuyStrategiesState();
  renderAll();
  return true;
}

function renameActiveStrategy(nameRaw) {
  const name = normalizeStrategyName(nameRaw);
  if (!name) return false;
  if (!strategyModalEditingId || strategyModalEditingId === DEFAULT_BUY_STRATEGY_ID) return false;
  const current = buyStrategies.find((s) => s.id === strategyModalEditingId);
  if (!current) return false;

  const exists = buyStrategies.some(
    (s) => s.id !== strategyModalEditingId && s.name.toLocaleLowerCase("ru") === name.toLocaleLowerCase("ru")
  );
  if (exists) return false;

  current.name = name;
  renderBuyStrategySelect();
  persistBuyStrategiesState();
  renderAll();
  return true;
}

function deleteBuyStrategyById(toDeleteRaw) {
  const toDelete = String(toDeleteRaw || "").trim();
  if (!toDelete || toDelete === DEFAULT_BUY_STRATEGY_ID) return false;
  if (!buyStrategies.some((s) => s.id === toDelete)) return false;

  saveCurrentBuysToActiveStrategy();

  buyStrategies = buyStrategies.filter((s) => s.id !== toDelete);
  buyRowsByStrategyId.delete(toDelete);

  if (!buyStrategies.some((s) => s.id === DEFAULT_BUY_STRATEGY_ID)) {
    buyStrategies.unshift(getDefaultBuyStrategy());
    if (!buyRowsByStrategyId.has(DEFAULT_BUY_STRATEGY_ID)) {
      buyRowsByStrategyId.set(DEFAULT_BUY_STRATEGY_ID, []);
    }
  }

  if (activeBuyStrategyId === toDelete) {
    activeBuyStrategyId = DEFAULT_BUY_STRATEGY_ID;
    const rows = getActiveStrategyBuysRows();
    writeRows(buysTbody, buyTpl, Array.isArray(rows) && rows.length ? rows : defaultBuyRows());
    syncBuySummaries();
    syncStaticTableEmptyStates();
  }

  renderBuyStrategySelect();
  persistBuyStrategiesState();
  renderAll();
  return true;
}

function deleteActiveBuyStrategy() {
  return deleteBuyStrategyById(activeBuyStrategyId);
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
      const priceInput = row.querySelector('input[data-buy-field="price"]');
      if (bondSelect) {
        const want = String(item.bond || "").trim().toUpperCase();
        const opt = Array.from(bondSelect.options).find((o) => o.value.toUpperCase() === want);
        if (opt) bondSelect.value = opt.value;
        else if (want) bondSelect.value = want;
      }
      if (bondSelect && priceInput && !buyModalItems[idx].priceUserEdited) {
        const b = String(bondSelect.value || "").trim();
        if (b && !String(priceInput.value || "").trim()) {
          const ref = getBondReferencePrice(b);
          if (Number.isFinite(ref)) {
            const formatted = formatAmount(ref);
            priceInput.value = formatted;
            buyModalItems[idx].price = formatted;
          }
        }
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
  buyModalItems = [newBuyModalItem()];
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
  if (!tr.classList.contains("buyTable__dataRow")) return;
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

    if (field === "price" && el instanceof HTMLInputElement) {
      buyModalItems[idx].priceUserEdited = String(el.value || "").trim() !== "";
    }

    if (field === "bond" && el instanceof HTMLSelectElement && buyItemsWrap) {
      const priceInput = buyItemsWrap.querySelector(`input[data-buy-field="price"][data-buy-idx="${idx}"]`);
      if (!priceInput || buyModalItems[idx].priceUserEdited) return;
      const b = String(el.value || "").trim();
      if (!b) {
        priceInput.value = "";
        buyModalItems[idx].price = "";
        return;
      }
      const ref = getBondReferencePrice(b);
      if (Number.isFinite(ref)) {
        const formatted = formatAmount(ref);
        priceInput.value = formatted;
        buyModalItems[idx].price = formatted;
      } else {
        priceInput.value = "";
        buyModalItems[idx].price = "";
      }
    }
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
    if (!buyModalItems.length) buyModalItems.push(newBuyModalItem());
    renderBuyModalItems();
  });
}

if (buyAddItemBtn) {
  buyAddItemBtn.addEventListener("click", () => {
    syncBuyModalItemsFromDom();
    buyModalItems.push(newBuyModalItem());
    renderBuyModalItems();
  });
}

if (buyStrategyNewBtn) {
  buyStrategyNewBtn.addEventListener("click", () => {
    openStrategyModalCreate();
  });
}

if (buyStrategyEditBtn) {
  buyStrategyEditBtn.addEventListener("click", () => {
    openStrategyModalEdit();
  });
}

if (buyStrategyDuplicateBtn) {
  buyStrategyDuplicateBtn.addEventListener("click", () => {
    openStrategyModalDuplicate();
  });
}

if (buyStrategyDeleteBtn) {
  buyStrategyDeleteBtn.addEventListener("click", () => {
    if (activeBuyStrategyId === DEFAULT_BUY_STRATEGY_ID) return;
    const current = buyStrategies.find((s) => s.id === activeBuyStrategyId);
    const label = current?.name ? `«${String(current.name)}»` : "эту стратегию";
    const ok = window.confirm(`Удалить стратегию ${label} и все запланированные в ней покупки?`);
    if (!ok) return;
    deleteActiveBuyStrategy();
  });
}

if (strategyModalClose) {
  strategyModalClose.addEventListener("click", () => closeStrategyModal());
}

if (strategyModalOverlay) {
  strategyModalOverlay.addEventListener("click", (e) => {
    if (e.target === strategyModalOverlay) closeStrategyModal();
  });
}

if (strategySaveBtn) {
  strategySaveBtn.addEventListener("click", () => {
    let ok = false;
    if (strategyModalDuplicatingFromId) {
      ok = duplicateBuyStrategy(strategyModalDuplicatingFromId, strategyNameInput?.value || "");
    } else if (strategyModalEditingId) {
      ok = renameActiveStrategy(strategyNameInput?.value || "");
    } else {
      ok = createBuyStrategy(strategyNameInput?.value || "");
    }
    if (!ok) return;
    closeStrategyModal();
  });
}

if (strategyDeleteBtn) {
  strategyDeleteBtn.addEventListener("click", () => {
    if (!strategyModalEditingId) return;
    const ok = deleteBuyStrategyById(strategyModalEditingId);
    if (!ok) return;
    closeStrategyModal();
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
    const purchaseYmd = normalizeYMD(row.purchaseDate) || String(row.purchaseDate || "").trim();
    const maturityYmd = normalizeYMD(row.maturityDate) || String(row.maturityDate || "").trim();
    const purchaseTs = ymdToUTCms(purchaseYmd);
    const maturityTs = ymdToUTCms(maturityYmd);

    if (!String(row.bondName || "").trim()) {
      window.alert("Укажите название или тикер облигации.");
      return;
    }
    if (purchaseTs == null || !Number.isFinite(purchaseTs)) {
      window.alert("Укажите корректную дату покупки.");
      return;
    }
    if (maturityTs == null || !Number.isFinite(maturityTs)) {
      window.alert("Укажите корректную дату погашения.");
      return;
    }
    if (maturityTs <= purchaseTs) {
      window.alert("Дата погашения должна быть позже даты покупки.");
      return;
    }

    const rowToSave = {
      ...row,
      purchaseDate: purchaseYmd,
      maturityDate: maturityYmd,
    };

    if (yieldModalEditingTr) {
      const tr = yieldModalEditingTr;
      yieldModalEditingTr = null;
      tr.querySelectorAll("input[data-field]").forEach((el) => {
        const key = el.getAttribute("data-field");
        if (key && rowToSave[key] !== undefined) el.value = rowToSave[key];
      });
      saveYieldCalcState();
      renderYieldCalculator();
      closeYieldModal();
      return;
    }

    if (!yieldTpl || !yieldTbody) return;
    const node = yieldTpl.content.cloneNode(true);
    const tr = node.querySelector("tr");
    if (!tr) return;
    tr.querySelectorAll("input[data-field]").forEach((el) => {
      const key = el.getAttribute("data-field");
      if (key && rowToSave[key] !== undefined) el.value = rowToSave[key];
    });
    clearTableEmptyState(yieldTbody);
    yieldTbody.appendChild(node);
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
  if (strategyModalOverlay && !strategyModalOverlay.hidden) {
    closeStrategyModal();
    e.preventDefault();
    return;
  }
  if (bondModalOverlay && !bondModalOverlay.hidden) {
    closeBondModal();
    e.preventDefault();
    return;
  }
  if (monthModalOverlay && !monthModalOverlay.hidden) {
    monthPickerState = null;
    setMonthModalOpen(false);
    e.preventDefault();
  }
});

