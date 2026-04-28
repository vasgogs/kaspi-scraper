const state = {
  filters: null,
  activeTab: "overview",
  masterPage: 1,
  rawPage: 1,
  comparePage: 1,
  overview: null,
  master: null,
  raw: null,
  compare: null,
  skuAnalysis: null,
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  bootstrap().catch((error) => {
    renderErrorAll(error);
  });
});

function cacheElements() {
  [
    "hero-source",
    "hero-source-note",
    "hero-updated",
    "hero-status",
    "reload-data",
    "filter-start-month",
    "filter-end-month",
    "filter-cohort",
    "filter-level1",
    "filter-level2",
    "filter-level3",
    "filter-brand",
    "filter-search",
    "download-dataset",
    "download-format",
    "apply-filters",
    "reset-filters",
    "download-current",
    "overview-metrics",
    "overview-granularity",
    "overview-trend",
    "overview-share-trend",
    "overview-level2",
    "overview-level3",
    "overview-brands",
    "overview-featured-categories",
    "overview-best-gap",
    "overview-weak-gap",
    "master-sort",
    "master-page-size",
    "masterfile-meta",
    "masterfile-table",
    "master-prev",
    "master-next",
    "master-pagination",
    "sku-my-input",
    "sku-compare-input",
    "sku-granularity",
    "sku-apply",
    "sku-meta",
    "sku-summary",
    "sku-price-chart",
    "sku-revenue-chart",
    "sku-qty-chart",
    "sku-compare-table",
    "raw-sort",
    "raw-page-size",
    "raw-meta",
    "raw-table",
    "raw-prev",
    "raw-next",
    "raw-pagination",
    "compare-level",
    "compare-search",
    "compare-sort",
    "compare-granularity",
    "compare-chart",
    "compare-table",
    "compare-prev",
    "compare-next",
    "compare-pagination",
  ].forEach((id) => {
    els[id] = document.getElementById(id);
  });
  els.tabs = Array.from(document.querySelectorAll(".tab"));
  els.tabPanels = Array.from(document.querySelectorAll(".tab-panel"));
}

function bindEvents() {
  els["reload-data"].addEventListener("click", async () => {
    await reloadData();
  });

  els["apply-filters"].addEventListener("click", async () => {
    resetPages();
    await refreshAll();
  });

  els["reset-filters"].addEventListener("click", async () => {
    applyDefaults();
    resetPages();
    await refreshAll();
  });

  els["download-current"].addEventListener("click", () => {
    window.location.href = buildDownloadUrl();
  });

  els["filter-level1"].addEventListener("change", () => {
    populateCategorySelects();
  });

  els["filter-level2"].addEventListener("change", () => {
    populateLevel3Select();
  });

  els["master-sort"].addEventListener("change", async () => {
    state.masterPage = 1;
    await refreshMasterfile();
  });

  els["master-page-size"].addEventListener("change", async () => {
    state.masterPage = 1;
    await refreshMasterfile();
  });

  els["raw-sort"].addEventListener("change", async () => {
    state.rawPage = 1;
    await refreshRaw();
  });

  els["raw-page-size"].addEventListener("change", async () => {
    state.rawPage = 1;
    await refreshRaw();
  });

  els["compare-level"].addEventListener("change", async () => {
    state.comparePage = 1;
    await refreshCompare();
  });

  els["compare-sort"].addEventListener("change", async () => {
    state.comparePage = 1;
    await refreshCompare();
  });

  els["compare-search"].addEventListener("keydown", async (event) => {
    if (event.key === "Enter") {
      state.comparePage = 1;
      await refreshCompare();
    }
  });

  els["overview-granularity"].addEventListener("change", () => {
    renderOverview();
  });

  els["compare-granularity"].addEventListener("change", () => {
    renderCompare();
  });

  els["master-prev"].addEventListener("click", async () => {
    if ((state.master?.pagination?.page || 1) > 1) {
      state.masterPage -= 1;
      await refreshMasterfile();
    }
  });

  els["master-next"].addEventListener("click", async () => {
    if ((state.master?.pagination?.page || 1) < (state.master?.pagination?.total_pages || 1)) {
      state.masterPage += 1;
      await refreshMasterfile();
    }
  });

  els["sku-apply"].addEventListener("click", async () => {
    await refreshSkuAnalysis();
  });

  [els["sku-my-input"], els["sku-compare-input"]].forEach((input) => {
    input.addEventListener("keydown", async (event) => {
      if (event.key === "Enter") {
        await refreshSkuAnalysis();
      }
    });
  });

  els["sku-granularity"].addEventListener("change", () => {
    renderSkuAnalysis();
  });

  els["raw-prev"].addEventListener("click", async () => {
    if ((state.raw?.pagination?.page || 1) > 1) {
      state.rawPage -= 1;
      await refreshRaw();
    }
  });

  els["raw-next"].addEventListener("click", async () => {
    if ((state.raw?.pagination?.page || 1) < (state.raw?.pagination?.total_pages || 1)) {
      state.rawPage += 1;
      await refreshRaw();
    }
  });

  els["compare-prev"].addEventListener("click", async () => {
    if ((state.compare?.pagination?.page || 1) > 1) {
      state.comparePage -= 1;
      await refreshCompare();
    }
  });

  els["compare-next"].addEventListener("click", async () => {
    if ((state.compare?.pagination?.page || 1) < (state.compare?.pagination?.total_pages || 1)) {
      state.comparePage += 1;
      await refreshCompare();
    }
  });

  [els["filter-brand"], els["filter-search"]].forEach((input) => {
    input.addEventListener("keydown", async (event) => {
      if (event.key === "Enter") {
        resetPages();
        await refreshAll();
      }
    });
  });

  els.tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      state.activeTab = tab.dataset.tab;
      els.tabs.forEach((item) => item.classList.toggle("is-active", item === tab));
      els.tabPanels.forEach((panel) =>
        panel.classList.toggle("is-active", panel.id === `tab-${state.activeTab}`)
      );
    });
  });
}

async function bootstrap() {
  renderLoadingAll();
  state.filters = await fetchJson("/algatop/api/filters");
  populateStaticControls();
  applyDefaults();
  updateHeroMeta(state.filters);
  setHeroStatus("Ready", "ready");
  await refreshAll();
}

function populateStaticControls() {
  const { months, cohorts, level1_options, download_datasets } = state.filters;
  setSelectOptions(els["filter-start-month"], months.map((value) => ({ value, label: value })));
  setSelectOptions(els["filter-end-month"], months.map((value) => ({ value, label: value })));
  setSelectOptions(els["filter-cohort"], cohorts);
  setSelectOptions(els["filter-level1"], level1_options, { includeBlank: true, blankLabel: "All" });
  setSelectOptions(els["download-dataset"], download_datasets);
  populateCategorySelects();
  populateSkuDatalists();
}

function applyDefaults() {
  const defaults = state.filters.defaults;
  els["filter-start-month"].value = defaults.start_month;
  els["filter-end-month"].value = defaults.end_month;
  els["filter-cohort"].value = defaults.cohort;
  els["filter-level1"].value = defaults.level1_code;
  els["filter-level2"].value = defaults.level2_code;
  els["filter-level3"].value = defaults.level3_code;
  els["filter-brand"].value = defaults.brand_query;
  els["filter-search"].value = defaults.search_query;
  els["master-sort"].value = defaults.master_sort;
  els["raw-sort"].value = defaults.raw_sort;
  els["compare-level"].value = String(defaults.compare_level);
  els["compare-sort"].value = defaults.compare_sort;
  els["overview-granularity"].value = "month";
  els["compare-granularity"].value = "month";
  els["sku-granularity"].value = "month";
  els["download-dataset"].value = "masterfile";
  els["download-format"].value = "xlsx";
  setSkuInputValue(els["sku-my-input"], "my", defaults.my_sku);
  setSkuInputValue(els["sku-compare-input"], "competitor", defaults.compare_sku);
  populateCategorySelects();
}

function populateSkuDatalists() {
  populateDatalist("sku-my-list", state.filters.sku_options?.my || []);
  populateDatalist("sku-compare-list", state.filters.sku_options?.competitor || []);
}

function populateDatalist(id, options) {
  const element = document.getElementById(id);
  element.innerHTML = options
    .map(
      (option) =>
        `<option value="${escapeHtml(skuDisplayValue(option))}" label="${escapeHtml(option.label || option.display)}"></option>`
    )
    .join("");
}

function skuDisplayValue(option) {
  return option ? `${option.sku} - ${option.product_name}` : "";
}

function extractSkuInput(value) {
  const match = String(value || "").match(/\d{6,}/);
  return match ? match[0] : "";
}

function findSkuOption(cohort, sku) {
  const options = state.filters?.sku_options?.[cohort] || [];
  return options.find((option) => option.sku === sku) || null;
}

function setSkuInputValue(input, cohort, sku) {
  const option = findSkuOption(cohort, sku);
  input.value = option ? skuDisplayValue(option) : sku || "";
}

function populateCategorySelects() {
  const selectedLevel1 = els["filter-level1"].value;
  const level2Options = state.filters.level2_options.filter(
    (row) => !selectedLevel1 || row.parent_code === selectedLevel1
  );
  setSelectOptions(els["filter-level2"], level2Options, { includeBlank: true, blankLabel: "All" });
  if (!level2Options.find((row) => row.code === els["filter-level2"].value)) {
    els["filter-level2"].value = "";
  }
  populateLevel3Select();
}

function populateLevel3Select() {
  const selectedLevel2 = els["filter-level2"].value;
  const level3Options = state.filters.level3_options.filter(
    (row) => !selectedLevel2 || row.parent_code === selectedLevel2
  );
  setSelectOptions(els["filter-level3"], level3Options, { includeBlank: true, blankLabel: "All" });
  if (!level3Options.find((row) => row.code === els["filter-level3"].value)) {
    els["filter-level3"].value = "";
  }
}

function setSelectOptions(select, options, config = {}) {
  const { includeBlank = false, blankLabel = "All" } = config;
  const current = select.value;
  const normalizedOptions = includeBlank ? [{ value: "", label: blankLabel }] : [];
  options.forEach((option) => {
    normalizedOptions.push({
      value: option.value || option.code,
      label: option.label || option.path || option.name,
    });
  });
  select.innerHTML = normalizedOptions
    .map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  if (normalizedOptions.find((option) => option.value === current)) {
    select.value = current;
  }
}

function resetPages() {
  state.masterPage = 1;
  state.rawPage = 1;
  state.comparePage = 1;
}

function snapshotControls() {
  return {
    ...globalParams(),
    master_sort: els["master-sort"].value,
    raw_sort: els["raw-sort"].value,
    compare_level: els["compare-level"].value,
    compare_sort: els["compare-sort"].value,
    compare_search: els["compare-search"].value.trim(),
    overview_granularity: els["overview-granularity"].value,
    compare_granularity: els["compare-granularity"].value,
    sku_granularity: els["sku-granularity"].value,
    my_sku: els["sku-my-input"].value.trim(),
    compare_sku: els["sku-compare-input"].value.trim(),
    download_dataset: els["download-dataset"].value,
    download_format: els["download-format"].value,
  };
}

function restoreControls(snapshot) {
  if (!snapshot) {
    return;
  }
  els["filter-start-month"].value = snapshot.start_month;
  els["filter-end-month"].value = snapshot.end_month;
  els["filter-cohort"].value = snapshot.cohort;
  els["filter-level1"].value = snapshot.level1_code;
  populateCategorySelects();
  els["filter-level2"].value = snapshot.level2_code;
  populateLevel3Select();
  els["filter-level3"].value = snapshot.level3_code;
  els["filter-brand"].value = snapshot.brand_query;
  els["filter-search"].value = snapshot.search_query;
  els["master-sort"].value = snapshot.master_sort;
  els["raw-sort"].value = snapshot.raw_sort;
  els["compare-level"].value = snapshot.compare_level;
  els["compare-sort"].value = snapshot.compare_sort;
  els["compare-search"].value = snapshot.compare_search;
  els["overview-granularity"].value = snapshot.overview_granularity || "month";
  els["compare-granularity"].value = snapshot.compare_granularity || "month";
  els["sku-granularity"].value = snapshot.sku_granularity || "month";
  els["sku-my-input"].value = snapshot.my_sku;
  els["sku-compare-input"].value = snapshot.compare_sku;
  els["download-dataset"].value = snapshot.download_dataset;
  els["download-format"].value = snapshot.download_format;
}

function updateHeroMeta(payload) {
  const sourceFiles = payload?.source_files || {};
  const master = sourceFiles.master || document.body.dataset.masterFilename || "";
  const compare = sourceFiles.compare || "";
  const category = sourceFiles.category || "";
  const myMonthly = sourceFiles.my_monthly || "";
  const competitorMonthly = sourceFiles.competitor_monthly || "";
  els["hero-source"].textContent = master;
  els["hero-updated"].textContent = payload?.updated_at || document.body.dataset.updatedAt || "";
  els["hero-source-note"].textContent = [compare, category, myMonthly, competitorMonthly].filter(Boolean).join(" · ");
}

function setHeroStatus(label, tone = "ready") {
  const element = els["hero-status"];
  element.textContent = label;
  element.className = `status-badge status-badge--${tone}`;
}

async function reloadData() {
  const snapshot = snapshotControls();
  els["reload-data"].disabled = true;
  setHeroStatus("Reloading", "busy");
  try {
    await fetchJson("/algatop/api/reload", { method: "POST" });
    state.filters = await fetchJson("/algatop/api/filters");
    populateStaticControls();
    restoreControls(snapshot);
    resetPages();
    updateHeroMeta(state.filters);
    await refreshAll();
    setHeroStatus("Fresh", "ready");
  } catch (error) {
    setHeroStatus("Reload failed", "error");
    renderErrorAll(error);
  } finally {
    els["reload-data"].disabled = false;
  }
}

function globalParams() {
  return {
    start_month: els["filter-start-month"].value,
    end_month: els["filter-end-month"].value,
    cohort: els["filter-cohort"].value,
    level1_code: els["filter-level1"].value,
    level2_code: els["filter-level2"].value,
    level3_code: els["filter-level3"].value,
    brand_query: els["filter-brand"].value.trim(),
    search_query: els["filter-search"].value.trim(),
  };
}

function compareParams() {
  return {
    start_month: els["filter-start-month"].value,
    end_month: els["filter-end-month"].value,
    level: els["compare-level"].value,
    search_query: els["compare-search"].value.trim(),
    sort: els["compare-sort"].value,
    page: state.comparePage,
    page_size: 25,
  };
}

function skuParams() {
  return {
    start_month: els["filter-start-month"].value,
    end_month: els["filter-end-month"].value,
    my_sku: extractSkuInput(els["sku-my-input"].value),
    compare_sku: extractSkuInput(els["sku-compare-input"].value),
  };
}

function selectedGranularity(id, fallback) {
  return els[id]?.value || fallback;
}

function numeric(value) {
  return Number(value || 0);
}

function pct(part, total) {
  const whole = Number(total || 0);
  return whole ? (Number(part || 0) / whole) * 100 : 0;
}

function parseMonthLabel(month) {
  const [year, monthNumber] = String(month || "").split("-").map(Number);
  return new Date(Date.UTC(year || 1970, Math.max((monthNumber || 1) - 1, 0), 1));
}

function parseDateLabel(value) {
  const [year, month, day] = String(value || "").split("-").map(Number);
  return new Date(Date.UTC(year || 1970, Math.max((month || 1) - 1, 0), day || 1));
}

function isoDate(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(
    date.getUTCDate()
  ).padStart(2, "0")}`;
}

function monthLabel(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function quarterLabel(date) {
  return `${date.getUTCFullYear()}-Q${Math.floor(date.getUTCMonth() / 3) + 1}`;
}

function startOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function startOfQuarter(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), Math.floor(date.getUTCMonth() / 3) * 3, 1));
}

function startOfYear(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
}

function isoWeekInfo(date) {
  const target = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNumber = (target.getUTCDay() + 6) % 7;
  target.setUTCDate(target.getUTCDate() - dayNumber + 3);
  const isoYear = target.getUTCFullYear();
  const firstThursday = new Date(Date.UTC(isoYear, 0, 4));
  const firstDayNumber = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNumber + 3);
  const week = 1 + Math.round((target - firstThursday) / 604800000);
  return { isoYear, week };
}

function startOfIsoWeek(date) {
  const result = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const dayNumber = (result.getUTCDay() + 6) % 7;
  result.setUTCDate(result.getUTCDate() - dayNumber);
  return result;
}

function monthBucketInfo(month, granularity) {
  const date = parseMonthLabel(month);
  if (granularity === "year") {
    const start = startOfYear(date);
    return { key: String(start.getUTCFullYear()), label: String(start.getUTCFullYear()), sortKey: start.getTime() };
  }
  if (granularity === "quarter") {
    const start = startOfQuarter(date);
    return { key: quarterLabel(start), label: quarterLabel(start), sortKey: start.getTime() };
  }
  return { key: monthLabel(date), label: monthLabel(date), sortKey: date.getTime() };
}

function dateBucketInfo(value, granularity) {
  const date = parseDateLabel(value);
  if (granularity === "year") {
    const start = startOfYear(date);
    return { key: String(start.getUTCFullYear()), label: String(start.getUTCFullYear()), sortKey: start.getTime() };
  }
  if (granularity === "quarter") {
    const start = startOfQuarter(date);
    return { key: quarterLabel(start), label: quarterLabel(start), sortKey: start.getTime() };
  }
  if (granularity === "month") {
    const start = startOfMonth(date);
    return { key: monthLabel(start), label: monthLabel(start), sortKey: start.getTime() };
  }
  if (granularity === "week") {
    const start = startOfIsoWeek(date);
    const info = isoWeekInfo(start);
    return {
      key: `${info.isoYear}-W${String(info.week).padStart(2, "0")}`,
      label: `${info.isoYear}-W${String(info.week).padStart(2, "0")}`,
      sortKey: start.getTime(),
    };
  }
  return { key: isoDate(date), label: isoDate(date), sortKey: date.getTime() };
}

function upsertBucket(map, bucket) {
  if (!map.has(bucket.key)) {
    map.set(bucket.key, { key: bucket.key, label: bucket.label, sortKey: bucket.sortKey });
  }
  return map.get(bucket.key);
}

function accumulateWeightedMetric(entry, prefix, value, basis) {
  const amount = numeric(value);
  const weight = numeric(basis);
  if (amount <= 0) {
    return;
  }
  entry[`_${prefix}_weighted_sum`] = numeric(entry[`_${prefix}_weighted_sum`]) + amount * (weight > 0 ? weight : 1);
  entry[`_${prefix}_weighted_basis`] = numeric(entry[`_${prefix}_weighted_basis`]) + (weight > 0 ? weight : 1);
  entry[`_${prefix}_last_value`] = amount;
}

function finalizeWeightedMetric(entry, prefix, targetField) {
  const basis = numeric(entry[`_${prefix}_weighted_basis`]);
  const weightedSum = numeric(entry[`_${prefix}_weighted_sum`]);
  entry[targetField] = basis ? weightedSum / basis : numeric(entry[`_${prefix}_last_value`]);
  delete entry[`_${prefix}_weighted_sum`];
  delete entry[`_${prefix}_weighted_basis`];
  delete entry[`_${prefix}_last_value`];
}

function accumulateAverageMetric(entry, prefix, value) {
  const amount = numeric(value);
  if (amount <= 0) {
    return;
  }
  entry[`_${prefix}_sum`] = numeric(entry[`_${prefix}_sum`]) + amount;
  entry[`_${prefix}_count`] = numeric(entry[`_${prefix}_count`]) + 1;
}

function finalizeAverageMetric(entry, prefix, targetField) {
  const count = numeric(entry[`_${prefix}_count`]);
  const sum = numeric(entry[`_${prefix}_sum`]);
  entry[targetField] = count ? sum / count : 0;
  delete entry[`_${prefix}_sum`];
  delete entry[`_${prefix}_count`];
}

function aggregateOverviewTrend(rows, granularity) {
  if (granularity === "month") {
    return rows.map((row) => ({ ...row, label: row.month }));
  }
  const buckets = new Map();
  rows.forEach((row) => {
    const bucket = monthBucketInfo(row.month, granularity);
    const entry = upsertBucket(buckets, bucket);
    entry.my_sale_amount = numeric(entry.my_sale_amount) + numeric(row.my_sale_amount);
    entry.competitor_sale_amount = numeric(entry.competitor_sale_amount) + numeric(row.competitor_sale_amount);
    entry.total_sale_amount = numeric(entry.total_sale_amount) + numeric(row.total_sale_amount);
    entry.total_sale_qty = numeric(entry.total_sale_qty) + numeric(row.total_sale_qty);
  });
  return Array.from(buckets.values()).sort((a, b) => a.sortKey - b.sortKey);
}

function aggregateShareTrend(rows, granularity) {
  if (granularity === "month") {
    return rows.map((row) => ({
      ...row,
      label: row.month,
      my_share_amount_pct: numeric(row.my_share_amount_pct),
      competitor_share_amount_pct: numeric(row.competitor_share_amount_pct),
    }));
  }
  const buckets = new Map();
  rows.forEach((row) => {
    const bucket = monthBucketInfo(row.month, granularity);
    const entry = upsertBucket(buckets, bucket);
    entry.my_sale_amount = numeric(entry.my_sale_amount) + numeric(row.my_sale_amount);
    entry.competitor_sale_amount = numeric(entry.competitor_sale_amount) + numeric(row.competitor_sale_amount);
    entry.category_sale_amount = numeric(entry.category_sale_amount) + numeric(row.category_sale_amount);
  });
  return Array.from(buckets.values())
    .sort((a, b) => a.sortKey - b.sortKey)
    .map((entry) => ({
      ...entry,
      my_share_amount_pct: pct(entry.my_sale_amount, entry.category_sale_amount),
      competitor_share_amount_pct: pct(entry.competitor_sale_amount, entry.category_sale_amount),
    }));
}

function aggregateSkuMonthlyRows(payload, granularity) {
  const months = payload?.monthly_trend?.months || [];
  const myRows = payload?.monthly_trend?.my || [];
  const compareRows = payload?.monthly_trend?.compare || [];

  if (granularity === "month") {
    return months.map((month, index) => {
      const myRow = myRows[index] || {};
      const compareRow = compareRows[index] || {};
      return {
        label: month,
        sortKey: parseMonthLabel(month).getTime(),
        my_sale_amount: numeric(myRow.sale_amount),
        compare_sale_amount: numeric(compareRow.sale_amount),
        my_sale_qty: numeric(myRow.sale_qty),
        compare_sale_qty: numeric(compareRow.sale_qty),
        my_sale_price: numeric(myRow.sale_price),
        compare_sale_price: numeric(compareRow.sale_price),
        my_merchant_count: numeric(myRow.merchant_count),
        compare_merchant_count: numeric(compareRow.merchant_count),
        my_review_qty: numeric(myRow.review_qty),
        compare_review_qty: numeric(compareRow.review_qty),
        revenue_gap: numeric(myRow.sale_amount) - numeric(compareRow.sale_amount),
        price_gap: numeric(myRow.sale_price) - numeric(compareRow.sale_price),
      };
    });
  }

  const buckets = new Map();
  months.forEach((month, index) => {
    const bucket = monthBucketInfo(month, granularity);
    const entry = upsertBucket(buckets, bucket);
    const myRow = myRows[index] || {};
    const compareRow = compareRows[index] || {};

    entry.my_sale_amount = numeric(entry.my_sale_amount) + numeric(myRow.sale_amount);
    entry.compare_sale_amount = numeric(entry.compare_sale_amount) + numeric(compareRow.sale_amount);
    entry.my_sale_qty = numeric(entry.my_sale_qty) + numeric(myRow.sale_qty);
    entry.compare_sale_qty = numeric(entry.compare_sale_qty) + numeric(compareRow.sale_qty);
    accumulateWeightedMetric(entry, "my_price", myRow.sale_price, myRow.sale_qty);
    accumulateWeightedMetric(entry, "compare_price", compareRow.sale_price, compareRow.sale_qty);
    accumulateAverageMetric(entry, "my_merchant_count", myRow.merchant_count);
    accumulateAverageMetric(entry, "compare_merchant_count", compareRow.merchant_count);
    entry.my_review_qty = numeric(myRow.review_qty) || numeric(entry.my_review_qty);
    entry.compare_review_qty = numeric(compareRow.review_qty) || numeric(entry.compare_review_qty);
  });

  return Array.from(buckets.values())
    .sort((a, b) => a.sortKey - b.sortKey)
    .map((entry) => {
      finalizeWeightedMetric(entry, "my_price", "my_sale_price");
      finalizeWeightedMetric(entry, "compare_price", "compare_sale_price");
      finalizeAverageMetric(entry, "my_merchant_count", "my_merchant_count");
      finalizeAverageMetric(entry, "compare_merchant_count", "compare_merchant_count");
      entry.revenue_gap = numeric(entry.my_sale_amount) - numeric(entry.compare_sale_amount);
      entry.price_gap = numeric(entry.my_sale_price) - numeric(entry.compare_sale_price);
      return entry;
    });
}

function aggregateSkuDailyRows(rows, granularity) {
  const normalizedGranularity = granularity || "day";
  if (normalizedGranularity === "day") {
    return rows.map((row) => ({
      label: row.event_date,
      sortKey: parseDateLabel(row.event_date).getTime(),
      my_sale_price: numeric(row.my_sale_price),
      compare_sale_price: numeric(row.compare_sale_price),
      my_sale_amount: numeric(row.my_sale_amount),
      compare_sale_amount: numeric(row.compare_sale_amount),
      my_sale_qty: numeric(row.my_sale_qty),
      compare_sale_qty: numeric(row.compare_sale_qty),
      my_merchant_count: numeric(row.my_merchant_count),
      compare_merchant_count: numeric(row.compare_merchant_count),
      my_review_qty: numeric(row.my_review_qty),
      compare_review_qty: numeric(row.compare_review_qty),
      revenue_gap: numeric(row.my_sale_amount) - numeric(row.compare_sale_amount),
      price_gap: numeric(row.my_sale_price) - numeric(row.compare_sale_price),
    }));
  }

  const buckets = new Map();
  rows.forEach((row) => {
    const bucket = dateBucketInfo(row.event_date, normalizedGranularity);
    const entry = upsertBucket(buckets, bucket);
    entry.my_sale_amount = numeric(entry.my_sale_amount) + numeric(row.my_sale_amount);
    entry.compare_sale_amount = numeric(entry.compare_sale_amount) + numeric(row.compare_sale_amount);
    entry.my_sale_qty = numeric(entry.my_sale_qty) + numeric(row.my_sale_qty);
    entry.compare_sale_qty = numeric(entry.compare_sale_qty) + numeric(row.compare_sale_qty);
    accumulateWeightedMetric(entry, "my_price", row.my_sale_price, row.my_sale_qty);
    accumulateWeightedMetric(entry, "compare_price", row.compare_sale_price, row.compare_sale_qty);
    accumulateAverageMetric(entry, "my_merchant_count", row.my_merchant_count);
    accumulateAverageMetric(entry, "compare_merchant_count", row.compare_merchant_count);
    entry.my_review_qty = numeric(row.my_review_qty) || numeric(entry.my_review_qty);
    entry.compare_review_qty = numeric(row.compare_review_qty) || numeric(entry.compare_review_qty);
  });

  return Array.from(buckets.values())
    .sort((a, b) => a.sortKey - b.sortKey)
    .map((entry) => {
      finalizeWeightedMetric(entry, "my_price", "my_sale_price");
      finalizeWeightedMetric(entry, "compare_price", "compare_sale_price");
      finalizeAverageMetric(entry, "my_merchant_count", "my_merchant_count");
      finalizeAverageMetric(entry, "compare_merchant_count", "compare_merchant_count");
      entry.revenue_gap = numeric(entry.my_sale_amount) - numeric(entry.compare_sale_amount);
      entry.price_gap = numeric(entry.my_sale_price) - numeric(entry.compare_sale_price);
      return entry;
    });
}

function buildTooltipRows(items) {
  return items
    .filter((item) => item && item.value !== undefined && item.value !== null)
    .map((item) => ({
      label: item.label,
      value: item.value,
      color: item.color || "",
      emphasis: item.emphasis || false,
    }));
}

function initLineCharts(root = document) {
  if (!root) {
    return;
  }
  root.querySelectorAll(".line-chart").forEach((chart) => {
    if (chart.dataset.ready === "1") {
      return;
    }
    chart.dataset.ready = "1";
    const stage = chart.querySelector(".line-chart__stage");
    const tooltip = chart.querySelector(".line-chart__tooltip");
    const hoverLine = chart.querySelector(".line-chart__hover-line");
    if (!stage || !tooltip) {
      return;
    }
    const payloadElement = chart.querySelector(".line-chart__payload");
    const tooltipPayload = payloadElement ? JSON.parse(payloadElement.textContent || "[]") : [];
    const slices = Array.from(chart.querySelectorAll(".line-chart__hit-slice"));

    const hideTooltip = () => {
      tooltip.hidden = true;
      hoverLine?.setAttribute("opacity", "0");
      slices.forEach((slice) => slice.classList.remove("is-active"));
    };

    const showTooltip = (slice, event) => {
      const index = Number(slice.dataset.index || 0);
      const data = tooltipPayload[index];
      if (!data) {
        hideTooltip();
        return;
      }
      const rect = slice.getBoundingClientRect();
      const stageRect = stage.getBoundingClientRect();
      const centerX = rect.left + rect.width / 2 - stageRect.left;
      const preferredX = (event?.clientX || rect.left + rect.width / 2) - stageRect.left;
      const xPosition = Math.max(12, Math.min(stageRect.width - 12, preferredX));
      tooltip.innerHTML = `
        <div class="line-chart__tooltip-title">${escapeHtml(data.title)}</div>
        ${data.rows
          .map(
            (row) => `
              <div class="line-chart__tooltip-row${row.emphasis ? " is-emphasis" : ""}">
                <span class="line-chart__tooltip-label">
                  ${row.color ? `<span class="line-chart__tooltip-dot" style="background:${escapeHtml(row.color)}"></span>` : ""}
                  ${escapeHtml(row.label)}
                </span>
                <strong class="mono">${escapeHtml(row.value)}</strong>
              </div>
            `
          )
          .join("")}
      `;
      tooltip.hidden = false;
      const halfWidth = tooltip.offsetWidth / 2;
      tooltip.style.left = `${Math.max(halfWidth + 12, Math.min(stageRect.width - halfWidth - 12, xPosition))}px`;
      tooltip.style.top = "10px";
      hoverLine?.setAttribute("x1", String(centerX));
      hoverLine?.setAttribute("x2", String(centerX));
      hoverLine?.setAttribute("opacity", "1");
      slices.forEach((item) => item.classList.toggle("is-active", item === slice));
    };

    slices.forEach((slice) => {
      slice.addEventListener("mouseenter", (event) => showTooltip(slice, event));
      slice.addEventListener("mousemove", (event) => showTooltip(slice, event));
      slice.addEventListener("mouseleave", hideTooltip);
      slice.addEventListener("focus", (event) => showTooltip(slice, event));
      slice.addEventListener("blur", hideTooltip);
    });

    stage.addEventListener("mouseleave", hideTooltip);
  });
}

async function refreshAll() {
  renderLoadingAll();
  const overviewPromise = fetchJson(`/algatop/api/overview?${new URLSearchParams(globalParams())}`);
  const masterPromise = fetchJson(
    `/algatop/api/masterfile?${new URLSearchParams({
      ...globalParams(),
      sort: els["master-sort"].value,
      page: String(state.masterPage),
      page_size: els["master-page-size"].value,
    })}`
  );
  const rawPromise = fetchJson(
    `/algatop/api/raw?${new URLSearchParams({
      ...globalParams(),
      sort: els["raw-sort"].value,
      page: String(state.rawPage),
      page_size: els["raw-page-size"].value,
    })}`
  );
  const comparePromise = fetchJson(`/algatop/api/category-compare?${new URLSearchParams(compareParams())}`);
  const skuPromise = fetchJson(`/algatop/api/sku-analysis?${new URLSearchParams(skuParams())}`);

  try {
    const [overview, master, raw, compare, skuAnalysis] = await Promise.all([
      overviewPromise,
      masterPromise,
      rawPromise,
      comparePromise,
      skuPromise,
    ]);
    state.overview = overview;
    state.master = master;
    state.raw = raw;
    state.compare = compare;
    state.skuAnalysis = skuAnalysis;
    renderOverview();
    renderMasterfile();
    renderRaw();
    renderCompare();
    renderSkuAnalysis();
  } catch (error) {
    renderErrorAll(error);
  }
}

async function refreshMasterfile() {
  showLoading(els["masterfile-table"]);
  try {
    state.master = await fetchJson(
      `/algatop/api/masterfile?${new URLSearchParams({
        ...globalParams(),
        sort: els["master-sort"].value,
        page: String(state.masterPage),
        page_size: els["master-page-size"].value,
      })}`
    );
    renderMasterfile();
  } catch (error) {
    showError(els["masterfile-table"], error);
  }
}

async function refreshRaw() {
  showLoading(els["raw-table"]);
  try {
    state.raw = await fetchJson(
      `/algatop/api/raw?${new URLSearchParams({
        ...globalParams(),
        sort: els["raw-sort"].value,
        page: String(state.rawPage),
        page_size: els["raw-page-size"].value,
      })}`
    );
    renderRaw();
  } catch (error) {
    showError(els["raw-table"], error);
  }
}

async function refreshCompare() {
  showLoading(els["compare-table"]);
  showLoading(els["compare-chart"]);
  try {
    state.compare = await fetchJson(`/algatop/api/category-compare?${new URLSearchParams(compareParams())}`);
    renderCompare();
  } catch (error) {
    showError(els["compare-table"], error);
    showError(els["compare-chart"], error);
  }
}

async function refreshSkuAnalysis() {
  showLoading(els["sku-summary"]);
  showLoading(els["sku-price-chart"]);
  showLoading(els["sku-revenue-chart"]);
  showLoading(els["sku-qty-chart"]);
  showLoading(els["sku-compare-table"]);
  try {
    state.skuAnalysis = await fetchJson(`/algatop/api/sku-analysis?${new URLSearchParams(skuParams())}`);
    renderSkuAnalysis();
  } catch (error) {
    showError(els["sku-summary"], error);
    showError(els["sku-price-chart"], error);
    showError(els["sku-revenue-chart"], error);
    showError(els["sku-qty-chart"], error);
    showError(els["sku-compare-table"], error);
  }
}

function renderOverview() {
  const overview = state.overview;
  const granularity = selectedGranularity("overview-granularity", "month");
  const revenueTrend = aggregateOverviewTrend(overview.monthly_trend || [], granularity);
  const featuredCategories = (overview.featured_categories || []).map((item) => ({
    ...item,
    aggregatedTrend: aggregateShareTrend(item.trend || [], granularity),
  }));

  els["overview-metrics"].innerHTML = overview.metrics
    .map(
      (metric) => `
        <article class="metric-card metric-card--${escapeHtml(metric.tone || "neutral")}">
          <span class="metric-card__label">${escapeHtml(metric.label)}</span>
          <strong class="metric-card__value mono">${formatValue(metric.value, metric.label)}</strong>
          <span class="metric-card__note">${escapeHtml(metric.note || "")}</span>
        </article>
      `
    )
    .join("");

  els["overview-trend"].innerHTML = revenueTrend.length
    ? buildLineChartMarkup({
        labels: revenueTrend.map((row) => row.label),
        series: [
          { name: "My", color: "#146356", values: revenueTrend.map((row) => row.my_sale_amount) },
          {
            name: "Competitor",
            color: "#c06014",
            values: revenueTrend.map((row) => row.competitor_sale_amount),
          },
          { name: "Total", color: "#734bce", values: revenueTrend.map((row) => row.total_sale_amount) },
        ],
        formatter: shortMoney,
        tooltipFormatter: formatMoney,
        extraRows: revenueTrend.map((row) =>
          buildTooltipRows([
            { label: "Sales qty", value: formatNumber(row.total_sale_qty), emphasis: true },
            { label: "My share", value: formatPct(pct(row.my_sale_amount, row.total_sale_amount)) },
            {
              label: "Competitor share",
              value: formatPct(pct(row.competitor_sale_amount, row.total_sale_amount)),
            },
          ])
        ),
      })
    : emptyState("Нет данных для графика");

  els["overview-share-trend"].innerHTML = featuredCategories.length
    ? featuredCategories
        .map(
          (item) => `
            <article class="category-card">
              <div class="category-card__title">${escapeHtml(item.category_name)}</div>
              <div class="category-card__stats">
                <div class="category-card__stat">
                  <span>My share</span>
                  <strong class="mono">${formatPct(item.my_share_amount_pct)}</strong>
                </div>
                <div class="category-card__stat">
                  <span>Competitor share</span>
                  <strong class="mono">${formatPct(item.competitor_share_amount_pct)}</strong>
                </div>
              </div>
              ${buildLineChartMarkup({
                labels: item.aggregatedTrend.map((row) => row.label),
                compact: true,
                formatter: (value) => `${round1(value)}%`,
                tooltipFormatter: (value) => `${exactNumber(value)}%`,
                series: [
                  {
                    name: "My share",
                    color: "#146356",
                    values: item.aggregatedTrend.map((row) => row.my_share_amount_pct),
                  },
                  {
                    name: "Competitor share",
                    color: "#c06014",
                    values: item.aggregatedTrend.map((row) => row.competitor_share_amount_pct),
                  },
                ],
                extraRows: item.aggregatedTrend.map((row) =>
                  buildTooltipRows([
                    { label: "Market revenue", value: formatMoney(row.category_sale_amount), emphasis: true },
                    { label: "My revenue", value: formatMoney(row.my_sale_amount) },
                    { label: "Competitor revenue", value: formatMoney(row.competitor_sale_amount) },
                    {
                      label: "Gap",
                      value: `${row.my_share_amount_pct - row.competitor_share_amount_pct > 0 ? "+" : ""}${exactNumber(
                        row.my_share_amount_pct - row.competitor_share_amount_pct
                      )} pp`,
                    },
                  ])
                ),
              })}
            </article>
          `
        )
        .join("")
    : emptyState("Нет данных для share trend");

  renderTable(
    els["overview-level2"],
    [
      { label: "Category", render: (row) => escapeHtml(row.label), title: (row) => row.label },
      { label: "Revenue", render: (row) => formatMoney(row.sale_amount), title: (row) => exactNumber(row.sale_amount) },
      { label: "SKU", render: (row) => formatNumber(row.sku_count), title: (row) => exactNumber(row.sku_count) },
    ],
    overview.top_level2,
    { compact: true }
  );

  renderTable(
    els["overview-level3"],
    [
      { label: "Category", render: (row) => escapeHtml(row.label), title: (row) => row.label },
      { label: "Revenue", render: (row) => formatMoney(row.sale_amount), title: (row) => exactNumber(row.sale_amount) },
      { label: "SKU", render: (row) => formatNumber(row.sku_count), title: (row) => exactNumber(row.sku_count) },
    ],
    overview.top_level3,
    { compact: true }
  );

  renderTable(
    els["overview-brands"],
    [
      { label: "Brand", render: (row) => escapeHtml(row.label), title: (row) => row.label },
      { label: "Revenue", render: (row) => formatMoney(row.sale_amount), title: (row) => exactNumber(row.sale_amount) },
      { label: "SKU", render: (row) => formatNumber(row.sku_count), title: (row) => exactNumber(row.sku_count) },
    ],
    overview.top_brands,
    { compact: true }
  );

  els["overview-featured-categories"].innerHTML = overview.featured_categories
    .map(
      (item) => `
        <article class="category-card">
          <div class="category-card__title">${escapeHtml(item.category_name)}</div>
          <div class="category-card__stats">
            <div class="category-card__stat">
              <span>Market revenue</span>
              <strong class="mono">${formatMoney(item.latest_market_revenue)}</strong>
            </div>
            <div class="category-card__stat">
              <span>Orders</span>
              <strong class="mono">${formatNumber(item.latest_market_orders)}</strong>
            </div>
            <div class="category-card__stat">
              <span>My share</span>
              <strong class="mono">${formatPct(item.my_share_amount_pct)}</strong>
            </div>
            <div class="category-card__stat">
              <span>Competitor share</span>
              <strong class="mono">${formatPct(item.competitor_share_amount_pct)}</strong>
            </div>
          </div>
        </article>
      `
    )
    .join("");

  renderTable(
    els["overview-best-gap"],
    [
      { label: "Subcategory", render: (row) => escapeHtml(row.category_name), title: (row) => row.full_path || row.category_name },
      { label: "My", render: (row) => formatPct(row.my_share_amount_pct), title: (row) => exactNumber(row.my_share_amount_pct) },
      { label: "Comp", render: (row) => formatPct(row.competitor_share_amount_pct), title: (row) => exactNumber(row.competitor_share_amount_pct) },
      { label: "Gap", render: (row) => toneDelta(row.share_gap_amount_pct), title: (row) => exactNumber(row.share_gap_amount_pct) },
    ],
    overview.best_vs_competitors,
    { compact: true }
  );

  renderTable(
    els["overview-weak-gap"],
    [
      { label: "Subcategory", render: (row) => escapeHtml(row.category_name), title: (row) => row.full_path || row.category_name },
      { label: "My", render: (row) => formatPct(row.my_share_amount_pct), title: (row) => exactNumber(row.my_share_amount_pct) },
      { label: "Comp", render: (row) => formatPct(row.competitor_share_amount_pct), title: (row) => exactNumber(row.competitor_share_amount_pct) },
      { label: "Gap", render: (row) => toneDelta(row.share_gap_amount_pct), title: (row) => exactNumber(row.share_gap_amount_pct) },
    ],
    overview.weak_vs_competitors,
    { compact: true }
  );

  initLineCharts(document.getElementById("tab-overview"));
}

function renderMasterfile() {
  const payload = state.master;
  els["masterfile-meta"].innerHTML = [
    metaPill(`Period: ${payload.meta.start_month} - ${payload.meta.end_month}`),
    metaPill(`Latest month: ${payload.meta.latest_month}`),
    metaPill(`SKU: ${formatNumber(payload.meta.unique_sku_count)}`),
    metaPill(`Revenue: ${formatMoney(payload.meta.total_sale_amount)}`),
    metaPill(`Qty: ${formatNumber(payload.meta.total_sale_qty)}`),
  ].join("");

  renderTable(
    els["masterfile-table"],
    [
      { label: "Cohort", render: (row) => escapeHtml(row.cohort), title: (row) => row.cohort },
      { label: "SKU", render: (row) => `<span class="mono">${escapeHtml(row.sku)}</span>`, title: (row) => row.sku },
      {
        label: "Product",
        render: (row) =>
          row.product_url
            ? `<a href="${escapeHtml(row.product_url)}" target="_blank" rel="noreferrer">${escapeHtml(row.product_name)}</a>`
            : escapeHtml(row.product_name),
        title: (row) => row.product_name,
      },
      { label: "Brand", render: (row) => escapeHtml(row.brand_name), title: (row) => row.brand_name },
      { label: "L2", render: (row) => escapeHtml(row.level_2_name), title: (row) => row.level_2_name },
      { label: "L3", render: (row) => escapeHtml(row.level_3_name || ""), title: (row) => row.level_3_name || "" },
      { label: "Price start", render: (row) => formatMoney(row.price_first_value), title: (row) => exactNumber(row.price_first_value) },
      { label: "Price latest", render: (row) => formatMoney(row.price_last_value), title: (row) => exactNumber(row.price_last_value) },
      { label: "Price Δ%", render: (row) => tonePct(row.price_change_pct), title: (row) => exactNumber(row.price_change_pct) },
      {
        label: "Price min/max",
        render: (row) => `${formatMoney(row.price_min_value)} / ${formatMoney(row.price_max_value)}`,
        title: (row) => `${exactNumber(row.price_min_value)} / ${exactNumber(row.price_max_value)}`,
      },
      { label: "Period revenue", render: (row) => formatMoney(row.selected_period_sale_amount), title: (row) => exactNumber(row.selected_period_sale_amount) },
      { label: "Period qty", render: (row) => formatNumber(row.selected_period_sale_qty), title: (row) => exactNumber(row.selected_period_sale_qty) },
      { label: "Latest revenue", render: (row) => formatMoney(row.latest_month_sale_amount), title: (row) => exactNumber(row.latest_month_sale_amount) },
      { label: "Latest price", render: (row) => formatMoney(row.latest_month_sale_price), title: (row) => exactNumber(row.latest_month_sale_price) },
      { label: "Active months", render: (row) => formatNumber(row.active_months_qty), title: (row) => exactNumber(row.active_months_qty) },
      { label: "Share", render: (row) => formatPct(row.selected_share_pct), title: (row) => exactNumber(row.selected_share_pct) },
    ],
    payload.rows,
    { wide: true }
  );

  els["master-pagination"].textContent = formatPagination(payload.pagination);
  els["master-prev"].disabled = payload.pagination.page <= 1;
  els["master-next"].disabled = payload.pagination.page >= payload.pagination.total_pages;
}

function renderRaw() {
  const payload = state.raw;
  els["raw-meta"].innerHTML = [
    metaPill(`Rows: ${formatNumber(payload.pagination.total)}`),
    metaPill(`Revenue: ${formatMoney(payload.meta.total_sale_amount)}`),
    metaPill(`Qty: ${formatNumber(payload.meta.total_sale_qty)}`),
  ].join("");

  renderTable(
    els["raw-table"],
    [
      { label: "Month", render: (row) => escapeHtml(row.month), title: (row) => row.month },
      { label: "Cohort", render: (row) => escapeHtml(row.cohort), title: (row) => row.cohort },
      { label: "SKU", render: (row) => `<span class="mono">${escapeHtml(row.sku)}</span>`, title: (row) => row.sku },
      { label: "Product", render: (row) => escapeHtml(row.product_name), title: (row) => row.product_name },
      { label: "Brand", render: (row) => escapeHtml(row.brand_name), title: (row) => row.brand_name },
      { label: "L2", render: (row) => escapeHtml(row.level_2_name), title: (row) => row.level_2_name },
      { label: "L3", render: (row) => escapeHtml(row.level_3_name || ""), title: (row) => row.level_3_name || "" },
      { label: "Revenue", render: (row) => formatMoney(row.sale_amount), title: (row) => exactNumber(row.sale_amount) },
      { label: "Qty", render: (row) => formatNumber(row.sale_qty), title: (row) => exactNumber(row.sale_qty) },
      { label: "Sellers", render: (row) => formatNumber(row.merchant_count), title: (row) => exactNumber(row.merchant_count) },
      { label: "Reviews", render: (row) => formatNumber(row.review_qty), title: (row) => exactNumber(row.review_qty) },
    ],
    payload.rows,
    { wide: true }
  );

  els["raw-pagination"].textContent = formatPagination(payload.pagination);
  els["raw-prev"].disabled = payload.pagination.page <= 1;
  els["raw-next"].disabled = payload.pagination.page >= payload.pagination.total_pages;
}

function renderCompare() {
  const payload = state.compare;
  const granularity = selectedGranularity("compare-granularity", "month");
  els["compare-chart"].innerHTML = payload.featured_trend.length
    ? payload.featured_trend
        .map(
          (item) => {
            const seriesRows = aggregateShareTrend(item.series || [], granularity);
            return `
            <article class="category-card">
              <div class="category-card__title">${escapeHtml(item.category_name)}</div>
              ${buildLineChartMarkup({
                labels: seriesRows.map((row) => row.label),
                compact: true,
                formatter: (value) => `${round1(value)}%`,
                tooltipFormatter: (value) => `${exactNumber(value)}%`,
                series: [
                  {
                    name: "My share",
                    color: "#146356",
                    values: seriesRows.map((row) => row.my_share_amount_pct),
                  },
                  {
                    name: "Competitor share",
                    color: "#c06014",
                    values: seriesRows.map((row) => row.competitor_share_amount_pct),
                  },
                ],
                extraRows: seriesRows.map((row) =>
                  buildTooltipRows([
                    { label: "Market revenue", value: formatMoney(row.category_sale_amount), emphasis: true },
                    { label: "My revenue", value: formatMoney(row.my_sale_amount) },
                    { label: "Competitor revenue", value: formatMoney(row.competitor_sale_amount) },
                    {
                      label: "Gap",
                      value: `${row.my_share_amount_pct - row.competitor_share_amount_pct > 0 ? "+" : ""}${exactNumber(
                        row.my_share_amount_pct - row.competitor_share_amount_pct
                      )} pp`,
                    },
                  ])
                ),
              })}
            </article>
          `;
          }
        )
        .join("")
    : emptyState("Нет данных для category compare");

  renderTable(
    els["compare-table"],
    [
      { label: "Category", render: (row) => escapeHtml(row.category_name), title: (row) => row.category_name },
      { label: "Path", render: (row) => escapeHtml(row.full_path), title: (row) => row.full_path },
      { label: "Market revenue", render: (row) => formatMoney(row.selected_period_market_revenue), title: (row) => exactNumber(row.selected_period_market_revenue) },
      { label: "My revenue", render: (row) => formatMoney(row.selected_period_my_revenue), title: (row) => exactNumber(row.selected_period_my_revenue) },
      { label: "My share", render: (row) => formatPct(row.selected_period_my_share_pct), title: (row) => exactNumber(row.selected_period_my_share_pct) },
      { label: "Comp revenue", render: (row) => formatMoney(row.selected_period_competitor_revenue), title: (row) => exactNumber(row.selected_period_competitor_revenue) },
      { label: "Comp share", render: (row) => formatPct(row.selected_period_competitor_share_pct), title: (row) => exactNumber(row.selected_period_competitor_share_pct) },
      { label: "Gap", render: (row) => toneDelta(row.share_gap_amount_pct), title: (row) => exactNumber(row.share_gap_amount_pct) },
    ],
    payload.rows,
    { wide: true }
  );

  els["compare-pagination"].textContent = formatPagination(payload.pagination);
  els["compare-prev"].disabled = payload.pagination.page <= 1;
  els["compare-next"].disabled = payload.pagination.page >= payload.pagination.total_pages;
  initLineCharts(document.getElementById("tab-category-compare"));
}

function renderSkuAnalysis() {
  const payload = state.skuAnalysis;
  if (!payload) {
    return;
  }
  const granularity = selectedGranularity("sku-granularity", "month");

  const myCard = payload.summary_cards?.my || {};
  const compareCard = payload.summary_cards?.compare || {};
  const compareSummary = payload.compare_summary || {};
  const priceRows = aggregateSkuDailyRows(payload.price_trend || [], granularity);
  const periodRows =
    granularity === "day" || granularity === "week"
      ? priceRows
      : aggregateSkuMonthlyRows(payload, granularity);

  setSkuInputValue(els["sku-my-input"], "my", payload.selected?.my_sku);
  setSkuInputValue(els["sku-compare-input"], "competitor", payload.selected?.compare_sku);

  els["sku-meta"].innerHTML = [
    metaPill(`Period: ${payload.period.start_month} - ${payload.period.end_month}`),
    metaPill(`Granularity: ${granularity}`),
    metaPill(`My SKU: ${payload.selected.my_sku || "-"}`),
    metaPill(`Compare SKU: ${payload.selected.compare_sku || "-"}`),
    metaPill(`Revenue gap: ${formatMoney(compareSummary.revenue_gap_abs || 0)}`),
    metaPill(`Price gap: ${formatMoney(compareSummary.price_gap_abs || 0)}`),
    metaPill(`Price gap %: ${round1(compareSummary.price_gap_pct || 0)}%`),
  ].join("");

  els["sku-summary"].innerHTML = [
    buildSkuCardMarkup(myCard, "My SKU", "success"),
    buildSkuCardMarkup(compareCard, "Compare SKU", "warning"),
  ].join("");

  els["sku-price-chart"].innerHTML = priceRows.length
    ? buildLineChartMarkup({
        labels: priceRows.map((row) => row.label),
        formatter: formatMoney,
        tooltipFormatter: formatMoney,
        series: [
          {
            name: "My price",
            color: "#146356",
            values: priceRows.map((row) => row.my_sale_price),
          },
          {
            name: "Compare price",
            color: "#c06014",
            values: priceRows.map((row) => row.compare_sale_price),
          },
        ],
        extraRows: priceRows.map((row) =>
          buildTooltipRows([
            { label: "My revenue", value: formatMoney(row.my_sale_amount) },
            { label: "Compare revenue", value: formatMoney(row.compare_sale_amount) },
            { label: "My qty", value: formatNumber(row.my_sale_qty), emphasis: true },
            { label: "Compare qty", value: formatNumber(row.compare_sale_qty), emphasis: true },
            { label: "My sellers", value: formatNumber(row.my_merchant_count) },
            { label: "Compare sellers", value: formatNumber(row.compare_merchant_count) },
            { label: "My reviews", value: formatNumber(row.my_review_qty) },
            { label: "Compare reviews", value: formatNumber(row.compare_review_qty) },
          ])
        ),
      })
    : emptyState("Нет данных по динамике цены");

  els["sku-revenue-chart"].innerHTML = periodRows.length
    ? buildLineChartMarkup({
        labels: periodRows.map((row) => row.label),
        formatter: shortMoney,
        tooltipFormatter: formatMoney,
        series: [
          { name: "My revenue", color: "#146356", values: periodRows.map((row) => row.my_sale_amount) },
          { name: "Compare revenue", color: "#c06014", values: periodRows.map((row) => row.compare_sale_amount) },
        ],
        extraRows: periodRows.map((row) =>
          buildTooltipRows([
            { label: "Revenue gap", value: signedMoneyText(row.revenue_gap), emphasis: true },
            { label: "My qty", value: formatNumber(row.my_sale_qty) },
            { label: "Compare qty", value: formatNumber(row.compare_sale_qty) },
            { label: "My price", value: formatMoney(row.my_sale_price) },
            { label: "Compare price", value: formatMoney(row.compare_sale_price) },
          ])
        ),
      })
    : emptyState("Нет данных по выручке");

  els["sku-qty-chart"].innerHTML = periodRows.length
    ? buildLineChartMarkup({
        labels: periodRows.map((row) => row.label),
        formatter: formatNumber,
        tooltipFormatter: formatNumber,
        series: [
          { name: "My qty", color: "#146356", values: periodRows.map((row) => row.my_sale_qty) },
          { name: "Compare qty", color: "#c06014", values: periodRows.map((row) => row.compare_sale_qty) },
        ],
        extraRows: periodRows.map((row) =>
          buildTooltipRows([
            { label: "My revenue", value: formatMoney(row.my_sale_amount), emphasis: true },
            { label: "Compare revenue", value: formatMoney(row.compare_sale_amount), emphasis: true },
            { label: "My price", value: formatMoney(row.my_sale_price) },
            { label: "Compare price", value: formatMoney(row.compare_sale_price) },
            { label: "My sellers", value: formatNumber(row.my_merchant_count) },
            { label: "Compare sellers", value: formatNumber(row.compare_merchant_count) },
          ])
        ),
      })
    : emptyState("Нет данных по продажам");

  renderTable(
    els["sku-compare-table"],
    [
      { label: "Period", render: (row) => escapeHtml(row.label), title: (row) => row.label },
      { label: "My revenue", render: (row) => formatMoney(row.my_sale_amount), title: (row) => exactNumber(row.my_sale_amount) },
      { label: "Comp revenue", render: (row) => formatMoney(row.compare_sale_amount), title: (row) => exactNumber(row.compare_sale_amount) },
      { label: "Revenue gap", render: (row) => toneMoney(row.revenue_gap), title: (row) => exactNumber(row.revenue_gap) },
      { label: "My qty", render: (row) => formatNumber(row.my_sale_qty), title: (row) => exactNumber(row.my_sale_qty) },
      { label: "Comp qty", render: (row) => formatNumber(row.compare_sale_qty), title: (row) => exactNumber(row.compare_sale_qty) },
      { label: "My price", render: (row) => formatMoney(row.my_sale_price), title: (row) => exactNumber(row.my_sale_price) },
      { label: "Comp price", render: (row) => formatMoney(row.compare_sale_price), title: (row) => exactNumber(row.compare_sale_price) },
      { label: "Price gap", render: (row) => toneMoney(row.price_gap), title: (row) => exactNumber(row.price_gap) },
      { label: "My sellers", render: (row) => formatNumber(row.my_merchant_count), title: (row) => exactNumber(row.my_merchant_count) },
      { label: "Comp sellers", render: (row) => formatNumber(row.compare_merchant_count), title: (row) => exactNumber(row.compare_merchant_count) },
      { label: "My reviews", render: (row) => formatNumber(row.my_review_qty), title: (row) => exactNumber(row.my_review_qty) },
      { label: "Comp reviews", render: (row) => formatNumber(row.compare_review_qty), title: (row) => exactNumber(row.compare_review_qty) },
    ],
    periodRows || [],
    { wide: true }
  );

  initLineCharts(document.getElementById("tab-sku-analysis"));
}

function buildSkuCardMarkup(card, label, tone) {
  if (!card || !card.sku) {
    return emptyState(`Нет данных для ${label}`);
  }
  return `
    <article class="category-card category-card--${escapeHtml(tone)}">
      <div class="category-card__eyebrow">${escapeHtml(label)}</div>
      <div class="category-card__title">${escapeHtml(card.sku)} - ${escapeHtml(card.product_name)}</div>
      <div class="category-card__subtitle">${escapeHtml(card.brand_name)} · ${escapeHtml(card.level_2_name)} · ${escapeHtml(card.level_3_name || "")}</div>
      <div class="category-card__stats">
        <div class="category-card__stat"><span>Period revenue</span><strong class="mono">${formatMoney(card.period_sale_amount)}</strong></div>
        <div class="category-card__stat"><span>Period qty</span><strong class="mono">${formatNumber(card.period_sale_qty)}</strong></div>
        <div class="category-card__stat"><span>Latest price</span><strong class="mono">${formatMoney(card.price_last_value)}</strong></div>
        <div class="category-card__stat"><span>Price change</span><strong class="mono">${round1(card.price_change_pct)}%</strong></div>
        <div class="category-card__stat"><span>Min / max</span><strong class="mono">${formatMoney(card.price_min_value)} / ${formatMoney(card.price_max_value)}</strong></div>
        <div class="category-card__stat"><span>Price days</span><strong class="mono">${formatNumber(card.price_days_qty)}</strong></div>
      </div>
    </article>
  `;
}

function buildDownloadUrl() {
  const dataset = els["download-dataset"].value;
  const fileFormat = els["download-format"].value;
  return `/algatop/download?${new URLSearchParams({
    dataset,
    file_format: fileFormat,
    ...globalParams(),
    master_sort: els["master-sort"].value,
    raw_sort: els["raw-sort"].value,
    compare_level: els["compare-level"].value,
    compare_search_query: els["compare-search"].value.trim(),
    compare_sort: els["compare-sort"].value,
  })}`;
}

function renderTable(container, columns, rows, options = {}) {
  if (!rows || !rows.length) {
    container.innerHTML = emptyState("Нет данных по текущему фильтру");
    return;
  }

  container.innerHTML = `
    <div class="table-wrap ${options.compact ? "compact-table" : ""}">
      <table>
        <thead>
          <tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr>
                  ${columns
                    .map((column) => {
                      const title = escapeHtml(
                        String(
                          column.title ? column.title(row) : stripHtml(column.render(row)) || ""
                        )
                      );
                      return `<td title="${title}">${column.render(row)}</td>`;
                    })
                    .join("")}
                </tr>
              `
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function buildLineChartMarkup(config) {
  const labels = config.labels || config.months || [];
  const series = config.series || [];
  const formatter = config.formatter || shortMoney;
  const tooltipFormatter = config.tooltipFormatter || formatter;
  const extraRows = config.extraRows || [];
  const compact = Boolean(config.compact);
  const showPoints = config.showPoints !== false;

  if (!labels.length || !series.length) {
    return emptyState("Нет данных для графика");
  }

  const values = series.flatMap((item) => (item.values || []).map((value) => numeric(value)));
  let maxValue = Math.max(...values, 0);
  let minValue = Math.min(...values, 0);
  if (maxValue === minValue) {
    const padding = Math.max(Math.abs(maxValue) * 0.15, 1);
    maxValue += padding;
    minValue -= padding;
  }
  const range = Math.max(maxValue - minValue, 1);
  const width = 760;
  const height = compact ? 180 : 300;
  const margin = compact
    ? { top: 18, right: 18, bottom: 30, left: 42 }
    : { top: 24, right: 22, bottom: 38, left: 52 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const labelStride = labels.length > 8 ? Math.ceil(labels.length / (compact ? 5 : 7)) : 1;
  const yTicks = 4;

  function x(index) {
    if (labels.length === 1) {
      return margin.left + plotWidth / 2;
    }
    return margin.left + (plotWidth * index) / (labels.length - 1);
  }

  function y(value) {
    return margin.top + ((maxValue - value) / range) * plotHeight;
  }

  function pathFor(valuesList) {
    return valuesList
      .map((value, index) => `${index === 0 ? "M" : "L"} ${x(index).toFixed(2)} ${y(value).toFixed(2)}`)
      .join(" ");
  }

  const grid = Array.from({ length: yTicks + 1 }, (_, index) => {
    const tickValue = maxValue - (range * index) / yTicks;
    const lineY = y(tickValue);
    return `
      <line class="line-chart__grid" x1="${margin.left}" y1="${lineY}" x2="${width - margin.right}" y2="${lineY}"></line>
      <text class="line-chart__label" x="6" y="${lineY + 4}">${escapeHtml(formatter(tickValue))}</text>
    `;
  }).join("");

  const xLabels = labels
    .map((label, index) =>
      index % labelStride === 0 || index === labels.length - 1
        ? `<text class="line-chart__label" x="${x(index)}" y="${height - 8}" text-anchor="middle">${escapeHtml(label)}</text>`
        : ""
    )
    .join("");

  const seriesMarkup = series
    .map(
      (item) => `
        <path class="line-chart__path" d="${pathFor(item.values)}" style="stroke:${item.color}"></path>
        ${showPoints
          ? item.values
          .map(
            (value, index) => `
              <circle
                class="line-chart__point"
                cx="${x(index)}"
                cy="${y(value)}"
                r="${compact ? 3 : 4}"
                fill="${item.color}"
              ></circle>
            `
          )
          .join("")
          : ""}
      `
    )
    .join("");

  const tooltipPayload = labels.map((label, index) => ({
    title: typeof config.titleFormatter === "function" ? config.titleFormatter(label, index) : label,
    rows: [
      ...series.map((item) => ({
        label: item.name,
        value: String(tooltipFormatter(item.values[index], item, index)),
        color: item.color,
      })),
      ...(extraRows[index] || []).map((row) => ({
        label: row.label,
        value: String(row.value),
        color: row.color || "",
        emphasis: Boolean(row.emphasis),
      })),
    ],
  }));

  const hitSlices = labels
    .map((_label, index) => {
      const left = index === 0 ? margin.left : (x(index - 1) + x(index)) / 2;
      const right = index === labels.length - 1 ? width - margin.right : (x(index) + x(index + 1)) / 2;
      return `
        <rect
          class="line-chart__hit-slice"
          x="${left}"
          y="${margin.top}"
          width="${Math.max(right - left, 8)}"
          height="${plotHeight}"
          data-index="${index}"
          tabindex="0"
        ></rect>
      `;
    })
    .join("");

  const legend = series
    .map(
      (item) => `
        <span class="legend-item">
          <span class="legend-swatch" style="background:${item.color}"></span>
          ${escapeHtml(item.name)}
        </span>
      `
    )
    .join("");

  return `
    <div class="line-chart">
      <div class="line-chart__legend">${legend}</div>
      <div class="line-chart__stage">
        <svg class="line-chart__frame" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
          ${grid}
          <line class="line-chart__axis" x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}"></line>
          <line class="line-chart__axis" x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}"></line>
          <line class="line-chart__hover-line" x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}" opacity="0"></line>
          ${seriesMarkup}
          ${xLabels}
          ${hitSlices}
        </svg>
        <div class="line-chart__tooltip" hidden></div>
      </div>
      <script type="application/json" class="line-chart__payload">${JSON.stringify(tooltipPayload).replaceAll("<", "\\u003c")}</script>
    </div>
  `;
}

function metaPill(text) {
  return `<span class="meta-pill">${escapeHtml(text)}</span>`;
}

function toneDelta(value) {
  const numeric = Number(value || 0);
  const tone = numeric > 0 ? "up" : numeric < 0 ? "down" : "flat";
  const sign = numeric > 0 ? "+" : "";
  return `<span class="tone-${tone} mono">${sign}${round1(numeric)} pp</span>`;
}

function tonePct(value) {
  const numeric = Number(value || 0);
  const tone = numeric > 0 ? "up" : numeric < 0 ? "down" : "flat";
  const sign = numeric > 0 ? "+" : "";
  return `<span class="tone-${tone} mono">${sign}${round1(numeric)}%</span>`;
}

function toneMoney(value) {
  const numeric = Number(value || 0);
  const tone = numeric > 0 ? "up" : numeric < 0 ? "down" : "flat";
  const sign = numeric > 0 ? "+" : numeric < 0 ? "-" : "";
  return `<span class="tone-${tone} mono">${sign}${formatMoney(Math.abs(numeric))}</span>`;
}

function signedMoneyText(value) {
  const numeric = Number(value || 0);
  const sign = numeric > 0 ? "+" : numeric < 0 ? "-" : "";
  return `${sign}${formatMoney(Math.abs(numeric))}`;
}

function formatPagination(pagination) {
  return `Page ${pagination.page} / ${pagination.total_pages} · ${formatNumber(pagination.total)} rows`;
}

function renderLoadingAll() {
  [
    els["overview-trend"],
    els["overview-share-trend"],
    els["overview-level2"],
    els["overview-level3"],
    els["overview-brands"],
    els["overview-featured-categories"],
    els["overview-best-gap"],
    els["overview-weak-gap"],
    els["masterfile-table"],
    els["sku-summary"],
    els["sku-price-chart"],
    els["sku-revenue-chart"],
    els["sku-qty-chart"],
    els["sku-compare-table"],
    els["raw-table"],
    els["compare-chart"],
    els["compare-table"],
  ].forEach(showLoading);
  els["overview-metrics"].innerHTML = "";
  els["sku-meta"].innerHTML = "";
}

function renderErrorAll(error) {
  [
    els["overview-trend"],
    els["overview-share-trend"],
    els["overview-level2"],
    els["overview-level3"],
    els["overview-brands"],
    els["overview-featured-categories"],
    els["overview-best-gap"],
    els["overview-weak-gap"],
    els["masterfile-table"],
    els["sku-summary"],
    els["sku-price-chart"],
    els["sku-revenue-chart"],
    els["sku-qty-chart"],
    els["sku-compare-table"],
    els["raw-table"],
    els["compare-chart"],
    els["compare-table"],
  ].forEach((element) => showError(element, error));
}

function showLoading(container) {
  container.innerHTML = `<div class="loading-state">Загрузка данных...</div>`;
}

function showError(container, error) {
  container.innerHTML = `<div class="empty-state">Ошибка: ${escapeHtml(error.message || String(error))}</div>`;
}

function emptyState(text) {
  return `<div class="empty-state">${escapeHtml(text)}</div>`;
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    cache: "no-store",
    ...options,
  });
  if (!response.ok) {
    let detail = response.statusText;
    try {
      const payload = await response.json();
      detail = payload.detail || JSON.stringify(payload);
    } catch (_error) {
      detail = response.statusText;
    }
    throw new Error(detail || `HTTP ${response.status}`);
  }
  return response.json();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function stripHtml(value) {
  return String(value || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

function formatValue(value, label) {
  if (/revenue/i.test(label)) {
    return formatMoney(value);
  }
  return formatNumber(value);
}

function formatMoney(value) {
  return new Intl.NumberFormat("ru-RU", {
    maximumFractionDigits: 0,
  }).format(Number(value || 0));
}

function formatNumber(value) {
  return new Intl.NumberFormat("ru-RU", {
    maximumFractionDigits: Number.isInteger(Number(value || 0)) ? 0 : 2,
  }).format(Number(value || 0));
}

function exactNumber(value) {
  return new Intl.NumberFormat("ru-RU", {
    maximumFractionDigits: 4,
  }).format(Number(value || 0));
}

function formatPct(value) {
  return `${round1(value)}%`;
}

function round1(value) {
  return Number(value || 0).toFixed(1);
}

function shortMoney(value) {
  const numeric = Number(value || 0);
  if (Math.abs(numeric) >= 1_000_000_000) {
    return `${(numeric / 1_000_000_000).toFixed(1)}B`;
  }
  if (Math.abs(numeric) >= 1_000_000) {
    return `${(numeric / 1_000_000).toFixed(1)}M`;
  }
  if (Math.abs(numeric) >= 1_000) {
    return `${(numeric / 1_000).toFixed(0)}K`;
  }
  return formatNumber(numeric);
}
