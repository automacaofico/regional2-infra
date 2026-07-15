(function () {
  "use strict";

  var WORKBOOK_URL = "Report_Regional02_Infra.xlsx";
  var REQUIRED_SHEETS = ["Report Consolidado Semanal", "Gráficos", "Quadros"];
  var PACKAGE_KEYS = ["p04", "p05", "p06", "p0708"];
  var PACKAGE_LABELS = { p04: "Pacote 04", p05: "Pacote 05", p06: "Pacote 06", p0708: "Pacote 07/08" };
  var SELECTED_PACKAGE_LABELS = { p04: "Pacote 04", p05: "Pacote 05", p06: "Pacote 06", p07: "Pacote 07", p08: "Pacote 08", bridge: "Ponte do Araguaia" };
  var ACTIVITY_LABELS = {
    terra: "Terraplanagem",
    drenagem: "Drenagem superficial",
    ponte: "Ponte do Araguaia",
    dsh: "DSH / DHP"
  };

  var state = {
    mode: "weekly",
    package: "p04",
    activity: "terra",
    dayIndex: 0,
    data: null
  };

  var elements = {
    dashboard: document.getElementById("dashboard"),
    loading: document.getElementById("loading-stage"),
    error: document.getElementById("error-state"),
    errorMessage: document.getElementById("error-message"),
    syncState: document.getElementById("sync-state"),
    syncLabel: document.getElementById("sync-label"),
    refreshButton: document.getElementById("refresh-button"),
    fileInput: document.getElementById("file-input"),
    modeSwitch: document.getElementById("mode-switch"),
    packageSwitch: document.getElementById("package-switch"),
    activitySwitch: document.getElementById("activity-switch"),
    disciplineControl: document.getElementById("discipline-control"),
    dataNote: document.getElementById("data-note"),
    dayControl: document.getElementById("day-control"),
    dayPicker: document.getElementById("day-picker"),
    referenceDate: document.getElementById("reference-date"),
    referenceContext: document.getElementById("reference-context"),
    kpis: document.getElementById("kpi-grid"),
    trendKicker: document.getElementById("trend-kicker"),
    trendTitle: document.getElementById("trend-title"),
    trendChart: document.getElementById("trend-chart"),
    attention: document.getElementById("attention-list"),
    frontsKicker: document.getElementById("fronts-kicker"),
    frontsTitle: document.getElementById("fronts-title"),
    frontsNote: document.getElementById("fronts-note"),
    frontsTable: document.getElementById("fronts-table"),
    driversKicker: document.getElementById("drivers-kicker"),
    driversTitle: document.getElementById("drivers-title"),
    drivers: document.getElementById("drivers-content"),
    monthBoard: document.getElementById("month-board"),
    monthTitle: document.getElementById("month-title")
  };

  function isBlank(value) {
    return value === null || value === undefined || (typeof value === "string" && value.trim() === "");
  }

  function cell(sheet, row, column) {
    var address = XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
    return sheet && sheet[address] ? sheet[address].v : null;
  }

  function number(value) {
    if (isBlank(value) || value === "-") return null;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    var normalized = String(value).replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
    var parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function sum(values) {
    var valid = values.filter(function (value) { return value !== null && Number.isFinite(value); });
    return valid.length ? valid.reduce(function (total, value) { return total + value; }, 0) : null;
  }

  function safeDivide(numerator, denominator) {
    return denominator && numerator !== null && Number.isFinite(numerator) ? numerator / denominator : null;
  }

  function excelDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") return new Date(Math.round((value - 25569) * 86400 * 1000));
    if (typeof value === "string") {
      var parsed = new Date(value);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
    return null;
  }

  function formatDate(value, options) {
    var date = excelDate(value);
    if (!date) return "—";
    var formatOptions = Object.assign({ timeZone: "UTC" }, options || { day: "2-digit", month: "long", year: "numeric" });
    return new Intl.DateTimeFormat("pt-BR", formatOptions).format(date);
  }

  function shortDate(value) {
    return formatDate(value, { day: "2-digit", month: "2-digit" });
  }

  function weekDay(value) {
    var text = formatDate(value, { weekday: "short" }).replace(".", "");
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  function periodLabel(value) {
    if (isBlank(value)) return "Semana";
    return String(value)
      .replace(/ACOMPANHAMENTO\s+SEMANAL/gi, "")
      .replace(/PRODUÇÃO\s+TOTAL\s+S-\d/gi, "")
      .trim();
  }

  function formatNumber(value, decimals) {
    if (value === null || !Number.isFinite(value)) return "—";
    var abs = Math.abs(value);
    var digits = decimals;
    if (digits === undefined) digits = abs > 100 ? 0 : abs > 10 ? 1 : 2;
    return new Intl.NumberFormat("pt-BR", {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits
    }).format(value);
  }

  function formatCompact(value) {
    if (value === null || !Number.isFinite(value)) return "—";
    return new Intl.NumberFormat("pt-BR", {
      notation: Math.abs(value) >= 10000 ? "compact" : "standard",
      maximumFractionDigits: 1
    }).format(value);
  }

  function formatPercent(value) {
    if (value === null || !Number.isFinite(value)) return "—";
    return new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 0 }).format(value);
  }

  function escapeHtml(value) {
    return String(value === null || value === undefined ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function readBlock(sheet, plannedRow, metadata) {
    var realRow = plannedRow + 1;
    return {
      name: cell(sheet, plannedRow, 4),
      unit: cell(sheet, plannedRow, 5),
      weeks: {
        labels: metadata.weekLabels.slice(),
        planned: [8, 12, 16, 20].map(function (column) { return number(cell(sheet, plannedRow, column)); }),
        actual: [8, 12, 16, 20].map(function (column) { return number(cell(sheet, realRow, column)); })
      },
      days: {
        dates: metadata.dayDates.slice(),
        planned: [27, 28, 29, 30, 31, 32, 33].map(function (column) { return number(cell(sheet, plannedRow, column)); }),
        actual: [27, 28, 29, 30, 31, 32, 33].map(function (column) { return number(cell(sheet, realRow, column)); }),
        totalPlanned: number(cell(sheet, plannedRow, 34)),
        totalActual: number(cell(sheet, realRow, 34)),
        adherence: number(cell(sheet, plannedRow, 35))
      },
      month: {
        planned: number(cell(sheet, plannedRow, 37)),
        plannedToDate: number(cell(sheet, plannedRow, 38)),
        actual: number(cell(sheet, plannedRow, 39)),
        adherence: number(cell(sheet, plannedRow, 40))
      },
      baseline: {
        actual: number(cell(sheet, plannedRow, 42)),
        planned: number(cell(sheet, plannedRow, 43)),
        adherence: number(cell(sheet, plannedRow, 44))
      },
      scope: {
        total: number(cell(sheet, plannedRow, 46)),
        balance: number(cell(sheet, plannedRow, 47))
      },
      average: {
        planned: number(cell(sheet, plannedRow, 49)),
        actual: number(cell(sheet, plannedRow, 50)),
        needed: number(cell(sheet, plannedRow, 51))
      }
    };
  }

  function parseConsolidated(workbook) {
    var sheet = workbook.Sheets["Report Consolidado Semanal"];
    var metadata = {
      referenceDate: excelDate(cell(sheet, 4, 4)),
      weekLabels: [7, 11, 15, 19].map(function (column) { return periodLabel(cell(sheet, 4, column)); }),
      dayDates: [27, 28, 29, 30, 31, 32, 33].map(function (column) { return excelDate(cell(sheet, 7, column)); })
    };

    return {
      metadata: metadata,
      cut: { total: readBlock(sheet, 9, metadata), p04: readBlock(sheet, 13, metadata), p05: readBlock(sheet, 17, metadata), p06: readBlock(sheet, 21, metadata), p0708: readBlock(sheet, 25, metadata) },
      fill: { total: readBlock(sheet, 29, metadata), p04: readBlock(sheet, 33, metadata), p05: readBlock(sheet, 37, metadata), p06: readBlock(sheet, 41, metadata), p0708: readBlock(sheet, 45, metadata) },
      drainage: { total: readBlock(sheet, 49, metadata), p04: readBlock(sheet, 53, metadata), p05: readBlock(sheet, 57, metadata), p06: readBlock(sheet, 61, metadata), p0708: readBlock(sheet, 65, metadata) },
      bridge: {
        pillars: readBlock(sheet, 73, metadata),
        crossbeams: readBlock(sheet, 77, metadata),
        segments: readBlock(sheet, 81, metadata),
        beams: readBlock(sheet, 85, metadata),
        precast: readBlock(sheet, 89, metadata)
      },
      dsh: { total: readBlock(sheet, 93, metadata), p04: readBlock(sheet, 97, metadata), p05: readBlock(sheet, 101, metadata), p06: readBlock(sheet, 105, metadata), p0708: readBlock(sheet, 109, metadata) }
    };
  }

  function parseTables(workbook) {
    var sheet = workbook.Sheets.Quadros;
    function value(row, column) { return number(cell(sheet, row, column)); }
    return {
      terra: {
        excavationProductivity: value(28, 14),
        excavators: value(29, 14),
        fillProductivity: value(31, 14),
        rollers: value(32, 14),
        groups: {
          p0405: { excavationProductivity: value(28, 10), excavators: value(29, 10), fillProductivity: value(31, 10), rollers: value(32, 10) },
          p06: { excavationProductivity: value(28, 11), excavators: value(29, 11), fillProductivity: value(31, 11), rollers: value(32, 11) },
          p0708: { excavationProductivity: value(28, 12), excavators: value(29, 12), fillProductivity: value(31, 12), rollers: value(32, 12) }
        },
        progress: {
          p04: { target: value(36, 10), actual: value(38, 10), rate: value(40, 10) },
          p06: { target: value(36, 11), actual: value(38, 11), rate: value(40, 11) },
          p0708: { target: value(36, 12), actual: value(38, 12), rate: value(40, 12) },
          total: { target: value(36, 14), actual: value(38, 14), rate: value(40, 14) }
        }
      },
      drainage: {
        productivity: value(28, 25),
        workforce: value(29, 25),
        monthProduction: value(31, 25),
        groups: {
          p0405: { productivity: value(28, 21), workforce: value(29, 21), monthProduction: value(31, 21) },
          p06: { productivity: value(28, 22), workforce: value(29, 22), monthProduction: value(31, 22) },
          p0708: { productivity: value(28, 23), workforce: value(29, 23), monthProduction: value(31, 23) }
        },
        progress: {
          p04: { target: value(36, 21), actual: value(38, 21), rate: value(40, 21) },
          p06: { target: value(36, 22), actual: value(38, 22), rate: value(40, 22) },
          p0708: { target: value(36, 23), actual: value(38, 23), rate: value(40, 23) },
          total: { target: value(36, 25), actual: value(38, 25), rate: value(40, 25) }
        }
      },
      bridge: {
        pillars: { target: value(36, 31), actual: value(38, 31), rate: value(40, 31) },
        crossbeams: { target: value(36, 32), actual: value(38, 32), rate: value(40, 32) },
        segments: { target: value(36, 33), actual: value(38, 33), rate: value(40, 33) },
        beams: { target: value(36, 34), actual: value(38, 34), rate: value(40, 34) },
        precast: { target: value(36, 35), actual: value(38, 35), rate: value(40, 35) }
      }
    };
  }

  function parseCharts(workbook) {
    var sheet = workbook.Sheets["Gráficos"];
    var contributions = [];
    for (var row = 168; row <= 194; row += 1) {
      var name = cell(sheet, row, 4);
      if (isBlank(name)) continue;
      contributions.push({
        name: String(name).trim(),
        weight: number(cell(sheet, row, 2)),
        total: number(cell(sheet, row, 6)),
        planned: number(cell(sheet, row, 7)),
        actual: number(cell(sheet, row, 8)),
        balance: number(cell(sheet, row, 9)),
        plannedRate: number(cell(sheet, row, 10)),
        actualRate: number(cell(sheet, row, 11))
      });
    }
    return { bridgeContributions: contributions };
  }

  function parseWorkbook(buffer) {
    if (!window.XLSX) throw new Error("O leitor de planilhas não foi carregado. Verifique a conexão e tente novamente.");
    var workbook = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true });
    var missing = REQUIRED_SHEETS.filter(function (sheet) { return workbook.SheetNames.indexOf(sheet) < 0; });
    if (missing.length) throw new Error("A planilha não contém: " + missing.join(", ") + ".");
    return {
      consolidated: parseConsolidated(workbook),
      tables: parseTables(workbook),
      charts: parseCharts(workbook)
    };
  }

  function combineBlocks(name, unit, blocks) {
    var first = blocks[0];
    function sumArray(path, length) {
      return Array.from({ length: length }, function (_, index) {
        return sum(blocks.map(function (block) { return path(block)[index]; }));
      });
    }
    var combined = {
      name: name,
      unit: unit || first.unit,
      weeks: {
        labels: first.weeks.labels.slice(),
        planned: sumArray(function (block) { return block.weeks.planned; }, 4),
        actual: sumArray(function (block) { return block.weeks.actual; }, 4)
      },
      days: {
        dates: first.days.dates.slice(),
        planned: sumArray(function (block) { return block.days.planned; }, 7),
        actual: sumArray(function (block) { return block.days.actual; }, 7)
      },
      month: {
        planned: sum(blocks.map(function (block) { return block.month.planned; })),
        plannedToDate: sum(blocks.map(function (block) { return block.month.plannedToDate; })),
        actual: sum(blocks.map(function (block) { return block.month.actual; }))
      },
      baseline: {
        actual: sum(blocks.map(function (block) { return block.baseline.actual; })),
        planned: sum(blocks.map(function (block) { return block.baseline.planned; }))
      },
      scope: {
        total: sum(blocks.map(function (block) { return block.scope.total; })),
        balance: sum(blocks.map(function (block) { return block.scope.balance; }))
      },
      average: {
        planned: sum(blocks.map(function (block) { return block.average.planned; })),
        actual: sum(blocks.map(function (block) { return block.average.actual; })),
        needed: sum(blocks.map(function (block) { return block.average.needed; }))
      }
    };
    combined.days.totalPlanned = sum(blocks.map(function (block) { return block.days.totalPlanned; }));
    combined.days.totalActual = sum(blocks.map(function (block) { return block.days.totalActual; }));
    combined.days.adherence = safeDivide(combined.days.totalActual, combined.days.totalPlanned);
    combined.month.adherence = safeDivide(combined.month.actual, combined.month.plannedToDate);
    combined.baseline.adherence = safeDivide(combined.baseline.actual, combined.baseline.planned);
    return combined;
  }

  function activityModel(activity) {
    var consolidated = state.data.consolidated;
    if (activity === "terra") {
      var terraFronts = PACKAGE_KEYS.map(function (key) {
        return { name: PACKAGE_LABELS[key], block: combineBlocks(PACKAGE_LABELS[key], "m³", [consolidated.cut[key], consolidated.fill[key]]) };
      });
      return {
        key: activity,
        title: ACTIVITY_LABELS[activity],
        unit: "m³",
        total: combineBlocks("Regional 02", "m³", [consolidated.cut.total, consolidated.fill.total]),
        fronts: terraFronts
      };
    }
    if (activity === "drenagem") {
      return {
        key: activity,
        title: ACTIVITY_LABELS[activity],
        unit: "m",
        total: consolidated.drainage.total,
        fronts: PACKAGE_KEYS.map(function (key) { return { name: PACKAGE_LABELS[key], block: consolidated.drainage[key] }; })
      };
    }
    if (activity === "dsh") {
      return {
        key: activity,
        title: ACTIVITY_LABELS[activity],
        unit: "m",
        total: consolidated.dsh.total,
        fronts: PACKAGE_KEYS.map(function (key) { return { name: PACKAGE_LABELS[key], block: consolidated.dsh[key] }; })
      };
    }
    var bridge = consolidated.bridge;
    var components = [
      { name: "Pilares", block: bridge.pillars },
      { name: "Travessas", block: bridge.crossbeams },
      { name: "Aduelas", block: bridge.segments },
      { name: "Vigas pré-moldadas", block: bridge.beams },
      { name: "Pré-moldados", block: bridge.precast }
    ];
    return {
      key: activity,
      title: ACTIVITY_LABELS[activity],
      unit: "etapas",
      total: combineBlocks("Ponte do Araguaia", "etapas", components.map(function (item) { return item.block; })),
      fronts: components
    };
  }

  function canonicalPackageKey() {
    if (state.package === "p07" || state.package === "p08") return "p0708";
    return state.package;
  }

  function packageDisciplineBlocks(packageKey) {
    var consolidated = state.data.consolidated;
    return {
      terra: combineBlocks(SELECTED_PACKAGE_LABELS[state.package], "m³", [consolidated.cut[packageKey], consolidated.fill[packageKey]]),
      drenagem: consolidated.drainage[packageKey],
      dsh: consolidated.dsh[packageKey]
    };
  }

  function packageModel() {
    if (state.package === "bridge") return activityModel("ponte");
    var packageKey = canonicalPackageKey();
    var blocks = packageDisciplineBlocks(packageKey);
    var units = { terra: "m³", drenagem: "m", dsh: "m" };
    return {
      key: state.activity,
      packageKey: packageKey,
      packageLabel: SELECTED_PACKAGE_LABELS[state.package],
      title: SELECTED_PACKAGE_LABELS[state.package] + " · " + ACTIVITY_LABELS[state.activity],
      unit: units[state.activity],
      total: blocks[state.activity],
      fronts: [
        { name: "Terraplanagem", block: blocks.terra, unit: "m³" },
        { name: "Drenagem superficial", block: blocks.drenagem, unit: "m" },
        { name: "DSH / DHP", block: blocks.dsh, unit: "m" }
      ]
    };
  }

  function periodMetric(block) {
    if (state.mode === "daily") {
      var planned = block.days.planned[state.dayIndex];
      var actual = block.days.actual[state.dayIndex];
      return {
        planned: planned,
        actual: actual,
        adherence: safeDivide(actual, planned),
        variance: planned === null || actual === null ? null : actual - planned
      };
    }
    var weekPlanned = block.days.totalPlanned !== undefined ? block.days.totalPlanned : sum(block.days.planned);
    var weekActual = block.days.totalActual !== undefined ? block.days.totalActual : sum(block.days.actual);
    return {
      planned: weekPlanned,
      actual: weekActual,
      adherence: safeDivide(weekActual, weekPlanned),
      variance: weekPlanned === null || weekActual === null ? null : weekActual - weekPlanned
    };
  }

  function statusClass(rate) {
    if (rate === null || !Number.isFinite(rate)) return "is-empty";
    if (rate >= 1) return "is-good";
    if (rate >= 0.85) return "is-warn";
    return "is-bad";
  }

  function statusLabel(rate) {
    if (rate === null || !Number.isFinite(rate)) return "Sem base";
    if (rate >= 1) return "No ritmo";
    if (rate >= 0.85) return "Atenção";
    return "Crítico";
  }

  function defaultDayIndex() {
    var dates = state.data.consolidated.metadata.dayDates;
    var reference = state.data.consolidated.metadata.referenceDate;
    if (!reference) return 0;
    var referenceDay = Date.UTC(reference.getUTCFullYear(), reference.getUTCMonth(), reference.getUTCDate());
    var best = 0;
    dates.forEach(function (value, index) {
      if (!value) return;
      var day = Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate());
      if (day <= referenceDay) best = index;
    });
    return best;
  }

  function isSelectedDayFuture() {
    var metadata = state.data.consolidated.metadata;
    var selected = metadata.dayDates[state.dayIndex];
    var reference = metadata.referenceDate;
    if (!selected || !reference) return false;
    var selectedDay = Date.UTC(selected.getUTCFullYear(), selected.getUTCMonth(), selected.getUTCDate());
    var referenceDay = Date.UTC(reference.getUTCFullYear(), reference.getUTCMonth(), reference.getUTCDate());
    return selectedDay > referenceDay;
  }

  function renderDayPicker() {
    var dates = state.data.consolidated.metadata.dayDates;
    elements.dayPicker.innerHTML = dates.map(function (date, index) {
      var active = index === state.dayIndex;
      return '<button type="button" class="day-picker__button' + (active ? " is-active" : "") + '" data-day="' + index + '" aria-pressed="' + active + '">' +
        "<b>" + escapeHtml(weekDay(date)) + "</b>" +
        "<span>" + escapeHtml(shortDate(date)) + "</span>" +
      "</button>";
    }).join("");
  }

  function kpiCard(label, value, caption, className, deltaText, deltaClass) {
    return '<article class="kpi-card ' + (className || "") + '">' +
      '<span class="kpi-card__label">' + escapeHtml(label) + "</span>" +
      '<div class="kpi-card__value">' + escapeHtml(value) + "</div>" +
      '<div class="kpi-card__caption">' + escapeHtml(caption) + "</div>" +
      (deltaText ? '<span class="delta ' + (deltaClass || "") + '">' + escapeHtml(deltaText) + "</span>" : "") +
    "</article>";
  }

  function renderKpis(model) {
    var metric = periodMetric(model.total);
    var varianceRate = safeDivide(metric.variance, metric.planned);
    var monthRate = model.total.month.adherence;
    var scopeRate = safeDivide(model.total.baseline.actual, model.total.scope.total);
    var modeLabel = state.mode === "weekly" ? "na semana corrente" : "no dia selecionado";
    var future = state.mode === "daily" && isSelectedDayFuture();

    elements.kpis.innerHTML =
      kpiCard("Planejado", formatNumber(metric.planned) + " " + model.unit, modeLabel, "", state.mode === "weekly" ? "7 dias de programação" : weekDay(model.total.days.dates[state.dayIndex]), "") +
      kpiCard("Realizado", future ? "—" : formatNumber(metric.actual) + " " + model.unit, future ? "Data ainda não realizada" : modeLabel, future ? "" : statusClass(metric.adherence), future ? "Sem apontamento" : statusLabel(metric.adherence), future ? "" : statusClass(metric.adherence)) +
      kpiCard("Aderência", future ? "—" : formatPercent(metric.adherence), future ? "Aguardando execução" : "realizado ÷ planejado", future ? "" : statusClass(metric.adherence), future || varianceRate === null ? "" : (metric.variance >= 0 ? "+" : "") + formatPercent(varianceRate) + " de desvio", metric.variance >= 0 ? "is-positive" : "is-negative") +
      kpiCard("Aderência no mês", formatPercent(monthRate), "realizado ÷ planejado até a data", statusClass(monthRate), "Avanço físico: " + formatPercent(scopeRate), statusClass(monthRate));
  }

  function renderBarChart(model) {
    var labels;
    var planned;
    var actual;
    if (state.mode === "weekly") {
      labels = model.total.weeks.labels;
      planned = model.total.weeks.planned;
      actual = model.total.weeks.actual;
      elements.trendKicker.textContent = "Histórico de quatro semanas";
      elements.trendTitle.textContent = model.title + " · produção semanal";
    } else {
      labels = model.total.days.dates.map(function (date) { return weekDay(date) + " " + shortDate(date); });
      planned = model.total.days.planned;
      actual = model.total.days.actual;
      elements.trendKicker.textContent = "Semana corrente por dia";
      elements.trendTitle.textContent = model.title + " · ritmo diário";
    }
    var maxValue = Math.max.apply(null, planned.concat(actual).filter(function (value) { return value !== null && Number.isFinite(value); }).concat([1]));
    var points = labels.map(function (label, index) {
      var p = planned[index] || 0;
      var a = actual[index] || 0;
      var rate = safeDivide(a, p);
      var planHeight = Math.max(1, p / maxValue * 215);
      var actualHeight = Math.max(1, a / maxValue * 215);
      return '<div class="bar-chart__point">' +
        '<div class="bar-chart__bars">' +
          '<div class="bar-chart__bar bar-chart__bar--plan" style="height:' + planHeight + 'px" title="Planejado: ' + escapeHtml(formatNumber(p)) + '"><span>' + escapeHtml(formatCompact(p)) + "</span></div>" +
          '<div class="bar-chart__bar bar-chart__bar--real" style="height:' + actualHeight + 'px" title="Realizado: ' + escapeHtml(formatNumber(a)) + '"><span>' + escapeHtml(formatCompact(a)) + "</span></div>" +
        "</div>" +
        '<div class="bar-chart__label">' + escapeHtml(label || "—") + "<b>" + escapeHtml(formatPercent(rate)) + "</b></div>" +
      "</div>";
    }).join("");
    elements.trendChart.innerHTML = '<div class="bar-chart" style="--points:' + labels.length + '" role="img" aria-label="Comparação entre produção planejada e realizada">' + points + "</div>";
  }

  function renderAttention(model) {
    var future = state.mode === "daily" && isSelectedDayFuture();
    if (future) {
      elements.attention.innerHTML = '<div class="empty-note">A data selecionada é futura. Use esta visão para conferir a programação das frentes antes do início do turno.</div>';
      return;
    }
    var issues = model.fronts.map(function (front) {
      var metric = periodMetric(front.block);
      return { name: front.name, metric: metric, block: front.block, unit: front.unit || front.block.unit || model.unit };
    }).filter(function (item) {
      return item.metric.planned !== null && item.metric.planned > 0;
    }).sort(function (a, b) {
      return (a.metric.adherence === null ? 9 : a.metric.adherence) - (b.metric.adherence === null ? 9 : b.metric.adherence);
    });

    var monthRate = model.total.month.adherence;
    var cards = issues.slice(0, 3).map(function (item) {
      var cls = statusClass(item.metric.adherence);
      var detail = item.metric.variance < 0
        ? "Faltaram " + formatNumber(Math.abs(item.metric.variance)) + " " + item.unit + " para o plano"
        : "Produção acima do plano em " + formatNumber(item.metric.variance) + " " + item.unit;
      return '<div class="attention-item ' + cls + '">' +
        '<i class="attention-item__signal" aria-hidden="true"></i>' +
        "<div><strong>" + escapeHtml(item.name) + "</strong><span>" + escapeHtml(detail) + "</span></div>" +
        '<b class="attention-item__value">' + escapeHtml(formatPercent(item.metric.adherence)) + "</b>" +
      "</div>";
    });

    if (monthRate !== null && monthRate < 1) {
      cards.push('<div class="attention-item ' + statusClass(monthRate) + '">' +
        '<i class="attention-item__signal" aria-hidden="true"></i>' +
        '<div><strong>Ritmo mensal abaixo do planejado</strong><span>Produção realizada comparada ao plano até a data.</span></div>' +
        '<b class="attention-item__value">' + escapeHtml(formatPercent(monthRate)) + "</b>" +
      "</div>");
    }

    elements.attention.innerHTML = cards.length ? cards.join("") : '<div class="empty-note">Nenhuma frente com desvio relevante no período selecionado.</div>';
  }

  function renderFrontsTable(model) {
    var future = state.mode === "daily" && isSelectedDayFuture();
    var rows = model.fronts.map(function (front) {
      return { name: front.name, block: front.block, metric: periodMetric(front.block), unit: front.unit || front.block.unit || model.unit };
    }).sort(function (a, b) {
      var aRate = a.metric.adherence === null ? 9 : a.metric.adherence;
      var bRate = b.metric.adherence === null ? 9 : b.metric.adherence;
      return aRate - bRate;
    });

    var periodLabel = state.mode === "weekly" ? "Semana corrente" : "Dia selecionado";
    elements.frontsKicker.textContent = state.package === "bridge"
      ? periodLabel
      : SELECTED_PACKAGE_LABELS[state.package] + " · " + periodLabel;
    elements.frontsNote.textContent = future ? "Programação futura" : "Ordenado por aderência";
    var body = rows.map(function (row) {
      var metric = row.metric;
      var cls = future ? "is-empty" : statusClass(metric.adherence);
      return "<tr>" +
        "<td>" + escapeHtml(row.name) + "</td>" +
        "<td>" + escapeHtml(row.unit) + "</td>" +
        "<td>" + escapeHtml(formatNumber(metric.planned)) + "</td>" +
        "<td>" + escapeHtml(future ? "—" : formatNumber(metric.actual)) + "</td>" +
        "<td>" + escapeHtml(future ? "—" : (metric.variance >= 0 ? "+" : "") + formatNumber(metric.variance)) + "</td>" +
        '<td><span class="status-badge ' + cls + '">' + escapeHtml(future ? "Futuro" : formatPercent(metric.adherence)) + "</span></td>" +
        "<td>" + escapeHtml(formatPercent(row.block.month.adherence)) + "</td>" +
      "</tr>";
    }).join("");

    elements.frontsTable.innerHTML = '<table class="data-table">' +
      "<thead><tr><th>Atividade</th><th>Un.</th><th>Planejado</th><th>Realizado</th><th>Desvio</th><th>Aderência</th><th>Mês</th></tr></thead>" +
      "<tbody>" + body + "</tbody></table>";
  }

  function driver(label, value, unit) {
    return '<div class="driver"><span>' + escapeHtml(label) + "</span><strong>" + escapeHtml(formatNumber(value)) + "</strong><small>" + escapeHtml(unit) + "</small></div>";
  }

  function resourceGroupKey() {
    if (state.package === "p04" || state.package === "p05") return "p0405";
    if (state.package === "p06") return "p06";
    if (state.package === "p07" || state.package === "p08") return "p0708";
    return null;
  }

  function renderDrivers(model) {
    var tables = state.data.tables;
    if (model.key === "terra") {
      var terraGroupKey = resourceGroupKey();
      var terraMetrics = terraGroupKey ? tables.terra.groups[terraGroupKey] : tables.terra;
      elements.driversKicker.textContent = "Capacidade instalada";
      elements.driversTitle.textContent = terraGroupKey === "p0405" ? "Produtividade · PCTE 4/5" : terraGroupKey === "p0708" ? "Produtividade · PCTE 7/8" : "Produtividade e equipamentos";
      elements.drivers.innerHTML = '<div class="driver-grid">' +
        driver("Produtividade de escavação", terraMetrics.excavationProductivity, "m³ / escavadeira / dia") +
        driver("Escavadeiras mobilizadas", terraMetrics.excavators, "equipamentos") +
        driver("Produtividade de aterro", terraMetrics.fillProductivity, "m³ / rolo / dia") +
        driver("Rolos compactadores", terraMetrics.rollers, "equipamentos") +
      "</div>";
      return;
    }
    if (model.key === "drenagem") {
      var drainageGroupKey = resourceGroupKey();
      var drainageMetrics = drainageGroupKey ? tables.drainage.groups[drainageGroupKey] : tables.drainage;
      elements.driversKicker.textContent = "Capacidade instalada";
      elements.driversTitle.textContent = drainageGroupKey === "p0405" ? "Produtividade · PCTE 4/5" : drainageGroupKey === "p0708" ? "Produtividade · PCTE 7/8" : "Produtividade e efetivo";
      elements.drivers.innerHTML = '<div class="driver-grid">' +
        driver("Produtividade de drenagem", drainageMetrics.productivity, "m / homem / dia") +
        driver("Efetivo mobilizado", drainageMetrics.workforce, "pessoas") +
        driver("Produção no mês", drainageMetrics.monthProduction, "m") +
        driver("Média necessária", model.total.average.needed, "m / dia") +
      "</div>";
      return;
    }
    if (model.key === "dsh") {
      elements.driversKicker.textContent = "Ritmo diário";
      elements.driversTitle.textContent = "Médias de produção";
      elements.drivers.innerHTML = '<div class="driver-grid">' +
        driver("Média planejada", model.total.average.planned, "m / dia") +
        driver("Média realizada", model.total.average.actual, "m / dia") +
        driver("Média necessária", model.total.average.needed, "m / dia") +
        driver("Saldo contratual", model.total.scope.balance, "m") +
      "</div>";
      return;
    }

    elements.driversKicker.textContent = "Contribuição física";
    elements.driversTitle.textContent = "Atividades da ponte";
    var contributions = state.data.charts.bridgeContributions
      .filter(function (item) {
        return item.name.toUpperCase() !== "TOTAL" && (item.weight !== null || item.actualRate !== null);
      })
      .sort(function (a, b) { return (b.weight || 0) - (a.weight || 0); })
      .slice(0, 7);
    var max = Math.max.apply(null, contributions.map(function (item) { return item.actualRate || item.weight || 0; }).concat([1]));
    elements.drivers.innerHTML = '<div class="contribution-list">' + contributions.map(function (item) {
      var value = item.actualRate !== null ? item.actualRate : item.weight;
      return '<div class="contribution-row"><span>' + escapeHtml(item.name) + '</span><div class="contribution-row__track"><div class="contribution-row__fill" style="width:' + Math.min(100, value / max * 100) + '%"></div></div><b>' + escapeHtml(formatPercent(value)) + "</b></div>";
    }).join("") + "</div>";
  }

  function renderMonthBoard(model) {
    var cards = model.fronts.map(function (front) {
      var month = front.block.month;
      var rate = month.adherence;
      var need = front.block.average.needed;
      var actualAverage = front.block.average.actual;
      var progress = month.planned ? Math.min(100, (month.actual || 0) / month.planned * 100) : 0;
      var unit = front.unit || front.block.unit || model.unit;
      return '<article class="month-card">' +
        '<div class="month-card__header"><strong>' + escapeHtml(front.name) + '</strong><span class="status-badge ' + statusClass(rate) + '">' + escapeHtml(formatPercent(rate)) + "</span></div>" +
        '<div class="month-card__numbers"><span>Real ' + escapeHtml(formatCompact(month.actual)) + '</span><span>Meta ' + escapeHtml(formatCompact(month.planned)) + "</span></div>" +
        '<div class="month-card__track"><div class="month-card__fill" style="width:' + progress + '%"></div></div>' +
        '<div class="month-card__need">Média atual: <b>' + escapeHtml(formatNumber(actualAverage)) + '</b> · necessária: <b>' + escapeHtml(formatNumber(need)) + "</b> " + escapeHtml(unit) + "/dia</div>" +
      "</article>";
    }).join("");
    elements.monthBoard.innerHTML = cards;
  }

  function renderReference() {
    var metadata = state.data.consolidated.metadata;
    elements.referenceDate.textContent = formatDate(metadata.referenceDate);
    var validDates = metadata.dayDates.filter(Boolean);
    elements.referenceContext.textContent = validDates.length
      ? "Semana de " + shortDate(validDates[0]) + " a " + shortDate(validDates[validDates.length - 1])
      : "Semana corrente";
  }

  function renderDashboard() {
    if (!state.data) return;
    var model = packageModel();
    elements.dayControl.hidden = state.mode !== "daily";
    elements.disciplineControl.hidden = state.package === "bridge";
    var usesCombinedPackage = state.package === "p07" || state.package === "p08";
    elements.dataNote.hidden = !usesCombinedPackage;
    elements.dataNote.textContent = usesCombinedPackage
      ? "A planilha consolida os Pacotes 07 e 08 como PCTE 7/8. Por isso, os valores exibidos representam o conjunto 07/08."
      : "";
    elements.frontsTitle.textContent = state.package === "bridge"
      ? "Desempenho por componente"
      : "Desempenho por atividade";
    elements.monthTitle.textContent = state.package === "bridge"
      ? "Ritmo mensal por componente"
      : "Ritmo mensal do " + SELECTED_PACKAGE_LABELS[state.package];
    renderReference();
    renderDayPicker();
    renderKpis(model);
    renderBarChart(model);
    renderAttention(model);
    renderFrontsTable(model);
    renderDrivers(model);
    renderMonthBoard(model);
  }

  function setSync(message, status) {
    elements.syncLabel.textContent = message;
    elements.syncState.classList.toggle("is-ready", status === "ready");
    elements.syncState.classList.toggle("is-error", status === "error");
  }

  function showLoading(message) {
    setSync(message || "Lendo a base da obra…", "loading");
    elements.loading.hidden = false;
    elements.dashboard.hidden = true;
    elements.error.hidden = true;
    elements.refreshButton.disabled = true;
  }

  function showError(error) {
    elements.loading.hidden = true;
    elements.dashboard.hidden = true;
    elements.error.hidden = false;
    elements.errorMessage.textContent = error && error.message ? error.message : "Verifique a planilha e tente novamente.";
    elements.refreshButton.disabled = false;
    setSync("Falha ao ler a base", "error");
    window.__dashboardReady = false;
  }

  function applyBuffer(buffer, sourceLabel) {
    state.data = parseWorkbook(buffer);
    state.dayIndex = defaultDayIndex();
    elements.loading.hidden = true;
    elements.error.hidden = true;
    elements.dashboard.hidden = false;
    elements.refreshButton.disabled = false;
    setSync(sourceLabel, "ready");
    renderDashboard();
    window.__dashboardReady = true;
    window.__dashboardData = state.data;
  }

  async function loadRemote() {
    showLoading("Buscando a planilha mais recente…");
    try {
      var response = await fetch(WORKBOOK_URL + "?v=" + Date.now(), { cache: "no-store" });
      if (!response.ok) throw new Error("A base online respondeu com erro " + response.status + ".");
      var buffer = await response.arrayBuffer();
      applyBuffer(buffer, "Base atualizada agora");
    } catch (error) {
      showError(error);
    }
  }

  elements.modeSwitch.addEventListener("click", function (event) {
    var button = event.target.closest("[data-mode]");
    if (!button || !state.data) return;
    state.mode = button.dataset.mode;
    elements.modeSwitch.querySelectorAll("[data-mode]").forEach(function (item) {
      var active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-pressed", String(active));
    });
    renderDashboard();
  });

  elements.packageSwitch.addEventListener("click", function (event) {
    var button = event.target.closest("[data-package]");
    if (!button || !state.data) return;
    state.package = button.dataset.package;
    elements.packageSwitch.querySelectorAll("[data-package]").forEach(function (item) {
      var active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-selected", String(active));
    });
    renderDashboard();
  });

  elements.activitySwitch.addEventListener("click", function (event) {
    var button = event.target.closest("[data-activity]");
    if (!button || !state.data) return;
    state.activity = button.dataset.activity;
    elements.activitySwitch.querySelectorAll("[data-activity]").forEach(function (item) {
      var active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-selected", String(active));
    });
    renderDashboard();
  });

  elements.dayPicker.addEventListener("click", function (event) {
    var button = event.target.closest("[data-day]");
    if (!button || !state.data) return;
    state.dayIndex = Number(button.dataset.day);
    renderDashboard();
  });

  elements.refreshButton.addEventListener("click", loadRemote);

  elements.fileInput.addEventListener("change", async function (event) {
    var file = event.target.files && event.target.files[0];
    if (!file) return;
    showLoading("Lendo " + file.name + "…");
    try {
      applyBuffer(await file.arrayBuffer(), "Planilha local · " + file.name);
    } catch (error) {
      showError(error);
    } finally {
      event.target.value = "";
    }
  });

  loadRemote();
})();
