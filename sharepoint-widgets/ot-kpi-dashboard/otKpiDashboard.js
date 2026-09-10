/* eslint-disable no-undef */
(function ($) {
  "use strict";

  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var CATALOG_LIST_NAME = "OT KPI Catalog";
  var VALUES_LIST_NAME = "OT KPI Values";
  var PUBLISHED_STATUS = "Published";

  /**
   * When true (or when the Values list has no Published rows), the dashboard
   * uses SEED_VALUES below. Set to false once real SharePoint data is ready
   * and you no longer want automatic fallback.
   */
  var USE_DUMMY_VALUES = true;

  var CATEGORY_ORDER = [
    "Organization & Governance",
    "Training & Compliance",
    "Service Catalog",
    "Asset & Inventory",
    "DPI & Infrastructure",
    "Disaster Recovery",
    "Lifecycle Management"
  ];

  /**
   * Built-in seed used when the Catalog list is empty or unreachable,
   * so the dashboard layout is ready before data exists.
   */
  var SEED_CATALOG = [
    {
      code: "ORG-SITE-COVERAGE",
      title: "Sites OT Organization Coverage",
      category: "Organization & Governance",
      sortOrder: 10,
      description:
        "Percentage of OT-relevant sites with a formal Site OT organization created and HR Job Family Operational Technology assigned.",
      formula:
        "Sites with official Site OT organization and HR Job Family OT applied / Total OT-related sites",
      targetPercent: null
    },
    {
      code: "ORG-SITE-STAFFING",
      title: "Site OT Staffing Adequacy",
      category: "Organization & Governance",
      sortOrder: 20,
      description:
        "Percentage of OT-relevant sites that meet the defined minimum staffing threshold for day-to-day OT services.",
      formula:
        "Sites with official Site OT organization meeting ratio 1 OT FTE : 20 applications / Total OT-related sites",
      targetPercent: null
    },
    {
      code: "ORG-SECTOR-ROLES",
      title: "Sector OT Organization Coverage",
      category: "Organization & Governance",
      sortOrder: 30,
      description:
        "Percentage of required Sector OT roles assigned to named individuals with HR Job Family OT assigned.",
      formula:
        "Head of OT, Sector Architect, Sector Service Manager, Sector Portfolio Manager dedicated and HR Job Family applied / 16",
      targetPercent: null
    },
    {
      code: "ORG-SECTOR-GOV",
      title: "Sector OT Governance Coverage",
      category: "Organization & Governance",
      sortOrder: 40,
      description:
        "Percentage of OT-relevant sites integrated into Sector OT governance forums (Architecture, Portfolio, Service Management CoPs).",
      formula:
        "Sites integrated in Architecture, Portfolio, and Service Management CoPs / Total OT-related sites",
      targetPercent: null
    },
    {
      code: "TRN-FOUNDATIONAL",
      title: "OT Foundational Training Compliance",
      category: "Training & Compliance",
      sortOrder: 10,
      description:
        "Percentage of OT Job Family employees who completed mandatory foundational OT training within the required timeframe.",
      formula:
        "(OT employees assigned − OT employees overdue) / OT employees assigned",
      targetPercent: null
    },
    {
      code: "TRN-CURRICULUM-ASSIGN",
      title: "OT Role-Based Curriculum Assignment",
      category: "Training & Compliance",
      sortOrder: 20,
      description:
        "Percentage of OT Job Family employees assigned applicable role-based curricula according to the OT Training Matrix.",
      formula:
        "OT employees assigned one or more OTSKM curricula / OT employees",
      targetPercent: null
    },
    {
      code: "TRN-ROLE-COMPLIANCE",
      title: "OT Role-Based Training Compliance",
      category: "Training & Compliance",
      sortOrder: 30,
      description:
        "Percentage of employees with assigned OT role-based curricula who completed required training on time.",
      formula:
        "(OT employees assigned OTSKM curricula − OT employees overdue) / Employees assigned",
      targetPercent: null
    },
    {
      code: "SVC-APPROVAL",
      title: "OT Service Offering Approval Compliance",
      category: "Service Catalog",
      sortOrder: 10,
      description:
        "Percentage of OT Service Offerings that completed required approval workflow and are in an approved lifecycle state.",
      formula:
        "OT Service Offerings in approved/updating statuses / Total Catalog Service Offerings with Service Type OT",
      targetPercent: null
    },
    {
      code: "SVC-MANDATORY-ATTR",
      title: "OT Service Offering Mandatory Attributes",
      category: "Service Catalog",
      sortOrder: 20,
      description:
        "Percentage of OT Service Offerings in catalogue phase where all mandatory attributes are completed.",
      formula:
        "OT Service Offerings in Phase Catalog with mandatory fields completed / Total OT Service Offerings",
      targetPercent: null
    },
    {
      code: "SVC-BOP-INSTANCE",
      title: "BoP OT Service Offering-to-Instance",
      category: "Service Catalog",
      sortOrder: 30,
      description:
        "Whether BoP OT Business Service Offerings have at least one related BoP production Service Instance.",
      formula:
        "BoP OT BSOs with related Installed Service Instance / Total BoP OT BSOs in scope",
      targetPercent: null
    },
    {
      code: "INV-CEAD-COMPLETE",
      title: "OT Application Inventory Completeness",
      category: "Asset & Inventory",
      sortOrder: 10,
      description:
        "Percentage of OT-relevant sites with a completed and approved OT application inventory registered in CEAD.",
      formula:
        "Sites with full OT inventory registered in CEAD / Total OT-related sites",
      targetPercent: null
    },
    {
      code: "INV-CMDB-REGISTER",
      title: "OT Asset CMDB Registration Completeness",
      category: "Asset & Inventory",
      sortOrder: 20,
      description:
        "Percentage of OT assets connected to the DM network that are registered as configuration items in CMDB.",
      formula:
        "OT-connected assets registered in CMDB / Total endpoints connected to DM network",
      targetPercent: null
    },
    {
      code: "INV-SI-RELATIONSHIP",
      title: "OT Service Instance Relationship Completeness",
      category: "Asset & Inventory",
      sortOrder: 30,
      description:
        "Percentage of OT Service Instances that have a Technical Service Offering assigned.",
      formula:
        "Installed OT Service Instances with Technical Service Offering / Total installed OT Service Instances in scope",
      targetPercent: null
    },
    {
      code: "DPI-SITE-ENABLE",
      title: "DPI Site Enablement Coverage",
      category: "DPI & Infrastructure",
      sortOrder: 10,
      description:
        "Percentage of OT-relevant sites where DPI infrastructure is deployed and available for OT service integration.",
      formula: "Sites with DPI infrastructure / Total OT-related sites",
      targetPercent: null
    },
    {
      code: "DPI-APP-INTEGRATED",
      title: "Known OT Applications Integrated with DPI",
      category: "DPI & Infrastructure",
      sortOrder: 20,
      description:
        "Percentage of known OT applications integrated with DPI infrastructure based on related OT Service Instances.",
      formula:
        "OT Service Instances using DPI-integrated CIs / Total in-scope OT applications / Service Instances",
      targetPercent: null
    },
    {
      code: "DR-BSO-PLAN",
      title: "OT Service DR Plan Compliance",
      category: "Disaster Recovery",
      sortOrder: 10,
      description:
        "Percentage of OT Business Service Offerings requiring DR that have an approved and current DR plan.",
      formula:
        "OT BSOs with approved DR Plan / Total OT BSOs in Operations requiring DR",
      targetPercent: null
    },
    {
      code: "DR-BSO-TEST",
      title: "OT Service DR Test Evidence Compliance",
      category: "Disaster Recovery",
      sortOrder: 20,
      description:
        "Percentage of OT BSOs requiring DR that have an approved and current DR test report.",
      formula:
        "OT BSOs with approved DR Test Report / Total OT BSOs in Operations requiring DR",
      targetPercent: null
    },
    {
      code: "DR-DPI-SITE-PLAN",
      title: "DPI Site DR Plan Compliance",
      category: "Disaster Recovery",
      sortOrder: 30,
      description:
        "Percentage of DPI-enabled sites with an approved and current site-level DR plan.",
      formula: "DPI sites with approved Site DR Plan / Number of DPI sites",
      targetPercent: null
    },
    {
      code: "LIFECYCLE-SERVER-OS",
      title: "OT Server Operating System Lifecycle",
      category: "Lifecycle Management",
      sortOrder: 10,
      description:
        "Percentage of OT servers within OT BSOs that run a supported and approved OS version.",
      formula:
        "Servers in OT BSOs with supported OS / Total servers within OT BSOs",
      targetPercent: null
    },
    {
      code: "LIFECYCLE-CLIENT-OS",
      title: "OT Clients Operating System Lifecycle",
      category: "Lifecycle Management",
      sortOrder: 20,
      description:
        "Percentage of OT clients within OT BSOs that run a supported and approved OS version.",
      formula:
        "Clients in OT BSOs with supported OS / Total clients within OT BSOs",
      targetPercent: null
    }
  ];

  /**
   * Dummy period values for UI preview. Keyed by period → KPICode.
   * Remove or set USE_DUMMY_VALUES = false when live data is available.
   */
  var SEED_VALUES = {
    "2026-08": {
      "ORG-SITE-COVERAGE": { numerator: 38, denominator: 60 },
      "ORG-SITE-STAFFING": { numerator: 31, denominator: 60 },
      "ORG-SECTOR-ROLES": { numerator: 11, denominator: 16 },
      "ORG-SECTOR-GOV": { numerator: 34, denominator: 60 },
      "TRN-FOUNDATIONAL": { numerator: 412, denominator: 480 },
      "TRN-CURRICULUM-ASSIGN": { numerator: 390, denominator: 480 },
      "TRN-ROLE-COMPLIANCE": { numerator: 340, denominator: 390 },
      "SVC-APPROVAL": { numerator: 72, denominator: 95 },
      "SVC-MANDATORY-ATTR": { numerator: 61, denominator: 95 },
      "SVC-BOP-INSTANCE": { numerator: 18, denominator: 24 },
      "INV-CEAD-COMPLETE": { numerator: 29, denominator: 60 },
      "INV-CMDB-REGISTER": { numerator: 1850, denominator: 2400 },
      "INV-SI-RELATIONSHIP": { numerator: 410, denominator: 520 },
      "DPI-SITE-ENABLE": { numerator: 22, denominator: 60 },
      "DPI-APP-INTEGRATED": { numerator: 140, denominator: 310 },
      "DR-BSO-PLAN": { numerator: 48, denominator: 70 },
      "DR-BSO-TEST": { numerator: 36, denominator: 70 },
      "DR-DPI-SITE-PLAN": { numerator: 15, denominator: 22 },
      "LIFECYCLE-SERVER-OS": { numerator: 620, denominator: 780 },
      "LIFECYCLE-CLIENT-OS": { numerator: 4100, denominator: 5200 }
    },
    "2026-09": {
      "ORG-SITE-COVERAGE": { numerator: 42, denominator: 60 },
      "ORG-SITE-STAFFING": { numerator: 35, denominator: 60 },
      "ORG-SECTOR-ROLES": { numerator: 13, denominator: 16 },
      "ORG-SECTOR-GOV": { numerator: 39, denominator: 60 },
      "TRN-FOUNDATIONAL": { numerator: 438, denominator: 485 },
      "TRN-CURRICULUM-ASSIGN": { numerator: 410, denominator: 485 },
      "TRN-ROLE-COMPLIANCE": { numerator: 365, denominator: 410 },
      "SVC-APPROVAL": { numerator: 78, denominator: 98 },
      "SVC-MANDATORY-ATTR": { numerator: 70, denominator: 98 },
      "SVC-BOP-INSTANCE": { numerator: 20, denominator: 25 },
      "INV-CEAD-COMPLETE": { numerator: 33, denominator: 60 },
      "INV-CMDB-REGISTER": { numerator: 1980, denominator: 2450 },
      "INV-SI-RELATIONSHIP": { numerator: 445, denominator: 535 },
      "DPI-SITE-ENABLE": { numerator: 26, denominator: 60 },
      "DPI-APP-INTEGRATED": { numerator: 165, denominator: 320 },
      "DR-BSO-PLAN": { numerator: 54, denominator: 72 },
      "DR-BSO-TEST": { numerator: 41, denominator: 72 },
      "DR-DPI-SITE-PLAN": { numerator: 18, denominator: 26 },
      "LIFECYCLE-SERVER-OS": { numerator: 655, denominator: 790 },
      "LIFECYCLE-CLIENT-OS": { numerator: 4350, denominator: 5300 }
    }
  };

  var catalog = [];
  var valuesByPeriod = {};
  var selectedPeriod = "";
  var selectedCategory = "all";
  var usingSeedCatalog = false;
  var usingSeedValues = false;
  var expandedCategories = {};

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function trimOrEmpty(v) {
    if (v == null) return "";
    return String(v).trim();
  }

  function safeListTitle(name) {
    return String(name).replace(/'/g, "''");
  }

  function toNumber(v) {
    if (v == null || v === "") return null;
    var n = Number(v);
    return isFinite(n) ? n : null;
  }

  function formatPercent(n) {
    if (n == null || !isFinite(n)) return "—";
    return (
      Number(n).toLocaleString("en-US", {
        minimumFractionDigits: 0,
        maximumFractionDigits: 1
      }) + "%"
    );
  }

  function formatNumber(n) {
    if (n == null || !isFinite(n)) return "—";
    return Math.round(n).toLocaleString("en-US");
  }

  function showError(message) {
    $("#errorContainer").html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
  }

  function showInfo(message) {
    $("#infoContainer").html(
      '<div class="info-banner" role="status">' + escapeHtml(message) + "</div>"
    );
  }

  function clearBanners() {
    $("#errorContainer").empty();
    $("#infoContainer").empty();
  }

  function fetchAll(url) {
    var deferred = $.Deferred();
    var collected = [];

    function next(nextUrl) {
      $.ajax({
        url: nextUrl,
        type: "GET",
        headers: { Accept: "application/json;odata=verbose" }
      })
        .done(function (data) {
          var page =
            (data && data.value) || (data && data.d && data.d.results) || [];
          collected = collected.concat(page);
          var link =
            (data && data["@odata.nextLink"]) ||
            (data && data.d && data.d.__next) ||
            null;
          if (link) next(link);
          else deferred.resolve(collected);
        })
        .fail(function (xhr, status, err) {
          deferred.reject(xhr, status, err);
        });
    }

    next(url);
    return deferred.promise();
  }

  function scoreTone(percent, target) {
    if (percent == null || !isFinite(percent)) return "";
    var threshold = target != null && isFinite(target) ? target : 80;
    if (percent >= threshold) return "is-good";
    if (percent >= threshold * 0.75) return "is-warn";
    return "is-bad";
  }

  function computePercent(valueRow) {
    if (!valueRow) return null;
    if (valueRow.percentOverride != null && isFinite(valueRow.percentOverride)) {
      return valueRow.percentOverride;
    }
    if (
      valueRow.numerator == null ||
      valueRow.denominator == null ||
      valueRow.denominator === 0
    ) {
      return null;
    }
    return (100 * valueRow.numerator) / valueRow.denominator;
  }

  function normalizeCatalogItem(item) {
    var active = item.IsActive;
    if (active === false || active === 0 || String(active).toLowerCase() === "no") {
      return null;
    }

    return {
      code: trimOrEmpty(item.KPICode || item.Title),
      title: trimOrEmpty(item.Title) || trimOrEmpty(item.KPICode),
      category: trimOrEmpty(item.Category) || "Uncategorized",
      sortOrder: toNumber(item.SortOrder) != null ? toNumber(item.SortOrder) : 100,
      description: trimOrEmpty(item.Description),
      formula: trimOrEmpty(item.Formula),
      targetPercent: toNumber(item.TargetPercent),
      numeratorLabel: trimOrEmpty(item.NumeratorLabel),
      denominatorLabel: trimOrEmpty(item.DenominatorLabel),
      sourceSystem: trimOrEmpty(item.SourceSystem)
    };
  }

  function normalizeValueItem(item) {
    var status = trimOrEmpty(item.Status) || PUBLISHED_STATUS;
    if (status.toLowerCase() !== PUBLISHED_STATUS.toLowerCase()) {
      return null;
    }

    var period = trimOrEmpty(item.Period);
    var code = trimOrEmpty(item.KPICode);
    if (!period || !code) return null;

    return {
      code: code,
      period: period,
      periodDate: item.PeriodDate || null,
      numerator: toNumber(item.Numerator),
      denominator: toNumber(item.Denominator),
      percentOverride: toNumber(item.PercentOverride),
      notes: trimOrEmpty(item.Notes),
      dataAsOf: item.DataAsOf || null
    };
  }

  function loadCatalog() {
    var url =
      SITE_URL.replace(/\/$/, "") +
      "/_api/web/lists/getbytitle('" +
      safeListTitle(CATALOG_LIST_NAME) +
      "')/items?$select=" +
      encodeURIComponent(
        [
          "Id",
          "Title",
          "KPICode",
          "Category",
          "Description",
          "Formula",
          "NumeratorLabel",
          "DenominatorLabel",
          "TargetPercent",
          "SortOrder",
          "IsActive",
          "SourceSystem"
        ].join(",")
      ) +
      "&$top=5000";

    return fetchAll(url).then(
      function (items) {
        var rows = (items || [])
          .map(normalizeCatalogItem)
          .filter(Boolean)
          .filter(function (r) {
            return r.code;
          });

        if (!rows.length) {
          usingSeedCatalog = true;
          catalog = SEED_CATALOG.slice();
          return catalog;
        }

        usingSeedCatalog = false;
        catalog = rows.sort(function (a, b) {
          if (a.category === b.category) return a.sortOrder - b.sortOrder;
          return (
            CATEGORY_ORDER.indexOf(a.category) -
            CATEGORY_ORDER.indexOf(b.category)
          );
        });
        return catalog;
      },
      function () {
        usingSeedCatalog = true;
        catalog = SEED_CATALOG.slice();
        return catalog;
      }
    );
  }

  function applySeedValues() {
    usingSeedValues = true;
    valuesByPeriod = {};
    Object.keys(SEED_VALUES).forEach(function (period) {
      valuesByPeriod[period] = {};
      Object.keys(SEED_VALUES[period]).forEach(function (code) {
        var row = SEED_VALUES[period][code];
        valuesByPeriod[period][code] = {
          code: code,
          period: period,
          periodDate: null,
          numerator: row.numerator,
          denominator: row.denominator,
          percentOverride: null,
          notes: "Dummy data",
          dataAsOf: null
        };
      });
    });
    return valuesByPeriod;
  }

  function loadValues() {
    if (USE_DUMMY_VALUES) {
      return $.Deferred().resolve(applySeedValues()).promise();
    }

    var url =
      SITE_URL.replace(/\/$/, "") +
      "/_api/web/lists/getbytitle('" +
      safeListTitle(VALUES_LIST_NAME) +
      "')/items?$select=" +
      encodeURIComponent(
        [
          "Id",
          "Title",
          "KPICode",
          "Period",
          "PeriodDate",
          "Numerator",
          "Denominator",
          "PercentOverride",
          "Status",
          "Notes",
          "DataAsOf"
        ].join(",")
      ) +
      "&$top=5000";

    return fetchAll(url).then(
      function (items) {
        valuesByPeriod = {};
        usingSeedValues = false;
        (items || []).forEach(function (raw) {
          var row = normalizeValueItem(raw);
          if (!row) return;
          if (!valuesByPeriod[row.period]) valuesByPeriod[row.period] = {};
          valuesByPeriod[row.period][row.code] = row;
        });
        if (!Object.keys(valuesByPeriod).length) {
          return applySeedValues();
        }
        return valuesByPeriod;
      },
      function () {
        return applySeedValues();
      }
    );
  }

  function getPeriods() {
    return Object.keys(valuesByPeriod).sort(function (a, b) {
      return a < b ? 1 : a > b ? -1 : 0;
    });
  }

  function getCategories() {
    var seen = {};
    var out = [];
    CATEGORY_ORDER.forEach(function (c) {
      seen[c] = true;
      out.push(c);
    });
    catalog.forEach(function (k) {
      if (!seen[k.category]) {
        seen[k.category] = true;
        out.push(k.category);
      }
    });
    return out;
  }

  function populateFilters() {
    var periods = getPeriods();
    var $period = $("#periodSelect");
    $period.empty();

    if (!periods.length) {
      $period.append('<option value="">No periods yet</option>');
      selectedPeriod = "";
    } else {
      periods.forEach(function (p) {
        $period.append($("<option></option>").attr("value", p).text(p));
      });
      if (!selectedPeriod || periods.indexOf(selectedPeriod) === -1) {
        selectedPeriod = periods[0];
      }
      $period.val(selectedPeriod);
    }

    var $cat = $("#categorySelect");
    $cat.empty();
    $cat.append('<option value="all">All categories</option>');
    getCategories().forEach(function (c) {
      $cat.append($("<option></option>").attr("value", c).text(c));
    });
    $cat.val(selectedCategory || "all");
  }

  function filteredCatalog() {
    if (selectedCategory === "all") return catalog.slice();
    return catalog.filter(function (k) {
      return k.category === selectedCategory;
    });
  }

  function renderSummary(rows) {
    var withData = 0;
    var sum = 0;
    rows.forEach(function (k) {
      var value =
        selectedPeriod && valuesByPeriod[selectedPeriod]
          ? valuesByPeriod[selectedPeriod][k.code]
          : null;
      var pct = computePercent(value);
      if (pct == null) return;
      withData += 1;
      sum += pct;
    });

    $("#summaryTotal").text(formatNumber(rows.length));
    $("#summaryTotalSub").text(
      usingSeedCatalog ? "Built-in catalog (seed)" : "Active catalog KPIs"
    );
    $("#summaryWithData").text(formatNumber(withData));
    $("#summaryWithDataSub").text(
      selectedPeriod
        ? "Published values for " + selectedPeriod
        : "No reporting period selected"
    );
    $("#summaryAverage").text(
      withData ? formatPercent(sum / withData) : "—"
    );
    $("#summaryAverageSub").text(
      withData ? "Mean of KPIs with data" : "Awaiting published values"
    );
  }

  function getKpiValue(kpi) {
    return selectedPeriod && valuesByPeriod[selectedPeriod]
      ? valuesByPeriod[selectedPeriod][kpi.code]
      : null;
  }

  function getCategoryStats(items) {
    var withData = 0;
    var sum = 0;
    (items || []).forEach(function (k) {
      var pct = computePercent(getKpiValue(k));
      if (pct == null) return;
      withData += 1;
      sum += pct;
    });
    return {
      total: (items || []).length,
      withData: withData,
      average: withData ? sum / withData : null
    };
  }

  function renderCategories() {
    var rows = filteredCatalog();
    var $container = $("#categoriesContainer");
    $container.empty();

    if (!rows.length) {
      $("#emptyCatalog").removeClass("is-hidden");
      renderSummary([]);
      return;
    }
    $("#emptyCatalog").addClass("is-hidden");
    renderSummary(rows);

    var byCategory = {};
    rows.forEach(function (k) {
      if (!byCategory[k.category]) byCategory[k.category] = [];
      byCategory[k.category].push(k);
    });

    var categories = getCategories().filter(function (c) {
      return byCategory[c] && byCategory[c].length;
    });

    // When filtering to one category, expand it by default
    if (selectedCategory !== "all" && categories.length === 1) {
      expandedCategories[categories[0]] = true;
    }

    categories.forEach(function (category) {
      var items = byCategory[category].sort(function (a, b) {
        return a.sortOrder - b.sortOrder;
      });
      var stats = getCategoryStats(items);
      var isExpanded = !!expandedCategories[category];
      var tone = scoreTone(stats.average, 80);
      var hasData = stats.average != null;

      var metaText =
        (hasData
          ? stats.withData + " of " + stats.total + " KPIs with data"
          : "No published data yet") +
        " · Click to " +
        (isExpanded ? "collapse" : "expand");

      var $block = $('<section class="category-block"></section>')
        .attr("data-category", category)
        .toggleClass("is-expanded", isExpanded);

      var $summary = $(
        '<button type="button" class="category-summary" aria-expanded="' +
          (isExpanded ? "true" : "false") +
          '"></button>'
      );

      var $main = $('<div class="category-summary-main"></div>');
      $main.append(
        $('<h2 class="category-summary-title"></h2>').text(category)
      );
      $main.append(
        $('<p class="category-summary-meta"></p>').text(metaText)
      );

      var $right = $('<div class="category-summary-right"></div>');
      $right.append(
        $('<div class="category-summary-score"></div>')
          .addClass(hasData ? tone : "is-empty")
          .text(hasData ? formatPercent(stats.average) : "—")
      );
      $right.append(
        $('<span class="category-chevron" aria-hidden="true">▾</span>')
      );

      $summary.append($main, $right);
      $summary.on("click", function () {
        var next = !$block.hasClass("is-expanded");
        expandedCategories[category] = next;
        $block.toggleClass("is-expanded", next);
        $summary.attr("aria-expanded", next ? "true" : "false");
        $block
          .find(".category-summary-meta")
          .first()
          .text(
            (hasData
              ? stats.withData + " of " + stats.total + " KPIs with data"
              : "No published data yet") +
              " · Click to " +
              (next ? "collapse" : "expand")
          );
      });

      var $details = $('<div class="category-details"></div>');
      var $grid = $('<div class="kpi-grid"></div>');
      items.forEach(function (kpi) {
        $grid.append(renderKpiCard(kpi));
      });
      $details.append($grid);

      $block.append($summary, $details);
      $container.append($block);
    });
  }

  function renderKpiCard(kpi) {
    var value = getKpiValue(kpi);
    var pct = computePercent(value);
    var tone = scoreTone(pct, kpi.targetPercent);
    var hasData = pct != null;

    var metaParts = [];
    if (hasData && value) {
      metaParts.push(
        formatNumber(value.numerator) + " / " + formatNumber(value.denominator)
      );
    } else {
      metaParts.push("Awaiting data");
    }
    if (kpi.targetPercent != null) {
      metaParts.push("Target " + formatPercent(kpi.targetPercent));
    }
    if (kpi.sourceSystem) {
      metaParts.push(kpi.sourceSystem);
    }

    var $card = $('<article class="kpi-card"></article>');
    var $top = $('<div class="kpi-card-top"></div>');
    var $titleRow = $('<div class="kpi-title-row"></div>');
    $titleRow.append($('<h3 class="kpi-name"></h3>').text(kpi.title));

    if (kpi.description || kpi.formula) {
      var $info = $('<div class="kpi-info"></div>');
      var $btn = $(
        '<button type="button" class="kpi-info-btn" aria-label="Show KPI details">i</button>'
      );
      var $bubble = $('<div class="kpi-info-bubble" role="tooltip"></div>');

      if (kpi.description) {
        $bubble.append(
          $('<span class="kpi-info-label"></span>').text("Description")
        );
        $bubble.append($("<p></p>").text(kpi.description));
      }
      if (kpi.formula) {
        $bubble.append(
          $('<span class="kpi-info-label"></span>').text("Formula")
        );
        $bubble.append($("<p></p>").text(kpi.formula));
      }

      $btn.on("click", function (e) {
        e.preventDefault();
        e.stopPropagation();
        var open = $info.hasClass("is-open");
        $(".ot-kpi-dashboard .kpi-info.is-open").removeClass("is-open");
        if (!open) $info.addClass("is-open");
      });

      $info.append($btn, $bubble);
      $titleRow.append($info);
    }

    $top.append($titleRow);
    $top.append(
      $('<div class="kpi-score"></div>')
        .addClass(hasData ? tone : "is-empty")
        .text(hasData ? formatPercent(pct) : "—")
    );
    $card.append($top);

    $card.append(
      $('<div class="kpi-meta"></div>').text(metaParts.join(" · "))
    );

    var $progress = $('<div class="kpi-progress" aria-hidden="true"></div>');
    var width = hasData ? Math.max(0, Math.min(100, pct)) : 0;
    $progress.append(
      $('<div class="kpi-progress-bar"></div>')
        .addClass(tone)
        .css("width", width + "%")
    );
    $card.append($progress);

    return $card;
  }

  function setLastUpdated() {
    var stamp = new Date();
    $("#lastUpdated").text(
      "Updated " +
        stamp.toLocaleString("en-GB", {
          year: "numeric",
          month: "short",
          day: "numeric",
          hour: "2-digit",
          minute: "2-digit"
        })
    );
  }

  function renderAll() {
    populateFilters();
    renderCategories();
    setLastUpdated();
  }

  function bindEvents() {
    $("#periodSelect")
      .off("change.otKpi")
      .on("change.otKpi", function () {
        selectedPeriod = $(this).val() || "";
        renderCategories();
      });

    $("#categorySelect")
      .off("change.otKpi")
      .on("change.otKpi", function () {
        selectedCategory = $(this).val() || "all";
        renderCategories();
      });

    $(document)
      .off("click.otKpiInfo")
      .on("click.otKpiInfo", function (e) {
        if (!$(e.target).closest(".kpi-info").length) {
          $(".ot-kpi-dashboard .kpi-info.is-open").removeClass("is-open");
        }
      });
  }

  function init() {
    clearBanners();
    bindEvents();

    $.when(loadCatalog(), loadValues())
      .done(function () {
        var messages = [];
        if (usingSeedCatalog) {
          messages.push(
            "Showing built-in KPI definitions. Create the OT KPI Catalog list (or add items) to manage definitions in SharePoint."
          );
        }
        if (usingSeedValues) {
          messages.push(
            "Showing dummy values for 2026-08 and 2026-09. Set USE_DUMMY_VALUES = false in otKpiDashboard.js when live OT KPI Values data is ready."
          );
        } else if (!getPeriods().length) {
          messages.push(
            "No published values yet. Create OT KPI Values rows with Period, KPICode, Numerator, Denominator, and Status = Published."
          );
        }
        if (messages.length) {
          showInfo(messages.join(" "));
        }

        renderAll();
        $("#loadingOverlay").addClass("is-hidden");
        $("#dashboardContent").removeClass("is-hidden");
      })
      .fail(function (xhr) {
        usingSeedCatalog = true;
        catalog = SEED_CATALOG.slice();
        applySeedValues();
        showError(
          "Could not load SharePoint lists. Showing seed KPI layout with dummy data. Check list names and permissions. HTTP " +
            ((xhr && xhr.status) || "?")
        );
        renderAll();
        $("#loadingOverlay").addClass("is-hidden");
        $("#dashboardContent").removeClass("is-hidden");
      });
  }

  $(init);
})(jQuery);
