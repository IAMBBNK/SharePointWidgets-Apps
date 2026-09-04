/* eslint-disable no-undef */
(function () {
  "use strict";

  /* =========================================================================
   * CONFIG
   * Adjust SharePoint internal field names below if your column was renamed
   * after creation (the internal name is frozen at creation time).
   * On first load the script logs the keys of the first returned row to the
   * console so any mismatch is easy to spot.
   * =========================================================================
   */
  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var LIST_NAME = "Site Master Data";

  var FIELDS = {
    Region: "field_4",
    Sector: "Sector",
    ReviewStatus: "Review_x0020_Status",
    OTSystemIntegration: "OTSystemIntegration",
    DPIOperations: "DPI_x0020_Operations",
    DPICoreImplementation: "DPICoreImplementation",
    OTAssessment: "OTAssessment",
    ITAssessment: "ITAssessment",
    DPIVariant: "DPIVariant",
    RoadmapMaturity: "RoadmapMaturity",
    TotalOtApplications: "_x0023_ofOTSystems",
    OtAppsToMigrateToDpi: "_x0023_ofOTSystemtobeintegrated",
    OtAppsMigrated: "_x0023_ofmigratedOTSystems"
  };

  var REVIEWED_STATUS_LABEL = "Completed / Updated";
  var NOT_SPECIFIED_LABEL = "Not specified";

  /* =========================================================================
   * KPI CATEGORY ORDER
   * Fixed display order so KPIs render categories in the requested
   * order, even when count = 0 for a given category.
   * =========================================================================
   */
  var KPI_CONFIG = {
    OTSystemIntegration: {
      title: "OT System Integration",
      categories: ["Completed", "Ongoing", "Planning", "N/A", NOT_SPECIFIED_LABEL]
    },
    DPIOperations: {
      title: "DPI Operations",
      categories: ["Merck-managed", "Cognizant-managed", "Planning", "N/A", NOT_SPECIFIED_LABEL]
    },
    DPIImplementation: {
      title: "DPI Implementation",
      categories: ["Completed", "Ongoing", "Planning", "N/A", NOT_SPECIFIED_LABEL]
    },
    OTAssessment: {
      title: "OT Assessment",
      categories: ["Completed", "Ongoing", "Planning", "N/A", NOT_SPECIFIED_LABEL]
    },
    ITAssessment: {
      title: "IT Assessment",
      categories: ["Completed", "Ongoing", "Planning", "N/A", NOT_SPECIFIED_LABEL]
    },
    DPIVariant: {
      title: "(Target) DPI Variant",
      categories: ["Platinum", "Premium", "Premium Lite", "Essential", "N/A", NOT_SPECIFIED_LABEL]
    },
    RoadmapMaturity: {
      title: "Roadmap Maturity",
      categories: [
        "Stage 1 | Site Vision Roadmap (idea)",
        "Stage 2 | Approved Site Masterplan (Site LT approved)",
        "Stage 3 | Full Site Masterplan (budgeted / ongoing)",
        NOT_SPECIFIED_LABEL
      ],
      shortLabels: ["Stage 1", "Stage 2", "Stage 3"]
    }
  };

  /* =========================================================================
   * STATE
   * =========================================================================
   */
  var allRows = [];

  /* =========================================================================
   * UTILITIES
   * =========================================================================
   */
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

  function showError(message) {
    var $err = $("#errorContainer");
    $err.html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
  }

  function clearError() {
    $("#errorContainer").empty();
  }

  /* =========================================================================
   * SHAREPOINT FETCH (paginated)
   * Mirrors the fetchAll() pattern in Test/projects-webpart.js, but runs on
   * jQuery so we stay consistent with trainings.js when used inside SP.
   * =========================================================================
   */
  function fetchAll(url) {
    var deferred = $.Deferred();
    var collected = [];

    function fetchNext(nextUrl) {
      $.ajax({
        url: nextUrl,
        type: "GET",
        headers: {
          // Use verbose to keep maximum compatibility with SharePoint expand shapes
          // (notably FieldValuesAsText can be unreliable with nometadata in some tenants).
          Accept: "application/json;odata=verbose"
        }
      })
        .done(function (data) {
          var pageItems =
            (data && data.value) || (data && data.d && data.d.results) || [];
          collected = collected.concat(pageItems);

          var nextLink =
            (data && data["@odata.nextLink"]) ||
            (data && data.d && data.d.__next) ||
            null;

          if (nextLink) {
            fetchNext(nextLink);
          } else {
            deferred.resolve(collected);
          }
        })
        .fail(function (xhr, status, err) {
          deferred.reject(xhr, status, err);
        });
    }

    fetchNext(url);
    return deferred.promise();
  }

  function buildListApiUrl() {
    // Legacy (items?$select=...) kept for debugging only.
    // We now use RenderListDataAsStream for reliability.
    var base = SITE_URL.replace(/\/$/, "");
    return (
      base +
      "/_api/web/lists/getbytitle('" +
      LIST_NAME.replace(/'/g, "''") +
      "')/items?$select=Id&$top=1"
    );
  }

  function buildRenderListDataUrl() {
    var base = SITE_URL.replace(/\/$/, "");
    return (
      base +
      "/_api/web/lists/getbytitle('" +
      LIST_NAME.replace(/'/g, "''") +
      "')/RenderListDataAsStream"
    );
  }

  function buildViewXml() {
    // Use internal names. RenderListDataAsStream returns row objects keyed by internal names.
    var viewFields = [
      "ID",
      FIELDS.Region,
      FIELDS.Sector,
      FIELDS.ReviewStatus,
      FIELDS.OTSystemIntegration,
      FIELDS.DPIOperations,
      FIELDS.DPICoreImplementation,
      FIELDS.OTAssessment,
      FIELDS.ITAssessment,
      FIELDS.DPIVariant,
      FIELDS.RoadmapMaturity,
      FIELDS.TotalOtApplications,
      FIELDS.OtAppsToMigrateToDpi,
      FIELDS.OtAppsMigrated
    ]
      .filter(Boolean)
      .map(function (n) {
        return "<FieldRef Name='" + String(n).replace(/'/g, "&apos;") + "' />";
      })
      .join("");

    return (
      "<View Scope='RecursiveAll'>" +
      "<ViewFields>" +
      viewFields +
      "</ViewFields>" +
      "<RowLimit>5000</RowLimit>" +
      "</View>"
    );
  }

  function fetchRenderListDataAll() {
    var d = $.Deferred();
    var rows = [];

    function fetchPage(nextHref) {
      var url = nextHref ? SITE_URL.replace(/\/$/, "") + nextHref : buildRenderListDataUrl();
      var payload = {
        parameters: {
          RenderOptions: 2,
          ViewXml: buildViewXml()
        }
      };

      $.ajax({
        url: url,
        type: "POST",
        data: JSON.stringify(payload),
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose"
        }
      })
        .done(function (data) {
          var d0 = data && data.d ? data.d : data;
          var pageRows =
            (d0 && d0.RenderListData && d0.RenderListData.Row) ||
            (d0 && d0.Row) ||
            [];
          rows = rows.concat(pageRows);

          var next = (d0 && d0.NextHref) || (d0 && d0.RenderListData && d0.RenderListData.NextHref) || null;
          if (next) fetchPage(next);
          else d.resolve(rows);
        })
        .fail(function (xhr, status, err) {
          d.reject(xhr, status, err);
        });
    }

    fetchPage(null);
    return d.promise();
  }

  function buildItemsApiUrl(selectFieldsArray, top) {
    var base = SITE_URL.replace(/\/$/, "");
    var selectFields = (selectFieldsArray || []).join(",");
    return (
      base +
      "/_api/web/lists/getbytitle('" +
      LIST_NAME.replace(/'/g, "''") +
      "')/items?$select=" +
      encodeURIComponent(selectFields) +
      "&$top=" +
      encodeURIComponent(String(top || 1))
    );
  }

  function fetchItemsWithSelect(selectFieldsArray, top) {
    return fetchAll(buildItemsApiUrl(selectFieldsArray, top || 5000));
  }

  function buildProbeApiUrl() {
    // Minimal request to validate list access.
    var base = SITE_URL.replace(/\/$/, "");
    return (
      base +
      "/_api/web/lists/getbytitle('" +
      LIST_NAME.replace(/'/g, "''") +
      "')/items?$select=" +
      encodeURIComponent("Id") +
      "&$top=1"
    );
  }

  function buildFieldsCatalogUrl() {
    // Pull field metadata so we can map display Title -> InternalName.
    var base = SITE_URL.replace(/\/$/, "");
    return (
      base +
      "/_api/web/lists/getbytitle('" +
      LIST_NAME.replace(/'/g, "''") +
      "')/fields?$select=" +
      encodeURIComponent("Title,InternalName") +
      "&$top=5000"
    );
  }

  function mapFieldsByTitle(fields) {
    var out = {};
    (fields || []).forEach(function (f) {
      var t = trimOrEmpty(f && f.Title);
      var i = trimOrEmpty(f && f.InternalName);
      if (!t || !i) return;
      out[t.toLowerCase()] = { Title: t, InternalName: i };
    });
    return out;
  }

  function suggestInternalNames(fieldsCatalog) {
    var byTitle = mapFieldsByTitle(fieldsCatalog);
    var wantedTitles = {
      Region: "Region",
      Sector: "Sector",
      ReviewStatus: "Review Status",
      OTSystemIntegration: "OT System Integration",
      DPIOperations: "DPI Operations",
      DPICoreImplementation: "DPI Core Implementation",
      OTAssessment: "OT Assessment",
      ITAssessment: "IT Assessment",
      DPIVariant: "(Target) DPI Variant",
      RoadmapMaturity: "Roadmap Maturity",
      TotalOtApplications: "Total Number of OT Applications",
      OtAppsToMigrateToDpi: "Number of OT Applications to be migrated to DPI",
      OtAppsMigrated: "Number of migrated OT Applications"
    };

    var suggested = {};
    Object.keys(wantedTitles).forEach(function (k) {
      var title = wantedTitles[k];
      var hit = byTitle[title.toLowerCase()];
      if (hit) suggested[k] = hit.InternalName;
    });
    return suggested;
  }

  function xhrText(xhr) {
    try {
      return xhr && xhr.responseText ? String(xhr.responseText) : "";
    } catch (_e) {
      return "";
    }
  }

  function diagnoseSelectFields(fieldsToTest) {
    // Try fields incrementally to pinpoint which internal name breaks $select.
    var d = $.Deferred();
    var ok = ["Id"];
    var bad = [];

    var i = 0;
    function next() {
      if (i >= fieldsToTest.length) {
        d.resolve({ ok: ok.slice(0), bad: bad.slice(0) });
        return;
      }

      var f = fieldsToTest[i++];
      if (!f || f === "Id") {
        next();
        return;
      }

      var url = buildItemsApiUrl(ok.concat([f]), 1);
      $.ajax({
        url: url,
        type: "GET",
        headers: { Accept: "application/json;odata=verbose" }
      })
        .done(function () {
          ok.push(f);
          next();
        })
        .fail(function () {
          bad.push(f);
          next();
        });
    }

    next();
    return d.promise();
  }

  /* =========================================================================
   * NORMALIZATION
   * Bucket empty / unknown values into "N/A" so they show up in the donut.
   * =========================================================================
   */
  function normalizeValue(raw) {
    var s = trimOrEmpty(raw);
    if (!s) return NOT_SPECIFIED_LABEL;
    if (s.toLowerCase() === "n/a" || s.toLowerCase() === "na") return "N/A";
    return s;
  }

  function normalizeRow(row) {
    // With RenderListDataAsStream, row keys are internal names directly.
    return {
      Region: trimOrEmpty(row[FIELDS.Region]),
      Sector: trimOrEmpty(row[FIELDS.Sector]),
      ReviewStatus: trimOrEmpty(row[FIELDS.ReviewStatus]),
      OTSystemIntegration: normalizeValue(row[FIELDS.OTSystemIntegration]),
      DPIOperations: normalizeValue(row[FIELDS.DPIOperations]),
      DPIImplementation: normalizeValue(row[FIELDS.DPICoreImplementation]),
      OTAssessment: normalizeValue(row[FIELDS.OTAssessment]),
      ITAssessment: normalizeValue(row[FIELDS.ITAssessment]),
      DPIVariant: normalizeValue(row[FIELDS.DPIVariant]),
      RoadmapMaturity: normalizeValue(row[FIELDS.RoadmapMaturity]),
      TotalOtApplications: trimOrEmpty(row[FIELDS.TotalOtApplications]),
      OtAppsToMigrateToDpi: trimOrEmpty(row[FIELDS.OtAppsToMigrateToDpi]),
      OtAppsMigrated: trimOrEmpty(row[FIELDS.OtAppsMigrated])
    };
  }

  /* =========================================================================
   * KPI COMPUTATION
   * For per-category KPIs, return counts in the *configured* order so charts
   * can render consistently.
   * =========================================================================
   */
  function countByConfiguredCategories(rows, kpiKey) {
    var cfg = KPI_CONFIG[kpiKey];
    var counts = {};
    cfg.categories.forEach(function (c) {
      counts[c] = 0;
    });

    rows.forEach(function (r) {
      var v = r[kpiKey];
      if (v == null) v = NOT_SPECIFIED_LABEL;

      if (Object.prototype.hasOwnProperty.call(counts, v)) {
        counts[v] += 1;
      } else {
        // Match case-insensitively in case SP returns subtle whitespace/case
        // variants of the configured labels.
        var matched = cfg.categories.find(function (c) {
          return c.toLowerCase() === String(v).toLowerCase();
        });
        if (matched) {
          counts[matched] += 1;
        } else if (cfg.categories.indexOf(NOT_SPECIFIED_LABEL) !== -1) {
          counts[NOT_SPECIFIED_LABEL] += 1;
        } else if (cfg.categories.indexOf("N/A") !== -1) {
          counts["N/A"] += 1;
        }
      }
    });

    return cfg.categories.map(function (c) {
      return counts[c];
    });
  }

  function computeKpis(rows) {
    var totalSites = rows.length;
    var reviewedSites = rows.filter(function (r) {
      return (
        r.ReviewStatus &&
        r.ReviewStatus.toLowerCase() === REVIEWED_STATUS_LABEL.toLowerCase()
      );
    }).length;

    function toInt(v) {
      var s = trimOrEmpty(v);
      if (!s) return 0;
      // Handle values stored as text like "1,234", "1.234", "1 234".
      // Keep only digits and leading minus sign.
      var cleaned = String(s).trim().replace(/[^\d-]/g, "");
      if (!cleaned || cleaned === "-") return 0;
      var n = Number(cleaned);
      return Number.isFinite(n) ? n : 0;
    }

    var totalOtApps = 0;
    var otAppsToMigrate = 0;
    var otAppsMigrated = 0;

    var otDebug = {
      total: { nonEmpty: 0, samples: [] },
      toMigrate: { nonEmpty: 0, samples: [] },
      migrated: { nonEmpty: 0, samples: [] }
    };

    rows.forEach(function (r) {
      var a = r.TotalOtApplications;
      var b = r.OtAppsToMigrateToDpi;
      var c = r.OtAppsMigrated;

      if (trimOrEmpty(a)) {
        otDebug.total.nonEmpty += 1;
        if (otDebug.total.samples.length < 8) otDebug.total.samples.push(a);
      }
      if (trimOrEmpty(b)) {
        otDebug.toMigrate.nonEmpty += 1;
        if (otDebug.toMigrate.samples.length < 8) otDebug.toMigrate.samples.push(b);
      }
      if (trimOrEmpty(c)) {
        otDebug.migrated.nonEmpty += 1;
        if (otDebug.migrated.samples.length < 8) otDebug.migrated.samples.push(c);
      }

      totalOtApps += toInt(a);
      otAppsToMigrate += toInt(b);
      otAppsMigrated += toInt(c);
    });

    // Only log when we actually have rows; keep it compact.
    if (rows.length > 0) {
      console.log("[kpiDashboard] OT Apps debug:", {
        rows: rows.length,
        TotalOtApplications: {
          nonEmpty: otDebug.total.nonEmpty,
          samples: otDebug.total.samples
        },
        OtAppsToMigrateToDpi: {
          nonEmpty: otDebug.toMigrate.nonEmpty,
          samples: otDebug.toMigrate.samples
        },
        OtAppsMigrated: {
          nonEmpty: otDebug.migrated.nonEmpty,
          samples: otDebug.migrated.samples
        },
        sums: {
          totalOtApps: totalOtApps,
          otAppsToMigrate: otAppsToMigrate,
          otAppsMigrated: otAppsMigrated
        }
      });
    }

    var pctAppsIntegrated =
      otAppsToMigrate > 0 ? (otAppsMigrated / otAppsToMigrate) * 100 : null;

    var perCategory = {};
    Object.keys(KPI_CONFIG).forEach(function (key) {
      perCategory[key] = countByConfiguredCategories(rows, key);
    });

    return {
      totalSites: totalSites,
      reviewedSites: reviewedSites,
      perCategory: perCategory,
      totalOtApps: totalOtApps,
      otAppsToMigrate: otAppsToMigrate,
      otAppsMigrated: otAppsMigrated,
      pctAppsIntegrated: pctAppsIntegrated
    };
  }

  /* =========================================================================
   * RENDERING
   * =========================================================================
   */
  function renderTiles(kpis) {
    $("#kpiTotalSites").text(kpis.totalSites);
    $("#kpiReviewedSites").text(kpis.reviewedSites);

    var pct =
      kpis.totalSites > 0
        ? Math.round((kpis.reviewedSites / kpis.totalSites) * 100)
        : 0;
    $("#kpiReviewedSitesSub").text(
      "Review Status = Completed / Updated  \u2022  " + pct + "% of total"
    );

    if ($("#kpiTotalOtApps").length) {
      $("#kpiTotalOtApps").text(kpis.totalOtApps);
      $("#kpiOtAppsToMigrate").text(kpis.otAppsToMigrate);

      if (kpis.pctAppsIntegrated == null) {
        $("#kpiPctAppsIntegrated").text("—");
        $("#kpiPctAppsIntegratedSub").text(
          kpis.otAppsMigrated + " / " + kpis.otAppsToMigrate + " migrated"
        );
      } else {
        $("#kpiPctAppsIntegrated").text(
          Math.round(kpis.pctAppsIntegrated) + "%"
        );
        $("#kpiPctAppsIntegratedSub").text(
          kpis.otAppsMigrated + " / " + kpis.otAppsToMigrate + " migrated"
        );
      }
    }
  }

  function renderLegend($card, cfg, counts) {
    var labels = cfg.shortLabels || cfg.categories;
    var total = counts.reduce(function (a, b) {
      return a + b;
    }, 0);

    var header =
      '<li style="padding: 6px 0 8px 0; font-weight: 800; color: #0f69af;">' +
      '<span class="legend-label">Total</span>' +
      '<span class="legend-count">' +
      total +
      "</span></li>";

    var lines = labels
      .map(function (label, i) {
        return (
          "<li>" +
          '<span class="swatch" style="background-color:#0f69af;"></span>' +
          '<span class="legend-label" title="' +
          escapeHtml(cfg.categories[i]) +
          '">' +
          escapeHtml(label) +
          "</span>" +
          '<span class="legend-count">' +
          counts[i] +
          "</span></li>"
        );
      })
      .join("");

    $card.find(".chart-legend").html(header + lines);
  }

  function renderCharts(kpis) {
    Object.keys(KPI_CONFIG).forEach(function (key) {
      var cfg = KPI_CONFIG[key];
      var $card = $('.chart-card[data-kpi="' + key + '"]');
      if (!$card.length) return;
      renderLegend($card, cfg, kpis.perCategory[key]);
    });
  }

  function renderAll(rows) {
    var kpis = computeKpis(rows);
    renderTiles(kpis);
    renderCharts(kpis);
  }

  /* =========================================================================
   * FILTERS
   * =========================================================================
   */
  function uniqueValues(rows, key) {
    var set = {};
    rows.forEach(function (r) {
      var v = trimOrEmpty(r[key]);
      if (v) set[v] = true;
    });
    return Object.keys(set).sort(function (a, b) {
      return a.localeCompare(b, undefined, { sensitivity: "base" });
    });
  }

  function populateFilterOptions(rows) {
    var regions = uniqueValues(rows, "Region");
    var sectors = uniqueValues(rows, "Sector");

    var $r = $("#filterRegion");
    $r.find("option:not([value=''])").remove();
    regions.forEach(function (v) {
      $r.append(
        '<option value="' + escapeHtml(v) + '">' + escapeHtml(v) + "</option>"
      );
    });

    var $s = $("#filterSector");
    $s.find("option:not([value=''])").remove();
    sectors.forEach(function (v) {
      $s.append(
        '<option value="' + escapeHtml(v) + '">' + escapeHtml(v) + "</option>"
      );
    });
  }

  function applyFilters() {
    var region = $("#filterRegion").val() || "";
    var sector = $("#filterSector").val() || "";

    var filtered = allRows.filter(function (r) {
      if (region && r.Region !== region) return false;
      if (sector && r.Sector !== sector) return false;
      return true;
    });

    renderAll(filtered);
  }

  function wireFilterEvents() {
    $("#filterRegion").on("change", applyFilters);
    $("#filterSector").on("change", applyFilters);
    $("#resetFilters").on("click", function () {
      $("#filterRegion").val("");
      $("#filterSector").val("");
      applyFilters();
    });
  }

  /* =========================================================================
   * BOOTSTRAP
   * =========================================================================
   */
  function setLastUpdated() {
    var d = new Date();
    var pad = function (n) {
      return n < 10 ? "0" + n : "" + n;
    };
    $("#lastUpdated").text(
      "Last updated: " +
        d.getFullYear() +
        "-" +
        pad(d.getMonth() + 1) +
        "-" +
        pad(d.getDate()) +
        " " +
        pad(d.getHours()) +
        ":" +
        pad(d.getMinutes())
    );
  }

  function init() {
    wireFilterEvents();

    var renderUrl = buildRenderListDataUrl();
    console.log("[kpiDashboard] RenderListDataAsStream:", renderUrl);

    fetchRenderListDataAll()
      .done(function (rawRows) {
        if (rawRows.length > 0) {
          console.log(
            "[kpiDashboard] Loaded " + rawRows.length + " rows. First-row keys:",
            Object.keys(rawRows[0])
          );
        } else {
          console.warn("[kpiDashboard] List returned 0 rows.");
        }

        allRows = rawRows.map(normalizeRow);

        // Debug the OT application numeric fields (they are stored as text).
        // Helps validate that the OT fields are coming through.
        if (allRows.length > 0) {
          var sample = allRows.slice(0, 5).map(function (r) {
            return {
              TotalOtApplications: r.TotalOtApplications,
              OtAppsToMigrateToDpi: r.OtAppsToMigrateToDpi,
              OtAppsMigrated: r.OtAppsMigrated
            };
          });
          console.log("[kpiDashboard] OT app field samples (first 5):", sample);
        }

        populateFilterOptions(allRows);

        $("#loadingOverlay").hide();
        $("#chartsContainer").show();

        renderAll(allRows);
        setLastUpdated();
      })
      .fail(function (xhr, status, err) {
        console.error("[kpiDashboard] RenderListDataAsStream failed:", status, err, xhr);
        var httpInfo = xhr && xhr.status ? xhr.status + " " + xhr.statusText : status;
        showError(
          "Failed to load list '" +
            LIST_NAME +
            "' via RenderListDataAsStream. HTTP " +
            httpInfo +
            "."
        );
        $("#loadingOverlay").text("Could not load data. See error message above.");
      });
  }

  $(document).ready(init);
})();
