/* eslint-disable no-undef */
(function ($) {
  "use strict";

  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var LIST_NAME = "Site Master Data";

  var SECTIONS = [
    { key: "general", label: "General" },
    { key: "people", label: "People & roles" },
    { key: "maturity", label: "Maturity & assessments" },
    { key: "counts", label: "OT applications" }
  ];

  /**
   * Same field set / sections as siteMasterDataForm.js
   */
  var FIELD_DEFS = [
    { key: "SiteId", titles: ["Site ID"], type: "text", section: "general" },
    { key: "SiteName", titles: ["Site Name"], type: "text", section: "general" },
    {
      key: "SiteAddress",
      titles: ["Site Address"],
      type: "text",
      section: "general"
    },
    { key: "City", titles: ["City"], type: "text", section: "general" },
    { key: "Country", titles: ["Country"], type: "text", section: "general" },
    {
      key: "Region",
      titles: ["Region"],
      knownInternal: "field_4",
      type: "choice",
      section: "general"
    },
    { key: "Sector", titles: ["Sector"], type: "choice", section: "general" },
    {
      key: "LegalEntityName",
      titles: ["Legal Entity Name"],
      type: "text",
      section: "general"
    },
    {
      key: "LegalEntityCode",
      titles: ["Legal Entity Code"],
      type: "text",
      section: "general"
    },
    {
      key: "BusinessUnit",
      titles: ["Business Unit"],
      type: "text",
      section: "general"
    },
    {
      key: "BusinessField",
      titles: ["Business Field / Type / Service"],
      type: "text",
      section: "general"
    },
    {
      key: "Employees",
      titles: ["# of Employees"],
      type: "text",
      section: "general"
    },

    {
      key: "SiteDataAdmin",
      titles: ["Site Data Admin"],
      type: "person",
      section: "people"
    },
    { key: "SiteHead", titles: ["Site Head"], type: "person", section: "people" },
    {
      key: "DigitalLead",
      titles: ["Digtial Lead / SPOC", "Digital Lead / SPOC", "Digital Lead"],
      type: "person",
      section: "people"
    },
    {
      key: "OtServiceDeliveryMgr",
      titles: ["OT Service Delivery Mgr."],
      type: "person",
      section: "people"
    },
    {
      key: "OtProjectDeliveryMgr",
      titles: ["OT Project Delivery Mgr."],
      type: "person",
      section: "people"
    },
    {
      key: "OtEngineers",
      titles: ["OT Engineers"],
      type: "person",
      section: "people"
    },
    {
      key: "OtSolutionArchitect",
      titles: ["OT Solution Architect"],
      type: "person",
      section: "people"
    },
    {
      key: "CyberSecurityMgr",
      titles: ["CyberSecurity Mgr."],
      type: "person",
      section: "people"
    },
    {
      key: "ItBusinessPartner",
      titles: ["IT Business Partner"],
      type: "person",
      section: "people"
    },

    {
      key: "RoadmapMaturity",
      titles: ["Roadmap Maturity"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "CurrentDigitalMaturity",
      titles: ["Current Maturity (BPOG)"],
      type: "text",
      section: "maturity"
    },
    {
      key: "TargetDigitalMaturity",
      titles: ["Target Maturity (BPOG)"],
      type: "text",
      section: "maturity"
    },
    {
      key: "DpiVariant",
      titles: ["(Target) DPI Variant"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "ItAssessment",
      titles: ["IT Assessment"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "OtAssessment",
      titles: ["OT Assessment"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "DpiCoreImplementation",
      titles: ["DPI Core Implementation"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "DpiOperations",
      titles: ["DPI Operations"],
      type: "choice",
      section: "maturity"
    },
    {
      key: "OtSystemIntegration",
      titles: ["OT System Integration"],
      type: "choice",
      section: "maturity"
    },

    {
      key: "TotalOtApplications",
      titles: ["Total Number of OT Applications"],
      type: "text",
      section: "counts"
    },
    {
      key: "OtSystemsInGear",
      titles: ["Gear"],
      type: "text",
      section: "counts"
    },
    {
      key: "OtAppsToMigrate",
      titles: ["Number of OT Applications to be migrated to DPI"],
      type: "text",
      section: "counts"
    },
    {
      key: "OtAppsMigrated",
      titles: ["Number of migrated OT Applications"],
      type: "text",
      section: "counts"
    }
  ];

  var fieldsByKey = {};
  var reviewStatusInternal = "Review_x0020_Status";
  var reviewStatusChoices = [];
  var allRows = [];
  var REVIEWED_STATUS_LABEL = "Completed / Updated";
  var sortState = { key: "overall", dir: "asc" };

  var SORTABLE_COLUMNS = [
    { key: "site", label: "Site", sortable: true },
    { key: "review", label: "Review Status", sortable: true },
    { key: "general", label: "General", sortable: true, section: "general" },
    { key: "people", label: "People & roles", sortable: true, section: "people" },
    {
      key: "maturity",
      label: "Maturity & assessments",
      sortable: true,
      section: "maturity"
    },
    {
      key: "counts",
      label: "OT applications",
      sortable: true,
      section: "counts"
    },
    {
      key: "overall",
      label: "Overall",
      sortable: true,
      overall: true,
      tipTitle: "All section fields",
      tipAllSections: true
    }
  ];

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

  function safeListTitle() {
    return LIST_NAME.replace(/'/g, "''");
  }

  function itemDisplayUrl(id) {
    return (
      SITE_URL +
      "/Lists/" +
      encodeURIComponent(LIST_NAME) +
      "/DispForm.aspx?ID=" +
      encodeURIComponent(String(id))
    );
  }

  function itemEditUrl(id) {
    return (
      SITE_URL +
      "/Lists/" +
      encodeURIComponent(LIST_NAME) +
      "/EditForm.aspx?ID=" +
      encodeURIComponent(String(id))
    );
  }

  function fieldLabelsForSection(sectionKey) {
    return FIELD_DEFS.filter(function (d) {
      return d.section === sectionKey;
    }).map(function (d) {
      var resolved = fieldsByKey[d.key];
      return (resolved && resolved.Title) || d.titles[0];
    });
  }

  function allFieldLabels() {
    var out = [];
    SECTIONS.forEach(function (s) {
      fieldLabelsForSection(s.key).forEach(function (label) {
        out.push(label);
      });
    });
    return out;
  }

  function infoTipHtml(title, labels) {
    var items = (labels || [])
      .map(function (l) {
        return "<li>" + escapeHtml(l) + "</li>";
      })
      .join("");
    return (
      '<span class="scd-info" tabindex="0" aria-label="' +
      escapeHtml(title) +
      ' fields">' +
      "i" +
      '<span class="scd-info-tip" role="tooltip">' +
      "<strong>" +
      escapeHtml(title) +
      "</strong>" +
      "<ul>" +
      items +
      "</ul>" +
      "</span></span>"
    );
  }

  function toODataSelectName(internalName) {
    var n = trimOrEmpty(internalName);
    if (!n) return n;
    if (n.indexOf("OData_") === 0) return n;
    if (n.charAt(0) === "_") return "OData_" + n;
    return n;
  }

  function readFieldFromItem(item, internalName) {
    if (!item || !internalName) return null;
    if (Object.prototype.hasOwnProperty.call(item, internalName)) {
      return item[internalName];
    }
    var odata = toODataSelectName(internalName);
    if (odata !== internalName && item[odata] !== undefined) return item[odata];
    if (item[internalName] !== undefined) return item[internalName];
    return null;
  }

  function showError(message) {
    $("#scdError").html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
  }

  function clearError() {
    $("#scdError").empty();
  }

  function fetchFieldsCatalog() {
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle() +
        "')/fields?$select=" +
        encodeURIComponent(
          "Title,InternalName,TypeAsString,AllowMultipleValues,Choices,Hidden"
        ) +
        "&$top=5000",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      return ((data.d && data.d.results) || []).filter(function (f) {
        return f && !f.Hidden;
      });
    });
  }

  function resolveFieldDefs(catalog) {
    var byTitle = {};
    var byInternal = {};
    (catalog || []).forEach(function (f) {
      var t = trimOrEmpty(f.Title);
      if (t) byTitle[t.toLowerCase()] = f;
      if (f.InternalName) byInternal[f.InternalName] = f;
    });

    var reviewHit =
      byTitle["review status"] || byInternal["Review_x0020_Status"];
    if (reviewHit) {
      reviewStatusInternal = reviewHit.InternalName;
      var choices = [];
      if (reviewHit.Choices && reviewHit.Choices.results) {
        choices = reviewHit.Choices.results;
      } else if (Array.isArray(reviewHit.Choices)) {
        choices = reviewHit.Choices;
      }
      reviewStatusChoices = (choices || []).map(trimOrEmpty).filter(Boolean);
      console.log(
        "[Completion Dashboard] Review Status choices:",
        reviewStatusChoices
      );
    } else {
      console.warn("[Completion Dashboard] Review Status column not found");
    }

    fieldsByKey = {};
    FIELD_DEFS.forEach(function (def) {
      var hit = null;
      for (var i = 0; i < def.titles.length; i++) {
        var cand = byTitle[def.titles[i].toLowerCase()];
        if (cand) {
          hit = cand;
          break;
        }
      }
      if (!hit && def.knownInternal && byInternal[def.knownInternal]) {
        hit = byInternal[def.knownInternal];
      }
      if (!hit) {
        console.warn("[Completion Dashboard] Unresolved field:", def.key);
        return;
      }
      var typeAs = hit.TypeAsString || "";
      fieldsByKey[def.key] = {
        key: def.key,
        def: def,
        Title: hit.Title,
        InternalName: hit.InternalName,
        TypeAsString: typeAs,
        isPerson: typeAs === "User" || typeAs === "UserMulti"
      };
    });

    console.log(
      "[Completion Dashboard] Resolved",
      Object.keys(fieldsByKey).length,
      "/",
      FIELD_DEFS.length,
      "fields"
    );
  }

  function fetchAll(url) {
    var deferred = $.Deferred();
    var collected = [];

    function fetchNext(nextUrl) {
      $.ajax({
        url: nextUrl,
        type: "GET",
        headers: { Accept: "application/json;odata=verbose" }
      })
        .done(function (data) {
          var pageItems =
            (data && data.value) || (data && data.d && data.d.results) || [];
          collected = collected.concat(pageItems);
          var nextLink =
            (data && data["@odata.nextLink"]) ||
            (data && data.d && data.d.__next) ||
            null;
          if (nextLink) fetchNext(nextLink);
          else deferred.resolve(collected);
        })
        .fail(function (xhr, status, err) {
          deferred.reject(xhr, status, err);
        });
    }

    fetchNext(url);
    return deferred.promise();
  }

  function buildItemsQueryUrl() {
    var select = ["Id", "Title"];
    var expand = [];

    Object.keys(fieldsByKey).forEach(function (k) {
      var f = fieldsByKey[k];
      var od = toODataSelectName(f.InternalName);
      if (f.isPerson) {
        select.push(od + "/Id", od + "/Title", od + "/EMail");
        expand.push(od);
      } else {
        select.push(od);
      }
    });

    if (reviewStatusInternal) {
      select.push(toODataSelectName(reviewStatusInternal));
    }

    function unique(arr) {
      var seen = {};
      return arr.filter(function (x) {
        if (seen[x]) return false;
        seen[x] = true;
        return true;
      });
    }

    select = unique(select);
    expand = unique(expand);

    var url =
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeListTitle() +
      "')/items?$select=" +
      encodeURIComponent(select.join(",")) +
      "&$top=5000";

    if (expand.length) {
      url += "&$expand=" + encodeURIComponent(expand.join(","));
    }
    return url;
  }

  function peopleFromField(personField) {
    if (!personField) return [];
    var arr = Array.isArray(personField)
      ? personField
      : personField.results
        ? personField.results
        : [personField];
    return arr.filter(Boolean);
  }

  function isFieldFilled(meta, raw) {
    if (!meta) return false;
    if (meta.isPerson) {
      return peopleFromField(raw).length > 0;
    }
    if (raw == null) return false;
    if (typeof raw === "object" && raw.results) {
      return (raw.results || []).some(function (v) {
        return trimOrEmpty(v) !== "";
      });
    }
    return trimOrEmpty(raw) !== "";
  }

  function completionColor(ratio) {
    var t = Math.max(0, Math.min(1, Number(ratio) || 0));
    // red (0) -> amber (0.5) -> green (1)
    var r;
    var g;
    var b;
    if (t <= 0.5) {
      var u = t / 0.5;
      r = Math.round(211 + (245 - 211) * u);
      g = Math.round(47 + (166 - 47) * u);
      b = Math.round(47 + (35 - 47) * u);
    } else {
      var v = (t - 0.5) / 0.5;
      r = Math.round(245 + (46 - 245) * v);
      g = Math.round(166 + (125 - 166) * v);
      b = Math.round(35 + (50 - 35) * v);
    }
    return "rgb(" + r + ", " + g + ", " + b + ")";
  }

  function sectorAbbr(sector) {
    var s = trimOrEmpty(sector).toLowerCase();
    if (!s) return "";
    if (s.indexOf("electronic") !== -1) return "EL";
    if (s.indexOf("life science") !== -1 || s.indexOf("life-science") !== -1) return "LS";
    if (s.indexOf("healthcare") !== -1 || s.indexOf("health care") !== -1) return "HC";
    if (s.indexOf("3rd party") !== -1 || s.indexOf("third party") !== -1 ||
        s.indexOf("3rd-party") !== -1 || s.indexOf("third-party") !== -1) return "LS";
    return "";
  }

  function sectorTagClass(abbr) {
    if (abbr === "EL") return "scd-sector-tag scd-sector-tag--el";
    if (abbr === "LS") return "scd-sector-tag scd-sector-tag--ls";
    if (abbr === "HC") return "scd-sector-tag scd-sector-tag--hc";
    return "scd-sector-tag";
  }

  function isDpiSite(dpiVariant) {
    var v = trimOrEmpty(dpiVariant);
    if (!v) return false;
    return v.toLowerCase() !== "n/a" && v.toLowerCase() !== "na";
  }

  function reviewStatusTone(status) {
    var s = trimOrEmpty(status).toLowerCase();
    if (!s) return "empty";
    if (
      s === REVIEWED_STATUS_LABEL.toLowerCase() ||
      s.indexOf("completed") !== -1 ||
      s.indexOf("updated") !== -1 ||
      s === "done" ||
      s === "complete"
    ) {
      return "done";
    }
    if (
      s.indexOf("pending") !== -1 ||
      s.indexOf("not reviewed") !== -1 ||
      s.indexOf("open") !== -1 ||
      s.indexOf("due") !== -1 ||
      s === "n/a" ||
      s === "na"
    ) {
      return "pending";
    }
    if (
      s.indexOf("progress") !== -1 ||
      s.indexOf("ongoing") !== -1 ||
      s.indexOf("review") !== -1 ||
      s.indexOf("started") !== -1
    ) {
      return "progress";
    }
    return "other";
  }

  function computeRow(item) {
    var siteIdMeta = fieldsByKey.SiteId;
    var siteNameMeta = fieldsByKey.SiteName;
    var sectorMeta = fieldsByKey.Sector;
    var regionMeta = fieldsByKey.Region;
    var dpiMeta = fieldsByKey.DpiVariant;

    var siteId = siteIdMeta
      ? trimOrEmpty(readFieldFromItem(item, siteIdMeta.InternalName))
      : "";
    var siteName = siteNameMeta
      ? trimOrEmpty(readFieldFromItem(item, siteNameMeta.InternalName))
      : "";
    if (!siteId && item.Title) siteId = trimOrEmpty(item.Title);

    var sector = sectorMeta
      ? trimOrEmpty(readFieldFromItem(item, sectorMeta.InternalName))
      : "";
    var region = regionMeta
      ? trimOrEmpty(readFieldFromItem(item, regionMeta.InternalName))
      : "";
    var dpiVariant = dpiMeta
      ? trimOrEmpty(readFieldFromItem(item, dpiMeta.InternalName))
      : "";
    var reviewStatus = reviewStatusInternal
      ? trimOrEmpty(readFieldFromItem(item, reviewStatusInternal))
      : "";

    var sections = {};
    var totalFilled = 0;
    var totalCount = 0;

    SECTIONS.forEach(function (sec) {
      var filled = 0;
      var count = 0;
      Object.keys(fieldsByKey).forEach(function (k) {
        var meta = fieldsByKey[k];
        if (meta.def.section !== sec.key) return;
        count += 1;
        var raw = readFieldFromItem(item, meta.InternalName);
        if (isFieldFilled(meta, raw)) filled += 1;
      });
      var ratio = count ? filled / count : 0;
      sections[sec.key] = {
        filled: filled,
        total: count,
        ratio: ratio,
        pct: count ? Math.round(ratio * 100) : 0
      };
      totalFilled += filled;
      totalCount += count;
    });

    var overallRatio = totalCount ? totalFilled / totalCount : 0;

    return {
      Id: item.Id,
      siteId: siteId,
      siteName: siteName,
      label:
        siteId && siteName
          ? siteId + " — " + siteName
          : siteId || siteName || "Item " + item.Id,
      sector: sector,
      region: region,
      dpiVariant: dpiVariant,
      isDpi: isDpiSite(dpiVariant),
      reviewStatus: reviewStatus,
      reviewTone: reviewStatusTone(reviewStatus),
      sections: sections,
      overall: {
        filled: totalFilled,
        total: totalCount,
        ratio: overallRatio,
        pct: totalCount ? Math.round(overallRatio * 100) : 0
      }
    };
  }

  function isThirdPartySector(name) {
    var s = trimOrEmpty(name).toLowerCase();
    if (!s) return false;
    return (
      s.indexOf("3rd party") !== -1 ||
      s.indexOf("third party") !== -1 ||
      s.indexOf("3rd-party") !== -1 ||
      s.indexOf("third-party") !== -1
    );
  }

  function uniqueSectors(rows) {
    var seen = {};
    var out = [];
    rows.forEach(function (r) {
      var s = trimOrEmpty(r.sector);
      if (!s || seen[s]) return;
      seen[s] = true;
      out.push(s);
    });
    out.sort(function (a, b) {
      var aLast = isThirdPartySector(a);
      var bLast = isThirdPartySector(b);
      if (aLast !== bLast) return aLast ? 1 : -1;
      return a.localeCompare(b);
    });
    return out;
  }

  function populateSectorFilter(rows) {
    var $sel = $("#scdFilterSector");
    var current = $sel.val() || "";
    $sel.empty().append('<option value="">All sectors</option>');
    uniqueSectors(rows).forEach(function (s) {
      $sel.append($("<option></option>").attr("value", s).text(s));
    });
    if (current) $sel.val(current);
  }

  function uniqueRegions(rows) {
    var seen = {};
    var out = [];
    rows.forEach(function (r) {
      var s = trimOrEmpty(r.region);
      if (!s || seen[s]) return;
      seen[s] = true;
      out.push(s);
    });
    out.sort(function (a, b) {
      return a.localeCompare(b);
    });
    return out;
  }

  function populateRegionFilter(rows) {
    var $sel = $("#scdFilterRegion");
    var current = $sel.val() || "";
    $sel.empty().append('<option value="">All regions</option>');
    uniqueRegions(rows).forEach(function (s) {
      $sel.append($("<option></option>").attr("value", s).text(s));
    });
    if (current) $sel.val(current);
  }

  function filteredRows() {
    var sector = $("#scdFilterSector").val() || "";
    var region = $("#scdFilterRegion").val() || "";
    var dpi = $("#scdFilterDpi").val() || "";
    var search = trimOrEmpty($("#scdSearch").val()).toLowerCase();
    return allRows.filter(function (r) {
      if (sector && r.sector !== sector) return false;
      if (region && r.region !== region) return false;
      if (dpi === "dpi" && !r.isDpi) return false;
      if (dpi === "non-dpi" && r.isDpi) return false;
      if (search) {
        var haystack = ((r.siteId || "") + " " + (r.siteName || "")).toLowerCase();
        if (haystack.indexOf(search) === -1) return false;
      }
      return true;
    });
  }

  function renderCompletionMeter(stats, isOverall) {
    var ratio = Math.max(0, Math.min(1, Number(stats.ratio) || 0));
    var pct = stats.pct;
    var widthPct = Math.round(ratio * 1000) / 10;
    var fill = completionColor(ratio);
    var detail =
      String(stats.filled) + " / " + String(stats.total) + " fields";
    var cellClass = isOverall ? "scd-cell scd-cell-overall" : "scd-cell";
    var meterClass = isOverall ? "scd-meter scd-meter--overall" : "scd-meter";
    var sub =
      '<div class="scd-meter-sub">' +
      escapeHtml(String(stats.filled)) +
      " / " +
      escapeHtml(String(stats.total)) +
      "</div>";

    return (
      '<td class="' +
      cellClass +
      '">' +
      '<div class="' +
      meterClass +
      '" title="' +
      escapeHtml(detail) +
      '">' +
      '<div class="scd-meter-pct">' +
      escapeHtml(String(pct)) +
      "%</div>" +
      '<div class="scd-meter-track" aria-hidden="true">' +
      '<div class="scd-meter-fill" style="width:' +
      widthPct +
      "%;background:" +
      fill +
      '"></div></div>' +
      sub +
      "</div></td>"
    );
  }

  function renderReviewPill(row) {
    var label = trimOrEmpty(row.reviewStatus) || "Not set";
    var tone = row.reviewTone || "empty";
    return (
      '<td class="scd-cell"><span class="scd-review-pill scd-review-pill--' +
      escapeHtml(tone) +
      '" title="' +
      escapeHtml(label) +
      '">' +
      escapeHtml(label) +
      "</span></td>"
    );
  }

  function sortIndicator(colKey) {
    if (sortState.key !== colKey) return "↕";
    return sortState.dir === "asc" ? "↑" : "↓";
  }

  function renderTableHead() {
    var html = SORTABLE_COLUMNS.map(function (col) {
      var tip = "";
      if (col.section) {
        tip = infoTipHtml(col.label + " fields", fieldLabelsForSection(col.section));
      } else if (col.tipAllSections) {
        tip = infoTipHtml("Included in overall", allFieldLabels());
      }

      var thClass = col.overall ? ' class="scd-th-overall"' : "";
      var btnClass =
        "scd-th-btn" + (sortState.key === col.key ? " is-active" : "");

      return (
        "<th scope=\"col\"" +
        thClass +
        ">" +
        '<button type="button" class="' +
        btnClass +
        '" data-sort="' +
        escapeHtml(col.key) +
        '" aria-label="Sort by ' +
        escapeHtml(col.label) +
        '">' +
        escapeHtml(col.label) +
        '<span class="scd-sort-ind" aria-hidden="true">' +
        sortIndicator(col.key) +
        "</span>" +
        "</button>" +
        tip +
        "</th>"
      );
    }).join("");
    $("#scdTableHead").html('<th scope="col" style="width:50px"></th>' + html);
  }

  function compareRows(a, b) {
    var key = sortState.key;
    var dir = sortState.dir === "asc" ? 1 : -1;
    var av;
    var bv;

    if (key === "site") {
      av = a.label || "";
      bv = b.label || "";
      return av.localeCompare(bv) * dir;
    }
    if (key === "review") {
      av = trimOrEmpty(a.reviewStatus).toLowerCase();
      bv = trimOrEmpty(b.reviewStatus).toLowerCase();
      if (av === bv) return (a.label || "").localeCompare(b.label || "");
      return av.localeCompare(bv) * dir;
    }
    if (key === "overall") {
      av = a.overall.ratio;
      bv = b.overall.ratio;
    } else if (a.sections[key] && b.sections[key]) {
      av = a.sections[key].ratio;
      bv = b.sections[key].ratio;
    } else {
      av = 0;
      bv = 0;
    }

    if (av !== bv) return (av - bv) * dir;
    return (a.label || "").localeCompare(b.label || "");
  }

  function renderOverallCard(rows) {
    var totalFilled = 0;
    var totalCount = 0;
    rows.forEach(function (r) {
      totalFilled += r.overall.filled;
      totalCount += r.overall.total;
    });
    var ratio = totalCount ? totalFilled / totalCount : 0;
    var pct = totalCount ? Math.round(ratio * 100) : 0;
    var bg = completionColor(ratio);

    $("#scdOverallPill").html(
      '<div class="scd-meter scd-meter--hero" title="' +
        escapeHtml(String(totalFilled) + " / " + String(totalCount) + " fields") +
        '">' +
        '<div class="scd-meter-pct">' +
        escapeHtml(String(pct)) +
        "%</div>" +
        '<div class="scd-meter-track" aria-hidden="true">' +
        '<div class="scd-meter-fill" style="width:' +
        Math.round(ratio * 1000) / 10 +
        "%;background:" +
        bg +
        '"></div></div>' +
        '<div class="scd-meter-sub">' +
        escapeHtml(String(totalFilled)) +
        " / " +
        escapeHtml(String(totalCount)) +
        " fields</div></div>"
    );

    var dpiCount = rows.filter(function (r) {
      return r.isDpi;
    }).length;
    var completedCount = rows.filter(function (r) {
      return r.reviewTone === "done";
    }).length;

    $("#scdOverallMeta").html(
      "<span>Sites in view: <strong>" +
        escapeHtml(String(rows.length)) +
        "</strong></span>" +
        "<span>DPI: <strong>" +
        escapeHtml(String(dpiCount)) +
        "</strong></span>" +
        "<span>Review completed: <strong>" +
        escapeHtml(String(completedCount)) +
        "</strong></span>"
    );
  }

  function renderTable() {
    renderTableHead();

    var rows = filteredRows().slice().sort(compareRows);

    var $body = $("#scdTableBody");
    $body.empty();

    renderOverallCard(rows);

    if (!rows.length) {
      $("#scdEmpty").removeClass("is-hidden");
      $("#scdSummary").html(
        "<span>Showing <strong>0</strong> of " +
          escapeHtml(String(allRows.length)) +
          " sites</span>"
      );
      return;
    }

    $("#scdEmpty").addClass("is-hidden");

    $("#scdSummary").html(
      "<span>Showing <strong>" +
        escapeHtml(String(rows.length)) +
        "</strong> of " +
        escapeHtml(String(allRows.length)) +
        " sites · click a column header to sort</span>"
    );

    var pencilSvg =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor">' +
      '<path d="M11.13 1.47a1.5 1.5 0 0 1 2.12 0l1.28 1.28a1.5 1.5 0 0 1 0 2.12L5.93 13.47a1 1 0 0 1-.46.26l-3.2.9a.5.5 0 0 1-.6-.6l.9-3.2a1 1 0 0 1 .26-.46L11.13 1.47ZM12.2 2.53 3.9 10.83l-.56 2 2-.56 8.3-8.3-1.44-1.44Z"/>' +
      "</svg>";

    rows.forEach(function (r) {
      var abbr = sectorAbbr(r.sector);
      var sectorTag = abbr
        ? '<span class="' + sectorTagClass(abbr) + '">' + escapeHtml(abbr) + "</span>"
        : "";
      var dpiTag = r.isDpi
        ? '<span class="scd-dpi-tag">' +
          escapeHtml(r.dpiVariant || "DPI") +
          "</span>"
        : '<span class="scd-dpi-tag scd-dpi-tag--no">Non-DPI</span>';
      var href = itemDisplayUrl(r.Id);
      var editHref = itemEditUrl(r.Id);

      var html =
        "<tr>" +
        '<td class="scd-edit-cell">' +
        '<a class="scd-edit-btn" href="' +
        escapeHtml(editHref) +
        '" target="_blank" rel="noopener noreferrer" title="Edit entry">' +
        pencilSvg +
        "</a></td>" +
        '<td class="scd-site-cell">' +
        '<a class="scd-site-id" href="' +
        escapeHtml(href) +
        '" target="_blank" rel="noopener noreferrer">' +
        escapeHtml(r.siteId || "—") +
        "</a>" +
        '<span class="scd-site-name">' +
        escapeHtml(r.siteName || "") +
        "</span>" +
        sectorTag +
        dpiTag +
        "</td>" +
        renderReviewPill(r) +
        renderCompletionMeter(r.sections.general, false) +
        renderCompletionMeter(r.sections.people, false) +
        renderCompletionMeter(r.sections.maturity, false) +
        renderCompletionMeter(r.sections.counts, false) +
        renderCompletionMeter(r.overall, true) +
        "</tr>";
      $body.append(html);
    });
  }

  function init() {
    clearError();
    $("#scdLoading").removeClass("is-hidden").text("Loading sites…");
    $("#scdMain").addClass("is-hidden");

    fetchFieldsCatalog()
      .then(function (catalog) {
        resolveFieldDefs(catalog);
        if (!Object.keys(fieldsByKey).length) {
          throw new Error(
            "Could not resolve Site Master Data columns for the completion dashboard."
          );
        }
        return fetchAll(buildItemsQueryUrl());
      })
      .then(function (items) {
        allRows = (items || []).map(computeRow);
        var seenStatuses = {};
        allRows.forEach(function (r) {
          var s = trimOrEmpty(r.reviewStatus) || "(empty)";
          seenStatuses[s] = (seenStatuses[s] || 0) + 1;
        });
        console.log(
          "[Completion Dashboard] Review Status values in data:",
          seenStatuses,
          "List choices:",
          reviewStatusChoices
        );
        populateSectorFilter(allRows);
        populateRegionFilter(allRows);
        $("#scdLoading").addClass("is-hidden");
        $("#scdMain").removeClass("is-hidden");
        renderTable();
      })
      .fail(function (xhr) {
        $("#scdLoading").addClass("is-hidden");
        console.error("[Completion Dashboard] Init failed", xhr);
        var msg =
          (xhr && xhr.message) ||
          "Could not load Site Master Data. Ensure you are signed in and have list access.";
        if (xhr && xhr.status === 404) {
          msg = "List '" + LIST_NAME + "' was not found.";
        }
        showError(msg);
      });

    $("#scdFilterSector, #scdFilterRegion, #scdFilterDpi").on(
      "change",
      renderTable
    );

    $("#scdSearch").on("input", renderTable);

    $("#scdWidget").on("click", ".scd-th-btn", function () {
      var key = $(this).attr("data-sort");
      if (!key) return;
      if (sortState.key === key) {
        sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
      } else {
        sortState.key = key;
        sortState.dir = "asc";
      }
      renderTable();
    });
  }

  $(function () {
    if (!$("#scdWidget").length) return;
    init();
  });
})(window.jQuery);
