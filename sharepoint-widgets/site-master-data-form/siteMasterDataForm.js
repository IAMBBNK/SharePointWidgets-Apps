/* eslint-disable no-undef */
(function ($) {
  "use strict";

  /* =========================================================================
   * CONFIG
   * ========================================================================= */
  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var LIST_NAME = "Site Master Data";
  var REVIEWED_STATUS_LABEL = "Completed / Updated";

  // When true, site dropdown lists every site (ignores access roles).
  // Set back to false for production.
  var TEST_SHOW_ALL_SITES = false;

  /**
   * Form fields only (order = layout order). `titles` = SharePoint column Titles.
   * Pair half-width fields in twos for a clean 2-column grid; `full` spans both.
   */
  var FIELD_DEFS = [
    // General — paired for 2 columns
    {
      key: "SiteId",
      titles: ["Site ID"],
      type: "text",
      section: "general",
      label: "Site ID",
      hint: "Official site code: 4 letters + 1 digit (e.g. CHAL1, FRCA1)",
      placeholder: "e.g. CHAL1",
      full: false
    },
    {
      key: "SiteName",
      titles: ["Site Name"],
      type: "text",
      section: "general",
      label: "Site Name",
      hint: "Common site name, with suffix if needed for identification",
      placeholder: "e.g. Altdorf",
      full: false
    },
    {
      key: "SiteAddress",
      titles: ["Site Address"],
      type: "text",
      section: "general",
      label: "Site Address",
      hint: "Full street address where the site is located",
      placeholder: "Street, postal code, city, country",
      full: true
    },
    {
      key: "City",
      titles: ["City"],
      type: "text",
      section: "general",
      label: "City",
      hint: "City in which the site resides",
      placeholder: "e.g. St. Louis",
      full: false
    },
    {
      key: "Country",
      titles: ["Country"],
      type: "text",
      section: "general",
      label: "Country",
      hint: "Country in which the site resides",
      placeholder: "e.g. France",
      full: false
    },
    {
      key: "Region",
      titles: ["Region"],
      knownInternal: "field_4",
      type: "choice",
      section: "general",
      label: "Region",
      hint: "Geographic region (choice)",
      choices: ["EMEA", "NASA", "APAC"],
      full: false
    },
    {
      key: "Sector",
      titles: ["Sector"],
      type: "choice",
      section: "general",
      label: "Sector",
      hint: "Business sector (choice)",
      choices: ["Electronics", "Healthcare", "Life Science", "Lifescience"],
      full: false
    },
    {
      key: "LegalEntityName",
      titles: ["Legal Entity Name"],
      type: "text",
      section: "general",
      label: "Legal Entity Name",
      hint: "Full official name of the legal entity",
      placeholder: "e.g. SAFC Hitech Taiwan Co., Ltd.",
      full: false
    },
    {
      key: "LegalEntityCode",
      titles: ["Legal Entity Code"],
      type: "text",
      section: "general",
      label: "Legal Entity Code",
      hint: "Official 4-digit numeric code",
      placeholder: "e.g. 2088",
      full: false
    },
    {
      key: "BusinessUnit",
      titles: ["Business Unit"],
      type: "text",
      section: "general",
      label: "Business Unit",
      hint: "Global operations unit / function (e.g. ISCO, DS&S, GBM, GHO)",
      placeholder: "e.g. ISCO",
      full: false
    },
    {
      key: "BusinessField",
      titles: ["Business Field / Type / Service"],
      type: "text",
      section: "general",
      label: "Business Field / Type / Service",
      hint: "Main activity / scope (e.g. Manufacturing, Lab, DC)",
      placeholder: "e.g. Manufacturing",
      full: false
    },
    {
      key: "Employees",
      titles: ["# of Employees"],
      type: "text",
      section: "general",
      label: "# of Employees",
      hint: "Total number of employees at the site (number)",
      placeholder: "e.g. 250",
      full: false
    },

    // People — 2 columns (cardinality comes from list: single vs multi)
    {
      key: "SiteDataAdmin",
      titles: ["Site Data Admin"],
      type: "person",
      section: "people",
      label: "Site Data Admin",
      hint: "Owner of site data updates & quality",
      full: false
    },
    {
      key: "SiteHead",
      titles: ["Site Head"],
      type: "person",
      section: "people",
      label: "Site Head",
      hint: "Overall site leader / manager",
      full: false
    },
    {
      key: "DigitalLead",
      titles: ["Digtial Lead / SPOC", "Digital Lead / SPOC", "Digital Lead"],
      type: "person",
      section: "people",
      label: "Digital Lead / SPOC",
      hint: "Main contact for IT/OT topics at the site",
      full: false
    },
    {
      key: "OtServiceDeliveryMgr",
      titles: ["OT Service Delivery Mgr."],
      type: "person",
      section: "people",
      label: "OT Service Delivery Mgr.",
      hint: "Per OT Standard",
      full: false
    },
    {
      key: "OtProjectDeliveryMgr",
      titles: ["OT Project Delivery Mgr."],
      type: "person",
      section: "people",
      label: "OT Project Delivery Mgr.",
      hint: "Per OT Standard",
      full: false
    },
    {
      key: "OtEngineers",
      titles: ["OT Engineers"],
      type: "person",
      section: "people",
      label: "OT Engineers",
      hint: "Per OT Standard",
      full: false
    },
    {
      key: "OtSolutionArchitect",
      titles: ["OT Solution Architect"],
      type: "person",
      section: "people",
      label: "OT Solution Architect",
      hint: "Per OT Standard",
      full: false
    },
    {
      key: "CyberSecurityMgr",
      titles: ["CyberSecurity Mgr."],
      type: "person",
      section: "people",
      label: "CyberSecurity Mgr.",
      hint: "Per OT Standard",
      full: false
    },
    {
      key: "ItBusinessPartner",
      titles: ["IT Business Partner"],
      type: "person",
      section: "people",
      label: "IT Business Partner",
      hint: "Per IT Standard",
      full: false
    },

    // Maturity / assessments
    {
      key: "RoadmapMaturity",
      titles: ["Roadmap Maturity"],
      type: "choice",
      section: "maturity",
      label: "Roadmap Maturity",
      hint: "Digitalization roadmap stage (select one)",
      choices: [
        "Stage 1 | Site Vision Roadmap (idea)",
        "Stage 2 | Approved Site Masterplan (Site LT approved)",
        "Stage 3 | Full Site Masterplan (budgeted / ongoing)"
      ],
      full: true
    },
    {
      key: "CurrentDigitalMaturity",
      titles: ["Current Maturity (BPOG)"],
      type: "text",
      section: "maturity",
      label: "Current Site Digital Maturity (BioPhorum DPMM)",
      hint: "BioPhorum level text, e.g. Level 3 | Connected Plant",
      placeholder: "e.g. Level 3 | Connected Plant (value 2.5 - 3.4)",
      full: false
    },
    {
      key: "TargetDigitalMaturity",
      titles: ["Target Maturity (BPOG)"],
      type: "text",
      section: "maturity",
      label: "Target Site Digital Maturity (BioPhorum DPMM)",
      hint: "Target BioPhorum level text",
      placeholder: "e.g. Level 4 | Predictive Plant (value: 3.5 - 4.5)",
      full: false
    },
    {
      key: "DpiVariant",
      titles: ["(Target) DPI Variant"],
      type: "choice",
      section: "maturity",
      label: "(Target) DPI Variant",
      hint: "DPI variant per IT Standards (select one)",
      choices: ["Platinum", "Premium", "Premium Lite", "Essential", "N/A"],
      full: false
    },
    {
      key: "ItAssessment",
      titles: ["IT Assessment"],
      type: "choice",
      section: "maturity",
      label: "IT Assessment Status",
      hint: "Official IT Assessment / WP1 status",
      choices: ["N/A", "Planning", "Ongoing", "Completed"],
      full: false
    },
    {
      key: "OtAssessment",
      titles: ["OT Assessment"],
      type: "choice",
      section: "maturity",
      label: "OT Assessment Status",
      hint: "Official OT Assessment status",
      choices: ["N/A", "Planning", "Ongoing", "Completed"],
      full: false
    },
    {
      key: "DpiCoreImplementation",
      titles: ["DPI Core Implementation"],
      type: "choice",
      section: "maturity",
      label: "DPI Core Implementation Status",
      hint: "DPI Core / WP2 status",
      choices: ["N/A", "Planning", "Ongoing", "Completed"],
      full: false
    },
    {
      key: "DpiOperations",
      titles: ["DPI Operations"],
      type: "choice",
      section: "maturity",
      label: "DPI Operations Status",
      hint: "DPI Operations / WP5 status",
      choices: ["N/A", "Planning", "Merck-managed", "Cognizant-managed"],
      full: false
    },
    {
      key: "OtSystemIntegration",
      titles: ["OT System Integration"],
      type: "choice",
      section: "maturity",
      label: "OT System Integration Status",
      hint: "OT migration into DPI / WP6 status",
      choices: ["N/A", "Planning", "Ongoing", "Completed"],
      full: false
    },

    // OT application counts
    {
      key: "TotalOtApplications",
      titles: ["Total Number of OT Applications"],
      type: "text",
      section: "counts",
      label: "Total Number of OT Applications",
      hint: "Estimated count of OT applications (number)",
      placeholder: "e.g. 12",
      full: false
    },
    {
      key: "OtSystemsInGear",
      titles: ["Gear"],
      type: "text",
      section: "counts",
      label: "Number OT Systems inventorized in GEAR",
      hint: "Count of OT systems in GEAR (number)",
      placeholder: "e.g. 10",
      full: false
    },
    {
      key: "OtAppsToMigrate",
      titles: ["Number of OT Applications to be migrated to DPI"],
      type: "text",
      section: "counts",
      label: "Number of OT Applications to be migrated to DPI",
      hint: "Target count to migrate into DPI (number)",
      placeholder: "e.g. 8",
      full: false
    },
    {
      key: "OtAppsMigrated",
      titles: ["Number of migrated OT Applications"],
      type: "text",
      section: "counts",
      label: "Number of migrated OT Applications",
      hint: "Already migrated into DPI (number)",
      placeholder: "e.g. 3",
      full: false
    }
  ];

  // Who may open/edit a site in production (TEST_SHOW_ALL_SITES bypasses this)
  var ACCESS_ROLE_KEYS = ["SiteDataAdmin", "SiteHead", "DigitalLead"];
  var EXTRA_ACCESS_TITLES = [];

  /* =========================================================================
   * STATE
   * ========================================================================= */
  var currentUser = null;
  var listItemEntityType = null;
  var fieldsByKey = {}; // key -> resolved field meta
  var reviewStatusInternal = "Review_x0020_Status";
  var matchedSites = []; // { Id, label, raw }
  var selectedSiteId = null;
  var peoplePickers = {}; // key -> picker instance
  var formDigestCache = { value: null, expires: 0 };
  var fieldCatalogTitles = []; // display titles from list (for suggestions)
  var fieldDiagnostics = {
    unresolved: [],
    typeMismatches: [],
    badSelect: [],
    resolvedCount: 0
  };

  /* =========================================================================
   * UTILITIES
   * ========================================================================= */
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

  /**
   * SharePoint REST $select/$expand: internal names that start with "_"
   * (e.g. columns originally named "# of …") must be prefixed with "OData_".
   * Write payloads still use the raw InternalName.
   */
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
    if (odata !== internalName && Object.prototype.hasOwnProperty.call(item, odata)) {
      return item[odata];
    }
    // verbose nested sometimes uses different casing
    if (item[odata] !== undefined) return item[odata];
    if (item[internalName] !== undefined) return item[internalName];
    return null;
  }

  function focusStatusBanner($el) {
    if (!$el || !$el.length) return;
    try {
      var node = $el[0];
      if (node && typeof node.scrollIntoView === "function") {
        node.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }
    } catch (_e) {
      /* ignore */
    }
  }

  function showSaveStatus(html) {
    var $sticky = $("#smdSaveStatus");
    if ($sticky.length) {
      $sticky.html(html);
      focusStatusBanner($sticky);
    }
  }

  function clearSaveStatus() {
    $("#smdSaveStatus").empty();
  }

  function showError(message, detailLines) {
    var html =
      '<div class="error-banner" role="alert"><div>' +
      escapeHtml(message) +
      "</div>";
    if (detailLines && detailLines.length) {
      html += '<ul class="smd-diag-list">';
      detailLines.forEach(function (line) {
        html += "<li>" + escapeHtml(line) + "</li>";
      });
      html += "</ul>";
    }
    html += "</div>";
    $("#smdError").html(html);
    showSaveStatus(html);
  }

  function clearError() {
    $("#smdError").empty();
    clearSaveStatus();
  }

  function showSuccess(message) {
    var html =
      '<div class="success-banner" role="status">' +
      escapeHtml(message) +
      "</div>";
    $("#smdSuccess").html(html);
    showSaveStatus(html);
  }

  function clearSuccess() {
    $("#smdSuccess").empty();
    clearSaveStatus();
  }

  function showDiagnosticPanel(title, detailLines) {
    if (!detailLines || !detailLines.length) {
      $("#smdDiag").empty();
      return;
    }
    var html =
      '<div class="info-banner smd-diag-banner" role="status"><strong>' +
      escapeHtml(title) +
      "</strong><ul class=\"smd-diag-list\">";
    detailLines.forEach(function (line) {
      html += "<li>" + escapeHtml(line) + "</li>";
    });
    html +=
      "</ul><p class=\"smd-diag-hint\">Update <code>FIELD_DEFS</code> titles in siteMasterDataForm.js to match the list column Title exactly. Open the browser console (F12) for a full mapping table and all list column titles.</p></div>";
    $("#smdDiag").html(html);
  }

  function suggestTitles(wantedTitles) {
    var suggestions = [];
    var needles = (wantedTitles || [])
      .map(function (t) {
        return trimOrEmpty(t).toLowerCase();
      })
      .filter(Boolean);

    fieldCatalogTitles.forEach(function (title) {
      var low = title.toLowerCase();
      for (var i = 0; i < needles.length; i++) {
        var n = needles[i];
        if (!n) continue;
        if (low === n || low.indexOf(n) !== -1 || n.indexOf(low) !== -1) {
          suggestions.push(title);
          return;
        }
        // token overlap (e.g. "Digital Lead" vs "Digtial Lead")
        var tokens = n.split(/[^a-z0-9]+/).filter(function (x) {
          return x.length >= 4;
        });
        for (var t = 0; t < tokens.length; t++) {
          if (low.indexOf(tokens[t]) !== -1) {
            suggestions.push(title);
            return;
          }
        }
      }
    });

    return suggestions.slice(0, 5);
  }

  function showInfo(message) {
    $("#smdInfo").html(
      '<div class="info-banner">' + escapeHtml(message) + "</div>"
    );
  }

  function clearInfo() {
    $("#smdInfo").empty();
  }

  function clearDiagnostics() {
    $("#smdDiag").empty();
  }

  function normalizeEmail(email) {
    return trimOrEmpty(email).toLowerCase();
  }

  function peopleFromField(personField) {
    if (!personField) return [];
    var arr = Array.isArray(personField)
      ? personField
      : personField.results
        ? personField.results
        : [personField];
    return arr
      .filter(Boolean)
      .map(function (p) {
        return {
          Id: p.Id != null ? p.Id : p.id != null ? p.id : null,
          Title: p.Title || p.title || "",
          Email: p.EMail || p.Email || p.email || ""
        };
      })
      .filter(function (p) {
        return p.Id != null || p.Email || p.Title;
      });
  }

  function userMatchesPersonList(user, people) {
    if (!user || !people || !people.length) return false;
    var uid = user.Id;
    var uemail = normalizeEmail(user.email);
    for (var i = 0; i < people.length; i++) {
      var p = people[i];
      if (uid != null && p.Id != null && Number(p.Id) === Number(uid)) {
        return true;
      }
      if (uemail && normalizeEmail(p.Email) === uemail) {
        return true;
      }
    }
    return false;
  }

  function mergeChoiceOptions(preferred, fromList) {
    var seen = {};
    var out = [];
    function add(v) {
      var t = trimOrEmpty(v);
      if (!t) return;
      var k = t.toLowerCase();
      if (seen[k]) return;
      seen[k] = true;
      out.push(t);
    }
    (fromList || []).forEach(add);
    (preferred || []).forEach(add);
    return out;
  }

  /* =========================================================================
   * SHAREPOINT API
   * ========================================================================= */
  function xhrErrorMessage(xhr, fallback) {
    var status = xhr && xhr.status ? "HTTP " + xhr.status : "";
    var body = "";
    try {
      if (xhr && xhr.responseText) {
        var parsed = JSON.parse(xhr.responseText);
        body =
          (parsed &&
            parsed.error &&
            parsed.error.message &&
            (parsed.error.message.value || parsed.error.message)) ||
          (parsed && parsed["odata.error"] && parsed["odata.error"].message &&
            (parsed["odata.error"].message.value ||
              parsed["odata.error"].message)) ||
          "";
      }
    } catch (_e) {
      body = trimOrEmpty(xhr && xhr.responseText).slice(0, 300);
    }
    if (xhr && xhr.message && !xhr.status) {
      return xhr.message;
    }
    var parts = [fallback];
    if (status) parts.push(status);
    if (body) parts.push(String(body));
    return parts.join(" — ");
  }

  function getFormDigest() {
    var now = Date.now();
    if (formDigestCache.value && formDigestCache.expires > now + 5000) {
      return $.Deferred().resolve(formDigestCache.value).promise();
    }

    return $.ajax({
      url: SITE_URL + "/_api/contextinfo",
      type: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
      }
    }).then(function (data) {
      var info = data.d.GetContextWebInformation;
      formDigestCache.value = info.FormDigestValue;
      formDigestCache.expires =
        now + (info.FormDigestTimeoutSeconds || 1800) * 1000;
      return formDigestCache.value;
    });
  }

  function getCurrentUser() {
    return $.ajax({
      url: SITE_URL + "/_api/web/currentuser?$select=Id,Title,Email,LoginName",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      var d = data.d;
      var email = d.Email || "";
      var loginName = d.LoginName || "";
      if (!email && loginName.indexOf("|") !== -1) {
        var parts = loginName.split("|");
        var candidate = parts[parts.length - 1];
        if (candidate.indexOf("@") !== -1) email = candidate;
      }
      return {
        Id: d.Id,
        name: d.Title || d.LoginName || "Unknown user",
        email: email,
        loginName: loginName
      };
    });
  }

  function fetchListEntityType() {
    if (listItemEntityType) {
      return $.Deferred().resolve(listItemEntityType).promise();
    }
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle() +
        "')?$select=ListItemEntityTypeFullName",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      listItemEntityType = data.d.ListItemEntityTypeFullName;
      return listItemEntityType;
    });
  }

  function fetchFieldsCatalog() {
    // Keep $select minimal — selecting Choices/AllowMultipleValues on all field
    // types can 400 on some tenants.
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle() +
        "')/fields?$select=" +
        encodeURIComponent(
          "Title,InternalName,TypeAsString,AllowMultipleValues,Choices,Hidden,ReadOnlyField"
        ) +
        "&$top=5000",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      var results = (data.d && data.d.results) || [];
      return results.filter(function (f) {
        return f && !f.Hidden;
      });
    });
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

  function resolveFieldDefs(catalog) {
    fieldCatalogTitles = [];
    (catalog || []).forEach(function (f) {
      var t = trimOrEmpty(f.Title);
      if (t) fieldCatalogTitles.push(t);
    });
    fieldCatalogTitles.sort(function (a, b) {
      return a.localeCompare(b);
    });

    var byTitle = {};
    (catalog || []).forEach(function (f) {
      var t = trimOrEmpty(f.Title);
      if (!t) return;
      byTitle[t.toLowerCase()] = f;
    });

    var byInternal = {};
    (catalog || []).forEach(function (f) {
      if (f.InternalName) byInternal[f.InternalName] = f;
    });

    // Review Status
    var reviewHit =
      byTitle["review status"] || byInternal["Review_x0020_Status"];
    if (reviewHit) {
      reviewStatusInternal = reviewHit.InternalName;
    } else {
      console.warn(
        "[Site Master Data Form] Review Status column not found. Save will try InternalName Review_x0020_Status."
      );
    }

    fieldsByKey = {};
    fieldDiagnostics = {
      unresolved: [],
      typeMismatches: [],
      badSelect: [],
      resolvedCount: 0
    };

    FIELD_DEFS.forEach(function (def) {
      var hit = null;
      var matchedVia = "";
      // Prefer display Title (column names may have been renamed; internal names stay old)
      for (var i = 0; i < def.titles.length; i++) {
        var cand = byTitle[def.titles[i].toLowerCase()];
        if (cand) {
          hit = cand;
          matchedVia = "title:" + def.titles[i];
          break;
        }
      }
      if (!hit && def.knownInternal && byInternal[def.knownInternal]) {
        hit = byInternal[def.knownInternal];
        matchedVia = "internal:" + def.knownInternal;
      }
      // Fuzzy: only if needle is reasonably specific (avoid short false matches)
      if (!hit) {
        var needle = def.titles[0].toLowerCase();
        if (needle.length >= 8) {
          Object.keys(byTitle).forEach(function (k) {
            if (hit) return;
            if (k.indexOf(needle) !== -1) {
              hit = byTitle[k];
              matchedVia = "fuzzy:" + hit.Title;
            }
          });
        }
      }

      if (!hit) {
        var suggestions = suggestTitles(def.titles);
        fieldDiagnostics.unresolved.push({
          key: def.key,
          label: def.label || def.key,
          titlesTried: def.titles.slice(),
          suggestions: suggestions
        });
        return;
      }

      var listChoices = [];
      if (hit.Choices && hit.Choices.results) {
        listChoices = hit.Choices.results;
      } else if (Array.isArray(hit.Choices)) {
        listChoices = hit.Choices;
      }
      var typeAs = hit.TypeAsString || "";
      var isPerson = typeAs === "User" || typeAs === "UserMulti";
      if (def.type === "person" && !isPerson) {
        fieldDiagnostics.typeMismatches.push({
          key: def.key,
          label: def.label || def.key,
          expected: "Person or Group",
          actual: typeAs || "(empty)",
          listTitle: hit.Title,
          internalName: hit.InternalName
        });
        console.warn(
          "[Site Master Data Form] Expected person field but got " +
            typeAs +
            " for '" +
            hit.Title +
            "' (" +
            hit.InternalName +
            "). Mapped as text instead of people picker."
        );
      }

      fieldsByKey[def.key] = {
        key: def.key,
        def: def,
        Title: hit.Title,
        InternalName: hit.InternalName,
        TypeAsString: typeAs,
        AllowMultipleValues: !!hit.AllowMultipleValues || typeAs === "UserMulti",
        choices: mergeChoiceOptions(def.choices, listChoices),
        isPerson: isPerson,
        matchedVia: matchedVia
      };
      fieldDiagnostics.resolvedCount += 1;
    });

    // Extra Site Director column (if distinct from Site Head)
    EXTRA_ACCESS_TITLES.forEach(function (title) {
      var f = byTitle[title.toLowerCase()];
      if (!f) return;
      var already = Object.keys(fieldsByKey).some(function (k) {
        return fieldsByKey[k].InternalName === f.InternalName;
      });
      if (already) return;
      if (f.TypeAsString !== "User" && f.TypeAsString !== "UserMulti") return;
      var extraKey = "AccessExtra_" + f.InternalName;
      fieldsByKey[extraKey] = {
        key: extraKey,
        def: {
          key: extraKey,
          type: "person",
          section: "people",
          label: f.Title,
          full: true
        },
        Title: f.Title,
        InternalName: f.InternalName,
        TypeAsString: f.TypeAsString,
        AllowMultipleValues:
          !!f.AllowMultipleValues || f.TypeAsString === "UserMulti",
        choices: [],
        isPerson: true,
        accessOnly: false,
        matchedVia: "title:" + title
      };
      fieldDiagnostics.resolvedCount += 1;
    });

    logFieldDiagnostics();
    return fieldsByKey;
  }

  function logFieldDiagnostics() {
    console.group("[Site Master Data Form] Field mapping");
    console.log(
      "Resolved " +
        fieldDiagnostics.resolvedCount +
        " / " +
        FIELD_DEFS.length +
        " configured fields"
    );
    console.table(
      Object.keys(fieldsByKey).map(function (k) {
        var f = fieldsByKey[k];
        return {
          key: k,
          listTitle: f.Title,
          internalName: f.InternalName,
          type: f.TypeAsString,
          matchedVia: f.matchedVia || "",
          person: f.isPerson
        };
      })
    );

    if (fieldDiagnostics.unresolved.length) {
      console.warn("Unresolved (no matching list column Title):");
      fieldDiagnostics.unresolved.forEach(function (u) {
        console.warn(
          "  • " +
            u.label +
            " — tried: [" +
            u.titlesTried.join(" | ") +
            "]" +
            (u.suggestions.length
              ? " — similar on list: [" + u.suggestions.join(" | ") + "]"
              : " — no similar titles found")
        );
      });
    }

    if (fieldDiagnostics.typeMismatches.length) {
      console.warn("Type mismatches (expected Person, got something else):");
      fieldDiagnostics.typeMismatches.forEach(function (m) {
        console.warn(
          "  • " +
            m.label +
            " → list '" +
            m.listTitle +
            "' (" +
            m.internalName +
            ") is " +
            m.actual
        );
      });
    }

    console.log("All list column titles (" + fieldCatalogTitles.length + "):");
    console.log(fieldCatalogTitles.slice().sort().join("\n"));
    console.groupEnd();
  }

  function fieldDiagLines() {
    var lines = [];
    fieldDiagnostics.unresolved.forEach(function (u) {
      var line =
        "Missing: \"" +
        u.label +
        "\" (tried: " +
        u.titlesTried.join(", ") +
        ")";
      if (u.suggestions.length) {
        line += " — similar list columns: " + u.suggestions.join(", ");
      }
      lines.push(line);
    });
    fieldDiagnostics.typeMismatches.forEach(function (m) {
      lines.push(
        "Type mismatch: \"" +
          m.label +
          "\" matched list column \"" +
          m.listTitle +
          "\" (" +
          m.internalName +
          ") which is " +
          m.actual +
          ", expected Person or Group"
      );
    });
    fieldDiagnostics.badSelect.forEach(function (b) {
      lines.push(
        "Bad $select/$expand: \"" +
          b.label +
          "\" → InternalName \"" +
          b.internalName +
          "\" (" +
          (b.type || "?") +
          ")" +
          (b.error ? " — " + b.error : "")
      );
    });
    return lines;
  }

  function buildProbeUrl(selectParts, expandParts) {
    var select = ["Id"].concat(selectParts || []);
    var seen = {};
    select = select.filter(function (x) {
      if (!x || seen[x]) return false;
      seen[x] = true;
      return true;
    });
    var url =
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeListTitle() +
      "')/items?$select=" +
      encodeURIComponent(select.join(",")) +
      "&$top=1";
    if (expandParts && expandParts.length) {
      url += "&$expand=" + encodeURIComponent(expandParts.join(","));
    }
    return url;
  }

  function probeSelect(selectParts, expandParts) {
    return $.ajax({
      url: buildProbeUrl(selectParts, expandParts),
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    });
  }

  /**
   * When the full items query fails, test each resolved field alone to find
   * which InternalName breaks $select / $expand.
   */
  function diagnoseBadSelectFields() {
    var deferred = $.Deferred();
    var keys = Object.keys(fieldsByKey);
    var bad = [];
    var i = 0;

    function next() {
      if (i >= keys.length) {
        fieldDiagnostics.badSelect = bad;
        deferred.resolve(bad);
        return;
      }

      var key = keys[i++];
      var f = fieldsByKey[key];
      var selectParts;
      var expandParts;

      if (f.isPerson) {
        var personOd = toODataSelectName(f.InternalName);
        selectParts = [
          personOd + "/Id",
          personOd + "/Title",
          personOd + "/EMail"
        ];
        expandParts = [personOd];
      } else {
        selectParts = [toODataSelectName(f.InternalName)];
        expandParts = [];
      }

      probeSelect(selectParts, expandParts)
        .done(function () {
          next();
        })
        .fail(function (xhr) {
          bad.push({
            key: key,
            label: (f.def && f.def.label) || f.Title || key,
            listTitle: f.Title,
            internalName: f.InternalName,
            type: f.TypeAsString,
            error: xhrErrorMessage(xhr, "")
          });
          console.warn(
            "[Site Master Data Form] Field breaks items query:",
            key,
            f.InternalName,
            xhr && xhr.responseText
          );
          next();
        });
    }

    // Also probe Review Status if set
    function start() {
      if (!reviewStatusInternal) {
        next();
        return;
      }
      probeSelect([toODataSelectName(reviewStatusInternal)], [])
        .done(function () {
          next();
        })
        .fail(function (xhr) {
          bad.push({
            key: "ReviewStatus",
            label: "Review Status",
            listTitle: "Review Status",
            internalName: reviewStatusInternal,
            type: "Choice?",
            error: xhrErrorMessage(xhr, "")
          });
          next();
        });
    }

    start();
    return deferred.promise();
  }

  function dropBadFieldsFromMap(badList) {
    (badList || []).forEach(function (b) {
      if (b.key === "ReviewStatus") {
        reviewStatusInternal = null;
        return;
      }
      if (fieldsByKey[b.key]) {
        delete fieldsByKey[b.key];
      }
    });
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

  function siteLabel(item) {
    var idField = fieldsByKey.SiteId;
    var nameField = fieldsByKey.SiteName;
    var siteId = idField
      ? trimOrEmpty(readFieldFromItem(item, idField.InternalName))
      : "";
    var siteName = nameField
      ? trimOrEmpty(readFieldFromItem(item, nameField.InternalName))
      : "";
    if (!siteId && item.Title) siteId = trimOrEmpty(item.Title);
    if (siteId && siteName) return siteId + " — " + siteName;
    return siteId || siteName || "Item " + item.Id;
  }

  function userCanEditItem(user, item) {
    for (var i = 0; i < ACCESS_ROLE_KEYS.length; i++) {
      var meta = fieldsByKey[ACCESS_ROLE_KEYS[i]];
      if (!meta) continue;
      var people = peopleFromField(readFieldFromItem(item, meta.InternalName));
      if (userMatchesPersonList(user, people)) return true;
    }
    return false;
  }

  function ensureUser(loginOrEmail) {
    return getFormDigest().then(function (digest) {
      return $.ajax({
        url: SITE_URL + "/_api/web/ensureuser",
        type: "POST",
        data: JSON.stringify({
          logonName: loginOrEmail
        }),
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest
        }
      }).then(function (data) {
        var d = data.d;
        return {
          Id: d.Id,
          Title: d.Title || "",
          Email: d.Email || "",
          LoginName: d.LoginName || loginOrEmail
        };
      });
    });
  }

  function searchPeople(query) {
    return getFormDigest().then(function (digest) {
      var payload = {
        queryParams: {
          __metadata: {
            type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"
          },
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: 8,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: query
        }
      };

      return $.ajax({
        url:
          SITE_URL +
          "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser",
        type: "POST",
        data: JSON.stringify(payload),
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest
        }
      }).then(function (data) {
        var raw = data.d.ClientPeoplePickerSearchUser;
        var parsed = [];
        try {
          parsed = typeof raw === "string" ? JSON.parse(raw) : raw || [];
        } catch (e) {
          parsed = [];
        }
        return (parsed || [])
          .map(function (ent) {
            var ed = ent.EntityData || {};
            return {
              Key: ent.Key || "",
              DisplayText: ent.DisplayText || ed.Title || "",
              Description:
                ent.Description || ed.Email || ed.SPUserEmail || "",
              Email: ed.Email || ed.SPUserEmail || ""
            };
          })
          .filter(function (e) {
            return e.Key;
          });
      });
    });
  }

  /* =========================================================================
   * PEOPLE PICKER CONTROL
   * ========================================================================= */
  function createPeoplePicker($container, options) {
    var multi = !!options.multi;
    var selected = [];
    var searchTimer = null;
    var activeIndex = -1;
    var placeholderText = multi
      ? "Search and add people…"
      : "Search and select one person…";

    var $wrap = $('<div class="pp-wrap"></div>');
    var $chips = $('<div class="pp-chips"></div>');
    var $input = $(
      '<input type="text" class="pp-input" autocomplete="off" />'
    ).attr("placeholder", placeholderText);
    var $suggestions = $('<div class="pp-suggestions is-hidden"></div>');

    $wrap.append($chips, $input, $suggestions);
    $container.empty().append($wrap);

    function renderChips() {
      $chips.empty();
      selected.forEach(function (p, idx) {
        var label = p.Title || p.Email || "User " + p.Id;
        if (p.Email && p.Title) label = p.Title;
        var $chip = $(
          '<span class="pp-chip"><span class="pp-chip-text"></span><button type="button" class="pp-chip-remove" aria-label="Remove">&times;</button></span>'
        );
        $chip.find(".pp-chip-text").text(label + (p.Email ? " (" + p.Email + ")" : ""));
        $chip.find(".pp-chip-remove").on("click", function () {
          selected.splice(idx, 1);
          renderChips();
          updateInputState();
        });
        $chips.append($chip);
      });
    }

    function updateInputState() {
      if (!multi && selected.length > 0) {
        $input.prop("disabled", true).attr("placeholder", "");
      } else {
        $input.prop("disabled", false).attr("placeholder", placeholderText);
      }
    }

    function hideSuggestions() {
      $suggestions.addClass("is-hidden").empty();
      $wrap.removeClass("pp-open-up");
      activeIndex = -1;
    }

    function positionSuggestions() {
      $wrap.removeClass("pp-open-up");
      $suggestions.removeClass("is-hidden");
      var rect = $wrap[0].getBoundingClientRect();
      var spaceBelow = window.innerHeight - rect.bottom;
      var needed = Math.min(220, $suggestions.outerHeight() || 180) + 12;
      if (spaceBelow < needed && rect.top > spaceBelow) {
        $wrap.addClass("pp-open-up");
      }
    }

    function showSuggestions(items) {
      $suggestions.empty();
      if (!items.length) {
        $suggestions.append('<div class="pp-empty">No matches</div>');
        positionSuggestions();
        return;
      }
      items.forEach(function (ent, i) {
        var $btn = $(
          '<button type="button" class="pp-suggestion"></button>'
        );
        $btn.append(document.createTextNode(ent.DisplayText || ent.Key));
        if (ent.Email || ent.Description) {
          $btn.append(
            $('<span class="pp-suggestion-email"></span>').text(
              ent.Email || ent.Description
            )
          );
        }
        $btn.on("mousedown", function (e) {
          e.preventDefault();
          pickEntity(ent);
        });
        if (i === activeIndex) $btn.addClass("is-active");
        $suggestions.append($btn);
      });
      positionSuggestions();
    }

    function pickEntity(ent) {
      hideSuggestions();
      $input.val("");
      var login = ent.Key;
      $input.prop("disabled", true).attr("placeholder", "Adding…");

      ensureUser(login)
        .then(function (user) {
          var person = {
            Id: user.Id,
            Title: user.Title || ent.DisplayText || "",
            Email: user.Email || ent.Email || ""
          };
          if (!multi) selected = [person];
          else {
            var exists = selected.some(function (s) {
              return Number(s.Id) === Number(person.Id);
            });
            if (!exists) selected.push(person);
          }
          renderChips();
          updateInputState();
        })
        .fail(function (xhr) {
          console.error("ensureUser failed", xhr && xhr.responseText);
          showError(
            "Could not add user. Try again or contact the site owner."
          );
          updateInputState();
        });
    }

    $input.on("input", function () {
      var q = trimOrEmpty($input.val());
      clearTimeout(searchTimer);
      if (q.length < 2) {
        hideSuggestions();
        return;
      }
      searchTimer = setTimeout(function () {
        searchPeople(q)
          .then(function (items) {
            if (trimOrEmpty($input.val()) !== q) return;
            activeIndex = -1;
            showSuggestions(items);
          })
          .fail(function () {
            hideSuggestions();
          });
      }, 300);
    });

    $input.on("keydown", function (e) {
      var $opts = $suggestions.find(".pp-suggestion");
      if (e.key === "Escape") {
        hideSuggestions();
        return;
      }
      if (!$opts.length || $suggestions.hasClass("is-hidden")) return;
      if (e.key === "ArrowDown") {
        e.preventDefault();
        activeIndex = Math.min(activeIndex + 1, $opts.length - 1);
        $opts.removeClass("is-active").eq(activeIndex).addClass("is-active");
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        activeIndex = Math.max(activeIndex - 1, 0);
        $opts.removeClass("is-active").eq(activeIndex).addClass("is-active");
      } else if (e.key === "Enter" && activeIndex >= 0) {
        e.preventDefault();
        $opts.eq(activeIndex).trigger("mousedown");
      }
    });

    $input.on("blur", function () {
      setTimeout(hideSuggestions, 150);
    });

    return {
      setValue: function (people) {
        selected = (people || []).map(function (p) {
          return {
            Id: p.Id,
            Title: p.Title || "",
            Email: p.Email || ""
          };
        });
        renderChips();
        updateInputState();
      },
      getValue: function () {
        return selected.slice();
      },
      getIds: function () {
        return selected
          .map(function (p) {
            return p.Id;
          })
          .filter(function (id) {
            return id != null;
          });
      }
    };
  }

  /* =========================================================================
   * FORM RENDER / PREFILL
   * ========================================================================= */
  function sectionContainer(section) {
    if (section === "general") return $("#smdSectionGeneral");
    if (section === "people") return $("#smdSectionPeople");
    if (section === "maturity") return $("#smdSectionMaturity");
    return $("#smdSectionCounts");
  }

  function renderFormFields() {
    $("#smdSectionGeneral, #smdSectionPeople, #smdSectionMaturity, #smdSectionCounts").empty();
    peoplePickers = {};

    // Preserve FIELD_DEFS order so 2-column pairs stay aligned
    FIELD_DEFS.forEach(function (def) {
      var meta = fieldsByKey[def.key];
      if (!meta || def.hideInForm) return;

      var $section = sectionContainer(def.section || "general");
      var fullClass = def.full ? " smd-field--full" : "";
      var $field = $(
        '<div class="smd-field' +
          fullClass +
          '" data-field-key="' +
          escapeHtml(def.key) +
          '"></div>'
      );

      var $labelWrap = $('<div class="smd-field-label"></div>');
      $labelWrap.append(
        $("<label></label>")
          .attr("for", "smd_" + def.key)
          .text(def.label || meta.Title)
      );

      var hintParts = [];
      if (meta.isPerson) {
        hintParts.push(
          meta.AllowMultipleValues
            ? "Multiple people allowed"
            : "One person only"
        );
      }
      if (def.hint) hintParts.push(def.hint);
      if (hintParts.length) {
        $labelWrap.append(
          $('<span class="field-hint"></span>').text(hintParts.join(" · "))
        );
      }
      $field.append($labelWrap);

      if (meta.isPerson) {
        var $pp = $(
          '<div class="pp-host" id="smd_' + escapeHtml(def.key) + '"></div>'
        );
        $field.append($pp);
        $section.append($field);
        peoplePickers[def.key] = createPeoplePicker($pp, {
          multi: meta.AllowMultipleValues
        });
        return;
      }

      var controlType = def.type;
      if (
        meta.TypeAsString === "Choice" ||
        meta.TypeAsString === "MultiChoice"
      ) {
        controlType = "choice";
      } else if (
        meta.TypeAsString === "Number" ||
        meta.TypeAsString === "Currency" ||
        meta.TypeAsString === "Integer"
      ) {
        controlType = "number";
      } else if (
        meta.TypeAsString === "Note" ||
        controlType === "textarea"
      ) {
        controlType = "textarea";
      }

      var $ctrl;
      if (controlType === "choice") {
        $ctrl = $('<select id="smd_' + escapeHtml(def.key) + '"></select>');
        $ctrl.append('<option value="">— Select —</option>');
        (meta.choices || []).forEach(function (c) {
          $ctrl.append($("<option></option>").attr("value", c).text(c));
        });
      } else if (controlType === "textarea") {
        $ctrl = $(
          '<textarea id="smd_' +
            escapeHtml(def.key) +
            '" rows="3"></textarea>'
        );
      } else if (controlType === "number") {
        $ctrl = $(
          '<input type="number" id="smd_' +
            escapeHtml(def.key) +
            '" step="' +
            escapeHtml(def.step || "1") +
            '" />'
        );
      } else {
        $ctrl = $(
          '<input type="text" id="smd_' + escapeHtml(def.key) + '" />'
        );
      }

      if (def.placeholder && !$ctrl.is("select")) {
        $ctrl.attr("placeholder", def.placeholder);
      }

      $field.append($ctrl);
      $section.append($field);
    });
  }

  function getItemValue(item, meta) {
    if (!item || !meta) return null;
    return readFieldFromItem(item, meta.InternalName);
  }

  function prefillForm(item) {
    Object.keys(fieldsByKey).forEach(function (key) {
      var meta = fieldsByKey[key];
      var raw = getItemValue(item, meta);

      if (meta.isPerson) {
        if (peoplePickers[key]) {
          peoplePickers[key].setValue(peopleFromField(raw));
        }
        return;
      }

      var $el = $("#smd_" + key);
      if (!$el.length) return;

      if (meta.TypeAsString === "MultiChoice" && raw && raw.results) {
        // MultiChoice not expected; take first / join
        $el.val(raw.results[0] || "");
        return;
      }

      if (raw == null) {
        $el.val("");
        return;
      }

      if ($el.is('input[type="number"]')) {
        $el.val(raw === "" ? "" : raw);
      } else {
        var str = trimOrEmpty(raw);
        if ($el.is("select") && str) {
          // Add option if missing so current value displays
          if (
            $el.find('option').filter(function () {
              return $(this).val() === str;
            }).length === 0
          ) {
            $el.append($("<option></option>").attr("value", str).text(str));
          }
        }
        $el.val(str);
      }
    });
  }

  function extractSpErrorText(xhr) {
    try {
      if (!xhr || !xhr.responseText) return "";
      var parsed = JSON.parse(xhr.responseText);
      return (
        (parsed &&
          parsed.error &&
          parsed.error.message &&
          (parsed.error.message.value || parsed.error.message)) ||
        (parsed &&
          parsed["odata.error"] &&
          parsed["odata.error"].message &&
          (parsed["odata.error"].message.value ||
            parsed["odata.error"].message)) ||
        ""
      );
    } catch (_e) {
      return trimOrEmpty(xhr && xhr.responseText).slice(0, 400);
    }
  }

  function labelForPayloadKey(payloadKey) {
    if (!payloadKey) return "";
    if (
      reviewStatusInternal &&
      (payloadKey === reviewStatusInternal ||
        payloadKey === toODataSelectName(reviewStatusInternal))
    ) {
      return "Review Status";
    }
    var keys = Object.keys(fieldsByKey);
    for (var i = 0; i < keys.length; i++) {
      var meta = fieldsByKey[keys[i]];
      var label = (meta.def && meta.def.label) || meta.Title || keys[i];
      var od = toODataSelectName(meta.InternalName);
      if (
        payloadKey === meta.InternalName ||
        payloadKey === od ||
        payloadKey === meta.InternalName + "Id" ||
        payloadKey === od + "Id"
      ) {
        return payloadKey.indexOf("Id") === payloadKey.length - 2
          ? label + " (person)"
          : label;
      }
    }
    return payloadKey;
  }

  function guessFieldsFromError(errorText) {
    var text = String(errorText || "");
    var hits = [];
    var keys = Object.keys(fieldsByKey);
    keys.forEach(function (k) {
      var meta = fieldsByKey[k];
      var label = (meta.def && meta.def.label) || meta.Title || k;
      if (
        text.indexOf(meta.InternalName) !== -1 ||
        text.indexOf(meta.InternalName + "Id") !== -1 ||
        (label && text.indexOf(label) !== -1)
      ) {
        hits.push(label + " [" + meta.InternalName + "]");
      }
    });
    if (
      reviewStatusInternal &&
      text.indexOf(reviewStatusInternal) !== -1
    ) {
      hits.push("Review Status [" + reviewStatusInternal + "]");
    }
    return hits;
  }

  function buildUpdatePayload() {
    var payload = {
      __metadata: { type: listItemEntityType }
    };

    Object.keys(fieldsByKey).forEach(function (key) {
      var meta = fieldsByKey[key];
      if (meta.def && meta.def.hideInForm) return;

      // Internal names starting with "_" must use OData_ prefix in REST write/read
      var writeName = toODataSelectName(meta.InternalName);

      if (meta.isPerson) {
        var ids = peoplePickers[key] ? peoplePickers[key].getIds() : [];
        var idField = writeName + "Id";
        // Do not send null/empty person clears — SharePoint often 400s on null UserId
        if (!ids.length) return;
        if (meta.AllowMultipleValues) {
          payload[idField] = { results: ids };
        } else {
          payload[idField] = ids[0];
        }
        return;
      }

      var $el = $("#smd_" + key);
      if (!$el.length) return;
      var val = $el.val();

      if (meta.def.type === "number" || meta.TypeAsString === "Number") {
        if (val === "" || val == null) return;
        var num = Number(val);
        if (isNaN(num)) return;
        payload[writeName] = num;
        return;
      }

      if (
        (meta.TypeAsString === "Choice" || meta.def.type === "choice") &&
        (val === "" || val == null)
      ) {
        return;
      }

      payload[writeName] = val == null ? "" : String(val);
    });

    if (reviewStatusInternal) {
      payload[toODataSelectName(reviewStatusInternal)] = REVIEWED_STATUS_LABEL;
    }
    return payload;
  }

  function mergeItem(itemId, payload, digest) {
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle() +
        "')/items(" +
        itemId +
        ")",
      type: "POST",
      data: JSON.stringify(payload),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      }
    });
  }

  /**
   * When full MERGE fails with 400, probe each payload property alone to find
   * which field SharePoint rejects.
   */
  function diagnoseSavePayload(itemId, digest, fullPayload) {
    var deferred = $.Deferred();
    var bad = [];
    var keys = Object.keys(fullPayload).filter(function (k) {
      return k !== "__metadata";
    });
    var i = 0;

    function next() {
      if (i >= keys.length) {
        deferred.resolve(bad);
        return;
      }
      var key = keys[i++];
      var probe = {
        __metadata: { type: listItemEntityType }
      };
      probe[key] = fullPayload[key];

      mergeItem(itemId, probe, digest)
        .done(function () {
          next();
        })
        .fail(function (xhr) {
          bad.push({
            key: key,
            label: labelForPayloadKey(key),
            value: fullPayload[key],
            error: extractSpErrorText(xhr) || xhrErrorMessage(xhr, "rejected")
          });
          next();
        });
    }

    next();
    return deferred.promise();
  }

  function formatProbeValue(v) {
    if (v == null) return "(null)";
    if (typeof v === "object") {
      try {
        return JSON.stringify(v);
      } catch (_e) {
        return String(v);
      }
    }
    var s = String(v);
    if (s.length > 80) s = s.slice(0, 80) + "…";
    return s;
  }

  /* =========================================================================
   * SITE PICKER / FLOW
   * ========================================================================= */
  function getSiteReviewStatus(item) {
    if (!item || !reviewStatusInternal) return "";
    return trimOrEmpty(readFieldFromItem(item, reviewStatusInternal));
  }

  function isReviewCompleted(status) {
    var s = trimOrEmpty(status).toLowerCase();
    if (!s) return false;
    return (
      s === REVIEWED_STATUS_LABEL.toLowerCase() ||
      s.indexOf("completed") !== -1
    );
  }

  function reviewStatusPillClass(status) {
    if (isReviewCompleted(status)) return "smd-status-pill--done";
    if (trimOrEmpty(status)) return "smd-status-pill--open";
    return "smd-status-pill--unknown";
  }

  function reviewStatusPillLabel(status) {
    var s = trimOrEmpty(status);
    if (isReviewCompleted(s)) return REVIEWED_STATUS_LABEL;
    if (s) return s;
    return "Not reviewed";
  }

  function populateSitePicker() {
    var $list = $("#smdSiteList");
    $list.empty();

    if (matchedSites.length === 0) {
      $("#smdSitePicker").addClass("is-hidden");
      return;
    }

    if (matchedSites.length === 1) {
      $("#smdSitePicker").addClass("is-hidden");
      selectedSiteId = matchedSites[0].Id;
      return;
    }

    // Sort: pending first, then completed — easier to see what still needs work
    var ordered = matchedSites.slice().sort(function (a, b) {
      var aDone = isReviewCompleted(getSiteReviewStatus(a.raw)) ? 1 : 0;
      var bDone = isReviewCompleted(getSiteReviewStatus(b.raw)) ? 1 : 0;
      if (aDone !== bDone) return aDone - bDone;
      return a.label.localeCompare(b.label);
    });

    var doneCount = 0;
    ordered.forEach(function (s) {
      var status = getSiteReviewStatus(s.raw);
      if (isReviewCompleted(status)) doneCount += 1;

      var $btn = $(
        '<button type="button" class="smd-site-card" role="listitem"></button>'
      )
        .attr("data-site-id", String(s.Id))
        .attr(
          "aria-label",
          s.label + ", review status " + reviewStatusPillLabel(status)
        );

      if (selectedSiteId != null && Number(selectedSiteId) === Number(s.Id)) {
        $btn.addClass("is-selected");
      }

      var $main = $('<span class="smd-site-card-main"></span>');
      $main.append(
        $('<span class="smd-site-card-name"></span>').text(s.label)
      );
      $main.append(
        $('<span class="smd-site-card-meta"></span>').text("Click to edit")
      );

      var $pill = $('<span class="smd-status-pill"></span>')
        .addClass(reviewStatusPillClass(status))
        .attr("title", reviewStatusPillLabel(status))
        .text(reviewStatusPillLabel(status));

      $btn.append($main, $pill);
      $list.append($btn);
    });

    $("#smdSitePickerTitle").text(
      "Your sites (" +
        matchedSites.length +
        ") — " +
        doneCount +
        " completed"
    );
    $("#smdSitePicker").removeClass("is-hidden");
  }

  function highlightSelectedSiteCard() {
    $("#smdSiteList .smd-site-card").each(function () {
      var id = $(this).attr("data-site-id");
      $(this).toggleClass(
        "is-selected",
        selectedSiteId != null && Number(id) === Number(selectedSiteId)
      );
    });
  }

  function selectSiteById(siteId) {
    selectedSiteId = siteId != null ? Number(siteId) : null;
    highlightSelectedSiteCard();
    clearInfo();
    var site = getSelectedSite();
    if (site) {
      showInfo("Editing: " + site.label);
    }
    showFormForSelectedSite();
    try {
      var formEl = document.getElementById("smdForm");
      if (formEl && !$(formEl).hasClass("is-hidden")) {
        formEl.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    } catch (_e) {
      /* ignore */
    }
  }

  function getSelectedSite() {
    if (selectedSiteId == null) return null;
    for (var i = 0; i < matchedSites.length; i++) {
      if (Number(matchedSites[i].Id) === Number(selectedSiteId)) {
        return matchedSites[i];
      }
    }
    return null;
  }

  function showFormForSelectedSite() {
    clearError();
    clearSuccess();
    var site = getSelectedSite();
    if (!site) {
      $("#smdForm").addClass("is-hidden");
      return;
    }
    prefillForm(site.raw);
    $("#smdForm").removeClass("is-hidden");
  }

  function handleSave(e) {
    e.preventDefault();
    clearError();
    clearSuccess();

    var site = getSelectedSite();
    if (!site) {
      showError("Please select a site first.");
      return;
    }

    var $btn = $("#smdSaveBtn");
    $btn.prop("disabled", true).text("Saving…");

    $.when(getFormDigest(), fetchListEntityType())
      .then(function (digest) {
        var payload = buildUpdatePayload();
        return mergeItem(site.Id, payload, digest).then(
          function () {
            return { digest: digest, payload: payload, ok: true };
          },
          function (xhr) {
            return { digest: digest, payload: payload, ok: false, xhr: xhr };
          }
        );
      })
      .then(function (result) {
        if (result.ok) {
          Object.keys(fieldsByKey).forEach(function (key) {
            var meta = fieldsByKey[key];
            if (meta.def && meta.def.hideInForm) return;
            var writeName = toODataSelectName(meta.InternalName);
            if (meta.isPerson) {
              site.raw[meta.InternalName] = peoplePickers[key]
                ? peoplePickers[key].getValue()
                : [];
            } else if (result.payload[writeName] !== undefined) {
              site.raw[meta.InternalName] = result.payload[writeName];
              site.raw[writeName] = result.payload[writeName];
            }
          });
          if (reviewStatusInternal) {
            var rsWrite = toODataSelectName(reviewStatusInternal);
            site.raw[reviewStatusInternal] = REVIEWED_STATUS_LABEL;
            site.raw[rsWrite] = REVIEWED_STATUS_LABEL;
          }

          if (matchedSites.length > 1) {
            populateSitePicker();
            highlightSelectedSiteCard();
          }

          showSuccess(
            "Saved successfully. Review Status set to \"" +
              REVIEWED_STATUS_LABEL +
              "\"."
          );
          $btn.prop("disabled", false).text("Save changes");
          return;
        }

        var xhr = result.xhr;
        console.error("Save failed", xhr && xhr.responseText);
        var spMsg = extractSpErrorText(xhr);
        var guessed = guessFieldsFromError(spMsg);

        if (xhr && xhr.status === 403) {
          showError(
            "You do not have permission to update this item. Contact the site owner.",
            spMsg ? [spMsg] : null
          );
          $btn.prop("disabled", false).text("Save changes");
          return;
        }

        $btn.text("Diagnosing fields…");
        return diagnoseSavePayload(site.Id, result.digest, result.payload).then(
          function (bad) {
            var details = [];
            if (spMsg) details.push("SharePoint: " + spMsg);
            if (guessed.length) {
              details.push("Mentioned in error: " + guessed.join("; "));
            }
            if (bad.length) {
              details.push("Fields that fail when saved alone:");
              bad.forEach(function (b) {
                details.push(
                  "• " +
                    b.label +
                    " (" +
                    b.key +
                    ") = " +
                    formatProbeValue(b.value) +
                    " — " +
                    b.error
                );
              });
            } else {
              details.push(
                "No single field failed alone (may be a combination issue, or Review Status / permissions)."
              );
              details.push(
                "Payload keys: " +
                  Object.keys(result.payload)
                    .filter(function (k) {
                      return k !== "__metadata";
                    })
                    .map(labelForPayloadKey)
                    .join(", ")
              );
            }
            showError(
              "Save rejected by SharePoint (invalid field value).",
              details
            );
            $btn.prop("disabled", false).text("Save changes");
          }
        );
      })
      .fail(function (xhr) {
        console.error("Save failed", xhr && xhr.responseText);
        showError(
          xhrErrorMessage(xhr, "Failed to save. Please try again."),
          extractSpErrorText(xhr) ? [extractSpErrorText(xhr)] : null
        );
        $btn.prop("disabled", false).text("Save changes");
      });
  }

  function handleReset() {
    clearError();
    clearSuccess();
    var site = getSelectedSite();
    if (!site) return;
    prefillForm(site.raw);
    showInfo("Form reverted to last loaded / saved values.");
  }

  /* =========================================================================
   * INIT
   * ========================================================================= */
  function init() {
    $("#smdLoading").removeClass("is-hidden").text("Loading your sites…");
    $("#smdForm").addClass("is-hidden");
    clearError();
    clearSuccess();
    clearInfo();
    clearDiagnostics();

    getCurrentUser()
      .then(function (user) {
        currentUser = user;
        $("#smdUserLine")
          .text(
            "Signed in as " +
              user.name +
              (user.email ? " (" + user.email + ")" : "")
          )
          .removeClass("is-hidden");
        return fetchFieldsCatalog();
      })
      .then(function (catalog) {
        resolveFieldDefs(catalog);
        renderFormFields();

        var diag = fieldDiagLines();
        if (diag.length) {
          showDiagnosticPanel(
            "Column mapping warnings (" + diag.length + ")",
            diag
          );
        }

        var accessMetas = ACCESS_ROLE_KEYS.map(function (k) {
          return fieldsByKey[k];
        }).filter(Boolean);

        if (!TEST_SHOW_ALL_SITES && !accessMetas.length) {
          return $.Deferred()
            .reject({
              message:
                "Could not resolve Site Data Admin / Site Head / Digital Lead columns on the list. Check column titles.",
              detailLines: fieldDiagLines()
            })
            .promise();
        }

        return fetchListEntityType();
      })
      .then(function () {
        return fetchAll(buildItemsQueryUrl()).then(null, function (xhr) {
          console.warn(
            "[Site Master Data Form] Full items query failed — diagnosing fields…",
            xhr && xhr.responseText
          );
          $("#smdLoading").text("Diagnosing which columns break the query…");

          return diagnoseBadSelectFields().then(function (bad) {
            if (bad.length) {
              dropBadFieldsFromMap(bad);
              renderFormFields();
              showDiagnosticPanel(
                "These columns broke the SharePoint items query and were skipped",
                fieldDiagLines()
              );
              // Retry without the bad fields
              return fetchAll(buildItemsQueryUrl()).then(null, function (xhr2) {
                console.error(
                  "[Site Master Data Form] Retry still failed:",
                  xhrErrorMessage(xhr2, "items query")
                );
                var fallbackUrl =
                  SITE_URL +
                  "/_api/web/lists/getbytitle('" +
                  safeListTitle() +
                  "')/items?$select=Id,Title&$top=5000";
                return fetchAll(fallbackUrl).then(function (items) {
                  showError(
                    "Could not load full site field data. Dropdown uses Site ID/Title only. See column diagnostics above.",
                    [
                      xhrErrorMessage(xhr, "Original query"),
                      xhrErrorMessage(xhr2, "Retry after dropping bad fields")
                    ].concat(fieldDiagLines())
                  );
                  return items;
                });
              });
            }

            // No single field failed alone (combo / URL length / other) — minimal fallback
            var fallbackUrl =
              SITE_URL +
              "/_api/web/lists/getbytitle('" +
              safeListTitle() +
              "')/items?$select=Id,Title&$top=5000";
            return fetchAll(fallbackUrl).then(function (items) {
              showError(
                "Full items query failed, but no single column failed alone (possible URL length or combined $expand issue).",
                [xhrErrorMessage(xhr, "items query")].concat(fieldDiagLines())
              );
              return items;
            });
          });
        });
      })
      .then(function (items) {
        matchedSites = (items || [])
          .filter(function (item) {
            if (TEST_SHOW_ALL_SITES) return true;
            return userCanEditItem(currentUser, item);
          })
          .map(function (item) {
            return {
              Id: item.Id,
              label: siteLabel(item),
              raw: item
            };
          })
          .sort(function (a, b) {
            return a.label.localeCompare(b.label);
          });

        $("#smdLoading").addClass("is-hidden");

        if (!matchedSites.length) {
          showError(
            TEST_SHOW_ALL_SITES
              ? "No sites were found in Site Master Data."
              : "You are not listed as Site Data Admin, Site Head, or Digital Lead / SPOC on any site. If this is unexpected, ask the list owner to add you.",
            fieldDiagLines()
          );
          return;
        }

        populateSitePicker();

        if (matchedSites.length === 1) {
          showInfo("Editing: " + matchedSites[0].label);
          showFormForSelectedSite();
        } else {
          showInfo(
            "You are responsible for " +
              matchedSites.length +
              " sites. Select one from the list to edit."
          );
        }
      })
      .fail(function (xhr) {
        $("#smdLoading").addClass("is-hidden");
        console.error("Init failed", xhr && (xhr.responseText || xhr));
        var msg =
          (xhr && xhr.message) ||
          xhrErrorMessage(
            xhr,
            "Could not load Site Master Data. Ensure you are signed in to the IT/OT Community Hub and have read access to the list."
          );
        if (xhr && xhr.status === 404) {
          msg = "List '" + LIST_NAME + "' was not found.";
        }
        var details = (xhr && xhr.detailLines) || fieldDiagLines();
        showError(msg, details);
      });

    $("#smdSiteList").on("click", ".smd-site-card", function () {
      var id = $(this).attr("data-site-id");
      if (!id) return;
      selectSiteById(id);
    });

    // Save only via button click — Enter in inputs must not submit the form
    $("#smdForm").on("submit", function (e) {
      e.preventDefault();
    });
    $("#smdForm").on("keydown", function (e) {
      if (e.key !== "Enter" && e.which !== 13) return;
      var tag = ((e.target && e.target.tagName) || "").toLowerCase();
      if (tag === "textarea") return;
      e.preventDefault();
    });
    $("#smdSaveBtn").on("click", handleSave);
    $("#smdResetBtn").on("click", handleReset);
  }

  $(function () {
    if (!$("#smdFormWidget").length) return;
    init();
  });
})(window.jQuery);
