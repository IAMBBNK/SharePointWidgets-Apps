/* eslint-disable no-undef */
(function ($) {
  "use strict";

  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var PM_LIST_NAME = "Predictive Maintenance";
  var MASTER_LIST_NAME = "Site Master Data";

  /**
   * Hardcoded admins — match Email or LoginName (case-insensitive for email).
   * Update before deploy.
   */
  var ADMIN_USERS = [
    "user1@merckgroup.com",
    "i:0#.f|membership|user2@merckgroup.com"
  ];

  var PM_FIELD_DEFS = [
    { key: "SiteId", titles: ["Site ID"], type: "siteSelect" },
    { key: "SiteName", titles: ["Site Name"], type: "text" },
    { key: "Sector", titles: ["Sector"], type: "choice" },
    { key: "Production", titles: ["Production"], type: "text" },
    {
      key: "TechnicalContact",
      titles: ["Technical Contact"],
      type: "person"
    },
    { key: "SystemSize", titles: ["System Size"], type: "text" },
    {
      key: "StartOfOperation",
      titles: ["Start of Operation"],
      type: "datetime"
    },
    {
      key: "TechnicalSpecials",
      titles: ["Technical Specials"],
      type: "text"
    }
  ];

  var MASTER_FIELD_DEFS = [
    { key: "SiteId", titles: ["Site ID"] },
    { key: "SiteName", titles: ["Site Name"] },
    { key: "Sector", titles: ["Sector"] }
  ];

  var pmFieldsByKey = {};
  var masterFieldsByKey = {};
  var pmListItemEntityType = null;
  var formDigestCache = { value: null, expires: 0 };
  var currentUser = null;
  var masterSites = [];
  var pmItemSummaries = [];
  var currentItemId = null;
  var technicalContactPicker = null;
  var sectorChoices = [];

  function trimOrEmpty(v) {
    if (v == null) return "";
    return String(v).trim();
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function safeListTitle(name) {
    return String(name).replace(/'/g, "''");
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
    if (odata !== internalName && Object.prototype.hasOwnProperty.call(item, odata)) {
      return item[odata];
    }
    return item[odata] != null ? item[odata] : item[internalName];
  }

  function normalizeEmail(email) {
    return trimOrEmpty(email).toLowerCase();
  }

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
          (parsed &&
            parsed["odata.error"] &&
            parsed["odata.error"].message &&
            (parsed["odata.error"].message.value ||
              parsed["odata.error"].message)) ||
          "";
      }
    } catch (_e) {
      body = trimOrEmpty(xhr && xhr.responseText).slice(0, 300);
    }
    if (xhr && xhr.message && !xhr.status) return xhr.message;
    var parts = [fallback];
    if (status) parts.push(status);
    if (body) parts.push(String(body));
    return parts.join(" — ");
  }

  function showError(message) {
    $("#pmError").html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
    $("#pmSuccess").empty();
  }

  function clearError() {
    $("#pmError").empty();
  }

  function showSuccess(message) {
    $("#pmSuccess").html(
      '<div class="success-banner" role="status">' +
        escapeHtml(message) +
        "</div>"
    );
    $("#pmError").empty();
  }

  function clearSuccess() {
    $("#pmSuccess").empty();
  }

  function showWarn(message) {
    $("#pmWarn").html(
      '<div class="warn-banner" role="status">' + escapeHtml(message) + "</div>"
    );
  }

  function clearWarn() {
    $("#pmWarn").empty();
  }

  function isUserAdmin(user) {
    if (!user) return false;
    var email = normalizeEmail(user.email);
    var login = trimOrEmpty(user.loginName);
    for (var i = 0; i < ADMIN_USERS.length; i++) {
      var entry = trimOrEmpty(ADMIN_USERS[i]);
      if (!entry) continue;
      if (entry.indexOf("@") !== -1 && email && email === normalizeEmail(entry)) {
        return true;
      }
      if (login && login.toLowerCase() === entry.toLowerCase()) {
        return true;
      }
    }
    return false;
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

  function fetchFieldsCatalog(listName) {
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle(listName) +
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

  function resolveFieldDefs(catalog, defs, logPrefix) {
    var byTitle = {};
    (catalog || []).forEach(function (f) {
      var t = trimOrEmpty(f.Title);
      if (t) byTitle[t.toLowerCase()] = f;
    });

    var resolved = {};
    var unresolved = [];

    defs.forEach(function (def) {
      var hit = null;
      for (var i = 0; i < def.titles.length; i++) {
        var cand = byTitle[def.titles[i].toLowerCase()];
        if (cand) {
          hit = cand;
          break;
        }
      }
      if (!hit) {
        unresolved.push(def.titles[0]);
        return;
      }
      var listChoices = [];
      if (hit.Choices && hit.Choices.results) {
        listChoices = hit.Choices.results;
      } else if (Array.isArray(hit.Choices)) {
        listChoices = hit.Choices;
      }
      var typeAs = hit.TypeAsString || "";
      resolved[def.key] = {
        key: def.key,
        def: def,
        Title: hit.Title,
        InternalName: hit.InternalName,
        TypeAsString: typeAs,
        AllowMultipleValues:
          !!hit.AllowMultipleValues || typeAs === "UserMulti",
        choices: listChoices,
        isPerson: typeAs === "User" || typeAs === "UserMulti"
      };
    });

    if (unresolved.length) {
      console.warn(
        "[" + logPrefix + "] Unresolved columns:",
        unresolved.join(", ")
      );
    }
    return resolved;
  }

  function fetchPmListEntityType() {
    if (pmListItemEntityType) {
      return $.Deferred().resolve(pmListItemEntityType).promise();
    }
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle(PM_LIST_NAME) +
        "')?$select=ListItemEntityTypeFullName",
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      pmListItemEntityType = data.d.ListItemEntityTypeFullName;
      return pmListItemEntityType;
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
          Id: p.Id != null ? p.Id : p.id,
          Title: p.Title || p.title || "",
          Email: p.EMail || p.Email || p.email || ""
        };
      })
      .filter(function (p) {
        return p.Id != null || p.Email || p.Title;
      });
  }

  function ensureUser(loginOrEmail) {
    return getFormDigest().then(function (digest) {
      return $.ajax({
        url: SITE_URL + "/_api/web/ensureuser",
        type: "POST",
        data: JSON.stringify({ logonName: loginOrEmail }),
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
        if (!raw) return [];
        try {
          var parsed = JSON.parse(raw);
          return parsed || [];
        } catch (_e) {
          return [];
        }
      });
    });
  }

  function createPeoplePicker($container, options) {
    var multi = !!(options && options.multi);
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
        var $chip = $(
          '<span class="pp-chip"><span class="pp-chip-text"></span><button type="button" class="pp-chip-remove" aria-label="Remove">&times;</button></span>'
        );
        $chip
          .find(".pp-chip-text")
          .text(label + (p.Email ? " (" + p.Email + ")" : ""));
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
        var $btn = $('<button type="button" class="pp-suggestion"></button>');
        $btn.append(document.createTextNode(ent.DisplayText || ent.Key));
        $btn.on("mousedown", function (e) {
          e.preventDefault();
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
            .fail(function () {
              showError("Could not add user. Try again.");
              updateInputState();
            });
        });
        $suggestions.append($btn);
      });
      positionSuggestions();
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
      getIds: function () {
        return selected
          .map(function (p) {
            return p.Id;
          })
          .filter(function (id) {
            return id != null;
          });
      },
      clear: function () {
        selected = [];
        renderChips();
        updateInputState();
        $input.val("");
      }
    };
  }

  function buildMasterSitesUrl() {
    var siteIdMeta = masterFieldsByKey.SiteId;
    var siteNameMeta = masterFieldsByKey.SiteName;
    var sectorMeta = masterFieldsByKey.Sector;
    if (!siteIdMeta) return null;

    var select = ["Id"];
    var parts = [toODataSelectName(siteIdMeta.InternalName)];
    if (siteNameMeta) parts.push(toODataSelectName(siteNameMeta.InternalName));
    if (sectorMeta) parts.push(toODataSelectName(sectorMeta.InternalName));
    select = select.concat(parts);

    return (
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeListTitle(MASTER_LIST_NAME) +
      "')/items?$select=" +
      encodeURIComponent(select.join(",")) +
      "&$top=5000"
    );
  }

  function loadMasterSites() {
    var url = buildMasterSitesUrl();
    if (!url) {
      return $.Deferred()
        .reject({ message: "Site Master Data columns could not be resolved." })
        .promise();
    }
    return fetchAll(url).then(function (items) {
      var siteIdMeta = masterFieldsByKey.SiteId;
      var siteNameMeta = masterFieldsByKey.SiteName;
      var sectorMeta = masterFieldsByKey.Sector;
      var seen = {};
      var sites = [];

      items.forEach(function (item) {
        var siteId = trimOrEmpty(
          readFieldFromItem(item, siteIdMeta.InternalName)
        );
        if (!siteId || seen[siteId]) return;
        seen[siteId] = true;
        sites.push({
          masterItemId: item.Id,
          siteId: siteId,
          siteName: siteNameMeta
            ? trimOrEmpty(
                readFieldFromItem(item, siteNameMeta.InternalName)
              )
            : "",
          sector: sectorMeta
            ? trimOrEmpty(readFieldFromItem(item, sectorMeta.InternalName))
            : ""
        });
      });

      sites.sort(function (a, b) {
        return a.siteId.localeCompare(b.siteId);
      });
      masterSites = sites;
      return sites;
    });
  }

  function populateSiteIdDropdown(selectedValue) {
    var $sel = $("#pmSiteId");
    var keep = trimOrEmpty(selectedValue);
    $sel.find("option:not(:first)").remove();

    masterSites.forEach(function (s) {
      $sel.append(
        $("<option></option>").attr("value", s.siteId).text(s.siteId)
      );
    });

    if (keep) {
      var found = false;
      $sel.find("option").each(function () {
        if ($(this).val() === keep) found = true;
      });
      if (!found) {
        $sel.append($("<option></option>").attr("value", keep).text(keep));
      }
    }

    $sel.val(keep || "");
  }

  function applySitePrefill(siteId) {
    var id = trimOrEmpty(siteId);
    if (!id) {
      $("#pmSiteName").val("");
      $("#pmSector").val("");
      clearWarn();
      return;
    }

    var match = null;
    for (var i = 0; i < masterSites.length; i++) {
      if (masterSites[i].siteId === id) {
        match = masterSites[i];
        break;
      }
    }

    if (!match) {
      $("#pmSiteName").val("");
      $("#pmSector").val("");
      showWarn(
        "Site ID \"" +
          id +
          "\" was not found in Site Master Data. Site Name and Sector cannot be prefilled."
      );
      return;
    }

    $("#pmSiteName").val(match.siteName);
    $("#pmSector").val(match.sector);
    checkSectorAlignment(match.sector);
  }

  function checkSectorAlignment(sector) {
    var s = trimOrEmpty(sector);
    if (!s) {
      clearWarn();
      return;
    }
    if (!sectorChoices.length) {
      clearWarn();
      return;
    }
    var ok = sectorChoices.some(function (c) {
      return trimOrEmpty(c).toLowerCase() === s.toLowerCase();
    });
    if (!ok) {
      showWarn(
        "Sector \"" +
          s +
          "\" from Site Master Data is not in the Predictive Maintenance Sector choices. Save may fail — align list choices with master data."
      );
    } else {
      clearWarn();
    }
  }

  function buildPmSummariesUrl() {
    var siteIdMeta = pmFieldsByKey.SiteId;
    var productionMeta = pmFieldsByKey.Production;
    if (!siteIdMeta) return null;

    var select = ["Id", "Title"];
    select.push(toODataSelectName(siteIdMeta.InternalName));
    if (productionMeta) {
      select.push(toODataSelectName(productionMeta.InternalName));
    }

    return (
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeListTitle(PM_LIST_NAME) +
      "')/items?$select=" +
      encodeURIComponent(select.join(",")) +
      "&$top=5000"
    );
  }

  function pmItemLabel(item) {
    var siteIdMeta = pmFieldsByKey.SiteId;
    var productionMeta = pmFieldsByKey.Production;
    var siteId = siteIdMeta
      ? trimOrEmpty(readFieldFromItem(item, siteIdMeta.InternalName))
      : "";
    var production = productionMeta
      ? trimOrEmpty(readFieldFromItem(item, productionMeta.InternalName))
      : "";
    if (siteId && production) return siteId + " — " + production;
    if (siteId) return siteId;
    if (production) return production;
    return trimOrEmpty(item.Title) || "Item " + item.Id;
  }

  function loadPmSummaries() {
    var url = buildPmSummariesUrl();
    if (!url) {
      return $.Deferred()
        .reject({ message: "Predictive Maintenance columns could not be resolved." })
        .promise();
    }
    return fetchAll(url).then(function (items) {
      pmItemSummaries = items
        .map(function (item) {
          return {
            id: item.Id,
            label: pmItemLabel(item)
          };
        })
        .sort(function (a, b) {
          return a.label.localeCompare(b.label);
        });
      return pmItemSummaries;
    });
  }

  function populateEntrySelector(selectedId) {
    var $sel = $("#pmEntrySelect");
    $sel.find("option:not(:first)").remove();
    pmItemSummaries.forEach(function (s) {
      $sel.append(
        $("<option></option>").attr("value", String(s.id)).text(s.label)
      );
    });
    if (selectedId) {
      $sel.val(String(selectedId));
    } else {
      $sel.val("");
    }
  }

  function buildPmItemLoadUrl(itemId) {
    var select = ["Id", "Title"];
    var expand = [];

    Object.keys(pmFieldsByKey).forEach(function (key) {
      var meta = pmFieldsByKey[key];
      if (!meta || meta.isPerson) return;
      select.push(toODataSelectName(meta.InternalName));
    });

    var contactMeta = pmFieldsByKey.TechnicalContact;
    if (contactMeta && contactMeta.isPerson) {
      var od = toODataSelectName(contactMeta.InternalName);
      select.push(od + "/Id", od + "/Title", od + "/EMail");
      expand.push(od);
    }

    var url =
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeListTitle(PM_LIST_NAME) +
      "')/items(" +
      itemId +
      ")?$select=" +
      encodeURIComponent(select.join(","));

    if (expand.length) {
      url += "&$expand=" + encodeURIComponent(expand.join(","));
    }
    return url;
  }

  function spDateToLocalInput(value) {
    if (!value) return "";
    var d = new Date(value);
    if (isNaN(d.getTime())) return "";
    var pad = function (n) {
      return String(n).padStart(2, "0");
    };
    return (
      d.getFullYear() +
      "-" +
      pad(d.getMonth() + 1) +
      "-" +
      pad(d.getDate()) +
      "T" +
      pad(d.getHours()) +
      ":" +
      pad(d.getMinutes())
    );
  }

  function localInputToSpDate(value) {
    var v = trimOrEmpty(value);
    if (!v) return null;
    var d = new Date(v);
    if (isNaN(d.getTime())) return null;
    return d.toISOString();
  }

  function clearManualFields() {
    $("#pmProduction").val("");
    $("#pmSystemSize").val("");
    $("#pmStartOfOperation").val("");
    $("#pmTechnicalSpecials").val("");
    if (technicalContactPicker) technicalContactPicker.clear();
  }

  function resetCreateForm() {
    currentItemId = null;
    clearError();
    clearSuccess();
    $("#pmSiteId").val("");
    $("#pmSiteName").val("");
    $("#pmSector").val("");
    clearWarn();
    clearManualFields();
  }

  function fillFormFromItem(item) {
    var siteIdMeta = pmFieldsByKey.SiteId;
    var siteId = siteIdMeta
      ? trimOrEmpty(readFieldFromItem(item, siteIdMeta.InternalName))
      : "";

    populateSiteIdDropdown(siteId);
    applySitePrefill(siteId);

    var productionMeta = pmFieldsByKey.Production;
    $("#pmProduction").val(
      productionMeta
        ? trimOrEmpty(readFieldFromItem(item, productionMeta.InternalName))
        : ""
    );

    var systemSizeMeta = pmFieldsByKey.SystemSize;
    $("#pmSystemSize").val(
      systemSizeMeta
        ? trimOrEmpty(readFieldFromItem(item, systemSizeMeta.InternalName))
        : ""
    );

    var startMeta = pmFieldsByKey.StartOfOperation;
    $("#pmStartOfOperation").val(
      startMeta
        ? spDateToLocalInput(
            readFieldFromItem(item, startMeta.InternalName)
          )
        : ""
    );

    var specialsMeta = pmFieldsByKey.TechnicalSpecials;
    $("#pmTechnicalSpecials").val(
      specialsMeta
        ? trimOrEmpty(readFieldFromItem(item, specialsMeta.InternalName))
        : ""
    );

    var contactMeta = pmFieldsByKey.TechnicalContact;
    if (contactMeta && technicalContactPicker) {
      var people = peopleFromField(
        readFieldFromItem(item, contactMeta.InternalName)
      );
      technicalContactPicker.setValue(people);
    }
  }

  function loadPmItem(itemId) {
    return $.ajax({
      url: buildPmItemLoadUrl(itemId),
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      return data.d;
    });
  }

  function buildSavePayload() {
    var payload = {
      __metadata: { type: pmListItemEntityType }
    };

    var siteId = trimOrEmpty($("#pmSiteId").val());
    var production = trimOrEmpty($("#pmProduction").val());
    var siteName = trimOrEmpty($("#pmSiteName").val());
    var sector = trimOrEmpty($("#pmSector").val());

    payload.Title =
      siteId && production
        ? siteId + " — " + production
        : siteId || production || "Predictive Maintenance entry";

    var siteIdMeta = pmFieldsByKey.SiteId;
    if (siteIdMeta) {
      payload[toODataSelectName(siteIdMeta.InternalName)] = siteId;
    }

    var siteNameMeta = pmFieldsByKey.SiteName;
    if (siteNameMeta) {
      payload[toODataSelectName(siteNameMeta.InternalName)] = siteName;
    }

    var sectorMeta = pmFieldsByKey.Sector;
    if (sectorMeta && sector) {
      payload[toODataSelectName(sectorMeta.InternalName)] = sector;
    }

    var productionMeta = pmFieldsByKey.Production;
    if (productionMeta) {
      payload[toODataSelectName(productionMeta.InternalName)] = production;
    }

    var systemSizeMeta = pmFieldsByKey.SystemSize;
    if (systemSizeMeta) {
      payload[toODataSelectName(systemSizeMeta.InternalName)] = trimOrEmpty(
        $("#pmSystemSize").val()
      );
    }

    var startMeta = pmFieldsByKey.StartOfOperation;
    if (startMeta) {
      var iso = localInputToSpDate($("#pmStartOfOperation").val());
      if (iso) {
        payload[toODataSelectName(startMeta.InternalName)] = iso;
      }
    }

    var specialsMeta = pmFieldsByKey.TechnicalSpecials;
    if (specialsMeta) {
      payload[toODataSelectName(specialsMeta.InternalName)] = trimOrEmpty(
        $("#pmTechnicalSpecials").val()
      );
    }

    var contactMeta = pmFieldsByKey.TechnicalContact;
    if (contactMeta && technicalContactPicker) {
      var ids = technicalContactPicker.getIds();
      if (ids.length) {
        var idField = toODataSelectName(contactMeta.InternalName) + "Id";
        payload[idField] = ids[0];
      }
    }

    return payload;
  }

  function createItem(payload, digest) {
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle(PM_LIST_NAME) +
        "')/items",
      type: "POST",
      data: JSON.stringify(payload),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      }
    });
  }

  function mergeItem(itemId, payload, digest) {
    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeListTitle(PM_LIST_NAME) +
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

  function validateForm() {
    clearError();
    if (!trimOrEmpty($("#pmSiteId").val())) {
      showError("Please select a Site ID.");
      return false;
    }
    if (!trimOrEmpty($("#pmProduction").val())) {
      showError("Please enter Production.");
      return false;
    }
    return true;
  }

  function handleSave(e) {
    e.preventDefault();
    if (!validateForm()) return;

    var $btn = $("#pmSaveBtn");
    $btn.prop("disabled", true).text("Saving…");

    getFormDigest()
      .then(function (digest) {
        var payload = buildSavePayload();
        if (currentItemId) {
          return mergeItem(currentItemId, payload, digest);
        }
        return createItem(payload, digest);
      })
      .then(function (data) {
        var wasCreate = !currentItemId;
        var newId =
          wasCreate && data && data.d && data.d.Id ? data.d.Id : currentItemId;

        return loadPmSummaries().then(function () {
          if (wasCreate && newId) {
            currentItemId = newId;
            populateEntrySelector(newId);
          } else if (currentItemId) {
            populateEntrySelector(currentItemId);
          }
          showSuccess(
            wasCreate
              ? "Entry created successfully."
              : "Entry updated successfully."
          );
        });
      })
      .fail(function (xhr) {
        var message = "Failed to save entry. Please try again.";
        if (xhr && xhr.status === 403) {
          message =
            "You do not have permission to save. Contact the site owner.";
        } else if (xhr && xhr.status === 400) {
          message =
            "Save failed — check that list columns match predictiveMaintenanceForm_LIST_SETUP.md.";
          console.error("PM save 400:", xhr.responseText);
        } else if (xhr && xhr.message) {
          message = xhr.message;
        } else if (xhr) {
          message = xhrErrorMessage(xhr, message);
        }
        showError(message);
      })
      .always(function () {
        $btn.prop("disabled", false).text("Save entry");
      });
  }

  function handleEntrySelectChange() {
    var val = trimOrEmpty($("#pmEntrySelect").val());
    clearError();
    clearSuccess();

    if (!val) {
      resetCreateForm();
      return;
    }

    currentItemId = Number(val);
    $("#pmLoading").removeClass("is-hidden");
    $("#pmMain").addClass("is-hidden");

    loadPmItem(currentItemId)
      .then(function (item) {
        fillFormFromItem(item);
        $("#pmLoading").addClass("is-hidden");
        $("#pmMain").removeClass("is-hidden");
      })
      .fail(function (xhr) {
        $("#pmLoading").addClass("is-hidden");
        $("#pmMain").removeClass("is-hidden");
        showError(xhrErrorMessage(xhr, "Could not load the selected entry."));
      });
  }

  function bindEvents() {
    $("#pmSiteId").on("change", function () {
      applySitePrefill($(this).val());
    });

    $("#pmEntrySelect").on("change", handleEntrySelectChange);

    $("#pmForm").on("submit", handleSave);
  }

  function initPeoplePicker() {
    technicalContactPicker = createPeoplePicker($("#pmTechnicalContact"), {
      multi: false
    });
  }

  function showUnauthorized() {
    $("#pmLoading").addClass("is-hidden");
    $("#pmMain").addClass("is-hidden");
    $("#pmUnauthorized").removeClass("is-hidden");
  }

  function showMain() {
    $("#pmLoading").addClass("is-hidden");
    $("#pmUnauthorized").addClass("is-hidden");
    $("#pmMain").removeClass("is-hidden");
  }

  function init() {
    bindEvents();
    initPeoplePicker();

    getCurrentUser()
      .then(function (user) {
        currentUser = user;
        $("#pmUserLine").text(
          "Signed in as " + user.name + (user.email ? " (" + user.email + ")" : "")
        );

        if (!isUserAdmin(user)) {
          showUnauthorized();
          return $.Deferred().reject().promise();
        }

        return $.when(
          fetchFieldsCatalog(PM_LIST_NAME),
          fetchFieldsCatalog(MASTER_LIST_NAME),
          fetchPmListEntityType()
        );
      })
      .then(function (pmCatalog, masterCatalog) {
        pmFieldsByKey = resolveFieldDefs(
          pmCatalog,
          PM_FIELD_DEFS,
          "PM Form"
        );
        masterFieldsByKey = resolveFieldDefs(
          masterCatalog,
          MASTER_FIELD_DEFS,
          "PM Form Master"
        );

        var sectorMeta = pmFieldsByKey.Sector;
        sectorChoices = sectorMeta ? sectorMeta.choices || [] : [];

        if (!pmFieldsByKey.SiteId) {
          throw { message: "Site ID column not found on Predictive Maintenance list." };
        }
        if (!masterFieldsByKey.SiteId) {
          throw { message: "Site ID column not found on Site Master Data list." };
        }

        return $.when(loadMasterSites(), loadPmSummaries());
      })
      .then(function () {
        populateSiteIdDropdown("");
        populateEntrySelector(null);
        showMain();
      })
      .fail(function (xhr) {
        if (!currentUser || isUserAdmin(currentUser)) {
          $("#pmLoading").addClass("is-hidden");
          $("#pmUnauthorized").addClass("is-hidden");
          var msg =
            xhr && xhr.message
              ? xhr.message
              : xhrErrorMessage(xhr, "Failed to load the form.");
          showError(msg);
          $("#pmMain").removeClass("is-hidden");
        }
      });
  }

  $(init);
})(jQuery);
