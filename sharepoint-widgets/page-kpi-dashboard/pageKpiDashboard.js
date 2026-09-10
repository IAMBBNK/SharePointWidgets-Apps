/* eslint-disable no-undef */
(function () {
  "use strict";

  /* =========================================================================
   * CONFIG
   * Adjust folder path / list name if your SharePoint setup differs.
   * =========================================================================
   */
  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  /* Primary folder path. The loader also tries common fallbacks automatically. */
  var PAGE_KPI_FOLDER_SERVER_RELATIVE_URL =
    "/sites/ITOTCommunityHub/Shared Documents/Page KPIs";
  var PAGE_KPI_FOLDER_CANDIDATES = [
    "/sites/ITOTCommunityHub/Shared Documents/Page KPIs",
    "/sites/ITOTCommunityHub/Documents/Page KPIs",
    "/sites/ITOTCommunityHub/SiteAssets/Page KPIs",
    "/sites/ITOTCommunityHub/Page KPIs"
  ];
  var NEWSLETTER_LIST_NAME = "Newsletter KPIs";
  var PAGE_VIEW_TRACKING_LIST_NAME = "Page View Tracking";
  var VIVA_SOURCE_VALUE = "viva";
  var TOP_PAGES_COUNT = 10;
  var TOP_STORIES_OVERALL_COUNT = 10;

  var MENU_HUB_SECTIONS = {
    Program: true,
    Portfolio: true,
    Architecture: true,
    General: true,
    "Service Management": true,
    "News & Connect": true,
    "Upskilling & Resources": true,
    "Digital Plant": true
  };
  var ARTICLE_HUB_SECTION = "Articles";
  var IGNORE_HUB_SECTIONS = {
    Templates: true,
    Archive: true
  };

  var NEWSLETTER_FIELDS = {
    Title: "Title",
    SentDateTime: "SentDateTime",
    DeliveredTo: "DeliveredTo",
    Notes: "Notes",
    Opens24hSeries: "Opens24hSeries",
    Clicks24hSeries: "Clicks24hSeries",
    Opens7dSeries: "Opens7dSeries",
    Clicks7dSeries: "Clicks7dSeries"
  };

  var SERIES_LEN_24H = 24;
  var SERIES_LEN_7D = 7;
  var CHART_COLOR_OPENS = "#0f69af";
  var CHART_COLOR_CLICKS = "#2e7d32";

  /* =========================================================================
   * STATE
   * =========================================================================
   */
  var pageRows = [];
  var pageRowsByUrl = {};
  var newsletterEditions = [];
  var pageKpiFileModified = null;
  /** Normalized page URL/path → count of Source=viva tracking rows */
  var vivaViewsByPageKey = {};

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

  function showError(message) {
    $("#errorContainer").html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
  }

  function clearError() {
    $("#errorContainer").empty();
  }

  function formatNumber(n) {
    if (n == null || !isFinite(n)) return "—";
    return Math.round(n).toLocaleString("en-US");
  }

  function formatDecimal(n, digits) {
    if (n == null || !isFinite(n)) return "—";
    var d = digits == null ? 1 : digits;
    return Number(n).toLocaleString("en-US", {
      minimumFractionDigits: 0,
      maximumFractionDigits: d
    });
  }

  function toInt(v) {
    if (v == null || v === "") return 0;
    var n = Number(String(v).replace(/[^\d.-]/g, ""));
    return isFinite(n) ? Math.round(n) : 0;
  }

  function averageOf(values) {
    if (!values || !values.length) return null;
    var sum = 0;
    values.forEach(function (v) {
      sum += v;
    });
    return sum / values.length;
  }

  function sumSeries(arr) {
    var total = 0;
    (arr || []).forEach(function (n) {
      total += n || 0;
    });
    return total;
  }

  function seriesHasValues(arr) {
    if (!arr || !arr.length) return false;
    for (var i = 0; i < arr.length; i++) {
      if ((arr[i] || 0) !== 0) return true;
    }
    return false;
  }

  /**
   * Parse comma/semicolon/whitespace/newline separated integers.
   * Pads or truncates to expectedLen. Empty input → empty array (no series).
   */
  function parseSeries(raw, expectedLen) {
    if (raw == null) return [];
    var text = String(raw).trim();
    if (!text) return [];

    var parts = text.split(/[,;\s\n\r]+/).filter(function (p) {
      return p !== "";
    });

    if (!parts.length) return [];

    var values = [];
    var i;
    for (i = 0; i < expectedLen; i++) {
      values.push(i < parts.length ? toInt(parts[i]) : 0);
    }
    return values;
  }

  function zeroSeries(len) {
    var arr = [];
    for (var i = 0; i < len; i++) arr.push(0);
    return arr;
  }

  function sumSeriesPointwise(editions, key, len) {
    var result = zeroSeries(len);
    var any = false;
    (editions || []).forEach(function (edition) {
      var series = edition[key];
      if (!series || !series.length) return;
      any = true;
      for (var i = 0; i < len; i++) {
        result[i] += series[i] || 0;
      }
    });
    return any ? result : [];
  }

  function averageSeriesPointwise(editions, key, len) {
    var summed = sumSeriesPointwise(editions, key, len);
    if (!summed.length) return [];

    var withSeries = 0;
    (editions || []).forEach(function (edition) {
      var series = edition[key];
      if (series && series.length) withSeries++;
    });
    if (!withSeries) return [];

    return summed.map(function (v) {
      return v / withSeries;
    });
  }

  function formatDateTime(value) {
    if (!value) return "";
    var d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return "";
    var pad = function (n) {
      return n < 10 ? "0" + n : "" + n;
    };
    return (
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

  function setLastUpdated(fileModified) {
    var label = fileModified
      ? "Page KPIs file: " + formatDateTime(fileModified)
      : "Last updated: " + formatDateTime(new Date());
    $("#lastUpdated").text(label);
  }

  /* =========================================================================
   * PAGE KPI TXT FETCH + PARSE
   * =========================================================================
   */
  function isPageKpiTxtFile(name) {
    if (!name) return false;
    // e.g. KPIs-2026-07-22T08_30_05.4838198Z.txt
    return /\.txt$/i.test(name) || /^KPIs-.*\.txt$/i.test(name);
  }

  function encodeServerRelativePath(path) {
    return String(path || "")
      .split("/")
      .map(function (part) {
        return encodeURIComponent(part);
      })
      .join("/");
  }

  function buildFolderFilesUrl(folderServerRelativeUrl) {
    var base = SITE_URL.replace(/\/$/, "");
    var encoded = encodeServerRelativePath(folderServerRelativeUrl);
    return (
      base +
      "/_api/web/GetFolderByServerRelativePath(decodedurl=@p)/Files" +
      "?@p='" +
      encoded.replace(/'/g, "%27") +
      "'" +
      "&$select=Name,ServerRelativeUrl,TimeLastModified"
    );
  }

  function fetchFolderFiles(folderServerRelativeUrl) {
    return $.ajax({
      url: buildFolderFilesUrl(folderServerRelativeUrl),
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      return (data.d && data.d.results) || [];
    });
  }

  function fetchFilesFromListRoot(listTitle) {
    var safeTitle = String(listTitle).replace(/'/g, "''");
    var url =
      SITE_URL.replace(/\/$/, "") +
      "/_api/web/lists/getbytitle('" +
      safeTitle +
      "')/RootFolder/Files?$select=Name,ServerRelativeUrl,TimeLastModified";

    return $.ajax({
      url: url,
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(function (data) {
      return (data.d && data.d.results) || [];
    });
  }

  function pickLatestTxt(files) {
    var txtFiles = (files || []).filter(function (f) {
      return isPageKpiTxtFile(f.Name);
    });

    txtFiles.sort(function (a, b) {
      var ta = new Date(a.TimeLastModified || 0).getTime();
      var tb = new Date(b.TimeLastModified || 0).getTime();
      return tb - ta;
    });

    return txtFiles;
  }

  function tryLoadFromFolder(folderPath) {
    return fetchFolderFiles(folderPath).then(function (files) {
      var txtFiles = pickLatestTxt(files);
      if (!txtFiles.length) {
        var names = (files || [])
          .slice(0, 5)
          .map(function (f) {
            return f.Name;
          })
          .join(", ");
        return $.Deferred()
          .reject({
            folder: folderPath,
            fileCount: (files || []).length,
            sampleNames: names,
            message:
              "No .txt files in " +
              folderPath +
              (files && files.length
                ? " (found " + files.length + " other file(s): " + names + ")"
                : " (folder empty)")
          })
          .promise();
      }

      return {
        folder: folderPath,
        latest: txtFiles[0],
        txtCount: txtFiles.length
      };
    });
  }

  function resolvePageKpiSource() {
    var candidates = [];
    if (PAGE_KPI_FOLDER_SERVER_RELATIVE_URL) {
      candidates.push(PAGE_KPI_FOLDER_SERVER_RELATIVE_URL);
    }
    (PAGE_KPI_FOLDER_CANDIDATES || []).forEach(function (p) {
      if (candidates.indexOf(p) === -1) candidates.push(p);
    });

    var attempts = [];
    var i = 0;

    function nextFolder() {
      if (i >= candidates.length) {
        // Last resort: document library / list titled "Page KPIs"
        return fetchFilesFromListRoot("Page KPIs").then(
          function (files) {
            var txtFiles = pickLatestTxt(files);
            if (!txtFiles.length) {
              return $.Deferred()
                .reject(
                  new Error(
                    "No KPI .txt files found. Tried folders: " +
                      candidates.join(" | ") +
                      ". Also tried list 'Page KPIs'. " +
                      (attempts.length ? attempts.join(" ") : "")
                  )
                )
                .promise();
            }
            return {
              folder: "list:Page KPIs",
              latest: txtFiles[0],
              txtCount: txtFiles.length
            };
          },
          function () {
            return $.Deferred()
              .reject(
                new Error(
                  "No KPI .txt files found. Tried folders: " +
                    candidates.join(" | ") +
                    ". Update PAGE_KPI_FOLDER_SERVER_RELATIVE_URL in pageKpiDashboard.js to the exact folder path from SharePoint (Details pane → Path). " +
                    (attempts.length ? attempts.join(" ") : "")
                )
              )
              .promise();
          }
        );
      }

      var path = candidates[i++];
      return tryLoadFromFolder(path).then(
        function (hit) {
          return hit;
        },
        function (err) {
          if (err && err.message) attempts.push(err.message);
          else if (err && err.status) {
            attempts.push(path + " → HTTP " + err.status + ".");
          }
          return nextFolder();
        }
      );
    }

    return nextFolder();
  }

  function fetchFileText(serverRelativeUrl) {
    var base = SITE_URL.replace(/\/$/, "");
    var encoded = encodeServerRelativePath(serverRelativeUrl);
    var url =
      base +
      "/_api/web/GetFileByServerRelativePath(decodedurl=@p)/$value" +
      "?@p='" +
      encoded.replace(/'/g, "%27") +
      "'";

    return $.ajax({
      url: url,
      type: "GET",
      dataType: "text"
    });
  }

  function parsePageKpiEntry(raw) {
    var cleaned = String(raw == null ? "" : raw)
      .replace(/^"\."/, "")
      .replace(/^"+/, "")
      .replace(/"+$/, "")
      .trim();

    var hubSection = "";
    var hubMatch = cleaned.match(/HubSection:(.*?),Link:/i);
    if (hubMatch) {
      var hubRaw = String(hubMatch[1] || "").trim();
      if (hubRaw) {
        try {
          var hubObj = JSON.parse(hubRaw);
          hubSection = hubObj && hubObj.Value != null ? String(hubObj.Value).trim() : "";
        } catch (e) {
          var valueMatch = hubRaw.match(/"Value"\s*:\s*"((?:\\.|[^"\\])*)"/);
          hubSection = valueMatch ? valueMatch[1].replace(/\\"/g, '"') : "";
        }
      }
    }

    var linkMatch = cleaned.match(/Link:(.*?),Title:/i);
    var titleMatch = cleaned.match(/Title:(.*?)(?:,Views:|,Viewers:|$)/i);
    var viewsMatch = cleaned.match(/Views:([^,]*)/i);
    var viewersMatch = cleaned.match(/Viewers:([^,]*)/i);

    var title = titleMatch ? String(titleMatch[1]).trim() : "";
    if (!title) return null;

    var link = linkMatch ? String(linkMatch[1]).trim() : "";
    if (link && !/^https?:\/\//i.test(link) && /^\/sites\//i.test(link)) {
      link = SITE_URL.replace(/\/sites\/.*$/, "") + link;
    }

    return {
      title: title,
      link: link,
      hubSection: hubSection,
      pageType: classifyHubSection(hubSection),
      views: toInt(viewsMatch ? viewsMatch[1] : 0),
      viewers: toInt(viewersMatch ? viewersMatch[1] : 0)
    };
  }

  function classifyHubSection(hubSection) {
    if (!hubSection) return "other";
    if (IGNORE_HUB_SECTIONS[hubSection]) return "ignore";
    if (hubSection === ARTICLE_HUB_SECTION) return "article";
    if (MENU_HUB_SECTIONS[hubSection]) return "menu";
    return "other";
  }

  function normalizeUrlKey(url) {
    if (!url) return "";
    var u = String(url).trim();
    if (!u) return "";

    try {
      u = decodeURIComponent(u);
    } catch (e) {
      /* keep original */
    }

    u = u.split("#")[0].split("?")[0];
    u = u.replace(/\/+$/, "");
    return u.toLowerCase();
  }

  function normalizePathKey(path) {
    if (!path) return "";
    var p = String(path).trim().split("#")[0].split("?")[0];
    p = p.replace(/\\/g, "/");
    if (p.indexOf("http://") === 0 || p.indexOf("https://") === 0) {
      return normalizeUrlKey(p);
    }
    if (p.charAt(0) !== "/") p = "/" + p;
    p = p.replace(/\/+$/, "");
    return p.toLowerCase();
  }

  /**
   * Canonical key shared across full URLs, server-relative paths, and sharing links.
   * Example: ".../SitePages/My-News.aspx?source=viva" → "/sitepages/my-news.aspx"
   */
  function extractSitePagesKey(urlOrPath) {
    if (!urlOrPath) return "";
    var s = String(urlOrPath).trim();
    if (!s) return "";

    try {
      s = decodeURIComponent(s);
    } catch (e) {
      /* keep original */
    }

    s = s
      .split("#")[0]
      .split("?")[0]
      .replace(/\\/g, "/")
      .toLowerCase();

    var idx = s.indexOf("/sitepages/");
    if (idx === -1) return "";

    var key = s.substring(idx).replace(/\/+$/, "");
    var aspx = key.indexOf(".aspx");
    if (aspx !== -1) {
      key = key.substring(0, aspx + 5);
    }
    return key;
  }

  function collectPageAliases(pageUrl, pagePath) {
    var aliases = {};
    function add(k) {
      if (k) aliases[k] = true;
    }

    add(extractSitePagesKey(pageUrl));
    add(extractSitePagesKey(pagePath));
    add(normalizeUrlKey(pageUrl));
    add(normalizePathKey(pagePath));

    if (pageUrl) {
      try {
        var pathFromUrl = String(pageUrl).replace(/^https?:\/\/[^/]+/i, "");
        add(normalizePathKey(pathFromUrl));
        add(extractSitePagesKey(pathFromUrl));
      } catch (e) {
        /* ignore */
      }
    }

    return Object.keys(aliases);
  }

  function pageLookupKeys(row) {
    return collectPageAliases(row && row.link, null);
  }

  function getVivaViewsForPage(row) {
    var keys = pageLookupKeys(row);
    var best = null;
    for (var i = 0; i < keys.length; i++) {
      var count = vivaViewsByPageKey[keys[i]];
      if (typeof count === "number" && (best == null || count > best)) {
        best = count;
      }
    }
    return best;
  }

  function formatVivaViewsCell(row) {
    var count = getVivaViewsForPage(row);
    if (count == null) return "N/A";
    return formatNumber(count);
  }

  function fetchAllListItems(url) {
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

  /**
   * Count each tracking row once under a canonical /sitepages/... key,
   * then publish that same total on every URL/path alias for lookup.
   */
  function buildVivaViewsLookup(items) {
    var countsByCanonical = {};
    var aliasesByCanonical = {};
    var totalVivaRows = 0;

    (items || []).forEach(function (item) {
      var source = String(item.Source || "")
        .trim()
        .toLowerCase();
      if (source !== VIVA_SOURCE_VALUE) return;

      var aliases = collectPageAliases(item.PageUrl, item.PagePath);
      if (!aliases.length) return;

      totalVivaRows += 1;

      var canonical =
        extractSitePagesKey(item.PageUrl) ||
        extractSitePagesKey(item.PagePath) ||
        normalizePathKey(item.PagePath) ||
        normalizeUrlKey(item.PageUrl) ||
        aliases[0];

      countsByCanonical[canonical] =
        (countsByCanonical[canonical] || 0) + 1;

      if (!aliasesByCanonical[canonical]) {
        aliasesByCanonical[canonical] = {};
      }
      aliases.forEach(function (a) {
        aliasesByCanonical[canonical][a] = true;
      });
    });

    // Merge canonical buckets that share any alias (same page, different URL shapes)
    var parent = {};
    function find(x) {
      parent[x] = parent[x] || x;
      if (parent[x] !== x) parent[x] = find(parent[x]);
      return parent[x];
    }
    function union(a, b) {
      var ra = find(a);
      var rb = find(b);
      if (ra !== rb) parent[rb] = ra;
    }

    Object.keys(aliasesByCanonical).forEach(function (canonical) {
      find(canonical);
      Object.keys(aliasesByCanonical[canonical]).forEach(function (alias) {
        Object.keys(aliasesByCanonical).forEach(function (other) {
          if (other === canonical) return;
          if (aliasesByCanonical[other][alias]) union(canonical, other);
        });
      });
    });

    var mergedCounts = {};
    var mergedAliases = {};
    Object.keys(countsByCanonical).forEach(function (canonical) {
      var root = find(canonical);
      mergedCounts[root] =
        (mergedCounts[root] || 0) + countsByCanonical[canonical];
      if (!mergedAliases[root]) mergedAliases[root] = {};
      Object.keys(aliasesByCanonical[canonical] || {}).forEach(function (a) {
        mergedAliases[root][a] = true;
      });
      mergedAliases[root][canonical] = true;
    });

    var map = {};
    Object.keys(mergedCounts).forEach(function (root) {
      var count = mergedCounts[root];
      Object.keys(mergedAliases[root] || {}).forEach(function (alias) {
        map[alias] = count;
      });
      map[root] = count;
    });

    map.__totalVivaRows = totalVivaRows;
    return map;
  }

  function loadVivaEngageViews() {
    var safeTitle = PAGE_VIEW_TRACKING_LIST_NAME.replace(/'/g, "''");
    // Load broadly, then filter client-side (avoids missing rows when Source casing/spacing differs)
    var url =
      SITE_URL.replace(/\/$/, "") +
      "/_api/web/lists/getbytitle('" +
      safeTitle +
      "')/items?$select=" +
      encodeURIComponent("Id,PageUrl,PagePath,Source") +
      "&$orderby=Id desc" +
      "&$top=5000";

    return fetchAllListItems(url).then(
      function (items) {
        vivaViewsByPageKey = buildVivaViewsLookup(items);
        return vivaViewsByPageKey;
      },
      function (xhr) {
        vivaViewsByPageKey = {};
        var msg =
          "Page View Tracking list '" +
          PAGE_VIEW_TRACKING_LIST_NAME +
          "' was not found (Viva Engage views will show as N/A).";
        if (xhr && xhr.status && xhr.status !== 404) {
          msg =
            "Could not load list '" +
            PAGE_VIEW_TRACKING_LIST_NAME +
            "' (HTTP " +
            xhr.status +
            "). Viva Engage views will show as N/A.";
        }
        return $.Deferred().reject(new Error(msg)).promise();
      }
    );
  }

  function parsePageKpiText(text) {
    var parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error(
        "Page KPIs file is not valid JSON. Check the Power Automate export format."
      );
    }

    if (!$.isArray(parsed)) {
      throw new Error("Page KPIs file must be a JSON array of strings.");
    }

    var byKey = {};
    parsed.forEach(function (entry) {
      var row = parsePageKpiEntry(entry);
      if (!row) return;

      var key = normalizeUrlKey(row.link) || ("title:" + row.title.toLowerCase());
      var existing = byKey[key];
      if (!existing || row.views > existing.views) {
        byKey[key] = row;
      }
    });

    return Object.keys(byKey)
      .map(function (k) {
        return byKey[k];
      })
      .sort(function (a, b) {
        return b.views - a.views;
      });
  }

  function buildPageUrlLookup(rows) {
    var map = {};
    (rows || []).forEach(function (row) {
      var key = normalizeUrlKey(row.link);
      if (!key) return;
      map[key] = row;
    });
    return map;
  }

  function findPageByUrl(url) {
    var key = normalizeUrlKey(url);
    if (!key) return null;
    return pageRowsByUrl[key] || null;
  }

  function getTrackedPageRows(rows) {
    return (rows || []).filter(function (r) {
      return r.pageType === "menu" || r.pageType === "article";
    });
  }

  function loadLatestPageKpis() {
    return resolvePageKpiSource().then(function (source) {
      var latest = source.latest;
      pageKpiFileModified = latest.TimeLastModified || null;
      PAGE_KPI_FOLDER_SERVER_RELATIVE_URL = source.folder;

      return fetchFileText(latest.ServerRelativeUrl).then(function (text) {
        return {
          rows: parsePageKpiText(text),
          fileName: latest.Name,
          modified: latest.TimeLastModified,
          folder: source.folder
        };
      });
    });
  }

  /* =========================================================================
   * NEWSLETTER LIST FETCH
   * =========================================================================
   */
  function buildNewsletterSelect() {
    var fields = [
      "Id",
      NEWSLETTER_FIELDS.Title,
      NEWSLETTER_FIELDS.SentDateTime,
      NEWSLETTER_FIELDS.DeliveredTo,
      NEWSLETTER_FIELDS.Notes,
      NEWSLETTER_FIELDS.Opens24hSeries,
      NEWSLETTER_FIELDS.Clicks24hSeries,
      NEWSLETTER_FIELDS.Opens7dSeries,
      NEWSLETTER_FIELDS.Clicks7dSeries
    ];

    for (var i = 1; i <= 6; i++) {
      fields.push("Story" + i + "Title");
      fields.push("Story" + i + "Url");
      fields.push("Story" + i + "Clicks");
    }

    return fields.join(",");
  }

  function normalizeStoryUrl(raw) {
    if (raw == null) return "";
    var url = String(raw).trim();
    if (!url) return "";
    // Allow http(s) only
    if (!/^https?:\/\//i.test(url)) {
      if (/^\/\//.test(url)) {
        url = "https:" + url;
      } else if (/^www\./i.test(url)) {
        url = "https://" + url;
      } else {
        return "";
      }
    }
    if (!/^https?:\/\//i.test(url)) return "";
    return url;
  }

  function extractStories(item) {
    var stories = [];
    for (var i = 1; i <= 6; i++) {
      var title = item["Story" + i + "Title"];
      title = title == null ? "" : String(title).trim();
      if (!title) continue;
      stories.push({
        title: title,
        url: normalizeStoryUrl(item["Story" + i + "Url"]),
        clicks: toInt(item["Story" + i + "Clicks"])
      });
    }
    return stories;
  }

  function normalizeNewsletterItem(item) {
    var opens24hSeries = parseSeries(
      item[NEWSLETTER_FIELDS.Opens24hSeries],
      SERIES_LEN_24H
    );
    var clicks24hSeries = parseSeries(
      item[NEWSLETTER_FIELDS.Clicks24hSeries],
      SERIES_LEN_24H
    );
    var opens7dSeries = parseSeries(
      item[NEWSLETTER_FIELDS.Opens7dSeries],
      SERIES_LEN_7D
    );
    var clicks7dSeries = parseSeries(
      item[NEWSLETTER_FIELDS.Clicks7dSeries],
      SERIES_LEN_7D
    );

    var opens24h = sumSeries(opens24hSeries);
    var clicks24h = sumSeries(clicks24hSeries);
    var opens7d = sumSeries(opens7dSeries);
    var clicks7d = sumSeries(clicks7dSeries);

    return {
      id: item.Id,
      title: item[NEWSLETTER_FIELDS.Title] || "Untitled edition",
      sentDateTime: item[NEWSLETTER_FIELDS.SentDateTime] || null,
      deliveredTo: toInt(item[NEWSLETTER_FIELDS.DeliveredTo]),
      notes: item[NEWSLETTER_FIELDS.Notes] || "",
      opens24h: opens24h,
      clicks24h: clicks24h,
      opens7d: opens7d,
      clicks7d: clicks7d,
      opens24hSeries: opens24hSeries,
      clicks24hSeries: clicks24hSeries,
      opens7dSeries: opens7dSeries,
      clicks7dSeries: clicks7dSeries,
      stories: extractStories(item)
    };
  }

  function getEditionSortTime(edition) {
    if (!edition.sentDateTime) return 0;
    var t = new Date(edition.sentDateTime).getTime();
    return isNaN(t) ? 0 : t;
  }

  function getEditionDisplayDateTime(edition) {
    return formatDateTime(edition.sentDateTime);
  }

  function loadNewsletterEditions() {
    var safeTitle = NEWSLETTER_LIST_NAME.replace(/'/g, "''");
    var url =
      SITE_URL.replace(/\/$/, "") +
      "/_api/web/lists/getbytitle('" +
      safeTitle +
      "')/items?$select=" +
      encodeURIComponent(buildNewsletterSelect()) +
      "&$orderby=" +
      encodeURIComponent(NEWSLETTER_FIELDS.SentDateTime + " desc") +
      "&$top=5000";

    return $.ajax({
      url: url,
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" }
    }).then(
      function (data) {
        var results = (data.d && data.d.results) || [];
        return results
          .map(normalizeNewsletterItem)
          .sort(function (a, b) {
            return getEditionSortTime(b) - getEditionSortTime(a);
          });
      },
      function (xhr) {
        var msg =
          "Newsletter list '" +
          NEWSLETTER_LIST_NAME +
          "' was not found (create it using pageKpiDashboard_LIST_SETUP.md). Page KPIs can still load.";
        if (xhr && xhr.status && xhr.status !== 404) {
          msg =
            "Could not load list '" +
            NEWSLETTER_LIST_NAME +
            "' (HTTP " +
            xhr.status +
            ").";
        }
        return $.Deferred().reject(new Error(msg)).promise();
      }
    );
  }

  /* =========================================================================
   * PAGE KPI RENDER
   * =========================================================================
   */
  function formatPageTitleCell(row) {
    var titleHtml = escapeHtml(row.title);
    if (row.link) {
      return (
        '<a class="story-link" href="' +
        escapeHtml(row.link) +
        '" target="_blank" rel="noopener noreferrer">' +
        titleHtml +
        "</a>"
      );
    }
    return titleHtml;
  }

  function renderPageTable($body, $table, $empty, rows, limit, options) {
    $body.empty();
    var list = rows || [];
    var top =
      limit == null || limit <= 0 ? list : list.slice(0, limit);
    var includeViva = !!(options && options.includeVivaViews);

    if (!top.length) {
      $table.addClass("is-hidden");
      $empty.removeClass("is-hidden");
      return;
    }

    $table.removeClass("is-hidden");
    $empty.addClass("is-hidden");

    top.forEach(function (row, index) {
      var cells =
        "<tr>" +
        "<td>" +
        (index + 1) +
        "</td>" +
        "<td>" +
        formatPageTitleCell(row) +
        "</td>" +
        "<td>" +
        escapeHtml(row.hubSection || "—") +
        "</td>" +
        '<td class="num">' +
        formatNumber(row.views) +
        "</td>" +
        '<td class="num">' +
        formatNumber(row.viewers) +
        "</td>";

      if (includeViva) {
        cells +=
          '<td class="num">' + formatVivaViewsCell(row) + "</td>";
      }

      cells += "</tr>";
      $body.append(cells);
    });
  }

  function pageMatchesQuery(row, query) {
    if (!query) return true;
    var haystack = (
      (row.title || "") +
      " " +
      (row.hubSection || "") +
      " " +
      (row.link || "")
    ).toLowerCase();
    return haystack.indexOf(query) !== -1;
  }

  function getPageSearchQuery() {
    return $.trim($("#pageSearchInput").val() || "").toLowerCase();
  }

  function renderPageKpis(allRows) {
    pageRows = allRows || pageRows || [];
    var query = getPageSearchQuery();
    var isSearch = Boolean(query);

    var tracked = getTrackedPageRows(pageRows);
    var menuRows = tracked.filter(function (r) {
      return r.pageType === "menu" && pageMatchesQuery(r, query);
    });
    var articleRows = tracked.filter(function (r) {
      return r.pageType === "article" && pageMatchesQuery(r, query);
    });
    var otherRows = pageRows.filter(function (r) {
      return (
        (r.pageType === "ignore" || r.pageType === "other") &&
        pageMatchesQuery(r, query)
      );
    });

    var totalViews = 0;
    var allTracked = getTrackedPageRows(pageRows);
    allTracked.forEach(function (r) {
      totalViews += r.views;
    });

    var allMenu = allTracked.filter(function (r) {
      return r.pageType === "menu";
    });
    var allArticles = allTracked.filter(function (r) {
      return r.pageType === "article";
    });

    $("#kpiTotalViews").text(formatNumber(totalViews));
    $("#kpiPageCount").text(formatNumber(allTracked.length));
    $("#kpiPageCountSub").text(
      allMenu.length +
        " menu · " +
        allArticles.length +
        " articles (Templates/Archive ignored)"
    );

    renderArticleStats(allArticles);

    var limit = isSearch ? 0 : TOP_PAGES_COUNT;
    renderPageTable(
      $("#menuPagesBody"),
      $("#menuPagesTable"),
      $("#menuPagesEmpty"),
      menuRows,
      limit
    );
    renderPageTable(
      $("#articlePagesBody"),
      $("#articlePagesTable"),
      $("#articlePagesEmpty"),
      articleRows,
      limit,
      { includeVivaViews: true }
    );

    if (isSearch) {
      $("#menuPagesCaption").text(
        "Menu pages matching \"" + query + "\" (" + menuRows.length + ")"
      );
      $("#articlePagesCaption").text(
        "News articles matching \"" + query + "\" (" + articleRows.length + ")"
      );
      $("#articlePagesTableCaption").text(
        "Matching articles (" + articleRows.length + ")"
      );
      $("#pageSearchHint").text(
        "Showing all matches for \"" +
          query +
          "\". Clear the search to return to top pages."
      );
      $("#otherPagesSearchBlock").removeClass("is-hidden");
      $("#otherPagesCaption").text(
        "Other matches — Templates / Archive / uncategorized (" +
          otherRows.length +
          ")"
      );
      renderPageTable(
        $("#otherPagesBody"),
        $("#otherPagesTable"),
        $("#otherPagesEmpty"),
        otherRows,
        0
      );
      $("#menuPagesEmpty").text("No menu pages match this search.");
      $("#articlePagesEmpty").text("No articles match this search.");
    } else {
      $("#menuPagesCaption").text("Top menu pages by views");
      $("#articlePagesCaption").text("News articles");
      $("#articlePagesTableCaption").text("Top articles by views");
      $("#pageSearchHint").text(
        "Showing top pages by views. Type to find a specific page."
      );
      $("#otherPagesSearchBlock").addClass("is-hidden");
      $("#menuPagesEmpty").text("No menu page KPI data found.");
      $("#articlePagesEmpty").text("No article KPI data found.");
    }
  }

  function renderArticleStats(articles) {
    var count = (articles || []).length;
    var views = (articles || []).map(function (r) {
      return r.views;
    });
    var viewers = (articles || []).map(function (r) {
      return r.viewers;
    });

    var avgViews = averageOf(views);
    var avgViewers = averageOf(viewers);

    var vivaTotal = 0;
    var vivaWithData = 0;
    var seenCanonical = {};
    (articles || []).forEach(function (r) {
      var viva = getVivaViewsForPage(r);
      if (viva == null) return;
      // Avoid double-counting if two article rows resolve to the same page
      var keys = pageLookupKeys(r);
      var dedupeKey = keys[0] || String(r.link || r.title || "");
      if (seenCanonical[dedupeKey]) return;
      keys.forEach(function (k) {
        seenCanonical[k] = true;
      });
      vivaWithData += 1;
      vivaTotal += viva;
    });

    var listVivaTotal =
      typeof vivaViewsByPageKey.__totalVivaRows === "number"
        ? vivaViewsByPageKey.__totalVivaRows
        : null;

    var countLabel =
      count === 1 ? "1 article (HubSection: Articles)" : count + " articles (HubSection: Articles)";

    $("#kpiArticleAvgViews").text(
      avgViews == null ? "—" : formatDecimal(avgViews, 1)
    );
    $("#kpiArticleAvgViewers").text(
      avgViewers == null ? "—" : formatDecimal(avgViewers, 1)
    );
    $("#kpiArticleVivaViews").text(
      listVivaTotal != null && listVivaTotal > 0
        ? formatNumber(listVivaTotal)
        : vivaWithData
          ? formatNumber(vivaTotal)
          : "N/A"
    );

    $("#kpiArticleAvgViewsSub").text(countLabel);
    $("#kpiArticleAvgViewersSub").text(countLabel);
    $("#kpiArticleVivaViewsSub").text(
      listVivaTotal != null && listVivaTotal > 0
        ? listVivaTotal +
            " tracked · " +
            vivaWithData +
            " of " +
            count +
            " articles matched"
        : vivaWithData
          ? vivaWithData +
              " of " +
              count +
              " articles have Viva source data"
          : "No Viva Engage views in Page View Tracking yet"
    );
  }

  function wirePageSearchEvents() {
    $("#pageSearchInput")
      .off("input.pageKpiSearch")
      .on("input.pageKpiSearch", function () {
        renderPageKpis(pageRows);
      });
  }

  /* =========================================================================
   * NEWSLETTER RENDER
   * =========================================================================
   */
  function populateNewsletterDropdown(editions) {
    var $select = $("#newsletterEditionSelect");
    $select.find("option:not([value='overall'])").remove();

    editions.forEach(function (edition) {
      var label = edition.title;
      var dateLabel = getEditionDisplayDateTime(edition);
      if (dateLabel) label += " (" + dateLabel + ")";
      $select.append(
        $("<option></option>")
          .attr("value", String(edition.id))
          .text(label)
      );
    });
  }

  function avgAcrossLabel(count) {
    if (!count) return "";
    return "Avg across " + count + " edition" + (count === 1 ? "" : "s");
  }

  function buildLinePath(points) {
    if (!points.length) return "";
    var d = "M " + points[0].x.toFixed(1) + " " + points[0].y.toFixed(1);
    for (var i = 1; i < points.length; i++) {
      d += " L " + points[i].x.toFixed(1) + " " + points[i].y.toFixed(1);
    }
    return d;
  }

  function renderDualLineChart($container, opensSeries, clicksSeries, labelPrefix) {
    if (!$container || !$container.length) return;

    var hasOpens = seriesHasValues(opensSeries);
    var hasClicks = seriesHasValues(clicksSeries);

    if (!hasOpens && !hasClicks) {
      $container.html(
        '<div class="chart-empty">No series data yet. Paste incremental values into the series columns.</div>'
      );
      return;
    }

    var len = Math.max(
      (opensSeries && opensSeries.length) || 0,
      (clicksSeries && clicksSeries.length) || 0
    );
    if (!len) {
      $container.html(
        '<div class="chart-empty">No series data yet. Paste incremental values into the series columns.</div>'
      );
      return;
    }

    var opens = opensSeries && opensSeries.length ? opensSeries : zeroSeries(len);
    var clicks = clicksSeries && clicksSeries.length ? clicksSeries : zeroSeries(len);

    var maxVal = 0;
    var i;
    for (i = 0; i < len; i++) {
      if (opens[i] > maxVal) maxVal = opens[i];
      if (clicks[i] > maxVal) maxVal = clicks[i];
    }
    if (maxVal <= 0) maxVal = 1;

    var width = 560;
    var height = 200;
    var padL = 40;
    var padR = 12;
    var padT = 12;
    var padB = 28;
    var plotW = width - padL - padR;
    var plotH = height - padT - padB;

    function xAt(idx) {
      if (len === 1) return padL + plotW / 2;
      return padL + (idx / (len - 1)) * plotW;
    }

    function yAt(val) {
      return padT + plotH - (val / maxVal) * plotH;
    }

    var openPts = [];
    var clickPts = [];
    for (i = 0; i < len; i++) {
      openPts.push({ x: xAt(i), y: yAt(opens[i] || 0) });
      clickPts.push({ x: xAt(i), y: yAt(clicks[i] || 0) });
    }

    var gridLines = "";
    var tickCount = 4;
    for (i = 0; i <= tickCount; i++) {
      var gy = padT + (plotH * i) / tickCount;
      var gVal = Math.round(maxVal * (1 - i / tickCount));
      gridLines +=
        '<line x1="' +
        padL +
        '" y1="' +
        gy.toFixed(1) +
        '" x2="' +
        (width - padR) +
        '" y2="' +
        gy.toFixed(1) +
        '" stroke="#edf0f3" stroke-width="1"/>';
      gridLines +=
        '<text x="' +
        (padL - 6) +
        '" y="' +
        (gy + 3).toFixed(1) +
        '" text-anchor="end" font-size="10" fill="#888">' +
        gVal +
        "</text>";
    }

    var xLabels = "";
    var labelStep = len > 12 ? 3 : 1;
    for (i = 0; i < len; i++) {
      if (i % labelStep !== 0 && i !== len - 1) continue;
      xLabels +=
        '<text x="' +
        xAt(i).toFixed(1) +
        '" y="' +
        (height - 8) +
        '" text-anchor="middle" font-size="10" fill="#888">' +
        escapeHtml(labelPrefix + (i + 1)) +
        "</text>";
    }

    var svg =
      '<svg viewBox="0 0 ' +
      width +
      " " +
      height +
      '" role="img" aria-label="Opens and clicks trend">' +
      gridLines +
      '<line x1="' +
      padL +
      '" y1="' +
      (padT + plotH) +
      '" x2="' +
      (width - padR) +
      '" y2="' +
      (padT + plotH) +
      '" stroke="#cfd6dd" stroke-width="1"/>' +
      '<path d="' +
      buildLinePath(openPts) +
      '" fill="none" stroke="' +
      CHART_COLOR_OPENS +
      '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>' +
      '<path d="' +
      buildLinePath(clickPts) +
      '" fill="none" stroke="' +
      CHART_COLOR_CLICKS +
      '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>' +
      xLabels +
      "</svg>";

    $container.html(svg);
  }

  function renderNewsletterCharts(opens24h, clicks24h, opens7d, clicks7d, modeLabel) {
    $("#nlChart24hSub").text(
      modeLabel
        ? "Incremental opens and clicks per hour · " + modeLabel
        : "Incremental opens and clicks per hour"
    );
    $("#nlChart7dSub").text(
      modeLabel
        ? "Incremental opens and clicks per day · " + modeLabel
        : "Incremental opens and clicks per day"
    );

    renderDualLineChart($("#nlChart24h"), opens24h, clicks24h, "H");
    renderDualLineChart($("#nlChart7d"), opens7d, clicks7d, "D");
  }

  function formatStoryTitleCell(story) {
    var titleHtml = escapeHtml(story.title);
    if (story.url) {
      return (
        '<a class="story-link" href="' +
        escapeHtml(story.url) +
        '" target="_blank" rel="noopener noreferrer">' +
        titleHtml +
        "</a>"
      );
    }
    return titleHtml;
  }

  function enrichStoriesWithPageKpis(stories) {
    return (stories || []).map(function (story) {
      var page = findPageByUrl(story.url);
      var pageViews = page ? page.views : null;
      var viewSharePct =
        page && pageViews > 0 ? (story.clicks / pageViews) * 100 : null;
      return {
        title: story.title,
        url: story.url,
        clicks: story.clicks,
        pageViews: pageViews,
        pageViewers: page ? page.viewers : null,
        pageMatched: Boolean(page),
        viewSharePct: viewSharePct
      };
    });
  }

  function averageNewsletterViewShare(stories) {
    var enriched = enrichStoriesWithPageKpis(stories);
    var percentages = [];
    enriched.forEach(function (story) {
      if (story.viewSharePct != null && isFinite(story.viewSharePct)) {
        percentages.push(story.viewSharePct);
      }
    });
    return {
      average: averageOf(percentages),
      matchedCount: percentages.length,
      storyCount: enriched.length
    };
  }

  function renderNewsletterViewShare(stories) {
    var stats = averageNewsletterViewShare(stories);
    if (stats.average == null) {
      $("#nlAvgViewShare").text("—");
      $("#nlAvgViewShareSub").text(
        stats.storyCount
          ? "No matched page URLs to compute share"
          : "No stories available"
      );
      return;
    }

    $("#nlAvgViewShare").text(formatDecimal(stats.average, 1) + "%");
    $("#nlAvgViewShareSub").text(
      "Avg across " +
        stats.matchedCount +
        " matched stor" +
        (stats.matchedCount === 1 ? "y" : "ies") +
        " (clicks ÷ page views)"
    );
  }

  function renderStoryTable(stories, caption, viewShareStories) {
    $("#storyTableCaption").text(caption || "Story engagement");
    var $body = $("#storyClicksBody").empty();
    var enriched = enrichStoriesWithPageKpis(stories);

    /* viewShareStories lets Overall average all stories while the table shows top N */
    renderNewsletterViewShare(
      viewShareStories != null ? viewShareStories : stories
    );

    if (!enriched.length) {
      $("#storyClicksTable").addClass("is-hidden");
      $("#storyClicksEmpty").removeClass("is-hidden");
      return;
    }

    $("#storyClicksTable").removeClass("is-hidden");
    $("#storyClicksEmpty").addClass("is-hidden");

    enriched.forEach(function (story, index) {
      $body.append(
        "<tr>" +
          "<td>" +
          (index + 1) +
          "</td>" +
          "<td>" +
          formatStoryTitleCell(story) +
          "</td>" +
          '<td class="num">' +
          formatNumber(story.clicks) +
          "</td>" +
          '<td class="num">' +
          (story.pageMatched ? formatNumber(story.pageViews) : "—") +
          "</td>" +
          '<td class="num">' +
          (story.pageMatched ? formatNumber(story.pageViewers) : "—") +
          "</td>" +
          '<td class="num">' +
          (story.viewSharePct == null
            ? "—"
            : formatDecimal(story.viewSharePct, 1) + "%") +
          "</td>" +
          "</tr>"
      );
    });
  }

  function aggregateOverallStories(editions) {
    var byTitle = {};
    editions.forEach(function (edition) {
      (edition.stories || []).forEach(function (story) {
        if (!byTitle[story.title]) {
          byTitle[story.title] = { clicks: 0, url: "" };
        }
        byTitle[story.title].clicks += story.clicks;
        if (!byTitle[story.title].url && story.url) {
          byTitle[story.title].url = story.url;
        }
      });
    });

    return Object.keys(byTitle)
      .map(function (title) {
        return {
          title: title,
          url: byTitle[title].url,
          clicks: byTitle[title].clicks
        };
      })
      .sort(function (a, b) {
        return b.clicks - a.clicks;
      });
  }

  function renderNewsletterOverall(editions) {
    var count = editions.length;
    var opens24h = 0;
    var clicks24h = 0;
    var opens7d = 0;
    var clicks7d = 0;
    var deliveredTo = 0;

    editions.forEach(function (e) {
      opens24h += e.opens24h;
      clicks24h += e.clicks24h;
      opens7d += e.opens7d;
      clicks7d += e.clicks7d;
      deliveredTo += e.deliveredTo || 0;
    });

    var deliveredAvg = count ? deliveredTo / count : 0;
    var opens24hAvg = count ? opens24h / count : 0;
    var clicks24hAvg = count ? clicks24h / count : 0;
    var opens7dAvg = count ? opens7d / count : 0;
    var clicks7dAvg = count ? clicks7d / count : 0;

    $("#newsletterSubtitle").text(
      count
        ? "Overall averages across " +
            count +
            " newsletter edition" +
            (count === 1 ? "" : "s") +
            "."
        : "No newsletter editions found yet. Add items to the Newsletter KPIs list."
    );

    $("#nlDeliveredTo").text(formatNumber(deliveredAvg));
    $("#nlOpens24h").text(formatNumber(opens24hAvg));
    $("#nlClicks24h").text(formatNumber(clicks24hAvg));
    $("#nlOpens7d").text(formatNumber(opens7dAvg));
    $("#nlClicks7d").text(formatNumber(clicks7dAvg));

    $("#nlDeliveredToSub").text(avgAcrossLabel(count));
    $("#nlOpens24hSub").text(avgAcrossLabel(count));
    $("#nlClicks24hSub").text(avgAcrossLabel(count));
    $("#nlOpens7dSub").text(avgAcrossLabel(count));
    $("#nlClicks7dSub").text(avgAcrossLabel(count));

    renderNewsletterCharts(
      averageSeriesPointwise(editions, "opens24hSeries", SERIES_LEN_24H),
      averageSeriesPointwise(editions, "clicks24hSeries", SERIES_LEN_24H),
      averageSeriesPointwise(editions, "opens7dSeries", SERIES_LEN_7D),
      averageSeriesPointwise(editions, "clicks7dSeries", SERIES_LEN_7D),
      count ? "avg across editions" : ""
    );

    var allStories = aggregateOverallStories(editions);
    renderStoryTable(
      allStories.slice(0, TOP_STORIES_OVERALL_COUNT),
      "Top stories across editions",
      allStories
    );
  }

  function renderNewsletterEdition(edition) {
    if (!edition) {
      renderNewsletterOverall([]);
      return;
    }

    var sentLabel = getEditionDisplayDateTime(edition);
    $("#newsletterSubtitle").text(
      edition.title + (sentLabel ? " · Sent " + sentLabel : "")
    );

    $("#nlDeliveredTo").text(formatNumber(edition.deliveredTo));
    $("#nlOpens24h").text(formatNumber(edition.opens24h));
    $("#nlClicks24h").text(formatNumber(edition.clicks24h));
    $("#nlOpens7d").text(formatNumber(edition.opens7d));
    $("#nlClicks7d").text(formatNumber(edition.clicks7d));

    $("#nlDeliveredToSub").text("Recipients for this send");
    $("#nlOpens24hSub").text("Sum of hourly increments");
    $("#nlClicks24hSub").text("Sum of hourly increments");
    $("#nlOpens7dSub").text("Sum of daily increments");
    $("#nlClicks7dSub").text("Sum of daily increments");

    renderNewsletterCharts(
      edition.opens24hSeries,
      edition.clicks24hSeries,
      edition.opens7dSeries,
      edition.clicks7dSeries,
      "this edition"
    );

    var stories = (edition.stories || []).slice().sort(function (a, b) {
      return b.clicks - a.clicks;
    });

    renderStoryTable(stories, "Stories for this edition");
  }

  function renderNewsletterSelection() {
    var value = $("#newsletterEditionSelect").val() || "overall";
    if (value === "overall") {
      renderNewsletterOverall(newsletterEditions);
      return;
    }

    var id = parseInt(value, 10);
    var edition = null;
    newsletterEditions.forEach(function (e) {
      if (e.id === id) edition = e;
    });
    renderNewsletterEdition(edition);
  }

  function wireNewsletterEvents() {
    $("#newsletterEditionSelect")
      .off("change.pageKpi")
      .on("change.pageKpi", function () {
        renderNewsletterSelection();
      });
  }

  /* =========================================================================
   * BOOTSTRAP
   * =========================================================================
   */
  function showContent() {
    $("#loadingOverlay").addClass("is-hidden");
    $("#dashboardContent").removeClass("is-hidden");
  }

  function init() {
    clearError();
    wireNewsletterEvents();
    wirePageSearchEvents();

    var messages = [];
    var pageDone = false;
    var newsletterDone = false;
    var vivaDone = false;

    function maybeFinish() {
      if (!pageDone || !newsletterDone || !vivaDone) return;
      pageRowsByUrl = buildPageUrlLookup(pageRows);
      renderPageKpis(pageRows);
      renderNewsletterSelection();
      if (messages.length) {
        showError(messages.join(" "));
      }
      setLastUpdated(pageKpiFileModified);
      showContent();
    }

    loadLatestPageKpis()
      .done(function (pageResult) {
        pageRows = (pageResult && pageResult.rows) || [];
        pageKpiFileModified = (pageResult && pageResult.modified) || null;
        pageRowsByUrl = buildPageUrlLookup(pageRows);
      })
      .fail(function (err) {
        pageRows = [];
        pageRowsByUrl = {};
        var msg =
          (err && err.message) ||
          (err && err.responseText) ||
          "Failed to load Page KPIs.";
        if (err && err.status) {
          msg =
            "Failed to load Page KPIs folder (" +
            PAGE_KPI_FOLDER_SERVER_RELATIVE_URL +
            "). HTTP " +
            err.status +
            ". Update PAGE_KPI_FOLDER_SERVER_RELATIVE_URL in pageKpiDashboard.js if needed.";
        }
        messages.push(String(msg));
      })
      .always(function () {
        pageDone = true;
        maybeFinish();
      });

    loadNewsletterEditions()
      .done(function (editions) {
        newsletterEditions = editions || [];
        populateNewsletterDropdown(newsletterEditions);
        $("#newsletterEditionSelect").val("overall");
      })
      .fail(function (err) {
        newsletterEditions = [];
        populateNewsletterDropdown([]);
        messages.push(
          (err && err.message) ||
            "Failed to load Newsletter KPIs list."
        );
      })
      .always(function () {
        newsletterDone = true;
        maybeFinish();
      });

    loadVivaEngageViews()
      .done(function () {
        /* vivaViewsByPageKey already set */
      })
      .fail(function (err) {
        vivaViewsByPageKey = {};
        messages.push(
          (err && err.message) ||
            "Failed to load Page View Tracking list."
        );
      })
      .always(function () {
        vivaDone = true;
        maybeFinish();
      });
  }

  $(document).ready(init);
})();
