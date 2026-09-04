(() => {
  const PREFIX = "[sp-projects]";

  function escapeHtml(str) {
    return String(str)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function splitLines(value) {
    return String(value || "")
      .split(/\r?\n/g)
      .map((s) => s.trim())
      .filter(Boolean);
  }

  /**
   * SharePoint "enhanced" multiline fields return HTML with no \\n between blocks — splitLines would keep one glued line.
   * Convert <a> to [label](url), then block tags to newlines, then strip tags.
   */
  function prepareKeyPointsForSplit(value) {
    let s = String(value || "");
    if (!s.includes("<")) return s;

    s = s.replace(
      /<a\s+[^>]*\bhref\s*=\s*(["'])([^"']*)\1[^>]*>([\s\S]*?)<\/a>/gi,
      (_m, _q, href, inner) => {
        const label = inner.replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
        const url = String(href).replace(/&amp;/g, "&");
        if (!/^https?:\/\//i.test(url)) return label || url;
        return `[${label || url}](${url})`;
      }
    );

    s = s.replace(/<br\s*\/?>/gi, "\n");
    s = s.replace(/<\/p>/gi, "\n");
    s = s.replace(/<p\b[^>]*>/gi, "\n");
    s = s.replace(/<\/div>/gi, "\n");
    s = s.replace(/<div\b[^>]*>/gi, "\n");
    // SharePoint often stores bullet lists as <ul>/<li>…; without this,
    // splitLines() would see one glued paragraph ("linkONE …").
    s = s.replace(/<\/li>/gi, "\n");
    s = s.replace(/<li\b[^>]*>/gi, "\n");
    s = s.replace(/<\/ul>/gi, "\n");
    s = s.replace(/<ul\b[^>]*>/gi, "\n");
    s = s.replace(/<\/ol>/gi, "\n");
    s = s.replace(/<ol\b[^>]*>/gi, "\n");
    s = s.replace(/<[^>]+>/g, "");
    s = s.replace(/&nbsp;/gi, " ");
    s = s.replace(/&amp;/g, "&");
    s = s.replace(/&lt;/g, "<");
    s = s.replace(/&gt;/g, ">");
    return s;
  }

  function normWs(str) {
    return String(str).replace(/\s+/g, " ").trim();
  }

  /** `Sentence linkSentence` (rich text artifact) → one `Sentence ` before [link](url). */
  function collapseLinkGlueDuplicate(before) {
    const t = before.trimEnd();
    const m = /^(.+?)\s+link\b(.+)$/i.exec(t);
    if (!m) return before;
    if (normWs(m[1]) === normWs(m[2])) {
      return normWs(m[1]) + " ";
    }
    return before;
  }

  function anchorTag(url, label) {
    return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
  }

  /**
   * Each `[label](https://...)` → link. Text after the last link is dropped.
   * Collapse `…Description link…Description` when both halves match (SharePoint rich-text + markdown).
   */
  function formatBulletLine(text) {
    const trimmed = String(text).trim();
    const re = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/gi;
    let lastMatch = null;
    let m;
    while ((m = re.exec(trimmed)) !== null) {
      lastMatch = m;
    }
    if (!lastMatch) return escapeHtml(trimmed);

    let before = trimmed.slice(0, lastMatch.index);
    const md = trimmed.slice(lastMatch.index, lastMatch.lastIndex);
    before = collapseLinkGlueDuplicate(before);
    const s = (before + md).trim();

    let out = "";
    let last = 0;
    const re2 = new RegExp(re.source, re.flags);
    while ((m = re2.exec(s)) !== null) {
      out += escapeHtml(s.slice(last, m.index));
      out += anchorTag(m[2], m[1]);
      last = m.lastIndex;
    }
    return out;
  }

  async function fetchAll(url) {
    const items = [];
    let next = url;
    while (next) {
      const res = await fetch(next, {
        method: "GET",
        headers: { Accept: "application/json;odata=nometadata" },
        credentials: "same-origin",
      });
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}\n\n${text}`);
      const data = JSON.parse(text);
      items.push(...(data.value || []));
      next = data["@odata.nextLink"] || null;
    }
    return items;
  }

  function renderCell(item, fields) {
    const title = item[fields.Title] || "";
    const keyPoints = splitLines(prepareKeyPointsForSplit(item[fields.KeyPoints]));

    const bulletsHtml = keyPoints.length
      ? `<ul class="sp-bullets">${keyPoints.map((t) => `<li>${formatBulletLine(t)}</li>`).join("")}</ul>`
      : `<ul class="sp-bullets"><li style="color: var(--sp-muted)">No key points.</li></ul>`;

    const chevronSvg = `
      <svg class="sp-chevron" viewBox="0 0 20 20" aria-hidden="true">
        <path
          fill="currentColor"
          d="M5.3 7.7a1 1 0 0 1 1.4 0L10 11l3.3-3.3a1 1 0 1 1 1.4 1.4l-4 4a1 1 0 0 1-1.4 0l-4-4a1 1 0 0 1 0-1.4Z"
        />
      </svg>
    `;

    return `
      <article class="sp-cell">
        <details class="sp-details">
          <summary class="sp-summary">
            <span class="sp-summaryTitle">${escapeHtml(title)}</span>
            ${chevronSvg}
          </summary>
          <div class="sp-panelInner">
            ${bulletsHtml}
          </div>
        </details>
      </article>
    `;
  }

  async function hydrate(root) {
    const gridEl = root.querySelector("[data-sp-grid]");

    const siteUrl = root.getAttribute("data-site-url") || "";
    const listTitle = root.getAttribute("data-list-title") || "";
    const yearFilter = root.getAttribute("data-year") || "";

    const fields = {
      Id: root.getAttribute("data-field-id") || "Id",
      Title: root.getAttribute("data-field-title") || "Title",
      Year: root.getAttribute("data-field-year") || "Year",
      KeyPoints: root.getAttribute("data-field-keypoints") || "KeyPoints",
    };

    if (!gridEl) return;
    if (!siteUrl || !listTitle) {
      return;
    }

    gridEl.innerHTML = `<div style="padding: 8px 0; color: var(--sp-muted)">Loading…</div>`;

    const base = siteUrl.replace(/\/$/, "");
    const selectFields = [fields.Id, fields.Title, fields.Year, fields.KeyPoints].join(",");

    const api =
      base +
      "/_api/web/lists/getbytitle('" +
      listTitle.replaceAll("'", "''") +
      "')/items?$select=" +
      encodeURIComponent(selectFields) +
      "&$top=5000&$orderby=" +
      encodeURIComponent(`${fields.Id} asc`);

    const raw = await fetchAll(api);

    const filtered =
      yearFilter === ""
        ? raw
        : raw.filter((it) => String(it[fields.Year] ?? "").trim() === String(yearFilter));

    const htmlParts = filtered.map((it) => renderCell(it, fields));

    gridEl.innerHTML =
      htmlParts.length > 0
        ? htmlParts.join("")
        : `<div style="padding: 8px 0; color: var(--sp-muted)">No items found.</div>`;
  }

  async function run() {
    const roots = document.querySelectorAll("[data-sp-projects]");
    if (!roots.length) return;
    for (const root of roots) {
      try {
        // Mark each instance so multiple webparts are safe
        if (root.__spProjectsHydrated) continue;
        root.__spProjectsHydrated = true;
        await hydrate(root);
      } catch (e) {
        const gridEl = root.querySelector("[data-sp-grid]");
        if (gridEl) {
          const msg = String(e && (e.stack || e.message) ? e.stack || e.message : e);
          gridEl.innerHTML =
            `<div style="padding:8px 0; color:#b00020; font-weight:700;">Failed to load list data</div>` +
            `<pre style="white-space:pre-wrap; font-size:12px; color:#b00020; margin:6px 0 0 0;">${escapeHtml(
              msg
            )}</pre>`;
        }
        // eslint-disable-next-line no-console
        console.error(PREFIX, e);
      }
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", run);
  } else {
    run();
  }
})();

