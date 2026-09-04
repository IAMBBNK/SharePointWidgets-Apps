(function () {
  const siteUrl = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  const listName = "IT/OT Governance Stakeholders";
  const DATA_AI_TITLE = "DATA & AI";
  // SharePoint list Titles in display order (Cyber Security item → Corporate Security label)
  const SECTOR_ORDER = [
    "Electronics",
    "Life Science",
    "Healthcare",
    "Site Management",
    "Merck Group IT (ETP Architecture)",
    "Merck Group IT (Network)",
    "IT Application Management (Manufacturing)",
    "IT Application Management (Laboratories)",
    "Cyber Security",
    DATA_AI_TITLE
  ];

  const url =
    `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Title,` +
    `TechnologyandArchitectureCoE/Title,TechnologyandArchitectureCoE/EMail,` +
    `TechnologyandArchitectureCoEBack/Title,TechnologyandArchitectureCoEBack/EMail` +
    `&$expand=TechnologyandArchitectureCoE,TechnologyandArchitectureCoEBack`;

  function getContainer() {
    return document.getElementById("techArchCoeContainer");
  }

  function showError(message) {
    const container = getContainer();
    if (!container) {
      return;
    }
    container.innerHTML =
      '<p class="techarch-load-error" role="alert">' + message + "</p>";
  }

  const getPhotoUrlFromEmail = (email) =>
    email
      ? `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`
      : null;

  const createProfilePic = (email) => {
    const photoUrl = getPhotoUrlFromEmail(email);
    if (!photoUrl) {
      return "";
    }
    return `
      <div class="techarch-profile-pic-wrapper">
        <img src="${photoUrl}" alt="Profile" class="techarch-profile-pic"
             onerror="this.style.display='none';">
      </div>
    `;
  };

  const formatName = (name, email) => {
    if (!name) {
      return `
        <div class="techarch-person-wrapper">
          <div class="techarch-name-text">TBD</div>
        </div>`;
    }
    const parts = name.split(" ");
    const displayName =
      parts.length < 2 ? name : parts[0] + "<br>" + parts.slice(1).join(" ");
    const pic = createProfilePic(email);
    const inner = `
      <div class="techarch-person-wrapper">
        ${pic}
        <div class="techarch-name-text">${displayName}</div>
      </div>`;
    return email
      ? `<a href="mailto:${email}" title="${email}" class="techarch-person-link">${inner}</a>`
      : inner;
  };

  const getPerson = (personField) => {
    if (!personField) {
      return { Title: null, EMail: null };
    }
    const person = Array.isArray(personField) ? personField[0] : personField;
    if (!person) {
      return { Title: null, EMail: null };
    }
    return {
      Title: person.Title || null,
      EMail: person.EMail || null
    };
  };

  const formatSectorTitle = (title) => {
    if (title === "Cyber Security" || title === "Corporate Security") {
      return "Corporate Security";
    }
    let display = title || "";
    // Compat: old SharePoint Title → new display name
    if (display.indexOf("IT Infrastructure") === 0) {
      display = "Merck Group IT" + display.slice("IT Infrastructure".length);
    }
    // Break long titles before the parenthetical so columns wrap cleanly
    const paren = display.indexOf(" (");
    if (paren > 0) {
      return display.slice(0, paren) + "<br>" + display.slice(paren + 1);
    }
    return display;
  };

  const findItemByTitle = (byTitle, title) => {
    if (byTitle[title]) {
      return byTitle[title];
    }
    if (title === "Cyber Security" || title === "Corporate Security") {
      return byTitle["Cyber Security"] || byTitle["Corporate Security"] || null;
    }
    if (title === "Merck Group IT (ETP Architecture)") {
      return byTitle["Merck Group IT (ETP Architecture)"] ||
        byTitle["IT Infrastructure (ETP Architecture)"] ||
        null;
    }
    if (title === "Merck Group IT (Network)") {
      return byTitle["Merck Group IT (Network)"] ||
        byTitle["IT Infrastructure (Network)"] ||
        null;
    }
    return null;
  };

  const orderColumns = (items) => {
    const byTitle = {};
    (items || []).forEach((item) => {
      if (item && item.Title) {
        byTitle[item.Title] = item;
      }
    });

    const ordered = [];
    SECTOR_ORDER.forEach((title) => {
      const found = findItemByTitle(byTitle, title);
      if (found) {
        ordered.push(found);
      } else if (title === DATA_AI_TITLE) {
        ordered.push({
          Title: DATA_AI_TITLE,
          TechnologyandArchitectureCoE: null,
          TechnologyandArchitectureCoEBack: null
        });
      }
    });
    return ordered;
  };

  async function loadTable() {
    const container = getContainer();
    if (!container) {
      console.error("Governance table container not found.");
      return;
    }

    try {
      const response = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata" },
        credentials: "same-origin"
      });

      if (!response.ok) {
        throw new Error("SharePoint request failed (HTTP " + response.status + ").");
      }

      const data = await response.json();
      const columns = orderColumns(data.value || []);

      if (!columns.length) {
        showError("No governance data found in the list.");
        return;
      }

      let html = `<table class="techarch-table"><tbody>`;

      html += `<tr><th>Sector</th>`;
      columns.forEach((item) => {
        html += `<td>${formatSectorTitle(item.Title)}</td>`;
      });
      html += `</tr>`;

      html += `<tr><th>Person</th>`;
      columns.forEach((item) => {
        const primary = getPerson(item.TechnologyandArchitectureCoE);
        html += `<td>${formatName(primary.Title, primary.EMail)}</td>`;
      });
      html += `</tr>`;

      html += `<tr><th>Backup</th>`;
      columns.forEach((item) => {
        const backup = getPerson(item.TechnologyandArchitectureCoEBack);
        html += `<td>${formatName(backup.Title, backup.EMail)}</td>`;
      });
      html += `</tr>`;

      html += `</tbody></table>`;
      container.innerHTML = html;
    } catch (err) {
      console.error("Error fetching Technology & Architecture CoE items:", err);
      showError(
        err && err.message
          ? err.message
          : "Error loading governance table data."
      );
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", loadTable);
  } else {
    loadTable();
  }
})();
