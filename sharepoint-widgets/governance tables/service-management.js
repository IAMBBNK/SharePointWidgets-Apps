(function () {
  const siteUrl = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  const listName = "IT/OT Governance Stakeholders";
  const SECTOR_ORDER = [
    "Electronics",
    "Life Science",
    "Healthcare",
    "Site Management",
    "Merck Group IT (ETP Architecture)",
    "IT Application Management (Manufacturing)",
    "Cyber Security"
  ];
  const SECTOR_DISPLAY_NAMES = {
    "Merck Group IT (ETP Architecture)": "Merck Group IT",
    "IT Infrastructure (ETP Architecture)": "Merck Group IT",
    "IT Application Management (Manufacturing)": "IT Application Management",
    "Cyber Security": "Corporate Security",
    "Corporate Security": "Corporate Security"
  };

  const url =
    `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=Title,` +
    `ServiceManagementCoE/Title,ServiceManagementCoE/EMail,` +
    `ServiceManagementCoEBackup/Title,ServiceManagementCoEBackup/EMail` +
    `&$expand=ServiceManagementCoE,ServiceManagementCoEBackup`;

  function getContainer() {
    return document.getElementById("serviceCoeContainer");
  }

  function showError(message) {
    const container = getContainer();
    if (!container) {
      return;
    }
    container.innerHTML =
      '<p class="service-load-error" role="alert">' + message + "</p>";
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
      <div class="service-profile-pic-wrapper">
        <img src="${photoUrl}" alt="Profile" class="service-profile-pic"
             onerror="this.style.display='none';">
      </div>
    `;
  };

  const formatName = (name, email) => {
    if (!name) {
      return `
        <div class="service-person-wrapper">
          <div class="service-name-text">TBD</div>
        </div>`;
    }
    const parts = name.split(" ");
    const displayName =
      parts.length < 2 ? name : parts[0] + "<br>" + parts.slice(1).join(" ");
    const pic = createProfilePic(email);
    const inner = `
      <div class="service-person-wrapper">
        ${pic}
        <div class="service-name-text">${displayName}</div>
      </div>`;
    return email
      ? `<a href="mailto:${email}" title="${email}" class="service-person-link">${inner}</a>`
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

  const formatSectorTitle = (title) => SECTOR_DISPLAY_NAMES[title] || title || "";

  const findItemByTitle = (byTitle, title) => {
    if (byTitle[title]) {
      return byTitle[title];
    }
    if (title === "Cyber Security" || title === "Corporate Security") {
      return byTitle["Cyber Security"] || byTitle["Corporate Security"] || null;
    }
    if (title === "Merck Group IT (ETP Architecture)") {
      return (
        byTitle["Merck Group IT (ETP Architecture)"] ||
        byTitle["IT Infrastructure (ETP Architecture)"] ||
        null
      );
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

      let html = `<table class="service-table"><tbody>`;

      html += `<tr><th>Sector</th>`;
      columns.forEach((item) => {
        html += `<td>${formatSectorTitle(item.Title)}</td>`;
      });
      html += `</tr>`;

      html += `<tr><th>Person</th>`;
      columns.forEach((item) => {
        const primary = getPerson(item.ServiceManagementCoE);
        html += `<td>${formatName(primary.Title, primary.EMail)}</td>`;
      });
      html += `</tr>`;

      html += `<tr><th>Backup</th>`;
      columns.forEach((item) => {
        const backup = getPerson(item.ServiceManagementCoEBackup);
        html += `<td>${formatName(backup.Title, backup.EMail)}</td>`;
      });
      html += `</tr>`;

      html += `</tbody></table>`;
      container.innerHTML = html;
    } catch (err) {
      console.error("Error fetching Service Management CoE items:", err);
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
