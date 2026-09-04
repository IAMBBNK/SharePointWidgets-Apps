// Flip button handler: flip all primary/backup cells
document.getElementById("flipButton").addEventListener("click", () => {
    const inners = document.querySelectorAll(".coe-flip-inner");
    if (!inners.length) return;
    inners.forEach((card) => card.classList.toggle("flipped"));
    const flipped = inners[0].classList.contains("flipped");
    document.getElementById("flipButton").textContent = flipped
      ? "Show Primary"
      : "Show Backups";
  });
  
  (async function () {
    const siteUrl = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
    const listName = "IT/OT Governance Stakeholders";
    // SharePoint Titles in display order (display renames applied when rendering)
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
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,` +
      `CyberSecurityCoE/Title,CyberSecurityCoE/EMail,` +
      `CyberSecurityCoEBackup/Title,CyberSecurityCoEBackup/EMail,` +
      `PortfolioManagementCoE/Title,PortfolioManagementCoE/EMail,` +
      `PortfolioManagementCoEBackup/Title,PortfolioManagementCoEBackup/EMail,` +
      `ServiceManagementCoE/Title,ServiceManagementCoE/EMail,` +
      `ServiceManagementCoEBackup/Title,ServiceManagementCoEBackup/EMail,` +
      `TechnologyandArchitectureCoE/Title,TechnologyandArchitectureCoE/EMail,` +
      `TechnologyandArchitectureCoEBack/Title,TechnologyandArchitectureCoEBack/EMail,` +
      `IT_x002f_OTCoreTeam/Title,IT_x002f_OTCoreTeam/EMail,` +
      `ITOTCoreTeamBackup/Title,ITOTCoreTeamBackup/EMail` +
      `&$expand=` +
      `CyberSecurityCoE,` +
      `CyberSecurityCoEBackup,` +
      `PortfolioManagementCoE,` +
      `PortfolioManagementCoEBackup,` +
      `ServiceManagementCoE,` +
      `ServiceManagementCoEBackup,` +
      `TechnologyandArchitectureCoE,` +
      `TechnologyandArchitectureCoEBack,` +
      `IT_x002f_OTCoreTeam,` +
      `ITOTCoreTeamBackup`;
  
    // Build SharePoint user photo URL from email
    const getPhotoUrlFromEmail = (email) =>
      email
        ? `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(email)}`
        : null;
  
    // Create profile picture element (no default circle)
    const createProfilePic = (email) => {
      const photoUrl = getPhotoUrlFromEmail(email);
      if (!photoUrl) {
        return "";
      }
      return `
        <div class="profile-pic-wrapper">
          <img src="${photoUrl}" alt="Profile" class="profile-pic"
               onerror="this.style.display='none';">
        </div>
      `;
    };
  
    // helper: insert <br> between first and remaining names, make email hyperlink, add picture
    const formatName = (name, email, emptyText = "TBD") => {
      if (!name) {
        return `
          <div class="person-wrapper">
            <div class="name-text">${emptyText}</div>
          </div>`;
      }
      const parts = name.split(" ");
      const displayName =
        parts.length < 2 ? name : parts[0] + "<br>" + parts.slice(1).join(" ");
      const pic = createProfilePic(email);
      const inner = `
        <div class="person-wrapper">
          ${pic}
          <div class="name-text">${displayName}</div>
        </div>`;
      return email
        ? `<a href="mailto:${email}" title="${email}" class="person-link">${inner}</a>`
        : inner;
    };
  
    // helper: get first person from array or single object (for single‑person fields)
    const getPerson = (persons) => {
      if (!persons || !persons.length) return { Title: null, EMail: null };
      const person = Array.isArray(persons) ? persons[0] : persons;
      return { Title: person.Title, EMail: person.EMail };
    };
  
    // helper: build EXACTLY n backups (pad with TBD) based on array
    const buildBackupsArray = (backupsArray, needed) => {
      const arr = Array.isArray(backupsArray)
        ? backupsArray
        : backupsArray
          ? [backupsArray]
          : [];
      const result = [];
      for (let i = 0; i < needed && i < arr.length; i++) {
        result.push({ Title: arr[i].Title, EMail: arr[i].EMail });
      }
      while (result.length < needed) {
        result.push({ Title: null, EMail: null });
      }
      if (result.length === 0 && needed === 0) {
        result.push({ Title: null, EMail: null });
      }
      return result;
    };

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

    const orderRows = (items) => {
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

    const formatSectorTitle = (title) =>
      SECTOR_DISPLAY_NAMES[title] || title || "";
  
    try {
      const response = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata" },
      });
      const data = await response.json();
  
      const headerHtml = `
        <thead>
          <tr>
            <th>Sector</th>
            <th>
              Group IT/OT Core Team<br />
              <span class="members">Members: Stream Leads</span>
            </th>
            <th>
              IT/OT Tech. and Architecture CoE<br />
              <span class="members">Members: Lead Experts</span>
            </th>
            <th>
              IT/OT Service Management CoE<br />
              <span class="members">Members: Lead Experts</span>
            </th>
            <th>
              IT/OT Portfolio Management CoE<br />
              <span class="members">Members: Lead Experts</span>
            </th>
            <th>
              IT/OT Cyber Security CoE<br />
              <span class="members">Members: Lead Experts</span>
            </th>
          </tr>
        </thead>`;
  
      let html = `<table>${headerHtml}<tbody>`;

      const rows = orderRows(data.value || []);

      rows.forEach((item) => {
        const emptyLabel = "TBD";

        const primary = {
          tech: getPerson(item.TechnologyandArchitectureCoE),
          svc: getPerson(item.ServiceManagementCoE),
          port: getPerson(item.PortfolioManagementCoE),
          cyber: getPerson(item.CyberSecurityCoE),
        };
        const backup = {
          tech: getPerson(item.TechnologyandArchitectureCoEBack),
          svc: getPerson(item.ServiceManagementCoEBackup),
          port: getPerson(item.PortfolioManagementCoEBackup),
          cyber: getPerson(item.CyberSecurityCoEBackup),
        };
  
        // Core team primary (max 3) – multi‑person field array
        const corePrimaryArr = [];
        if (item.IT_x002f_OTCoreTeam && item.IT_x002f_OTCoreTeam.length) {
          item.IT_x002f_OTCoreTeam.forEach((p, idx) => {
            if (idx < 3) {
              corePrimaryArr.push({ Title: p.Title, EMail: p.EMail });
            }
          });
        }
        const primaryCount = corePrimaryArr.length;
  
        // Core team backups: treat ITOTCoreTeamBackup as array, pad to primaryCount
        const coreBackupArr = buildBackupsArray(
          item.ITOTCoreTeamBackup,
          primaryCount,
        );
  
        // Build primary core HTML
        let primaryCoreHtml = "";
        if (corePrimaryArr.length) {
          corePrimaryArr.forEach((p) => {
            primaryCoreHtml += `<div>${formatName(p.Title, p.EMail, emptyLabel)}</div>`;
          });
        } else {
          primaryCoreHtml = `
            <div class="person-wrapper">
              <div class="name-text">${emptyLabel}</div>
            </div>`;
        }

        // Build backup core HTML, same number of blocks as primary (or 1 if no primary)
        let backupCoreHtml = "";
        coreBackupArr.forEach((p) => {
          backupCoreHtml += `<div>${formatName(p.Title, p.EMail, emptyLabel)}</div>`;
        });
  
        // Grid class based on primary count (fallback if no primary)
        const effectiveCount = primaryCount || 1;
        const gridClass =
          effectiveCount === 3
            ? "td-grid-3"
            : effectiveCount === 2
              ? "td-grid-2"
              : "td-grid-1";
  
        const sectorTitle = formatSectorTitle(item.Title);
  
        html += `
          <tr>
            <td>${sectorTitle}</td>
            <td class="coe-cell">
              <div class="coe-flip-card">
                <div class="coe-flip-inner">
                  <div class="coe-front">
                    <div class="${gridClass}">${primaryCoreHtml}</div>
                  </div>
                  <div class="coe-back">
                    <div class="${gridClass}">${backupCoreHtml}</div>
                  </div>
                </div>
              </div>
            </td>
            <td class="coe-cell">
              <div class="coe-flip-card">
                <div class="coe-flip-inner">
                  <div class="coe-front">${formatName(primary.tech.Title, primary.tech.EMail, emptyLabel)}</div>
                  <div class="coe-back">${formatName(backup.tech.Title, backup.tech.EMail, emptyLabel)}</div>
                </div>
              </div>
            </td>
            <td class="coe-cell">
              <div class="coe-flip-card">
                <div class="coe-flip-inner">
                  <div class="coe-front">${formatName(primary.svc.Title, primary.svc.EMail, emptyLabel)}</div>
                  <div class="coe-back">${formatName(backup.svc.Title, backup.svc.EMail, emptyLabel)}</div>
                </div>
              </div>
            </td>
            <td class="coe-cell">
              <div class="coe-flip-card">
                <div class="coe-flip-inner">
                  <div class="coe-front">${formatName(primary.port.Title, primary.port.EMail, emptyLabel)}</div>
                  <div class="coe-back">${formatName(backup.port.Title, backup.port.EMail, emptyLabel)}</div>
                </div>
              </div>
            </td>
            <td class="coe-cell">
              <div class="coe-flip-card">
                <div class="coe-flip-inner">
                  <div class="coe-front">${formatName(primary.cyber.Title, primary.cyber.EMail, emptyLabel)}</div>
                  <div class="coe-back">${formatName(backup.cyber.Title, backup.cyber.EMail, emptyLabel)}</div>
                </div>
              </div>
            </td>
          </tr>`;
      });
  
      html += `</tbody></table>`;
      document.getElementById("listContainer").innerHTML = html;
  
      setTimeout(() => {
        const cards = document.querySelectorAll(".coe-flip-card");
        cards.forEach((card) => {
          const front = card.querySelector(".coe-front");
          const back = card.querySelector(".coe-back");
          if (front && back) {
            const h = Math.max(front.scrollHeight, back.scrollHeight);
            card.style.height = h + "px";
          }
        });
      }, 250);
    } catch (err) {
      console.error("Error fetching list items:", err);
      document.getElementById("listContainer").innerHTML =
        "<p>Error loading data</p>";
    }
  })();
  