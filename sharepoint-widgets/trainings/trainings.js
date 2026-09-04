$(document).ready(function () {
    const siteUrl =
      "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
    const listName = "Trainings";
    // SharePoint "internal name" for your custom column.
    // If your internal name isn't exactly "Order", update this value accordingly.
    const orderColumnInternalName = "Order0";
    // SharePoint internal name for column "IsInternal?"
    const isInternalColumnInternalName = "IsInternal_x003f_";
    let allResources = [];
    let filteredResources = [];
    let selectedTypes = [];
    $("#operatingModelBtn").removeClass("active");
    $("#technologyBtn").removeClass("active");
    $("#cybersecurityBtn").removeClass("active");
    $("#serviceManagementBtn").removeClass("active");
  
    function loadDataFromSharePoint() {
      // Sort smallest -> largest by numeric SharePoint column "Order"
      const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,field_1,field_4,field_7,Link,Curricula,${orderColumnInternalName},${isInternalColumnInternalName}&$orderby=${orderColumnInternalName} asc`;
  
      console.log("Request to API:", apiUrl);
  
      $.ajax({
        url: apiUrl,
        type: "GET",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        success: function (data) {
          console.log("Data successfully loaded:", data);
          if (data.d && data.d.results && data.d.results.length > 0) {
            allResourcesWithOriginalNames = data.d.results.map((item) => ({
              OriginalName: item.Title || "No title",
              Role: extractRoleValues(item.field_1),
              Short_Insights: item.field_4 || "",
              CategoryType: item.field_7 || "",
              Link: extractLinkUrl(item.Link),
              Curricula: isCurriculaTrue(item.Curricula),
              IsInternal: isCurriculaTrue(item[isInternalColumnInternalName]),
              // Keep raw Order; normalization happens in getOrderValue()
              Order: item[orderColumnInternalName] ?? item.Order ?? null,
            }));

            const missingOrder = allResourcesWithOriginalNames.filter(
              (r) =>
                r.Order === null ||
                r.Order === undefined ||
                String(r.Order).trim() === "",
            ).length;
            console.log(
              `Order debug: missing/empty Order for ${missingOrder}/${allResourcesWithOriginalNames.length} items (column internal name: ${orderColumnInternalName})`,
            );
  
            allResources = allResourcesWithOriginalNames.map((item) => ({
              ...item,
              Name: removeCurriculaFromName(item.OriginalName),
            }));
  
            console.log("Processed resources:", allResources);
  
            selectedTypes = [];
  
            $("#operatingModelBtn").removeClass("active");
            $("#technologyBtn").removeClass("active");
            $("#cybersecurityBtn").removeClass("active");
            $("#serviceManagementBtn").removeClass("active");
  
            filterResources();
          } else {
            console.warn(
              "There are no results in the list or the list is empty",
            );
            showError(
              "There are no results in the list or the list is empty.",
            );
            loadFallbackData();
          }
        },
        error: function (error) {
          console.error("Error loading data:", error);
          console.error("Response text:", error.responseText);
          showError(
            `Error loading data: ${error.status} ${error.statusText}. Please check the console.`,
          );
          loadFallbackData();
        },
      });
    }
  
    function extractRoleValues(roleField) {
      if (!roleField) return "";
      if (roleField.results && Array.isArray(roleField.results)) {
        return roleField.results.join(", ");
      }
      if (typeof roleField === "string") {
        return roleField;
      }
      return "";
    }
  
    function extractLinkUrl(linkField) {
      if (!linkField) return "";
      if (typeof linkField === "string") {
        return linkField;
      } else if (linkField.Url) {
        return linkField.Url;
      }
      return "";
    }
  
    let allResourcesWithOriginalNames = [];
  
    function removeCurriculaFromName(name) {
      if (!name) return name;
      return name
        .replace(/\s*curricula\s*/gi, " ")
        .replace(/\s{2,}/g, " ")
        .trim();
    }
  
    function isCurriculaTrue(curriculaValue) {
      if (!curriculaValue) return false;
      if (typeof curriculaValue === "boolean") return curriculaValue;
      if (typeof curriculaValue === "string") {
        return (
          curriculaValue.toLowerCase() === "yes" ||
          curriculaValue === "1" ||
          curriculaValue.toLowerCase() === "true" ||
          curriculaValue.toLowerCase() === "да"
        );
      }
      return false;
    }
  
    function getCategoryClass(categoryType) {
      if (!categoryType) return "";
      switch (categoryType) {
        case "IT/OT Operating Model":
          return "";
        case "IT/OT Technology & Architecture":
          return "technology";
        case "IT/OT Cybersecurity":
          return "cybersecurity";
        case "IT/OT Service Management":
          return "serviceManagement";
        default:
          return "";
      }
    }
  
    function getCategoryColor(categoryType) {
      if (!categoryType) return "#96D7D2";
      switch (categoryType) {
        case "IT/OT Operating Model":
          return "#96D7D2";
        case "IT/OT Technology & Architecture":
          return "#E1C3CD";
        case "IT/OT Cybersecurity":
          return "#B4DC96";
        case "IT/OT Service Management":
          return "#CBBAEF";
        default:
          return "#96D7D2";
      }
    }
  
    function renderResources() {
      $("#resourcesContainer").empty();
  
      if (filteredResources.length === 0) {
        $("#resourcesContainer").html(
          '<div class="no-results">No trainings found matching your criteria</div>',
        );
        return;
      }
  
      function getOrderValue(item) {
        // Treat missing/invalid values as very large so they go last
        const raw = item && item.Order;
        if (raw === null || raw === undefined) return Number.MAX_SAFE_INTEGER;
        if (typeof raw === "number") return Number.isFinite(raw) ? raw : Number.MAX_SAFE_INTEGER;

        const s = String(raw).trim();
        if (!s) return Number.MAX_SAFE_INTEGER;

        // Handle cases like "1,000" or "1.000" coming as text (remove separators)
        // Assumes "Order" is an integer (not a decimal).
        const normalized = s.replace(/\s/g, "").replace(/[.,]/g, "");
        const n = Number(normalized);
        return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
      }

      const sortedResources = [...filteredResources].sort((a, b) => {
        // Primary sort: SharePoint "Order" ascending
        const orderDiff = getOrderValue(a) - getOrderValue(b);
        if (orderDiff !== 0) return orderDiff;

        // Secondary sort (deterministic): prefer a stable human-readable order
        const aKey = (a.OriginalName || a.Name || "").toString();
        const bKey = (b.OriginalName || b.Name || "").toString();
        return aKey.localeCompare(bKey, undefined, { numeric: true, sensitivity: "base" });
      });
  
      // Debug: verify sort order and incoming SharePoint "Order" values
      console.log("=== Trainings Order Debug (renderResources) ===");
      sortedResources.forEach((item, idx) => {
        const displayName = (item.OriginalName || item.Name || "").toString();
        console.log(
          `${idx + 1}. ${displayName} | Order(raw)=${item.Order} | Order(num)=${getOrderValue(item)} | Category=${item.CategoryType}`,
        );
      });
      console.log("=== End Trainings Order Debug ===");

      sortedResources.forEach((item) => {
        const categoryClass = getCategoryClass(item.CategoryType);
        const displayName = item.Name;
        const categoryText = item.CategoryType;
  
        const accordionItem = $(`
              <div class="accordionItem">
                  <div class="accordionContainer ${categoryClass}" style="border-left-color: ${getCategoryColor(item.CategoryType)}">
                      <div class="accordionContent">
                          <div class="accordionMainContent">
                              <div class="accordionText">${displayName}</div>
                              <div class="accordionCategory">${categoryText}</div>
                                  <div class="internalExternalBadge ${item.IsInternal ? "internal" : "external"}">
                                      ${item.IsInternal ? "Internal" : "External"}
                                  </div>
                              <div class="moreInfoContainer">
                                  <span class="arrow">▼</span>
                                  <div class="accordionDropdown">More Information</div>
                              </div>
                              <div class="accordionActions">
                                  <button class="goToButton" data-link="${item.Link}">Go to Training</button>
                              </div>
                          </div>
                      </div>
                  </div>
                  <div class="hiddenText">
                      ${item.Short_Insights}
                  </div>
              </div>
          `);
  
        $("#resourcesContainer").append(accordionItem);
      });
  
      console.log("renderResources called");
      console.log("selectedTypes:", selectedTypes);
      initEventHandlers();
    }
  
    function filterResources() {
      filteredResources = allResources.filter((item) => {
        if (
          selectedTypes.length > 0 &&
          !selectedTypes.includes(item.CategoryType)
        ) {
          return false;
        }
  
        return true;
      });
  
      renderResources();
    }
  
    function initEventHandlers() {
      $(".accordionContainer")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          const accordionItem = $(this).closest(".accordionItem");
          const hiddenText = accordionItem.find(".hiddenText");
          const moreInfoContainer =
            accordionItem.find(".moreInfoContainer");
  
          // Переключаем состояние
          moreInfoContainer.toggleClass("active");
          hiddenText.slideToggle();
        });
  
      $(".goToButton")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          const link = $(this).data("link");
          if (link) {
            window.open(link, "_blank");
          }
        });
  
      $("#operatingModelBtn")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          $(this).toggleClass("active");
  
          const categoryType = "IT/OT Operating Model";
          const index = selectedTypes.indexOf(categoryType);
          if (index === -1) {
            selectedTypes.push(categoryType);
          } else {
            selectedTypes.splice(index, 1);
          }
  
          filterResources();
        });
  
      $("#technologyBtn")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          $(this).toggleClass("active");
  
          const categoryType = "IT/OT Technology & Architecture";
          const index = selectedTypes.indexOf(categoryType);
          if (index === -1) {
            selectedTypes.push(categoryType);
          } else {
            selectedTypes.splice(index, 1);
          }
  
          filterResources();
        });
  
      $("#cybersecurityBtn")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          $(this).toggleClass("active");
  
          const categoryType = "IT/OT Cybersecurity";
          const index = selectedTypes.indexOf(categoryType);
          if (index === -1) {
            selectedTypes.push(categoryType);
          } else {
            selectedTypes.splice(index, 1);
          }
  
          filterResources();
        });
  
      $("#serviceManagementBtn")
        .off("click")
        .on("click", function (e) {
          e.stopPropagation();
          $(this).toggleClass("active");
  
          const categoryType = "IT/OT Service Management";
          const index = selectedTypes.indexOf(categoryType);
          if (index === -1) {
            selectedTypes.push(categoryType);
          } else {
            selectedTypes.splice(index, 1);
          }
  
          filterResources();
        });
  
    }
  
    function showError(message) {
      $(".resourcePortal-container").prepend(`
                      <div class="error-message" style="color: red; padding: 10px; margin-bottom: 15px; border: 1px solid red;">
                          ${message}
                      </div>
                  `);
    }
  
    function loadFallbackData() {
      console.log("Loading test data...");
      const fallbackData = [
        {
          OriginalName: "OT Principles (Test Data) Curricula",
          Name: removeCurriculaFromName(
            "OT Principles (Test Data) Curricula",
          ),
          Role: "OT Portfolio Manager, OT Project Delivery Manager, OT Service Delivery Manager, OT Engineer, OT Solution Architect, OT Cybersecurity Manager",
          Short_Insights:
            "This is TEST DATA. The 'OT Principles' Module is part of the Merck IT/OT Introduction and Expert Training Series.",
          CategoryType: "IT/OT Operating Model",
          Order: 1,
          IsInternal: true,
          Link: "https://performancemanager5.successfactors.eu/sf/learning?destUrl=https://merckkga.plateau.com/learning/user/deeplink_redirect.jsp?linkId%3dITEM_DETAILS%26componentTypeID%3dCOURSE%26revisionDate%3d1724333640000%26fromSF%3dY&company=merckgroup",
          Curricula: true,
        },
        {
          OriginalName: "Regular Training Without Curricula",
          Name: removeCurriculaFromName(
            "Regular Training Without Curricula",
          ),
          Role: "OT Portfolio Manager, OT Project Delivery Manager",
          Short_Insights:
            "This is a regular training without curricula flag.",
          CategoryType: "IT/OT Technology & Architecture",
          Order: 2,
          IsInternal: false,
          Link: "https://example.com",
          Curricula: false,
        },
      ];
  
      allResourcesWithOriginalNames = fallbackData;
      allResources = fallbackData;
  
      selectedTypes = [];
  
      $("#operatingModelBtn").removeClass("active");
      $("#technologyBtn").removeClass("active");
      $("#cybersecurityBtn").removeClass("active");
      $("#serviceManagementBtn").removeClass("active");
  
      filterResources();
    }
  
    loadDataFromSharePoint();
  });