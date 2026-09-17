$(document).ready(async function () {
    const allFilters = {
      sectorFilter: "all",
      unitFilter: []
    }
  
    const getSiteMasterData = async () => {
      const filter = `&$select=ID,Title,Sector,field_2,field_5,field_6,field_7,field_8,field_9,DPIVariant,TargetMaturity_x0028_BPOG_x0029_,ITAssessment,OTAssessment,OTSystemIntegration,DPICoreImplementation,CurrentMaturity_x0028_BPOG_x0029,Lattitude,Longitude,OData__x0023_ofOTSystems,OData__x0023_ofmigratedOTSystems,LegalEntityNumber`
      try {
        const response = await fetch(`https://mdigital.sharepoint.com/sites/ITOTCommunityHub/_api/web/lists/getbytitle('Site Master Data')/items?$top=1000${filter}`, {
          method: "GET",
          headers: {
            "accept": "application/json;odata=verbose"
          }
        });
        const data = await response.json();
        return data.d.results;
      } catch (error) {
        console.log("Error fetching list data", error);
        return dummyData.d.results;
      }
    }
  
    const siteMasterData = await getSiteMasterData();
    const formattedSiteMasterData = siteMasterData.map(item => {
      return {
        ID: item.ID,
        Title: item.Title ? item.Title : "-",
        Sector: item.Sector ? item.Sector : "-",
        SiteName: item.field_2 ? item.field_2 : "-",
        BusinessUnit: item.field_5 ? item.field_5 : "-",
        Location: item.field_8 ? item.field_8 : "-",
        LegalEntityName: item.field_7 ? item.field_7 : "-",
        LegalEntityCode: item.LegalEntityNumber ? item.LegalEntityNumber : "-",
        Services: item.field_9 ? item.field_9 : "-",
        Employees: item.field_6 ? item.field_6 : "-",
        TargetDpi: item.DPIVariant ? item.DPIVariant : "-",
        OtSystemsIntegrated: item.OData__x0023_ofOTSystems ? item.OData__x0023_ofOTSystems : "-",
        ItAssessment: item.ITAssessment ? item.ITAssessment : "-",
        OtAssessment: item.OTAssessment ? item.OTAssessment : "-",
        DpiCoreImplementation: item.DPICoreImplementation ? item.DPICoreImplementation : "-",
        OtSystemIntegration: item.OTSystemIntegration ? item.OTSystemIntegration : "-",
        CurrentMaturity: item.CurrentMaturity_x0028_BPOG_x0029 ? item.CurrentMaturity_x0028_BPOG_x0029 : "-",
        TargetMaturity: item.TargetMaturity_x0028_BPOG_x0029_ ? item.TargetMaturity_x0028_BPOG_x0029_ : "-",
        Lattitude: item.Lattitude,
        Longitude: item.Longitude
      }
    });
  
    console.log(formattedSiteMasterData);
  
    const renderUnitFilter = () => {
      const units = [...(new Set(formattedSiteMasterData.map(item => item.BusinessUnit)))].filter(item => item);
      const unitsHtml = units.map((item, i) => `<div class='panel-item'><input class="filter-option" type="checkbox" id="unit-${i}" name="${item}" value="${item}"><label for="unit-${i}">${item}</label></div>`);
      $("#dynamic-unit").prepend(unitsHtml.join(""));
    }
  
    renderUnitFilter();
  
    // bing api key
    const key = 'AgJMis6__ur7UfvExgEQb3R3EWCwZTTAYvhjRDPtG9ad-JlxHO3WzhzLIc9DHRzw';
  
    const hiddenCityLabels = {
      "version": "1.0",
      "elements": {
        "point": {
          "labelVisible": true,
          "visible": false
        },
        "road": {
          "visible": false,
          "labelVisible": false
        },
      }
    }
  
    const map = new Microsoft.Maps.Map(document.getElementById('map'), {
      credentials: key,
      center: new Microsoft.Maps.Location(51.165691, 10.451526),
      zoom: 2,
      disableScrollWheelZoom: false,
      mapTypeId: Microsoft.Maps.MapTypeId.aerial,
      customMapStyle: hiddenCityLabels,
    });
  
    const tooltip = new Microsoft.Maps.Infobox(map.getCenter(), {
      visible: true,
      showPointer: false,
      showCloseButton: false,
      offset: new Microsoft.Maps.Point(20, -260)
    });
  
    tooltip.setMap(map);
  
    const infobox = new Microsoft.Maps.Infobox(map.getCenter(), {
      visible: false,
    });
  
    infobox.setMap(map);
  
    // Function to clear existing pushpins
    const clearPushpins = () => {
      // Loop through map.entities and remove pushpins
      for (let i = map.entities.getLength() - 1; i >= 0; i--) {
        const entity = map.entities.get(i);
  
        // Check if the entity is a Pushpin (optional if entities contain only pushpins)
        if (entity instanceof Microsoft.Maps.Pushpin) {
          map.entities.removeAt(i);
        }
      }
    };
  
    const renderMap = (data) => {
      clearPushpins();
      try {
        for (const city of data) {
          if (city.Lattitude && city.Longitude) {
            const pin = new Microsoft.Maps.Pushpin(
              {
                latitude: city.Lattitude,
                longitude: city.Longitude,
              },
              {
                icon: createCircle(city.Sector)
              }
            );
            pin.metadata = {
              ID: city.ID,
              Title: city.Title,
              SiteName: city.SiteName,
              BusinessUnit: city.BusinessUnit,
              Location: city.Location,
              Sector: city.Sector,
              LegalEntityName: city.LegalEntityName,
              LegalEntityCode: city.LegalEntityCode,
              Services: city.Services,
              Employees: city.Employees,
              TargetDpi: city.TargetDpi,
              OtSystemsIntegrated: city.OtSystemsIntegrated,
              ItAssessment: city.ItAssessment,
              OtAssessment: city.OtAssessment,
              DpiCoreImplementation: city.DpiCoreImplementation,
              OtSystemIntegration: city.OtSystemIntegration,
              CurrentMaturity: city.CurrentMaturity,
              TargetMaturity: city.TargetMaturity,
            };
            Microsoft.Maps.Events.addHandler(pin, 'click', pushpinClicked);
            map.entities.push(pin);
          }
        }
        console.log(map.entities);
      } catch (error) {
        console.error('Error during geocoding:', error);
      }
    }
  
    const applyFiltersAndRender = () => {
      console.log("triggered");
      let filteredData = formattedSiteMasterData;
      if (allFilters.sectorFilter !== "all") {
        filteredData = filteredData.filter(item => item.Sector === allFilters.sectorFilter);
      }
  
      if (allFilters.unitFilter.length) {
        filteredData = filteredData.filter(item => allFilters.unitFilter.includes(item.BusinessUnit));
      }
  
      renderMap(filteredData);
    }
  
    function closeInfobox() {
      infobox.setOptions({ visible: false });
    }
  
    function createCircle(sector) {
      var c = document.createElement('canvas');
      c.width = 24;
      c.height = 24;
  
      var ctx = c.getContext('2d');
  
      if (sector === "Life Science") {
        ctx.fillStyle = '#ffc832';
      } else if (sector === "Healthcare") {
        ctx.fillStyle = '#eb3c96';
      } else {
        ctx.fillStyle = '#4ca5e9';
      }
  
      ctx.lineWidth = 2;
      ctx.strokeStyle = '#0f69af';
  
  
      //Draw a path in the shape of an arrow.
      ctx.beginPath();
      ctx.arc(c.width * 0.5, c.height * 0.5, 10, 0, 2 * Math.PI);
      ctx.fill();
      ctx.stroke();
  
      //Generate the base64 image URL from the canvas.
      return c.toDataURL();
    }
  
    // Function to smoothly move the map to the pushpin's location
    function smoothPanTo(location) {
      const animationDuration = 1000; // Duration in milliseconds
      const frames = 60; // Number of animation frames
      const interval = animationDuration / frames; // Interval between frames
  
      const startLocation = map.getCenter();
      const latDelta = location.latitude - startLocation.latitude;
      const lonDelta = location.longitude - startLocation.longitude;
  
      let frame = 0;
  
      function animate() {
        frame++;
        const progress = frame / frames;
  
        // Interpolate the latitude and longitude
        const currentLatitude =
          startLocation.latitude + latDelta * easeInOutQuad(progress);
        const currentLongitude =
          startLocation.longitude + lonDelta * easeInOutQuad(progress);
  
        // Set the map view to the interpolated location
        map.setView({
          center: new Microsoft.Maps.Location(currentLatitude, currentLongitude),
        });
  
        if (frame < frames) {
          setTimeout(animate, interval);
        }
      }
  
      animate();
    }
  
    // Ease-in-out function for smooth transitions
    function easeInOutQuad(t) {
      return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
    }
  
    // Function to ensure the tooltip always shows at the center of the map
    function updateTooltipPosition(pushpin) {
      // Set the map's center to the pushpin's location
      const pushpinLocation = pushpin.target.getLocation();
  
      // Smoothly move the map to the pushpin location
      smoothPanTo(pushpinLocation);
  
  
      if (pushpinLocation.latitude > 60) {
        tooltip.setOptions({
          offset: new Microsoft.Maps.Point(20, -400), // Center-aligned tooltip
        });
      } else {
        tooltip.setOptions({
          offset: new Microsoft.Maps.Point(20, -100), // Center-aligned tooltip
        });
      }
    }
  
    function pushpinClicked(e) {
  
      updateTooltipPosition(e);
  
      //Hide the infobox
      infobox.setOptions({ visible: false });
  
      //Make sure the infobox has metadata to display.
      if (e.target.metadata) {
        let borderColor;
        if (e.target.metadata.Sector === "Life Science") {
          borderColor = 'sector-yellow';
        } else if (e.target.metadata.Sector === "Healthcare") {
          borderColor = 'sector-pink';
        } else {
          borderColor = 'sector-blue';
        }
  
        setTimeout(() => {
          //Set the infobox options with the metadata of the pushpin.
          tooltip.setOptions({
            location: e.target.getLocation(),
            htmlContent: `<div class="map-tooltip ${borderColor}">
                                <div class="infobox-close" id="closeInfobox">x</div>
                                <div class="infobox-headline">${e.target.metadata.SiteName} / ${e.target.metadata.BusinessUnit}</div>
                                <div class="map-tooltip-items">
                                  <div class="map-tooltip-block">
                                    <div class="map-tooltip-item">
                                      <div>${e.target.metadata.LegalEntityName}</div>
                                    </div>
                                    <div class="map-tooltip-item">
                                      <div>
                                          ${e.target.metadata.Location}
                                        </div>
                                    </div>
                                  </div>
                                  <div class="map-tooltip-block">
                                    <div class='map-tooltip-item'>
                                      <img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/${(e.target.metadata.ItAssessment === "-" || e.target.metadata.ItAssessment === "N/A") ? "NA" : e.target.metadata.ItAssessment}.svg"><strong>IT Assessment</strong>
                                    </div>
                                    <div class='map-tooltip-item'>
                                      <img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/${(e.target.metadata.OtAssessment === "-" || e.target.metadata.OtAssessment === "N/A") ? "NA" : e.target.metadata.OtAssessment}.svg"><strong>OT Assessment</strong>
                                    </div>
                                    <div class='map-tooltip-item'>
                                      <img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/${(e.target.metadata.DpiCoreImplementation === "-" || e.target.metadata.DpiCoreImplementation === "N/A") ? "NA" : e.target.metadata.DpiCoreImplementation}.svg"><strong>DPI Core Implementation</strong>
                                    </div>
                                    <div class='map-tooltip-item'>
                                      <img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/${(e.target.metadata.OtSystemIntegration === "-" || e.target.metadata.OtSystemIntegration === "N/A") ? "NA" : e.target.metadata.OtSystemIntegration}.svg"><strong>OT System Integration</strong>
                                    </div>
                                  </div>
                                </div>
                                <div class='map-tooltip-items'>
                                  <div class='map-tooltip-block'>
                                    <div class="map-tooltip-item">
                                      <div><strong>Services: </strong>${e.target.metadata.Services}</div>
                                    </div>
                                    <div class="map-tooltip-item">
                                      <div>
                                        <div><strong>Employees: </strong>${e.target.metadata.Employees}</div>
                                      </div>
                                    </div>
                                  </div>
                                  <div class='map-tooltip-block'>
                                    <div class='map-tooltip-item'>
                                      <div><strong>Current Maturity (BPOG):</strong></div>
                                    </div>
                                    <div class='map-tooltip-item'>
                                      <div>${e.target.metadata.CurrentMaturity}</div>
                                    </div>
                                  </div>
                                </div>
                                <div class='map-tooltip-items'>
                                  <div class='map-tooltip-block'>
                                    <div class="map-tooltip-item">
                                      <div>
                                        <div><strong>Target DPI Variant: </strong>${e.target.metadata.TargetDpi}</div>
                                      </div>
                                    </div>
                                    <div class="map-tooltip-item">
                                      <div>
                                        <div><strong>OT Systems integrated: </strong>${e.target.metadata.OtSystemsIntegrated}</div>
                                      </div>
                                    </div>
                                  </div>
                                  <div class='map-tooltip-block'>
                                    <div class='map-tooltip-item'>
                                      <div><strong>Target Maturity (BPOG):</strong></div>
                                    </div>
                                    <div class='map-tooltip-item'>
                                      <div>${e.target.metadata.TargetMaturity}</div>
                                    </div>
                                  </div>
                                </div>
                                <div class='map-tooltip-item'>
                                  <div class='tooltip-legend'>
                                    <div><img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/NA.svg" /> N/A</div>
                                    <div><img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/Planning.svg" /> Planning</div>
                                    <div><img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/Ongoing.svg" /> Ongoing</div>
                                    <div><img src="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/SiteAssets/Visuals/Buttons%20%26%20Icons/Completed.svg" /> Completed</div>
                                  </div>
                                  <div >
                                    <a class="see-all" href="https://mdigital.sharepoint.com/sites/ITOTCommunityHub/Lists/Site%20Master%20Data/DispForm.aspx?ID=${e.target.metadata.ID}">See all details</a>
                                  </div>
                                  </div>
                              </div>`,
            visible: true
          });
  
          const closeButton = document.getElementById('closeInfobox');
          closeButton.addEventListener("click", () => {
            tooltip.setOptions({
              visible: false
            });
          })
        }, 500)
      }
    }
  
    const closeFilterAccordions = () => {
      $(".accordion").removeClass("active");
      $(".panel").css("display", "none");
    }
  
    $(document).on("click", ".accordion", function (e) {
      e.stopImmediatePropagation();
      if ($(this).hasClass("active")) {
        console.log("Closing this accordion");
        closeFilterAccordions();
      } else {
        closeFilterAccordions();
        console.log("opening this accordion");
        $(this).addClass("active");
        const panel = $(this).next();
        if (panel.css("display") === "none") {
          panel.css("display", "flex");
        } else {
          panel.css("display", "none");
        }
      }
    });
  
    $(document).on("click", ".sector-filter", function () {
      $(".sector-filter").removeClass("active-filter");
      $(this).addClass("active-filter");
      allFilters.sectorFilter = $(this).data("id");
      console.log(allFilters);
      applyFiltersAndRender();
    });
  
    $(".apply-filter").on("click", function () {
      allFilters.unitFilter = $('.filter-option:checkbox:checked').map(function () {
        return this.value;
      }).get();
      console.log(allFilters);
      applyFiltersAndRender();
      closeFilterAccordions();
    });
  
  
    applyFiltersAndRender();
  });
  