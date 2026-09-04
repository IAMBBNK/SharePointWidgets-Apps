/* eslint-disable no-undef */
(function () {
  "use strict";

  var SITE_URL = "https://mdigital.sharepoint.com/sites/ITOTCommunityHub";
  var LIST_NAME = "OT SI - Survey Reponses";
  var LOOKUP_LIST_NAME = "OT SI - Survey Dropdowns";
  var LOOKUP_TYPE_FIELD = "OptionType";
  var OPTION_TYPE_ROLE = "Role";

  /*
   * Legal request ID — shown in the survey intro (top of questionnaire).
   * Keep "LE00XXX" for the initial Legal upload; after approval, replace with
   * the real ID from Legal and re-upload this file to SharePoint.
   */
  var SURVEY_REQUEST_ID = "LE00XXX";
  var SURVEY_CONTACT_NAME = "[Contact person]";
  var SURVEY_CONTACT_EMAIL = "[contact email]";
  var SURVEY_CONTROLLER_NAME =
    "[insert name of the Merck subsidiary and address]";
  var SURVEY_OBJECTION_EMAIL = "[insert email address of survey team]";
  var SURVEY_RETENTION_TEXT =
    "Personal data provided in this survey (for example your role and, if you consent, your email for follow-up) will be retained only as long as we consider working with you on this topic, or until you object, unless overriding legitimate grounds apply under Article 21 (1) GDPR.";
  var SURVEY_PURPOSE_TEXT =
    "This survey collects feedback on how well OT System Integration milestones were achieved, so we can identify strengths, improvement areas, and priorities for future integrations across Merck sites.";
  var PRIVACY_STATEMENT_URL =
    "https://www.merckgroup.com/en/privacy-statement.html";

  var MILESTONES = [
    { id: "M1", name: "Service Design" },
    { id: "M2", name: "Demand Submission" },
    { id: "M3", name: "Solution Design" },
    { id: "M4", name: "Solution Deployment" },
    { id: "M5", name: "Solution Testing" },
    { id: "M6", name: "Handover and Hypercare" }
  ];

  var MILESTONE_IDS = MILESTONES.map(function (m) {
    return m.id;
  });

  var FIELDS = {
    ContactEmail: "ContactEmail",
    ContactRole: "ContactRole",
    ConsentFollowUp: "ConsentFollowUp",
    MilestoneRank: "MilestoneRank",
    BestMilestone: "BestMilestone",
    WorstMilestone: "WorstMilestone",
    BestPoint1: "BestPoint1",
    WorstPoint1: "WorstPoint1",
    AdditionalFeedback: "AdditionalFeedback"
  };

  var slotOrder = [];
  var poolMilestones = [];
  var extraFeedbackEntries = [];
  var extraFeedbackIdCounter = 0;
  var editingExtraFeedbackId = null;
  var listItemEntityType = null;
  var dragPayload = null;

  var RANK_LABELS = [
    "#1 (Best)",
    "#2",
    "#3",
    "#4",
    "#5",
    "#6 (Worst)"
  ];

  var SLOT_DISPLAY_ORDER = [0, 3, 1, 4, 2, 5];

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function getMilestoneLabel(id) {
    for (var i = 0; i < MILESTONES.length; i++) {
      if (MILESTONES[i].id === id) {
        return MILESTONES[i].id + ": " + MILESTONES[i].name;
      }
    }
    return id;
  }

  function getMilestoneIndex(id) {
    return MILESTONE_IDS.indexOf(id);
  }

  function showError(message) {
    var $error = $("#errorContainer");
    $error.html(
      '<div class="error-banner" role="alert">' + escapeHtml(message) + "</div>"
    );
    $("#successContainer").addClass("is-hidden");

    if ($error.length && $error[0].scrollIntoView) {
      $error[0].scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
  }

  function clearError() {
    $("#errorContainer").empty();
  }

  function showSuccess(message) {
    $("#successContainer")
      .html('<div class="success-banner" role="status">' + escapeHtml(message) + "</div>")
      .removeClass("is-hidden");
    clearError();
  }

  function getBestMilestone() {
    return slotOrder[0] ? getMilestoneLabel(slotOrder[0]) : "";
  }

  function getWorstMilestone() {
    return slotOrder[5] ? getMilestoneLabel(slotOrder[5]) : "";
  }

  function createMilestoneChip(id, sourceType, slotIndex) {
    return (
      '<div class="milestone-chip" draggable="true" data-id="' +
      escapeHtml(id) +
      '" data-source="' +
      escapeHtml(sourceType) +
      '" data-slot="' +
      (slotIndex == null ? "" : String(slotIndex)) +
      '">' +
      '<span class="milestone-drag-handle" aria-hidden="true">⠿</span>' +
      '<span class="milestone-label">' +
      escapeHtml(getMilestoneLabel(id)) +
      "</span>" +
      "</div>"
    );
  }

  function renderRankBoard() {
    var $board = $("#milestoneRankBoard");
    $board.empty();

    SLOT_DISPLAY_ORDER.forEach(function (index) {
      var milestoneId = slotOrder[index];
      var labelClass = "rank-slot-label";
      if (index === 0) {
        labelClass += " rank-slot-label--best";
      } else if (index === 5) {
        labelClass += " rank-slot-label--worst";
      }

      var content = milestoneId
        ? createMilestoneChip(milestoneId, "slot", index)
        : '<span class="rank-slot-placeholder">Drop milestone here</span>';

      var $slot = $(
        '<div class="rank-slot' +
          (milestoneId ? " is-filled" : "") +
          '" data-slot="' +
          index +
          '">' +
          '<span class="' +
          labelClass +
          '">' +
          escapeHtml(RANK_LABELS[index]) +
          "</span>" +
          '<div class="rank-slot-dropzone">' +
          content +
          "</div>" +
          "</div>"
      );

      $board.append($slot);
    });

    updateMilestoneLabels();
  }

  function renderPool() {
    var $pool = $("#milestonePool");
    $pool.empty();

    if (!poolMilestones.length) {
      $pool.append(
        '<li class="milestone-pool-empty" aria-hidden="true">All milestones ranked</li>'
      );
      return;
    }

    poolMilestones.forEach(function (id) {
      $pool.append(
        $("<li></li>").html(createMilestoneChip(id, "pool", null))
      );
    });
  }

  function renderMilestoneBoard() {
    renderRankBoard();
    renderPool();
    bindDragEvents();
  }

  function updateMilestoneLabels() {
    var bestId = slotOrder[0];
    var worstId = slotOrder[5];
    var best = bestId ? getMilestoneLabel(bestId) : "";
    var worst = worstId ? getMilestoneLabel(worstId) : "";

    $("#bestMilestoneChip").text("Best: " + (best || "—"));
    $("#worstMilestoneChip").text("Worst: " + (worst || "—"));
    $("#bestSectionMilestone").text(best ? "(" + best + ")" : "");
    $("#worstSectionMilestone").text(worst ? "(" + worst + ")" : "");
  }

  function removeFromPool(milestoneId) {
    poolMilestones = poolMilestones.filter(function (id) {
      return id !== milestoneId;
    });
  }

  function addToPool(milestoneId) {
    if (poolMilestones.indexOf(milestoneId) === -1) {
      poolMilestones.push(milestoneId);
      poolMilestones.sort(function (a, b) {
        return getMilestoneIndex(a) - getMilestoneIndex(b);
      });
    }
  }

  function placeInSlot(milestoneId, targetSlot, sourceType, sourceSlot) {
    var displaced = slotOrder[targetSlot];

    if (sourceType === "pool") {
      removeFromPool(milestoneId);
      if (displaced) {
        addToPool(displaced);
      }
      slotOrder[targetSlot] = milestoneId;
      return;
    }

    if (sourceType === "slot" && sourceSlot !== null && sourceSlot !== targetSlot) {
      if (displaced) {
        slotOrder[sourceSlot] = displaced;
        slotOrder[targetSlot] = milestoneId;
      } else {
        slotOrder[sourceSlot] = null;
        slotOrder[targetSlot] = milestoneId;
      }
    }
  }

  function returnToPool(milestoneId, sourceSlot) {
    if (sourceSlot !== null && sourceSlot >= 0) {
      slotOrder[sourceSlot] = null;
    }
    addToPool(milestoneId);
  }

  function readDragPayload($chip) {
    var sourceSlot = $chip.attr("data-slot");
    return {
      milestoneId: $chip.attr("data-id"),
      sourceType: $chip.attr("data-source"),
      sourceSlot: sourceSlot === "" ? null : parseInt(sourceSlot, 10)
    };
  }

  function bindDragEvents() {
    var $chips = $(".milestone-chip");
    var $slots = $(".rank-slot");
    var $pool = $("#milestonePool");

    $chips.off("dragstart dragend");
    $slots.off("dragover dragleave drop");
    $pool.off("dragover dragleave drop");

    $chips.on("dragstart", function (e) {
      dragPayload = readDragPayload($(this));
      $(this).addClass("is-dragging");
      if (e.originalEvent && e.originalEvent.dataTransfer) {
        e.originalEvent.dataTransfer.effectAllowed = "move";
        e.originalEvent.dataTransfer.setData(
          "text/plain",
          dragPayload.milestoneId
        );
      }
    });

    $chips.on("dragend", function () {
      dragPayload = null;
      $(".milestone-chip").removeClass("is-dragging");
      $(".rank-slot, .milestone-pool").removeClass("is-drag-over");
    });

    $slots.on("dragover", function (e) {
      e.preventDefault();
      if (e.originalEvent && e.originalEvent.dataTransfer) {
        e.originalEvent.dataTransfer.dropEffect = "move";
      }
      $(this).addClass("is-drag-over");
    });

    $slots.on("dragleave", function () {
      $(this).removeClass("is-drag-over");
    });

    $slots.on("drop", function (e) {
      e.preventDefault();
      var targetSlot = parseInt($(this).attr("data-slot"), 10);
      $(".rank-slot, .milestone-pool").removeClass("is-drag-over");

      if (!dragPayload || !dragPayload.milestoneId) {
        return;
      }

      if (
        dragPayload.sourceType === "slot" &&
        dragPayload.sourceSlot === targetSlot
      ) {
        return;
      }

      placeInSlot(
        dragPayload.milestoneId,
        targetSlot,
        dragPayload.sourceType,
        dragPayload.sourceSlot
      );
      dragPayload = null;
      renderMilestoneBoard();
    });

    $pool.on("dragover", function (e) {
      e.preventDefault();
      if (e.originalEvent && e.originalEvent.dataTransfer) {
        e.originalEvent.dataTransfer.dropEffect = "move";
      }
      $(this).addClass("is-drag-over");
    });

    $pool.on("dragleave", function () {
      $(this).removeClass("is-drag-over");
    });

    $pool.on("drop", function (e) {
      e.preventDefault();
      $(".rank-slot, .milestone-pool").removeClass("is-drag-over");

      if (!dragPayload || dragPayload.sourceType !== "slot") {
        return;
      }

      returnToPool(dragPayload.milestoneId, dragPayload.sourceSlot);
      dragPayload = null;
      renderMilestoneBoard();
    });
  }

  function clearFieldErrors() {
    $(".field-group textarea, .field-group input, .field-group select").removeClass(
      "is-invalid"
    );
    $(".field-error").attr("hidden", true).text("");
  }

  function setFieldError(fieldId, message) {
    $("#" + fieldId).addClass("is-invalid");
    $("#" + fieldId + "Error").removeAttr("hidden").text(message);
  }

  function populateMilestoneSelect($select, placeholder) {
    $select.empty();
    $select.append($("<option></option>").attr("value", "").text(placeholder));
    MILESTONES.forEach(function (milestone) {
      $select.append(
        $("<option></option>")
          .attr("value", milestone.id)
          .text(getMilestoneLabel(milestone.id))
      );
    });
  }

  function clearExtraFeedbackComposerErrors() {
    $("#extraFeedbackMilestone, #extraFeedbackText").removeClass("is-invalid");
    $("#extraFeedbackMilestoneError, #extraFeedbackTextError")
      .attr("hidden", true)
      .text("");
  }

  function resetExtraFeedbackComposer() {
    clearExtraFeedbackComposerErrors();
    $("#extraFeedbackMilestone").val("");
    $("#extraFeedbackText").val("");
    editingExtraFeedbackId = null;
    $("#saveExtraFeedbackBtn").text("Add feedback");
  }

  function showExtraFeedbackComposer(entry) {
    clearExtraFeedbackComposerErrors();
    editingExtraFeedbackId = entry ? entry.id : null;

    if (entry) {
      $("#extraFeedbackMilestone").val(entry.milestone);
      $("#extraFeedbackText").val(entry.text);
      $("#saveExtraFeedbackBtn").text("Save changes");
    } else {
      $("#extraFeedbackMilestone").val("");
      $("#extraFeedbackText").val("");
      $("#saveExtraFeedbackBtn").text("Add feedback");
    }

    $("#extraFeedbackComposer").removeClass("is-hidden");
    $("#addExtraFeedbackBtn").addClass("is-hidden");
    renderExtraFeedbackEntries();
  }

  function hideExtraFeedbackComposer() {
    $("#extraFeedbackComposer").addClass("is-hidden");
    $("#addExtraFeedbackBtn").removeClass("is-hidden");
    resetExtraFeedbackComposer();
    renderExtraFeedbackEntries();
    $("#extraFeedbackComposer").insertBefore($("#addExtraFeedbackBtn"));
  }

  function renderExtraFeedbackEntries() {
    var $container = $("#extraFeedbackEntries");
    var $composer = $("#extraFeedbackComposer");
    var $addBtn = $("#addExtraFeedbackBtn");

    $composer.detach();
    $container.empty();

    extraFeedbackEntries.forEach(function (entry) {
      if (entry.id === editingExtraFeedbackId) {
        $container.append($composer);
        return;
      }

      var $entry = $(
        '<article class="extra-feedback-entry" data-entry-id="' +
          entry.id +
          '">' +
          '<span class="extra-feedback-entry-milestone">' +
          escapeHtml(getMilestoneLabel(entry.milestone)) +
          "</span>" +
          '<p class="extra-feedback-entry-text"></p>' +
          '<div class="extra-feedback-entry-actions">' +
          '<button type="button" class="extra-feedback-edit-btn" aria-label="Edit feedback for ' +
          escapeHtml(getMilestoneLabel(entry.milestone)) +
          '">Edit</button>' +
          '<button type="button" class="extra-feedback-remove-btn" aria-label="Remove feedback for ' +
          escapeHtml(getMilestoneLabel(entry.milestone)) +
          '">×</button>' +
          "</div>" +
          "</article>"
      );

      $entry.find(".extra-feedback-entry-text").text(entry.text);
      $container.append($entry);
    });

    if (editingExtraFeedbackId === null || $composer.hasClass("is-hidden")) {
      $composer.insertBefore($addBtn);
    }
  }

  function validateExtraFeedbackComposer() {
    clearExtraFeedbackComposerErrors();
    var valid = true;
    var milestone = $("#extraFeedbackMilestone").val();
    var text = $.trim($("#extraFeedbackText").val());

    if (!milestone) {
      setFieldError("extraFeedbackMilestone", "Please select a milestone.");
      valid = false;
    }

    if (!text) {
      setFieldError("extraFeedbackText", "Please enter your feedback.");
      valid = false;
    }

    return valid;
  }

  function saveExtraFeedbackEntry() {
    if (!validateExtraFeedbackComposer()) {
      return;
    }

    var milestone = $("#extraFeedbackMilestone").val();
    var text = $.trim($("#extraFeedbackText").val());

    if (editingExtraFeedbackId !== null) {
      extraFeedbackEntries = extraFeedbackEntries.map(function (entry) {
        if (entry.id === editingExtraFeedbackId) {
          return {
            id: entry.id,
            milestone: milestone,
            text: text
          };
        }
        return entry;
      });
    } else {
      extraFeedbackEntries.push({
        id: ++extraFeedbackIdCounter,
        milestone: milestone,
        text: text
      });
    }

    hideExtraFeedbackComposer();
  }

  function editExtraFeedbackEntry(entryId) {
    var entry = null;

    extraFeedbackEntries.forEach(function (item) {
      if (item.id === entryId) {
        entry = item;
      }
    });

    if (!entry) {
      return;
    }

    showExtraFeedbackComposer(entry);
  }

  function removeExtraFeedbackEntry(entryId) {
    if (editingExtraFeedbackId === entryId) {
      hideExtraFeedbackComposer();
    }

    extraFeedbackEntries = extraFeedbackEntries.filter(function (entry) {
      return entry.id !== entryId;
    });
    renderExtraFeedbackEntries();
  }

  function formatAdditionalFeedback() {
    return extraFeedbackEntries
      .map(function (entry) {
        return "[" + getMilestoneLabel(entry.milestone) + "] " + entry.text;
      })
      .join("\n\n");
  }

  function bindExtraFeedbackEvents() {
    $("#extraFeedbackContainer")
      .off("click.extraFeedback")
      .on("click.extraFeedback", "#addExtraFeedbackBtn", function () {
        showExtraFeedbackComposer(null);
      })
      .on("click.extraFeedback", "#saveExtraFeedbackBtn", function () {
        saveExtraFeedbackEntry();
      })
      .on("click.extraFeedback", "#cancelExtraFeedbackBtn", function () {
        hideExtraFeedbackComposer();
      })
      .on("click.extraFeedback", ".extra-feedback-edit-btn", function () {
        var entryId = parseInt(
          $(this).closest(".extra-feedback-entry").attr("data-entry-id"),
          10
        );
        editExtraFeedbackEntry(entryId);
      })
      .on("click.extraFeedback", ".extra-feedback-remove-btn", function () {
        var entryId = parseInt(
          $(this).closest(".extra-feedback-entry").attr("data-entry-id"),
          10
        );
        removeExtraFeedbackEntry(entryId);
      });
  }

  function validateForm() {
    clearFieldErrors();
    var valid = true;
    var textFields = ["bestFeedback", "worstFeedback"];

    var role = $("#contactRole").val();
    if (!role) {
      setFieldError("contactRole", "Please select your role.");
      valid = false;
    }

    textFields.forEach(function (id) {
      var value = $.trim($("#" + id).val());
      if (!value) {
        setFieldError(id, "This field is required.");
        valid = false;
      }
    });

    if (!valid) {
      showError("Please complete all required fields before submitting.");
    } else if (!slotOrder.every(function (id) { return Boolean(id); })) {
      showError("Please rank all six milestones before submitting.");
      valid = false;
    } else if (!$("#extraFeedbackComposer").hasClass("is-hidden")) {
      var milestone = $("#extraFeedbackMilestone").val();
      var text = $.trim($("#extraFeedbackText").val());
      if (milestone || text) {
        showError(
          "Please add or cancel your additional feedback before submitting."
        );
        valid = false;
      }
    }

    return valid;
  }

  function getFormDigest() {
    return $.ajax({
      url: SITE_URL + "/_api/contextinfo",
      type: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
      }
    }).then(function (data) {
      return data.d.GetContextWebInformation.FormDigestValue;
    });
  }

  function getCurrentUser() {
    return $.ajax({
      url: SITE_URL + "/_api/web/currentuser",
      type: "GET",
      headers: {
        Accept: "application/json;odata=verbose"
      }
    }).then(function (data) {
      var email = data.d.Email || "";
      var loginName = data.d.LoginName || "";

      // Fallback: LoginName is often like i:0#.f|membership|user@domain.com
      if (!email && loginName.indexOf("|") !== -1) {
        var parts = loginName.split("|");
        var candidate = parts[parts.length - 1];
        if (candidate.indexOf("@") !== -1) {
          email = candidate;
        }
      }

      return {
        name: data.d.Title || data.d.LoginName || "Unknown user",
        email: email
      };
    });
  }

  function populateSelect($select, titles, placeholder, emptyLabel) {
    $select.empty();

    if (!titles.length) {
      $select.append(
        $("<option></option>").attr("value", "").text(emptyLabel)
      );
      $select.prop("disabled", true);
      return false;
    }

    $select.append($("<option></option>").attr("value", "").text(placeholder));
    titles.forEach(function (title) {
      $select.append($("<option></option>").attr("value", title).text(title));
    });
    $select.prop("disabled", false);
    return true;
  }

  function getOptionType(item) {
    return item[LOOKUP_TYPE_FIELD] || item.OptionType || "";
  }

  function loadLookupOptions() {
    var safeTitle = LOOKUP_LIST_NAME.replace(/'/g, "''");
    var url =
      SITE_URL +
      "/_api/web/lists/getbytitle('" +
      safeTitle +
      "')/items?$select=Title," +
      LOOKUP_TYPE_FIELD +
      "&$orderby=Title asc";

    return $.ajax({
      url: url,
      type: "GET",
      headers: {
        Accept: "application/json;odata=verbose"
      }
    }).then(
      function (data) {
        var items = (data.d && data.d.results) || [];
        var roles = [];

        items.forEach(function (item) {
          var title = item.Title || "";
          if (!title) {
            return;
          }
          var type = getOptionType(item);
          if (type === OPTION_TYPE_ROLE) {
            roles.push(title);
          }
        });

        var rolesOk = populateSelect(
          $("#contactRole"),
          roles,
          "Select your role…",
          "No roles available"
        );

        if (!rolesOk) {
          return $.Deferred()
            .reject(
              new Error(
                "Lookup list '" +
                  LOOKUP_LIST_NAME +
                  "' needs items with OptionType '" +
                  OPTION_TYPE_ROLE +
                  "'."
              )
            )
            .promise();
        }
      },
      function (xhr) {
        console.error("Failed to load lookup options:", xhr.responseText);
        $("#contactRole")
          .empty()
          .append('<option value="">Could not load</option>')
          .prop("disabled", true);
        return $.Deferred()
          .reject(
            new Error(
              "Lookup list '" +
                LOOKUP_LIST_NAME +
                "' was not found. See survey_tool_LIST_SETUP.md."
            )
          )
          .promise();
      }
    );
  }

  function fetchListEntityType() {
    if (listItemEntityType) {
      return $.Deferred().resolve(listItemEntityType).promise();
    }

    var safeTitle = LIST_NAME.replace(/'/g, "''");

    return $.ajax({
      url:
        SITE_URL +
        "/_api/web/lists/getbytitle('" +
        safeTitle +
        "')?$select=ListItemEntityTypeFullName",
      type: "GET",
      headers: {
        Accept: "application/json;odata=verbose"
      }
    })
      .then(function (data) {
        listItemEntityType = data.d.ListItemEntityTypeFullName;
        return listItemEntityType;
      })
      .fail(function (xhr) {
        console.error("Failed to load list entity type:", xhr.responseText);
        throw new Error(
          "Survey list '" + LIST_NAME + "' was not found. See survey_tool_LIST_SETUP.md."
        );
      });
  }

  function buildPayload(user) {
    var now = new Date();
    var allowFollowUp = $("#consentFollowUp").is(":checked");
    var title =
      "Survey – " +
      user.name +
      " – " +
      now.toISOString().slice(0, 10);

    var payload = {
      __metadata: { type: listItemEntityType },
      Title: title
    };

    payload[FIELDS.ContactEmail] = allowFollowUp ? user.email || "" : "";
    payload[FIELDS.ConsentFollowUp] = allowFollowUp ? "Yes" : "No";
    payload[FIELDS.ContactRole] = $("#contactRole").val();
    payload[FIELDS.MilestoneRank] = slotOrder
      .map(function (id) {
        return id ? getMilestoneLabel(id) : "";
      })
      .join(";");
    payload[FIELDS.BestMilestone] = getBestMilestone();
    payload[FIELDS.WorstMilestone] = getWorstMilestone();
    payload[FIELDS.BestPoint1] = $.trim($("#bestFeedback").val());
    payload[FIELDS.WorstPoint1] = $.trim($("#worstFeedback").val());
    payload[FIELDS.AdditionalFeedback] = formatAdditionalFeedback();

    return payload;
  }

  function submitToSharePoint(digest, payload) {
    var safeTitle = LIST_NAME.replace(/'/g, "''");

    return $.ajax({
      url: SITE_URL + "/_api/web/lists/getbytitle('" + safeTitle + "')/items",
      type: "POST",
      data: JSON.stringify(payload),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      }
    });
  }

  function handleSubmit(e) {
    e.preventDefault();
    clearError();

    if (!validateForm()) {
      return;
    }

    var $btn = $("#submitSurvey");
    $btn.prop("disabled", true).text("Submitting…");

    $.when(getFormDigest(), getCurrentUser(), fetchListEntityType())
      .then(function (digest, user) {
        var allowFollowUp = $("#consentFollowUp").is(":checked");
        if (allowFollowUp && !user.email) {
          return $.Deferred()
            .reject({
              message:
                "Could not determine your signed-in email for follow-up. Uncheck the follow-up option or refresh and try again."
            })
            .promise();
        }
        var payload = buildPayload(user);
        return submitToSharePoint(digest, payload);
      })
      .done(function () {
        $("#surveyForm").addClass("is-hidden");
        showSuccess(
          "Thank you — your feedback has been submitted successfully."
        );
      })
      .fail(function (xhr) {
        var message = "Failed to submit survey. Please try again.";

        if (xhr && xhr.status === 403) {
          message =
            "You do not have permission to submit. Contact the site owner.";
        } else if (xhr && xhr.status === 400) {
          console.error("Survey submit 400:", xhr.responseText);
          message =
            "Submission failed — check that the SharePoint list columns match survey_tool_LIST_SETUP.md.";
        } else if (xhr && xhr.message) {
          message = xhr.message;
        }

        showError(message);
        $btn.prop("disabled", false).text("Submit feedback");
      });
  }

  function applyIntroConfig() {
    $("#surveyRequestId").text(SURVEY_REQUEST_ID || "LE00XXX");
    $("#surveyContactName").text(SURVEY_CONTACT_NAME || "[Contact person]");
    $("#surveyPurposeText").text(SURVEY_PURPOSE_TEXT);
    $("#surveyRetentionText").text(SURVEY_RETENTION_TEXT);
    $("#surveyControllerName").text(
      SURVEY_CONTROLLER_NAME ||
        "[insert name of the Merck subsidiary and address]"
    );

    var contactEmail = SURVEY_CONTACT_EMAIL || "";
    var $contactLink = $("#surveyContactEmailLink");
    if (contactEmail && contactEmail.indexOf("@") !== -1) {
      $contactLink.attr("href", "mailto:" + contactEmail).text(contactEmail);
    } else {
      $contactLink.attr("href", "#").text(contactEmail || "[contact email]");
    }

    var objectionEmail = SURVEY_OBJECTION_EMAIL || "";
    var $objectionLink = $("#surveyObjectionEmailLink");
    if (objectionEmail && objectionEmail.indexOf("@") !== -1) {
      $objectionLink
        .attr("href", "mailto:" + objectionEmail)
        .text(objectionEmail);
    } else {
      $objectionLink
        .attr("href", "#")
        .text(objectionEmail || "[insert email address of survey team]");
    }

    $("#surveyPrivacyStatementLink").attr(
      "href",
      PRIVACY_STATEMENT_URL ||
        "https://www.merckgroup.com/en/privacy-statement.html"
    );
  }

  function init() {
    applyIntroConfig();
    slotOrder = [null, null, null, null, null, null];
    poolMilestones = MILESTONE_IDS.slice();
    extraFeedbackEntries = [];
    extraFeedbackIdCounter = 0;
    editingExtraFeedbackId = null;
    populateMilestoneSelect($("#extraFeedbackMilestone"), "Select milestone…");
    renderExtraFeedbackEntries();
    bindExtraFeedbackEvents();
    renderMilestoneBoard();

    $.when(fetchListEntityType(), loadLookupOptions())
      .fail(function (err) {
        var message =
          (err && err.message) ||
          (err && err.responseText) ||
          "Failed to initialize survey.";
        showError(String(message));
      });

    $("#surveyForm").on("submit", handleSubmit);
  }

  $(document).ready(init);
})();
