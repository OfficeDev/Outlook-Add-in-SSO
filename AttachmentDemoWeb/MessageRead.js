// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
(function () {
    "use strict";

    var messageBanner;
    var overlay;
    var spinner;
    var authenticator;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // For auth helper
            if (OfficeHelpers.Authenticator.isAuthDialog()) return;

            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric['MessageBanner'](element);

            var overlayComponent = document.querySelector(".ms-Overlay");
            // Override click so user can't dismiss overlay
            overlayComponent.addEventListener("click", function (e) {
                e.preventDefault();
                e.stopImmediatePropagation();
            });
            overlay = new window.fabric["Overlay"](overlayComponent);

            var spinnerElement = document.querySelector(".ms-Spinner");
            spinner = new window.fabric["Spinner"](spinnerElement);
            spinner.stop();

            $("#save-selected").on("click", function () {
                saveSelectedAttachments();
            });

            authenticator = new OfficeHelpers.Authenticator();
            authenticator.endpoints.registerMicrosoftAuth(authConfig.clientId, {
                redirectUrl: authConfig.redirectUrl,
                scope: authConfig.scopes
            });

            initializePane();
        });
    };

    // Initialize the pane
    function initializePane() {
        // Check if item has any attachments
        if (Office.context.mailbox.item.attachments.length > 0)
        {
            $("#main-content").show();
            populateList();
            initListItems();
        } else {
            $("#no-attachments").show();
        }
    }

    function populateList() {
        // Get the list
        var attachmentList = $(".ms-List");

        Office.context.mailbox.item.attachments.forEach(function (attachment) {
            var listItem = $("<li>")
                .addClass("ms-ListItem")
                .addClass("is-selectable")
                .attr("tabindex", "0")
                .appendTo(attachmentList);

            $("<div>")
                .addClass("attachment-id")
                .text(attachment.id)
                .appendTo(listItem);

            $("<span>")
                .addClass("ms-ListItem-secondaryText")
                .text(attachment.name)
                .appendTo(listItem);

            var contentType = attachment.attachmentType === "file" ?
                attachment.contentType : "Outlook item";
            if (contentType === null || contentType.length === 0) {
                contentType = "unknown";
            }

            $("<span>")
                .addClass("ms-ListItem-secondaryText")
                .text(contentType)
                .appendTo(listItem);

            $("<span>")
                .addClass("ms-ListItem-metaText")
                .text(generateFileSizeString(attachment.size))
                .appendTo(listItem);

            $("<div>")
                .addClass("ms-ListItem-selectionTarget")
                .appendTo(listItem);

            var actions = $("<div>")
                .addClass("ms-ListItem-actions")
                .appendTo(listItem);

            var saveAction = $("<div>")
                .addClass("ms-ListItem-action")
                .appendTo(actions);

            $("<i>")
                .addClass("ms-Icon")
                .addClass("ms-Icon--Save")
                .appendTo(saveAction);
        });
    }

    function generateFileSizeString(size) {
        if (size > 1048576) {
            var megString = Math.round(size / 1048576).toString() + " MB";
            return megString;
        }

        if (size > 1024) {
            var kbString = Math.round(size / 1024).toString() + " KB";
            return kbString;
        }

        else return size.toString() + " B";
    }

    function initListItems() {
        var ListElements = document.querySelectorAll(".ms-List");
        for (var i = 0; i < ListElements.length; i++) {
            new fabric['List'](ListElements[i]);
        }

        $(".ms-ListItem-selectionTarget").on("click", function () {
            var disableButton = $(".is-selected").length === 0;
            $("#save-selected").prop("disabled", disableButton);
        });

        $(".ms-ListItem-action").on("click", function () {
            var attachmentId = $(this).closest(".ms-ListItem").children(".attachment-id").text();
            saveAttachments([getRestId(attachmentId)]);
        });
    }

    function saveSelectedAttachments() {
        var selectedItems = $(".is-selected");
        if (selectedItems.length > 0) {
            var attachmentIds = [];

            for (var i = 0; i < selectedItems.length; i++) {
                var id = $(selectedItems[i]).children(".attachment-id").text();
                attachmentIds.push(getRestId(id));
            }

            saveAttachments(attachmentIds);
        }
    }

    function saveAttachments(attachmentIds) {
        showSpinner();

        // First attempt to get an SSO token
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync(function (result) {
                if (result.status === "succeeded") {
                    // No need to prompt user, use this token to call Web API
                    saveAttachmentsWithSSO(result.value, attachmentIds);
                } else {
                    // Could not get SSO token, proceed with authentication prompt
                    saveAttachmentsWithPrompt(attachmentIds);
                }
            });
        }
    }

    function saveAttachmentsWithSSO(accessToken, attachmentIds) {
        var saveAttachmentsRequest = {
            attachmentIds: attachmentIds,
            messageId: getRestId(Office.context.mailbox.item.itemId)
        };

        $.ajax({
            type: "POST",
            url: "/api/SaveAttachments",
            headers: {
                "Authorization": "Bearer " + accessToken
            },
            data: JSON.stringify(saveAttachmentsRequest),
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            showNotification("Success", "Attachments saved");
        }).fail(function (error) {
            showNotification("Error saving attachments", error.status);
        }).always(function () {
            hideSpinner();
        });
    }

    function saveAttachmentsWithPrompt(attachmentIds) {
        authenticator
            .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
            .then(function (token) {
                // Get callback token, which grants read access to the current message
                // via the Outlook API
                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                    if (result.status === "succeeded") {
                        var saveAttachmentsRequest = {
                            attachmentIds: attachmentIds,
                            messageId: getRestId(Office.context.mailbox.item.itemId),
                            outlookToken: result.value,
                            outlookRestUrl: getRestUrl(),
                            oneDriveToken: token.access_token
                        };

                        $.ajax({
                            type: "POST",
                            url: "/api/SaveAttachments",
                            data: JSON.stringify(saveAttachmentsRequest),
                            contentType: "application/json; charset=utf-8"
                        }).done(function (data) {
                            showNotification("Success", "Attachments saved");
                        }).fail(function (error) {
                            showNotification("Error saving attachments", error.status);
                        }).always(function () {
                            hideSpinner();
                        });
                    } else {
                        showNotification("Error getting callback token", JSON.stringify(result));
                        hideSpinner();
                    }
                });
            })
            .catch(function (error) {
                showNotification("Error authenticating to OneDrive", error);
                hideSpinner();
            });
    }

    // Helper function to show spinner
    function showSpinner() {
        spinner.start();
        overlay.show();
    }

    // Helper function to hide spinner
    function hideSpinner() {
        spinner.stop();
        overlay.hide();
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
    }
})();