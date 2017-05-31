// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
Office.initialize = function () {
}

function showProgress(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "progressIndicator",
        message: message
    });
}

function showError(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "errorMessage",
        message: message
    });
}

function showSuccess(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        message: message,
        icon: "icon16",
        persistent: false
    });
}

function saveAllAttachments(event) {
    showProgress("Retrieving OneDrive access token");

    var authenticator = new OfficeHelpers.Authenticator();
    authenticator.endpoints.registerMicrosoftAuth(authConfig.clientId, {
        redirectUrl: authConfig.redirectUrl,
        scope: authConfig.scopes
    });

    authenticator
        .authenticate("Microsoft")
        .then(function (token) {
            showProgress("Retrieving Outlook callback token");

            Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                if (result.status === "succeeded") {
                    showProgress("Saving attachments");

                    var attachmentIds = [];

                    Office.context.mailbox.item.attachments.forEach(function (attachment) {
                        attachmentIds.push(getRestId(attachment.id));
                    });

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
                        showSuccess("Attachments saved");
                    }).fail(function (error) {
                        showError("Error saving attachments");
                    }).always(function () {
                        event.completed();
                    });
                } else {
                    showError("Error getting callback token.");
                    event.completed();
                }
            });
        })
        .catch(function (error) {
            showError("Error authenticating to OneDrive.");
            event.completed();
        });
}