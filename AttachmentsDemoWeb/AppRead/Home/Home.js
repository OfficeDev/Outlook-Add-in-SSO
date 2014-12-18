// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#saveAttachment').click(saveAttachment);
            $('#providePermission').click(doOAuthFlow);

            checkForConsent();
            displayItemDetails();
        });
    };

    // Displays the attachments from the current mail item
    function displayItemDetails() {
        var item = Office.context.mailbox.item;

        for (var i = 0; i < item.attachments.length; i++) {
            $('<tr><td>' + item.attachments[i].name + '</td><td>' + item.attachments[i].size + '</td></tr>').appendTo('#attachments tbody');
        }
    }

    function showOneDriveUI(isAuthorized)
    {
        $('#checking_auth').hide();
      
        if (isAuthorized == true) {
            $('#auth_needed').hide();
            $('#providePermission').hide();
        }
        else {
            $('#auth_needed').show();
            $('#providePermission').show();
        }

        $('#onedrive_ui').show();
    }

    // Check if the WebAPI already has consent for the user
    function checkForConsent() {
        var user = Office.context.mailbox.userProfile;

        var authData = {
            UserEmail: user.emailAddress
        };
        $.ajax({
            url: '../../api/OAuth/IsConsentInPlace',
            type: 'POST',
            data: JSON.stringify(authData),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            showOneDriveUI(data);
        }).fail(function (status) {
            showOneDriveUI(false);
            app.showNotification('Error', JSON.stringify(status));
        });
    }

    function doOAuthFlow() {
        var user = Office.context.mailbox.userProfile;

        var authData = {
            UserEmail: user.emailAddress
        };

        $('.disable-while-sending').prop('disabled', true);

        $.ajax({
            url: '../../api/OAuth/GetAuthorizationUrl',
            type: 'POST',
            data: JSON.stringify(authData),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            window.open(data);
            
        }).fail(function (status) {
            app.showNotification('Error', JSON.stringify(status));
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }

    function saveAttachment() {
        $('.disable-while-sending').prop('disabled', true);

        var attachmentIds = [];
        for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
            attachmentIds[i] = Office.context.mailbox.item.attachments[i].id;
        }
        
        var ewsUrl = Office.context.mailbox.ewsUrl;
        Office.context.mailbox.getCallbackTokenAsync(function (ar) {
            var token = ar.value;

            var attachmentData = {
                UserEmail: Office.context.mailbox.userProfile.emailAddress,
                AuthToken: token,
                AttachmentIds: attachmentIds,
                EwsUrl: ewsUrl
            };

            sendRequest("GetAttachment/SaveAttachments", attachmentData);
        });
    }

    // Helper method
    function sendRequest(method, data) {
        $.ajax({
            url: '../../api/' + method,
            type: 'POST',
            data: JSON.stringify(data),
            contentType: 'application/json;charset=utf-8'
        }).done(function (data) {
            app.showNotification("Success", JSON.stringify(data));
        }).fail(function (status) {
            app.showNotification('Error', JSON.stringify(status));
        }).always(function () {
            $('.disable-while-sending').prop('disabled', false);
        });
    }
})();

// MIT License: 
 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 
 
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 
 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 