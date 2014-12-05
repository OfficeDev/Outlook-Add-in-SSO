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
            //$('#saveAttachment').show();
            $('#auth_needed').hide();
            $('#providePermission').hide();
        }
        else {
            //$('#saveAttachment').hide();
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
            
            //$.oauthpopup({
            //    path: encodeURIComponent(data),
            //    callback: function () {
            //        app.showNotification("Callback happened!");
            //    }
            //});
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

    //$.oauthpopup = function (options) {
    //    options.windowName = options.windowName || 'ConnectWithOAuth'; // should not include space for IE
    //    options.windowOptions = options.windowOptions || 'location=0,status=0,width=800,height=400';
    //    options.callback = options.callback || function () { window.location.reload(); };
    //    var that = this;
        
    //    var targetUrl = '../OAuthStart.html?oauthUrl=' + options.path;
    //    that._oauthWindow = window.open(targetUrl, options.windowName, options.windowOptions);
    //    that._oauthWindow.oauthUrl = options.path;
    //    that._oauthInterval = window.setInterval(function () {
    //        if (that._oauthWindow.closed) {
    //            window.clearInterval(that._oauthInterval);
    //            options.callback();
    //        }
    //    }, 1000);

    //};
})();