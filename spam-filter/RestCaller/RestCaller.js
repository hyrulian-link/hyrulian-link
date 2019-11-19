// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

(function () {
    'use strict';

    var rawToken = '';
    var parsedToken = '';

    var getItemSpinnerElement = null;
    var getItemSpinner = null;

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            var PivotElements = document.querySelectorAll('.ms-Pivot');
            for (var i = 0; i < PivotElements.length; i++) {
                new fabric['Pivot'](PivotElements[i]);
            }

            var ToggleElements = document.querySelectorAll('.ms-Toggle');
            for (var i = 0; i < ToggleElements.length; i++) {
                new fabric['Toggle'](ToggleElements[i]);
            }

            getItemSpinnerElement = document.querySelector('.get-item-spinner');
            getItemSpinner = new fabric['Spinner'](getItemSpinnerElement);
            getItemSpinner.stop();

            var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
            for (var i = 0; i < DropdownHTMLElements.length; ++i) {
                var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
            }

            $('#parse-token-toggle').click(function () {
                loadToken($('#parse-token-toggle').is(':checked'));
            });

            $('.get-item-button').click(function () {
                deleteJunk();
            });

            loadRestDetails();
        });
    };

    function loadRestDetails() {
        $('.hostname').text(Office.context.mailbox.diagnostics.hostName);
        $('.hostversion').text(Office.context.mailbox.diagnostics.hostVersion);
        $('.owaview').text(Office.context.mailbox.diagnostics.OWAView);

        var restId = '';
        if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
            // Loaded in non-mobile context, so ID needs to be converted
            restId = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.Beta
            );
        } else {
            restId = Office.context.mailbox.item.itemId;
        }

        // Build the URL to the item
        //var itemUrl = Office.context.mailbox.restUrl + 
        var itemUrl = 'https://outlook.office.com' +
            '/api/beta/me/messages/' + restId;

        $('.resturl-display code').text(itemUrl);

        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === "succeeded") {
                rawToken = result.value;
                loadToken($('#parse-token-toggle').is(':checked'));
                enableButtons();
            } else {
                rawToken = 'error';
            }
        });
    }

    function loadToken(parseToken) {
        var code = $('.token-display code');
        if (rawToken === 'error') {
            code.text('ERROR RETRIEVING TOKEN');
            return;
        }

        if (parseToken) {
            if (parsedToken === '') {
                parsedToken = jwt_decode(rawToken);
            }

            code.text(JSON.stringify(parsedToken, null, 2));
        } else {
            code.text(rawToken);
        }
    }

    function deleteJunk() {
        deleteJunkWithSearch('nsub');
        deleteJunkWithSearch('newsletter');
    }

    function deleteJunkWithSearch(search) {
        var junkEmailFolder = getJunkMailFolder();
        var junkUrl = 'https://outlook.office.com/api/beta/me/MailFolders/' + junkEmailFolder +
            '/messages/?$select=Id&search=%22' + search + '%22&$top=50';
        do {
            var junkMessageResult = getItem(junkUrl);
            var junkMessages = junkMessageResult.value;
            for (var i = 0; i < junkMessages.length; i++) {
                deleteItem('https://outlook.office.com/api/beta/me/messages/' + junkMessages[i].Id);
            }
            junkUrl = junkMessageResult['@odata.nextLink'];
        } while (junkUrl);
    }

    function getJunkMailFolder() {
        var foldersResult = getItem('https://outlook.office.com/api/beta/me/MailFolders?$top=50');
        var folders = foldersResult.value;
        for (var i = 0; i < folders.length; i++) {
            var folder = folders[i];
            if (folder.DisplayName.includes('Junk Email')) {
                return folder.Id;
            }
        }
    }

    function getItem(url) {
        return restRequest('GET', url, false);
    }

    function deleteItem(url) {
        restRequest('DELETE', url, true);
    }

    function restRequest(type, url, isAsync) {
        toggleGetItemSpinner(true);
        var result;
        $.ajax({
            type: type,
            url: url,
            dataType: 'json',
            async: isAsync,
            headers: { 'Authorization': 'Bearer ' + rawToken }
        }).done(function (item) {
            toggleGetItemSpinner(false);
            result = item;
        }).fail(function (error) {
            toggleGetItemSpinner(false);
            result = error;
        });
        return result;
    }

    function enableButtons() {
        $('.get-item-button').removeClass('is-disabled');
        $('.update-item-button').removeClass('is-disabled');
    }

    function toggleGetItemSpinner(showSpinner) {
        if (showSpinner) {
            getItemSpinner.start();
            getItemSpinnerElement.style.display = "block";
        } else {
            getItemSpinner.stop();
            getItemSpinnerElement.style.display = "none";
        }
    }

})();