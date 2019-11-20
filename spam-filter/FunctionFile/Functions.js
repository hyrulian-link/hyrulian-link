// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

'use strict';

var rawToken;
var parsedToken = '';

var Queue = function () {
    var previous = new $.Deferred().resolve();

    return function (fn, fail) {
        return previous = previous.then(fn, fail || fn);
    };
};
var queue = Queue();

var spamWords = ['NEWSLETTER', 'UNSUB', 'PW', 'EMAIL LIST', 'EMAILS LIST', 'NETFLIX']

Office.initialize = function () {
}

function loadRestDetails() {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            rawToken = result.value;
            loadToken($('#parse-token-toggle').is(':checked'));
        } else {
            rawToken = 'error';
        }
        deleteJunk();
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
    var junkEmailFolder = getJunkMailFolder();
    var junkUrl = 'https://outlook.office.com/api/beta/me/MailFolders/' + junkEmailFolder +
        '/messages/?$select=Sender,Body&$top=50';
    do {
        var junkMessageResult = getItem(junkUrl);
        var junkMessages = junkMessageResult.value;
        for (var i = 0; i < junkMessages.length; i++) {
            var message = junkMessages[i];
            var senderAddress = message.Sender.EmailAddress.Address;
            var bodyContent = message.Body.Content;
            var spamWordsRegExp = new RegExp(spamWords.join("|"));
            if (senderAddress && spamWordsRegExp.test(senderAddress.toUpperCase())
                || bodyContent && spamWordsRegExp.test(bodyContent.toUpperCase())) {
                deleteItem('https://outlook.office.com/api/beta/me/messages/' + message.Id);
            }
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
    var result;
    if (isAsync) {
        queue(function () {
            return ajaxRequest(type, url, isAsync);
        });
    } else {
        ajaxRequest(type, url, isAsync)
            .done(function (item) {
                result = item;
            });
    }
    return result;
}

function ajaxRequest(type, url, isAsync) {
    return $.ajax({
        type: type,
        url: url,
        dataType: 'json',
        async: isAsync,
        headers: { 'Authorization': 'Bearer ' + rawToken }
    });
}