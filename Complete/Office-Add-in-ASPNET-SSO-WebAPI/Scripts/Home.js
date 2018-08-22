// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/

// This value is used to prevent the user from being
// cycled repeatedly through prompts to rerun the operation.
var timesGetOneDriveFilesHasRun = 0;


// This value is used to record whether the add-in is running in an environment that support consent dialog.
var unsupportConsentDialog = false;

// This value is used to record how many times has been retried in the backend server for token swapping.
var retryOnServerMissingConsent = 0;

// Remember whether the user has consented successfully.
// This can avoid unnecessary consent prompt.
var consentGranted = false;


Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
        $("#getGraphAccessTokenButton").click(function () {
            timesGetOneDriveFilesHasRun = 0;
            $('#file-list').html("");
            getOneDriveFiles();
        });
    });
};

// Main function to do everything
function getOneDriveFiles() {
    // Reset all the variables
    retryOnServerMissingConsent = 0;
    timesGetOneDriveFilesHasRun++;
    unsupportConsentDialog = false;
    consentGranted = false;
    getAccessToken({});
}


function getAccessToken(options) {
    // If the add-in is running in a category that consent is not supported, the API should abort if hitting any consent required error.
    if (options["forceConsent"] == true && unsupportConsentDialog) {
        console.log("Cannot get access token for this catalog. Abort.");
        return;
    }

    // If consented has been granted before, reset the option to prevent unnecessary consent prompt
    if (consentGranted) {
        options["forceConsent"] = false;
    }

    console.log("Call Office.context.auth.getAccessTokenAsync()");
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                if (options["forceConsent"] == true) {
                    consentGranted = true;
                }

                var accessToken = result.value;
                getGraphData(accessToken);
            }
            else {
                handleClientSideErrors(result);
            }
        });
}

// Calls the specified URL or route (in the same domain as the add-in) 
// and includes the specified access token.
function getGraphData(accessToken) {

    console.log("Send request to add-in server for Graph data.");
    $.support.cors = true;
    $.ajax({
        type: "GET",
        url: "/api/values",
        headers: {
            "Authorization": "Bearer " + accessToken
        },
        dataType: "json"
    })
        .done(function (data) {
            showResult(data);

        })
        .fail(handleServerSideErrors);
}

function handleClientSideErrors(result) {

    switch (result.error.code) {

        case 13001:
            // The user is not logged in, or the user cancelled without responding a
            // prompt to provide a 2nd authentication factor. (See comment about two-
            // factor authentication in the fail callback of the getData method.)
            // Either way start over and force a sign-in. 
            getAccessToken({ forceAddAccount: true });
            break;
        case 13002:
            // User refuses to grant the consent. Stop.
            showResult(['Please grant the consent']);
            logApiError(result);
            break;
        case 13003:
            // The user is logged in with an account that is neither work or school, nor Micrososoft Account.
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
            break;
        case 13005:
            // Missing consent error. Need to prompt the consent dialog.
            getAccessToken({ forceConsent: true });
            break;
        case 13006:
            // Unspecified error in the Office host.
            showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
            break;
        case 13007:
            // The Office host cannot get an access token to the add-ins web service/application.
            showResult(['That operation cannot be done at this time. Please try again later.']);
            break;
        case 13008:
            // The user tiggered an operation that calls getAccessTokenAsync before a previous call of it completed.
            showResult(['Please try that operation again after the current operation has finished.']);
            break;
        case 13009:
            // The add-in does not support forcing consent. Try signing the user in without forcing consent, unless
            // that's already been tried.
            unsupportConsentDialog = true;
            getAccessToken({ forceConsent: false });
            break;
        case 13012:
            // The SSO API is not supported in this platform. Develooper should consider using alternative solution for authentication.
            showResult(['The SSO API is not supported in this platform.']);
            break;
        default:
            logApiError(result);
            break;
    }
}


function handleServerSideErrors(error) {
    if (error.status == 401) {
        var response = JSON.parse(error.responseText);
        var errorCode = response["errorCode"];

        // This indicates that the server fail to do the on-behalf-of flow because of MFA or permission problem.
        if (errorCode == "invalid_grant") {
            // For OrgID, claim string will be returned in all cases. Need to check the suberror for more info.
            if (response["claims"] != null && response["suberror"] != null && response["suberror"] == "basic_action") {
                // MFA required
                console.log("Add-in server response: need to do MFA");
                getAccessToken({ authChallenge: response["claims"] });

            }
            else {
                // Consent required
                console.log("Add-in server response: missing consent");
                if (retryOnServerMissingConsent == 10) {

                    console.log("Cannot get the access token after 10 retries, stop.");
                    console.log("Failed to get graph Data.");
                } else if (retryOnServerMissingConsent == 0) {
                    retryOnServerMissingConsent++;
                    getAccessToken({ forceConsent: true });
                } else {
                    retryOnServerMissingConsent++;

                    console.log("Wait for 5 seconds then retry.");
                    setTimeout(function () {
                        getAccessToken({ forceConsent: true });
                    }, 5000);
                }
            }
        } else if (errorCode == "invalid_graph_token") {
            // If the token sent to MS Graph is expired or invalid, start the whole process over.
            if (timesGetOneDriveFilesHasRun < 2) {
                getOneDriveFiles();
            }
        } else if (errorCode == "invalid_access_token") {
            showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
        } else {
            showResult(['Hit unknown error']);
        }
    }
    else {
        showResult(['Server encounter internal error']);
    }


}


// Displays the data, assumed to be an array.
function showResult(data) {
    $('#file-list').html("");
    for (var i = 0; i < data.length; i++) {
        $('#file-list').append('<li class="ms-ListItem">' +
            '<span class="ms-ListItem-secondaryText">' +
            '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' +
            '</span></li>');
    }
}

function logApiError(result) {
    console.log("Status: " + result.status);
    console.log("Code: " + result.error.code);
    console.log("Name: " + result.error.name);
    console.log("Message: " + result.error.message);
}



