// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters, 
// visit https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md

var userAgent = navigator.userAgent;
console.log("user agent", userAgent);

var redirurl = window.location.protocol + "//" + window.location.hostname;
if(window.location.port){
    redirurl += ":" + window.location.port;
}
console.log("redir", redirurl);

var msalConfig = {
    auth: {
        clientId: "38235898-fa02-45d3-a200-f657688ccec2",
        authority: "https://login.microsoftonline.com/5a98c1cc-eb85-4540-a57b-fc658c02f598",
        //redirectUri: "https://msalpopupfordynamics.azurewebsites.net",
        redirectUri: redirurl,
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: function(level, message, containsPii){
                if (containsPii) {	
                    return;	
                }	
                switch (level) {	
                    case msal.LogLevel.Error:	
                        console.error(message);	
                        return;	
                    case msal.LogLevel.Info:	
                        console.info(message);	
                        return;	
                    case msal.LogLevel.Verbose:	
                        console.debug(message);	
                        return;	
                    case msal.LogLevel.Warning:	
                        console.warn(message);	
                        return;	
                }
            }
        }
    }
};

// Add here the scopes that you would like the user to consent during sign-in
var loginRequest = {
    scopes: ["User.Read"]
};

// Add here the scopes to request when obtaining an access token for MS Graph API
var tokenRequest = {
    scopes: ["User.Read", "Mail.Read"],
    forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
};

// Add here the endpoints for MS Graph API services you would like to use.
var graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages"
};

// Select DOM elements to work with
var welcomeDiv = document.getElementById("WelcomeMessage");
var signInButton = document.getElementById("SignIn");
var cardDiv = document.getElementById("card-div");
var mailButton = document.getElementById("readMail");
var profileButton = document.getElementById("seeProfile");
var profileDiv = document.getElementById("profile-div");

function showWelcomeMessage(account) {
    // Reconfiguring DOM elements
    cardDiv.style.display = 'block';
    welcomeDiv.innerHTML = "Welcome " + account.username;
    signInButton.setAttribute("onclick", "signOut();");
    signInButton.setAttribute('class', "btn btn-success")
    signInButton.innerHTML = "Sign Out";
}

function updateUI(data, endpoint) {
    console.log('Graph API responded at: ' + new Date().toString());

    if (endpoint === graphConfig.graphMeEndpoint) {
        var title = document.createElement('p');
        title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
        var email = document.createElement('p');
        email.innerHTML = "<strong>Mail: </strong>" + data.mail;
        var phone = document.createElement('p');
        phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
        var address = document.createElement('p');
        address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
        profileDiv.appendChild(title);
        profileDiv.appendChild(email);
        profileDiv.appendChild(phone);
        profileDiv.appendChild(address);

    } else if (endpoint === graphConfig.graphMailEndpoint) {
        if (data.value.length < 1) {
            alert("Your mailbox is empty!")
        } else {
            var tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call
            var tabContent = document.getElementById("nav-tabContent");

            data.value.map(function(d, i){
                // Keeping it simple
                if (i < 10) {
                    var listItem = document.createElement("a");
                    listItem.setAttribute("class", "list-group-item list-group-item-action")
                    listItem.setAttribute("id", "list" + i + "list")
                    listItem.setAttribute("data-toggle", "list")
                    listItem.setAttribute("href", "#list" + i)
                    listItem.setAttribute("role", "tab")
                    listItem.setAttribute("aria-controls", i)
                    listItem.innerHTML = d.subject;
                    tabList.appendChild(listItem)

                    var contentItem = document.createElement("div");
                    contentItem.setAttribute("class", "tab-pane fade")
                    contentItem.setAttribute("id", "list" + i)
                    contentItem.setAttribute("role", "tabpanel")
                    contentItem.setAttribute("aria-labelledby", "list" + i + "list")
                    contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
                    tabContent.appendChild(contentItem);
                }
            });
        }
    }
}

// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
var myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = "";

function loadPage() {
    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    var currentAccounts = myMSALObj.getAllAccounts();
    console.log("currentAccounts", currentAccounts);
    if (currentAccounts === null) {
        return;
    } else if (currentAccounts.length > 1) {
        // Add choose account code here
        console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
        username = currentAccounts[0].username;
        showWelcomeMessage(currentAccounts[0]);
    }
}

function handleResponse(resp) {
    if (resp !== null) {
        username = resp.account.username;
        showWelcomeMessage(resp.account);
    } else {
        loadPage();
    }
}

function signIn() {
    myMSALObj.loginPopup(loginRequest).then(handleResponse).catch(function(error){
        console.error(error);
    });
}

function signOut() {
    var logoutRequest = {
        account: myMSALObj.getAccountByUsername(username)
    };

    myMSALObj.logout(logoutRequest);
}

function getTokenPopup(request) {
    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    request.account = myMSALObj.getAccountByUsername(username);
    
    return myMSALObj.acquireTokenSilent(request).catch(function(error) {
        console.warn("silent token acquisition fails. acquiring token using popup");
        if (error instanceof msal.InteractionRequiredAuthError) {
            // fallback to interaction when silent call fails
            return myMSALObj.acquireTokenPopup(request).then(function(tokenResponse) {
                console.log(tokenResponse);
                return tokenResponse;
            }).catch(function(error) {
                console.error(error);
            });
        } else {
            console.warn(error);   
        }
    });
}

function seeProfile() {
    getTokenPopup(loginRequest).then(function(response) {
        callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
        profileButton.classList.add('d-none');
        mailButton.classList.remove('d-none');
    }).catch(function(error) {
        console.error(error);
    });
}

function readMail() {
    getTokenPopup(tokenRequest).then(function(response) {
        callMSGraph(graphConfig.graphMailEndpoint, response.accessToken, updateUI);
    }).catch(function(error) {
        console.error(error);
    });
}

loadPage();

// Helper function to call MS Graph API endpoint 
// using authorization bearer token scheme
function callMSGraph(endpoint, token, callback) {
    var headers = new Headers();
    var bearer = "Bearer " + token;

    headers.append("Authorization", bearer);

    var options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(function(response){ return response.json()})
        .then(function(response){ return callback(response, endpoint)})
        .catch(function (error){ console.log(error)});
}
