// config
const msalConfig = {
    auth: {
        clientId: '0e1277ee-5c7f-4648-a5bf-09e839affe82',
        authority: "https://login.microsoftonline.com/common",
        redirectUri: 'http://localhost/ci4-office365-auth/public/home/office365Auth'
    }
};
const msalRequest = {
    scopes: [
        'user.read',
    ]
}

// Create the main MSAL instance
// configuration parameters are located in config.js
const msalClient = new msal.PublicClientApplication(msalConfig);

async function loginWithOffice365() {
    // Login
    try {
        // Use MSAL to login
        const authResult = await msalClient.loginPopup(msalRequest);
        // console.log('authResult ', authResult);
        // console.log('id_token acquired at: ' + new Date().toString());
        // Save the account username, needed for token acquisition
        sessionStorage.setItem('msalAccount', authResult.account.username);

        // login success
        callBackendToLogin();

    } catch (error) {
        alert(error);
    }
}

function callBackendToLogin() {
    authProvider.getAccessToken().then(accessToken=>{
        jQuery.ajax({
            url: msalConfig.auth.redirectUri,
            type: 'post',
            data: {accessToken: accessToken},
            dataType : "json",
            success: function(resp){
            console.log(resp);
            if ( resp.status=='success' ) {
                window.location.href = '';// set redirect url
            } else {
                alert(resp.msg);
            }
        }
    });
    });
}

async function getToken() {
    let account = sessionStorage.getItem('msalAccount');
    if (!account){
        alert('User account missing from session. Please sign out and sign in again.');
    }

    try {
        // First, attempt to get the token silently
        const silentRequest = {
            scopes: msalRequest.scopes,
            account: msalClient.getAccountByUsername(account)
        };

        const silentResult = await msalClient.acquireTokenSilent(silentRequest);
        return silentResult.accessToken;
    } catch (silentError) {
        // If silent requests fails with InteractionRequiredAuthError,
        // attempt to get the token interactively
        if (silentError instanceof msal.InteractionRequiredAuthError) {
            const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
            return interactiveResult.accessToken;
        } else {
            throw silentError;
        }
    }
}

// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
        return await getToken();
    }
};

// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});

async function getUser() {
    return await graphClient
        .api('/me')
        .select('id,displayName,userPrincipalName')
        .get();
}

function signOut() {
    account = null;
    // sessionStorage.removeItem('graphUser');
    sessionStorage.removeItem('msalAccount');
    msalClient.logout();
}