const msalConfig = {
  auth: {
    clientId: env.CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    validateAuthority: true
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const loginRequest = {
  scopes: ["openid", "profile", "User.Read"]
};

const tokenRequest = {
  scopes: ["Directory.Read.All"]
};

// resource endpoints
const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};

// instantiate MSAL
const myMSALObj = new Msal.UserAgentApplication(msalConfig);

// register callback for redirect usecases
myMSALObj.handleRedirectCallback(authRedirectCallBack);

// signin and acquire a token silently with redirect flow. Fall back in case of failure with silent acquisition to redirect
function signIn() {
  myMSALObj.loginRedirect(loginRequest);
}

// Call Graph to fetch data
function callMSGraph(theUrl, accessToken, callback) {
  axios.get(theUrl,
    {
      headers: {
        Authorization: 'Bearer ' + accessToken
      }
    })
    .then(function (response) {
      callback(response.data);
    })
    .catch(function (error) {
      console.log(error);
    })
}

function graphAPICallback(data) {
  document.getElementById("json").innerHTML = JSON.stringify(data, null, 2);

  if (data.value) {
    const list = data.value.map(element => '<li>' + element.displayName + ' <span class="font-weight-light">(' + element.jobTitle + ')</span><br/><span class="font-weight-light">UPN: ' + element.userPrincipalName + '</span></li>')
      .join('');
    document.getElementById("userList").innerHTML = '<ul>' + list + '</ul>';
  }
}

function dumpAccessToken(data) {
  document.getElementById("accessToken").innerHTML = data;
}

function showWelcomeMessage() {
  const divWelcome = document.getElementById('WelcomeMessage');
  divWelcome.innerHTML = 'Welcome ' + myMSALObj.getAccount().userName + " to Microsoft Graph API";
  const loginbutton = document.getElementById('SignIn');
  loginbutton.innerHTML = 'Sign Out';
  loginbutton.setAttribute('onclick', 'signOut();');
}

// search for a given email address
function searchGraph() {
  const email = document.getElementById('Email').value;
  const graphEndpoint = `https://graph.microsoft.com/v1.0/users?$filter=startswith(mail,'${encodeURIComponent(email)}')&$select=jobTitle,displayName,mail,userPrincipalName`;
  acquireTokenRedirectAndCallMSGraph(graphEndpoint, tokenRequest);
  return false;
}

// signout
function signOut() {
  myMSALObj.logout();
}

function acquireTokenRedirectAndCallMSGraph(endpoint, request) {
  //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
  myMSALObj.acquireTokenSilent(request).then(function (tokenResponse) {
    dumpAccessToken(tokenResponse.accessToken);
    callMSGraph(endpoint, tokenResponse.accessToken, graphAPICallback);
  }).catch(function (error) {
    console.log("error is: " + error);
    console.log("stack:" + error.stack);
    //Call acquireTokenRedirect in case of acquireToken Failure
    if (requiresInteraction(error.errorCode)) {
      myMSALObj.acquireTokenRedirect(request);
    }
  });
}

// redirect call back
function authRedirectCallBack(error, response) {
  if (error) {
    console.log(error);
  } else {
    if (response.tokenType === "id_token") {
      showWelcomeMessage();
      acquireTokenRedirectAndCallMSGraph(graphConfig.graphMeEndpoint, loginRequest);
    } else if (response.tokenType === "access_token") {
      dumpAccessToken(response.accessToken);
      callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, graphAPICallback);
    } else {
      console.log("token type is:" + response.tokenType);
    }
  }
}

// utils to handle standard error set that would need user interaction
function requiresInteraction(errorMessage) {
  if (!errorMessage || !errorMessage.length) {
    return false;
  }

  console.log("requiresinteraction is:" + errorMessage);
  return errorMessage.indexOf("consent_required") !== -1 ||
    errorMessage.indexOf("interaction_required") !== -1 ||
    errorMessage.indexOf("login_required") !== -1;
}

showWelcomeMessage();
acquireTokenRedirectAndCallMSGraph(graphConfig.graphMeEndpoint, loginRequest);
