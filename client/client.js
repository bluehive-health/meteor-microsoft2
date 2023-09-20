Microsoft = {

    serviceName: 'microsoft',

    graph: { authUrl: 'https://graph.microsoft.com/' },
    outlook: { authUrl: 'https://outlook.office.com/' },

    requestCredential: function (options, credentialRequestCompleteCallback) {

        if (!credentialRequestCompleteCallback && typeof options === 'function') {
            credentialRequestCompleteCallback = options;
            options = {};
        } else if (!options) {
            options = {};
        }

        var config = ServiceConfiguration.configurations.findOne({service: Microsoft.serviceName});
        if (!config) {
            credentialRequestCompleteCallback &&
            credentialRequestCompleteCallback(new ServiceConfiguration.ConfigError());
            return;
        }

        var credentialToken = Random.secret(),
            loginStyle = OAuth._loginStyle(Microsoft.serviceName, config, options);

        OAuth.launchLogin({
            loginService: Microsoft.serviceName,
            loginStyle: loginStyle,
            loginUrl: getLoginUrlOptions(loginStyle, credentialToken, config, options),
            credentialRequestCompleteCallback: credentialRequestCompleteCallback,
            credentialToken: credentialToken,
            popupOptions: { width: 445, height: 625 }
        });
    }
};

var getLoginUrlOptions = function(loginStyle, credentialToken, config, options) {

    var scope = [Microsoft.graph.authUrl + 'User.Read'];
    if (options.requestOfflineToken) {
        scope.push('offline_access');
    }
    if (options.requestPermissions) {
        scope = [...new Set(scope.concat(options.requestPermissions))];
    }

    if (options.requestGraphPermissions) {
      scope = [...new Set(scope.concat(options.requestGraphPermissions.map( scope => Microsoft.graph.authUrl + scope )))];
    }

    if (options.requestOutlookPermissions) {
      scope = [...new Set(scope.concat(options.requestOutlookPermissions.map( scope => Microsoft.outlook.authUrl + scope )))];
    }

    var loginUrlParameters = {};
    if (config.loginUrlParameters){
        Object.assign(loginUrlParameters, config.loginUrlParameters);
    }
    if (options.loginUrlParameters){
        Object.assign(loginUrlParameters, options.loginUrlParameters);
    }
    var illegal_parameters = ['response_type', 'client_id', 'scope', 'redirect_uri', 'state'];
    Object.keys(loginUrlParameters).forEach(function (key) {
        if (illegal_parameters.includes(key)) {
            throw new Error('Microsoft.requestCredential: Invalid loginUrlParameter: ' + key);
        }
    });

    Object.assign(loginUrlParameters, {
        response_type: 'code',
        client_id:  config.clientId,
        scope: scope.join(' '),
        redirect_uri: OAuth._redirectUri(Microsoft.serviceName, config),
        state: OAuth._stateParam(loginStyle, credentialToken, options.redirectUrl)
    });

    return 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?' +
        Object.entries(loginUrlParameters).map(([param, value]) => `${encodeURIComponent(param)}=${encodeURIComponent(value)}`).join('&');
};

Meteor.loginWithMicrosoft = function(options, callback) {
    if (! callback && typeof options === 'function') {
        callback = options;
        options = null;
    }

    if (typeof Accounts._options.restrictCreationByEmailDomain === 'string') {
        options = Object.assign({}, options || {});
        options.loginUrlParameters = Object.assign({}, options.loginUrlParameters || {});
        options.loginUrlParameters.hd = Accounts._options.restrictCreationByEmailDomain;
    }

    var credentialRequestCompleteCallback = Accounts.oauth.credentialRequestCompleteHandler(callback);
    Microsoft.requestCredential(options, credentialRequestCompleteCallback);
};