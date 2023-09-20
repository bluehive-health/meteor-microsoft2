import {
    ServiceConfiguration
} from 'meteor/service-configuration';

import {
    OAuth
} from 'meteor/oauth';

import {
    Accounts
} from 'meteor/accounts-base';


Microsoft = {
    serviceName: 'microsoft',
    tokenURI: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    authURI: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
    whitelistedFields: ['id', 'mail', 'displayName', 'givenName', 'surname', 'jobTitle', 'mobilePhone', 'userPrincipalName'],

    retrieveCredential: function (credentialToken, credentialSecret) {
        return OAuth.retrieveCredential(credentialToken, credentialSecret);
    },

    getConfiguration() {
        const config = ServiceConfiguration.configurations.findOne({ service: Microsoft.serviceName });

        if (!config && !returnNullIfMissing)
            throw new ServiceConfiguration.ConfigError();
        else if (!config && returnNullIfMissing)
            return null;

        config.loginStyle = 'popup';

        return config;
    }
};

OAuth.registerService(Microsoft.serviceName, 2, null, function ({ code }) {

    const response = getTokensFromCode(code),
        expiresAt = (+new Date) + (1000 * parseInt(response.expiresIn, 10)),
        identity = getIdentity(response.accessToken);
    let serviceData = {
        accessToken: response.accessToken,
        idToken: response.idToken,
        expiresAt: expiresAt,
        scope: response.scope
    };

    const fields = Microsoft.whitelistedFields.reduce((obj, key) => (identity[key] ? { ...obj, [key]: identity[key] } : obj), {});
    serviceData = { ...serviceData, ...fields };

    if (response.refreshToken)
        serviceData.refreshToken = response.refreshToken;

    return {
        serviceData: serviceData,
        options: { profile: { name: identity.displayName } }
    };
});

const getTokensFromCode = code => {
    return Microsoft.http.getTokensBase({
        grant_type: 'authorization_code',
        code,
    });
};

const getTokens = function (query) {
    const config = ServiceConfiguration.configurations.findOne({ service: Microsoft.serviceName });
    if (!config) {
        throw new ServiceConfiguration.ConfigError();
    }

    let response;
    try {
        response = HTTP.post(
            Microsoft.tokenURI, {
                params: {
                    code: query.code,
                    client_id: config.clientId,
                    client_secret: OAuth.openSecret(config.secret),
                    redirect_uri: OAuth._redirectUri(Microsoft.serviceName, config),
                    grant_type: 'authorization_code'
                }
        });
    } catch (err) {
        throw Object.assign(
            new Error('Failed to complete OAuth handshake with Microsoft. ' + err.message),
            { response: err.response }
        );
    }

    if (response.data.error) {
        throw new Error('Failed to complete OAuth handshake with Microsoft. ' + response.data.error);
    } else {
        return {
            accessToken: response.data.access_token,
            refreshToken: response.data.refresh_token,
            expiresIn: response.data.expires_in,
            idToken: response.data.id_token,
            scope: response.data.scope,
        };
    }
};

const getIdentity = function (accessToken) {
    try {
        return HTTP.get(
            'https://graph.microsoft.com/v1.0/me',
            { headers: { Authorization: 'Bearer ' + accessToken } }).data;
    } catch (err) {
        throw Object.assign(
            new Error('Failed to fetch identity from Microsoft. ' + err.message),
            { response: err.response }
        );
    }
};

Accounts.oauth.registerService('microsoft');