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

    async getConfiguration() {
        const config = await ServiceConfiguration.configurations.findOneAsync({ service: Microsoft.serviceName });

        if (!config)
            throw new ServiceConfiguration.ConfigError();

        config.loginStyle = 'popup';

        return config;
    }
};

OAuth.registerService(Microsoft.serviceName, 2, null, async function ({ code }) {

    const response = await getTokensFromCode(code),
        expiresAt = (+new Date) + (1000 * parseInt(response.expiresIn, 10)),
        identity = await getIdentity(response.accessToken);
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

const getIdentity = async function (accessToken) {
    try {
        return await Microsoft.http.callBase('GET',
            'https://graph.microsoft.com/v1.0/me',
            { headers: { Authorization: 'Bearer ' + accessToken } });
    } catch (err) {
        throw Object.assign(
            new Error('Failed to fetch identity from Microsoft. ' + err.message),
            { response: err.response }
        );
    }
};

Accounts.oauth.registerService('microsoft');