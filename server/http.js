import { fetch } from 'meteor/fetch';

Microsoft.http = {
  async callBase(method, url, options = {}) {
    let data;

    try {
      const headers = { ...options.headers };
      let requestUrl = url;
      let body;

      if (options.params) {
        const search = new URLSearchParams(options.params).toString();
        if (method.toUpperCase() === 'GET') {
          const parsedUrl = new URL(url);
          if (search) {
            parsedUrl.search = parsedUrl.search ? `${parsedUrl.search}&${search}` : search;
          }
          requestUrl = parsedUrl.toString();
        } else {
          body = search;
          headers['Content-Type'] = 'application/x-www-form-urlencoded';
        }
      }

      if (options.data) {
        body = JSON.stringify(options.data);
        headers['Content-Type'] = 'application/json';
      }

      const response = await fetch(requestUrl, { method, headers, body });
      const text = await response.text();
      try {
        data = JSON.parse(text);
      } catch (parseError) {
        data = text;
      }

      if (!response.ok && !(options.allowErrorResponse && data && data.error)) {
        throw new Error(`HTTP ${response.status} ${response.statusText}`);
      }
    } catch (err) {
      const details = JSON.stringify({ url, options, method });
      throw new Meteor.Error('microsoft.http-request.', err.message, details);
    }

    return data;
  },
  callAuthenticated(method, url, accessToken, options = {}) {
    options.headers = {
      ...options.headers,
      Authorization: `Bearer ${accessToken}`,
    };

    return Microsoft.http.callBase(method, url, options);
  },
  async callAsUser(method, url, user, options) {
    let accessToken = await Microsoft.getOrUpdateUserAccessToken(user);
    const requestOptions = { ...options, allowErrorResponse: true };
    let data = await Microsoft.http.callAuthenticated(method, url, accessToken, requestOptions);

    if (data.error && data.error.code === 'InvalidAuthenticationToken') {
      await Microsoft.updateUserAccessToken(user);
      data = await Microsoft.http.callAuthenticated(method, url, accessToken, requestOptions);
    }

    if (data.error) {
      throw new Meteor.Error(`microsoft.http-response.${data.error.code}`,
        `Error calling Microsoft API: ${data.error.message}`);
    }

    return data;
  },
  async callAsUserId(method, url, userId, options) {
    const user = await Meteor.users.findOneAsync(userId);
    return Microsoft.http.callAsUser(method, url, user, options);
  },
  async getTokensBase(additionalParams) {
    const config = await Microsoft.getConfiguration();
    const baseParams = {
      client_id: config.clientId,
      client_secret: OAuth.openSecret(config.secret),
      redirect_uri: OAuth._redirectUri(Microsoft.serviceName, config),
    };
    const requestBody = { ...baseParams, ...additionalParams };
    const response = await Microsoft.http.callBase('POST', Microsoft.tokenURI, { params: requestBody });
    return {
      accessToken: response.access_token,
      refreshToken: response.refresh_token,
      expiresIn: response.expires_in,
      idToken: response.id_token,
      scope: response.scope,
    };
  },
};