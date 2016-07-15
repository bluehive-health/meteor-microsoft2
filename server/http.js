Microsoft.http = {
  callBase(method, url, options) {
    // if (method === 'POST') {
    //   options = options || {};
    //   options.headers = {
    //     ...options.headers,
    //     'Content-type': 'application/json', // required for POST requests
    //   };
    // }

    let response;

    try {
      response = HTTP.call(method, url, options)
    } catch (err) {
      const details = JSON.stringify({ url, options, method });
      throw new Meteor.Error('microsoft.http-request.', err.message, details);
    }

    return response.data || response;
  },
  callAuthenticated(method, url, accessToken, options) {
    options = options || {};
    options.headers = {
      ...options.headers,
      Authorization: 'Bearer ' + accessToken,
    };

    return Microsoft.http.callBase(method, url, options);
  },
  callAsUser(method, url, user, options) {
    let accessToken = Microsoft.getOrUpdateUserAccessToken(user);
    let data = Microsoft.http.callAuthenticated(method, url, accessToken, options);

    if (data.error && data.error.code === 'InvalidAuthenticationToken') {
      // try once to get new access token
      Microsoft.updateUserAccessToken(user);
      data = Microsoft.http.callAuthenticated(method, url, accessToken, options);
    }

    if (data.error) {
      throw new Meteor.Error('microsoft.http-response.' + data.error.code,
        'Error calling Microsoft API: ' + data.error.message);
    }

    return data;
  },
  callAsUserId(method, url, userId, options) {
    var user = Meteor.users.findOne(userId);
    return Microsoft.http.callAsUser(method, url, user, options);
  },
  getTokensBase(additionalParams) {
    const config = Microsoft.getConfiguration();
    const baseParams = {
      client_id: config.clientId,
      client_secret: OAuth.openSecret(config.secret),
      redirect_uri: OAuth._redirectUri(Microsoft.serviceName, config),
    };
    const requestBody = _.extend(baseParams, additionalParams);
    const response = Microsoft.http.callBase('POST', Microsoft.tokenURI, { params: requestBody });
    return {
      accessToken: response.access_token,
      refreshToken: response.refresh_token,
      expiresIn: response.expires_in,
      idToken: response.id_token,
      scope: response.scope,
    };
  },
};
