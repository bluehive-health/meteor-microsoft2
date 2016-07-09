Microsoft.getOrUpdateUserAccessToken = user => {
  checkUserIsDefined(user);
  checkMicrosoftAccountIsLinked(user);

  if (isAccessTokenMissingOrExpired(user)) {
    user = Microsoft.updateUserAccessToken(user); // returns updated user object
  }

  return user.services.microsoft.accessToken;
};

Microsoft.updateUserAccessToken = user => {
  const tokens = getTokensWithRefresh(user);
  user = saveTokens({ tokens, user });
  return user;
};

function checkUserIsDefined(user) {
  if (!user) {
    throw new Meteor.Error('microsoft:User required',
      'The supplied user is null or undefined');
  }
}

function checkMicrosoftAccountIsLinked(user) {
  if (!user || !user.services || !user.services.microsoft) {
    throw new Meteor.Error('microsoft:Account not linked',
      'The user needs to link a Microsoft account prior to making this call.');
  }
}

function isAccessTokenMissingOrExpired(user) {
  return !user.services.microsoft.accessToken ||
    new Date() >= new Date(user.services.microsoft.expiresAt);
}

function getTokensWithRefresh(user) {
  return Microsoft.http.getTokensBase({
    grant_type: 'refresh_token',
    refresh_token: user.services.microsoft.refreshToken,
  });
}

function saveTokens({ tokens, user }) {
  user.services.microsoft = {
    ...user.services.microsoft,
    ...tokens, // keep additional user fields
  };
  const modifier = {
    $set: {
      'services.microsoft': user.services.microsoft,
    },
  };

  Meteor.users.update(user._id, modifier);
  return user;
}
