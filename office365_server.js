/* globals Office365, OAuth */
/* eslint-disable */
Office365 = {};

let userAgent = 'Meteor';
if (Meteor.release) {
  userAgent += `/${Meteor.release}`;
}

const getTokenResponse = function(query) {
  const config = ServiceConfiguration.configurations.findOne({ service: 'office365' });
  if (!config) {
    throw new ServiceConfiguration.ConfigError();
  }

  const redirectUri = OAuth._redirectUri('office365', config).replace('?close', '');

  let response;
  try {
    response = HTTP.post(`https://login.microsoftonline.com/${config.tenant || 'common'}/oauth2/v2.0/token`, {
      headers: {
        Accept: 'application/json',
        'User-Agent': userAgent,
      },
      params: {
        grant_type: 'authorization_code',
        code: query.code,
        client_id: config.clientId,
        client_secret: OAuth.openSecret(config.secret),
        redirect_uri: redirectUri,
        state: query.state,
      },
    });
  } catch (error) {
    throw _.extend(new Error(`Failed to complete OAuth handshake with Microsoft Office365. ${error.message}`), { response: error.response });
  }
  if (response.data.error) {
    throw new Error(`Failed to complete OAuth handshake with Microsoft Office365. ${response.data.error}`);
  } else {
    return response.data;
  }
};

const getIdentity = function(accessToken) {
  try {
    return HTTP.get('https://graph.microsoft.com/v1.0/me', {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: 'application/json',
        'User-Agent': userAgent,
      },
    }).data;
  } catch (error) {
    throw _.extend(new Error(`Failed to fetch identity from Microsoft Office365. ${error.message}`), { response: error.response });
  }
};

OAuth.registerService('office365', 2, null, function(query) {
  const { access_token: accessToken, refresh_token: refreshToken, expires_in: expiresIn } = getTokenResponse(query);
  const expiresAt = +new Date() + 1000 * expiresIn;
  const identity = getIdentity(accessToken);
  return {
    serviceData: {
      id: identity.id,
      accessToken: OAuth.sealSecret(accessToken),
      refreshToken,
      expiresAt,
      displayName: identity.displayName,
      givenName: identity.givenName,
      surname: identity.surname,
      username: identity.userPrincipalName && identity.userPrincipalName.split('@')[0],
      userPrincipalName: identity.userPrincipalName,
      mail: identity.mail,
      jobTitle: identity.jobTitle,
      mobilePhone: identity.mobilePhone,
      businessPhones: identity.businessPhones,
      officeLocation: identity.officeLocation,
      preferredLanguage: identity.preferredLanguage,
    },
    options: { profile: { name: identity.givenName } },
  };
});

Office365.retrieveCredential = function(credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};
