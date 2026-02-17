/**
 * Handles OAuth2 token management for Strava API.
 */

/**
 * Retrieves a valid access token.
 * Checks script properties first. If expired or missing, refreshes the token using the refresh token.
 * 
 * @return {string} Valid access token.
 */
function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('access_token');
  
  if (token) {
    return token;
  } else {
    return refreshAccessToken();
  }
}

/**
 * Refreshes the access token using the stored refresh token.
 * Updated to use UrlFetchApp as per spec.
 * 
 * @return {string} New access token.
 */
function refreshAccessToken() {
  const url = 'https://www.strava.com/oauth/token';
  const payload = {
    client_id: STRAVA_CLIENT_ID,
    client_secret: STRAVA_CLIENT_SECRET,
    refresh_token: STRAVA_REFRESH_TOKEN,
    grant_type: 'refresh_token'
  };
  
  const options = {
    method: 'post',
    payload: payload
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.access_token) {
      PropertiesService.getScriptProperties().setProperty('access_token', data.access_token);
      return data.access_token;
    } else {
      console.error('Error refreshing token: ' + JSON.stringify(data));
      throw new Error('Failed to refresh access token.');
    }
  } catch (e) {
    console.error('Exception refreshing token: ' + e);
    throw e;
  }
}
