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
  const expiresAt = parseInt(props.getProperty('expires_at') || '0');
  
  // Check if token exists AND is valid (with 5-minute buffer)
  const nowSeconds = Math.floor(Date.now() / 1000);
  
  if (token && expiresAt > (nowSeconds + 300)) {
    return token;
  } else {
    console.log('Access token missing or expired. Refreshing...');
    return refreshAccessToken();
  }
}

/**
 * Refreshes the access token using the stored refresh token.
 * Handles token rotation by storing new refresh tokens if provided.
 * 
 * @return {string} New access token.
 */
function refreshAccessToken() {
  const props = PropertiesService.getScriptProperties();
  // Prioritize the stored refresh token (from rotation) over the hardcoded one
  const currentRefreshToken = props.getProperty('refresh_token') || STRAVA_REFRESH_TOKEN;
  
  const url = 'https://www.strava.com/oauth/token';
  const payload = {
    client_id: STRAVA_CLIENT_ID,
    client_secret: STRAVA_CLIENT_SECRET,
    refresh_token: currentRefreshToken,
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
      props.setProperty('access_token', data.access_token);
      
      // Store expiration time if provided
      if (data.expires_at) {
        props.setProperty('expires_at', String(data.expires_at));
      }

      // HANDLE TOKEN ROTATION: Store the new refresh token if provided
      if (data.refresh_token) {
        console.log('New refresh token received. Updating...');
        props.setProperty('refresh_token', data.refresh_token);
      }
      
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
