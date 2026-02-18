/**
 * Handles interactions with the Strava API.
 */

/**
 * Fetches recent activities from Strava for the authenticated athlete.
 * 
 * @param {number} page - The page number to fetch (default: 1).
 * @param {number} perPage - Number of items per page (default: 30).
 * @return {Array} List of activity summary objects.
 */
function fetchActivities(page = 1, perPage = 30) {
  const token = getAccessToken();
  const url = `https://www.strava.com/api/v3/athlete/activities?page=${page}&per_page=${perPage}`;
  
  const options = {
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true // We need to check status code manually
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 429) {
      throw new Error('RATE_LIMIT_EXCEEDED');
    }
    
    if (code !== 200) {
      console.error(`Error fetching activities. Code: ${code}. Response: ${response.getContentText()}`);
      return [];
    }
    
    return JSON.parse(response.getContentText());
  } catch (e) {
    // Re-throw if it's our specific rate limit error
    if (e.message === 'RATE_LIMIT_EXCEEDED') throw e;
    
    console.error('Exception fetching activities: ' + e);
    return [];
  }
}

/**
 * Fetches detailed activity data and photos for multiple activities in parallel.
 * This optimizes performance by batching requests.
 * 
 * @param {Array<string>} activityIds - List of activity IDs to fetch.
 * @return {Object} Map where key is activityId and value is an object { details: Object, photos: Array }.
 */
function fetchActivitiesDetailsParallel(activityIds) {
  if (!activityIds || activityIds.length === 0) return {};
  
  const token = getAccessToken();
  const requests = [];
  
  // Create request objects for both Details and Photos for each activity
  activityIds.forEach(id => {
    // 1. Details Request
    requests.push({
      url: `https://www.strava.com/api/v3/activities/${id}?include_all_efforts=true`,
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    
    // 2. Photos Request
    requests.push({
      url: `https://www.strava.com/api/v3/activities/${id}/photos?size=5000`,
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  });
  
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    const results = {};
    
    // Check for rate limits in any of the responses
    for (const r of responses) {
      if (r.getResponseCode() === 429) {
        throw new Error('RATE_LIMIT_EXCEEDED');
      }
    }
    
    // Process responses in pairs (Details, Photos)
    for (let i = 0; i < activityIds.length; i++) {
      const id = activityIds[i];
      const detailsResp = responses[i * 2];
      const photosResp = responses[i * 2 + 1];
      
      let details = null;
      let photos = [];
      
      if (detailsResp.getResponseCode() === 200) {
        details = JSON.parse(detailsResp.getContentText());
      } else {
        console.error(`Error fetching details for ${id}: ${detailsResp.getResponseCode()}`);
      }
      
      if (photosResp.getResponseCode() === 200) {
        const photoData = JSON.parse(photosResp.getContentText());
        if (photoData && Array.isArray(photoData)) {
           photos = photoData.map(p => {
             if (p.urls) {
               const sizes = Object.keys(p.urls).sort((a, b) => Number(b) - Number(a));
               return p.urls[sizes[0]];
             }
             return null;
           }).filter(url => url !== null);
        }
      }
      
      results[id] = { details, photos };
    }
    
    return results;
  } catch (e) {
    if (e.message === 'RATE_LIMIT_EXCEEDED') throw e;
    console.error('Error in parallel fetch: ' + e);
    return {};
  }
}

/**
 * Kept for backward compatibility or single testing.
 */
function fetchActivityDetails(activityId) {
  const result = fetchActivitiesDetailsParallel([activityId]);
  return result[activityId] ? result[activityId].details : null;
}

function fetchActivityPhotos(activityId) {
  const result = fetchActivitiesDetailsParallel([activityId]);
  return result[activityId] ? result[activityId].photos : [];
}
