/**
 * Handles interactions with the Strava API.
 */

/**
 * Fetches recent activities from Strava for the authenticated athlete.
 * 
 * @return {Array} List of activity objects.
 */
function fetchActivities() {
  const token = getAccessToken();
  const url = 'https://www.strava.com/api/v3/athlete/activities?per_page=30'; // Fetch last 30 activities
  
  const options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options); // Using UrlFetchApp as per spec
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error('Error fetching activities: ' + e);
    return [];
  }
}

/**
 * Fetches photos for a specific activity.
 * 
 * @param {number} activityId - The ID of the activity.
 * @return {string|null} The URL of the primary photo, or null if none found.
 */
function fetchActivityPhotos(activityId) {
  const token = getAccessToken();
  // Ensure we get the high-res photo if available, though endpoint usually returns list
  const url = `https://www.strava.com/api/v3/activities/${activityId}/photos?size=5000`; 
  
  const options = {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const photos = JSON.parse(response.getContentText());
    
    if (photos && photos.length > 0) {
      // Return the first photo's URL. Prefer 'approx_width' > 0 if structure varies,
      // but standard response has urls property. 
      // Strava photo object structure: { urls: { "100": "...", "600": "..." } }
      // We want the largest available.
      const photo = photos[0];
      if (photo.urls) {
        // Get the largest size available key
        const sizes = Object.keys(photo.urls).sort((a, b) => Number(b) - Number(a));
        return photo.urls[sizes[0]];
      }
    }
    return null;
  } catch (e) {
    console.error(`Error fetching photos for activity ${activityId}: ` + e);
    return null;
  }
}
