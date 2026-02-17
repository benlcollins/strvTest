# Spec: Strava-to-Sheets Apps Script Sync Engine

## 1. Project Overview
A Google Apps Script (GAS) project designed to automate the extraction of activity data and media links from the **Strava API** directly into a Google Spreadsheet. This solution replaces the need for an external server or Python environment.

---

## 2. Security & Secret Management
**Critical Requirement:** Sensitive credentials must be isolated from the functional logic to allow for safe code sharing or versioning.

### 2.1 File: `Config.gs` (The Secret File)
This file acts as your local `.env` or `secrets.json`. It is **never** to be shared.
* **Constants to define:**
    * `STRAVA_CLIENT_ID`: Your Strava App ID.
    * `STRAVA_CLIENT_SECRET`: Your Strava App Secret.
    * `STRAVA_REFRESH_TOKEN`: Your persistent refresh token.
    * `TARGET_SHEET_NAME`: The name of the tab (e.g., "Activities").

### 2.2 Property Service
The script must use `PropertiesService.getScriptProperties()` to store the volatile `access_token`.
* **Logic:** If the stored `access_token` is null or returns a 401 error, the script must trigger a refresh using the `STRAVA_REFRESH_TOKEN` from `Config.gs`.

---

## 3. Data Architecture & Logic

### 3.1 Strava API Integration
* **Service:** Use Google's native `UrlFetchApp`.
* **Endpoints:**
    * `GET https://www.strava.com/api/v3/athlete/activities`
    * `GET https://www.strava.com/api/v3/activities/{id}/photos` (to retrieve high-res image URLs).

### 3.2 Google Sheets Implementation
* **Formula Injection:** To ensure links are clickable, use:
  `range.setFormula('=HYPERLINK("' + photoUrl + '", "View Photo")');`
* **Deduplication:** * The script must check Column A (Activity ID) before appending.
    * If `activity_id` exists in the sheet, skip that record.

---

## 4. Technical File Structure
1.  **`Config.gs`**: Static credential storage.
2.  **`Auth.gs`**: Handles OAuth2 token refresh logic and header generation.
3.  **`StravaAPI.gs`**: Handles the fetching and parsing of activity and photo data.
4.  **`Main.gs`**: Coordinates the sync, handles sheet writing, and creates a Custom Menu.

---

## 5. Implementation Prompt for Anti-Gravity
> "Act as a Google Apps Script Expert. Build a script based on `spec.md` that:
> 1. Uses `UrlFetchApp` to pull my Strava activities.
> 2. Implements a `getAccessToken()` function that refreshes tokens using the constants in `Config.gs`.
> 3. For every activity, fetches the primary photo URL and writes it to the sheet using the `=HYPERLINK` formula.
> 4. Prevents duplicate entries by checking for existing Activity IDs in the sheet.
> 5. Adds an `onOpen()` function to create a spreadsheet menu named 'Strava Tools' with a 'Sync Now' button."