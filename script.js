const env = 'https://xxxxxx.awmdm.com/api';
const clientId = 'XXXXXX';
const clientSecret = 'XXXXXX';
const apiKey = 'XXXXXX'
const tenantUrl = 'XXXXXX';


/**
 * Main Entry Point
 * ----------------
 * Runs the full pipeline in a clean, high-level way.
*/
function main(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Devices");
    deleteRows(sheet);

    const token = getToken();
    const devices = getDevices(token).Devices;

    const filteredDevices = filterRecentWindowsDevices(devices);
    const formattedData = formatDeviceData(filteredDevices);

    writeToSheet(sheet, formattedData);
}

/**
 * Delete Rows
 * -----------
 * Clears all rows in the "Devices" sheet except the header (first row).
 */
function deleteRows(sheet){
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 1) { 
        sheet.getRange(2, 1, lastRow - 1, lastColumn).range.clearContent();
    }
}

/**
 * Get OAuth token from WS1
 * ------------------------
 * Authenticates using the client ID and secret to retrieve a bearer token.
*/
function getToken() {
    const authHeader = Utilities.base64Encode(`${clientId}:${clientSecret}`);
    const tokenResponse = UrlFetchApp.fetch(`${tenantUrl}`,
    {
        method: 'post',
        headers: {
            'Authorization': `Basic ${authHeader}`,
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: {
            'grant_type': 'client_credentials'
        }
    });
    
    return JSON.parse(tokenResponse.getContentText()).access_token;
}

/**
 * Get Devices
 * -----------
 * Retrieves all devices from WS1 using the MDM search endpoint.
 */
function getDevices(token){
    var devices = JSON.parse(UrlFetchApp.fetch(`${env}/mdm/devices/search`, {
        method: 'get',
        headers: {
            'Accept': 'application/json',
            'Authorization': `Bearer ${token}`,
            'aw-tenant-code': `${apiKey}`
        }
    }));
    
    return devices;
}


/**
 * Filter Windows devices active within the last 14 days.
*/
function filterRecentWindowsDevices(devices) {
  const now = new Date();
  const fourteenDaysAgo = new Date(now);
  fourteenDaysAgo.setDate(now.getDate() - 14);

  return devices.filter(device => {
    const lastSeenDate = new Date(device.LastSeen.replace(" ", "T"));
    return device.Platform === "WinRT" && lastSeenDate >= fourteenDaysAgo;
  });
}


/**
 * Map WS1 device objects to simple rows for the sheet.
 */
function formatDeviceData(devices) {
  return devices.map(device => {
    const osVersion = getOSVersion(device.OperatingSystem, device.OSBuildVersion);
    const date = device.LastSeen.split("T")[0];
    return [
      device.SerialNumber,
      device.FriendlyNumber,
      device.UserName,
      device.Model,
      osVersion,
      date,
      device.EnrollmentStatus,
      device.ComplianceStatus,
      device.CompromisedStatus
    ];
  });
}

/**
 * Write formatted data to the Google Sheet.
 */
function writeToSheet(sheet, data) {
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}


/**
 * Get OS Version
 * --------------
 * Formats the operating system and build version into a single string.
*/
function getOSVersion(operatingSystem, buildVersion){
    const cleanOS = String(operatingSystem || "").substring(5);
    return `${cleanOS}.${buildVersion || ""}`;
}
