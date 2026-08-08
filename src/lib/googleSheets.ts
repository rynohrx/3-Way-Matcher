export interface DriveFile {
  id: string;
  name: string;
  modifiedTime?: string;
  webViewLink?: string;
}

export interface SheetTab {
  sheetId: number;
  title: string;
}

export interface SpreadsheetInfo {
  id: string;
  title: string;
  sheets: SheetTab[];
}

/**
 * Extracts spreadsheet ID from full Google Sheets URL or raw ID string.
 */
export function extractSpreadsheetId(input: string): string | null {
  if (!input) return null;
  const clean = input.trim();
  const urlMatch = clean.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }
  if (/^[a-zA-Z0-9-_]{25,}$/.test(clean)) {
    return clean;
  }
  return null;
}

/**
 * List Google Spreadsheets from user's Google Drive.
 */
export async function listDriveSpreadsheets(accessToken: string): Promise<DriveFile[]> {
  const query = encodeURIComponent("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name,modifiedTime,webViewLink)&pageSize=30&orderBy=modifiedTime%20desc`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || 'Failed to list spreadsheets from Google Drive.');
  }

  const data = await response.json();
  return data.files || [];
}

/**
 * Get Spreadsheet Title and List of Tabs.
 */
export async function getSpreadsheetDetails(
  accessToken: string,
  spreadsheetId: string
): Promise<SpreadsheetInfo> {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || 'Failed to fetch Google Sheet details.');
  }

  const data = await response.json();
  const title = data.properties?.title || 'Untitled Spreadsheet';
  const sheets: SheetTab[] = (data.sheets || []).map((s: any) => ({
    sheetId: s.properties?.sheetId,
    title: s.properties?.title || 'Sheet1',
  }));

  return { id: spreadsheetId, title, sheets };
}

/**
 * Get 2D Cell Values from a Spreadsheet range/tab.
 */
export async function getSpreadsheetValues(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string
): Promise<any[][]> {
  const range = encodeURIComponent(`'${sheetName}'!A1:ZZ500`);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || 'Failed to fetch spreadsheet cell values.');
  }

  const data = await response.json();
  return data.values || [];
}

/**
 * Create a new Google Spreadsheet in the user's Drive.
 */
export async function createGoogleSheet(
  accessToken: string,
  title: string,
  headers: string[],
  rows: any[][]
): Promise<{ spreadsheetId: string; spreadsheetUrl: string }> {
  const response = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      properties: {
        title,
      },
      sheets: [
        {
          properties: {
            title: 'Audit Report',
          },
          data: [
            {
              startRow: 0,
              startColumn: 0,
              rowData: [headers, ...rows].map((row) => ({
                values: row.map((val) => ({
                  userEnteredValue:
                    typeof val === 'number'
                      ? { numberValue: val }
                      : typeof val === 'boolean'
                      ? { boolValue: val }
                      : { stringValue: String(val ?? '') },
                })),
              })),
            },
          ],
        },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || 'Failed to create new Google Sheet.');
  }

  const data = await response.json();
  return {
    spreadsheetId: data.spreadsheetId,
    spreadsheetUrl: data.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${data.spreadsheetId}/edit`,
  };
}

/**
 * Append rows to an existing Google Spreadsheet tab.
 */
export async function appendGoogleSheetValues(
  accessToken: string,
  spreadsheetId: string,
  sheetName: string,
  rows: any[][]
): Promise<void> {
  const range = encodeURIComponent(`'${sheetName}'!A1`);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}:append?valueInputOption=USER_ENTERED`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        values: rows,
      }),
    }
  );

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || 'Failed to append rows to Google Sheet.');
  }
}
