// ============================================
// GLOBAL CONFIGURATION
// ============================================
const SESSION_TIMEOUT = 12 * 60 * 60 * 1000; // 12 hours

// ============================================
// AUTHENTICATION FUNCTIONS
// ============================================

function authenticateUser(email, password) {
  Logger.log('=== LOGIN ATTEMPT ===');
  Logger.log('Email: ' + email);
  Logger.log('Password: ' + password);
  
  try {
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✓ Spreadsheet accessed: ' + ss.getName());
    
    // Get Auth sheet
    const authSheet = ss.getSheetByName('Auth');
    if (!authSheet) {
      Logger.log('✗ Auth sheet not found!');
      return { success: false, error: 'Auth sheet not found. Create a sheet named "Auth".' };
    }
    Logger.log('✓ Auth sheet found');

    // Get all data
    const data = authSheet.getDataRange().getValues();
    Logger.log('✓ Total rows in Auth sheet: ' + data.length);
    Logger.log('Sheet data: ' + JSON.stringify(data.slice(0, 5))); // Log first 5 rows for debugging

    // Search for user (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userName = String(row[0] || '').trim();
      const role = String(row[1] || '').trim();
      const division = String(row[2] || '').trim();
      const userEmail = String(row[3] || '').trim();
      const userPassword = String(row[4] || '').trim();

      Logger.log('Row ' + i + ': Email="' + userEmail + '" vs "' + email + '"');

      // Check if email and password match
      if (userEmail.toLowerCase() === email.toLowerCase() && userPassword === password) {
        Logger.log('✓✓✓ USER FOUND AND AUTHENTICATED: ' + userName);
        
        // Create session
        const session = {
          email: userEmail,
          name: userName,
          role: role,
          division: division || 'N/A',
          createdAt: new Date().getTime(),
          expiresAt: new Date().getTime() + SESSION_TIMEOUT
        };

        // Store in Properties Service
        const props = PropertiesService.getUserProperties();
        props.setProperty('userSession', JSON.stringify(session));
        Logger.log('✓ Session stored');

        return {
          success: true,
          user: {
            email: userEmail,
            name: userName,
            role: role,
            division: division
          }
        };
      }
    }

    Logger.log('✗ User not found with email: ' + email);
    return { success: false, error: 'Invalid email or password' };
  } catch (e) {
    Logger.log('✗✗✗ AUTH ERROR: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return { success: false, error: 'Authentication error: ' + e.toString() };
  }
}

/**
 * Get current user session
 */
function getCurrentUserSession() {
  try {
    const props = PropertiesService.getUserProperties();
    const sessionData = props.getProperty('userSession');

    if (!sessionData) {
      Logger.log('No session found');
      return null;
    }

    const session = JSON.parse(sessionData);

    // Check if session expired
    if (new Date().getTime() > session.expiresAt) {
      Logger.log('Session expired');
      props.deleteProperty('userSession');
      return null;
    }

    Logger.log('Session valid for: ' + session.email);
    return session;
  } catch (e) {
    Logger.log('Session Error: ' + e.toString());
    return null;
  }
}

/**
 * Get current user info
 */
function getCurrentUserEmail() {
  try {
    const session = getCurrentUserSession();
    if (session) {
      Logger.log('getCurrentUserEmail - Authenticated: ' + session.email);
      return {
        isAuthenticated: true,
        email: session.email,
        name: session.name,
        role: session.role,
        division: session.division
      };
    }
    Logger.log('getCurrentUserEmail - Not authenticated');
    return { isAuthenticated: false };
  } catch (e) {
    Logger.log('getCurrentUserEmail Error: ' + e.toString());
    return { isAuthenticated: false };
  }
}

/**
 * Logout user
 */
function logoutUser() {
  try {
    const props = PropertiesService.getUserProperties();
    props.deleteProperty('userSession');
    Logger.log('User logged out');
    return { success: true };
  } catch (e) {
    Logger.log('Logout Error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================
// SERVE HTML
// ============================================

function doGet() {
  try {
    Logger.log('=== PAGE LOAD ===');
    const session = getCurrentUserSession();

    if (!session) {
      Logger.log('No session - showing Login');
      return HtmlService.createTemplateFromFile('Login')
        .evaluate()
        .setTitle('OUMA Document Management System - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    Logger.log('Session exists - showing Index');
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Document Management System')
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/128/1091/1091223.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    Logger.log('doGet Error: ' + e.toString());
    return HtmlService.createHtmlOutput('Error: ' + e.toString());
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// DOCUMENT OPERATIONS
// ============================================

function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let totalDocs = 0;
    let pending = 0;
    let approvedToday = 0;
    let released = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) totalDocs++;

      const status = String(row[7] || '').toUpperCase();
      if (status.includes('FOR_')) pending++;
      if (status === 'COMPLETED') released++;

      try {
        const createdDate = new Date(row[11]);
        createdDate.setHours(0, 0, 0, 0);
        if (createdDate.getTime() === today.getTime() && status === 'COMPLETED') {
          approvedToday++;
        }
      } catch (e) {}
    }

    return { totalDocs, pending, approvedToday, released };
  } catch (e) {
    Logger.log('getDashboardStats Error: ' + e.toString());
    return null;
  }
}

function getMyDocuments() {
  try {
    const session = getCurrentUserSession();
    if (!session) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const docs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[9]).trim() === session.name || String(row[9]).trim() === session.email) {
        docs.push({
          id: String(row[0]),
          docId: String(row[1]),
          type: String(row[2]),
          subject: String(row[3]),
          addr: String(row[4]),
          division: String(row[5]),
          priority: String(row[6]),
          status: String(row[7]),
          createdAt: new Date(row[11]).toISOString(),
          submittedBy: String(row[9]),
          remarks: String(row[10])
        });
      }
    }

    return docs;
  } catch (e) {
    Logger.log('getMyDocuments Error: ' + e.toString());
    return [];
  }
}

function getAllDocuments() {
  try {
    const session = getCurrentUserSession();
    if (!session) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const docs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      docs.push({
        id: String(row[0]),
        docId: String(row[1]),
        type: String(row[2]),
        subject: String(row[3]),
        addr: String(row[4]),
        division: String(row[5]),
        priority: String(row[6]),
        status: String(row[7]),
        createdAt: new Date(row[11]).toISOString(),
        submittedBy: String(row[9]),
        remarks: String(row[10])
      });
    }

    return docs;
  } catch (e) {
    Logger.log('getAllDocuments Error: ' + e.toString());
    return [];
  }
}

function saveDocument(doc) {
  try {
    const session = getCurrentUserSession();
    if (!session) {
      return { success: false, error: 'Not authenticated' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Documents');
    if (!sheet) {
      return { success: false, error: 'Documents sheet not found' };
    }

    sheet.appendRow([
      doc.id,
      doc.docId,
      doc.type,
      doc.subject,
      doc.addr,
      doc.division,
      doc.priority,
      doc.status,
      doc.fileLink,
      session.name,
      doc.remarks,
      new Date(),
      doc.endorseTo || 'Director'
    ]);

    return { success: true, message: 'Document saved successfully' };
  } catch (e) {
    Logger.log('saveDocument Error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function _extractDriveFileId(raw) {
  if (!raw) return '';
  raw = String(raw).trim();
  if (/^[a-zA-Z0-9_-]{25,}$/.test(raw)) return raw;
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /thumbnail\?id=([a-zA-Z0-9_-]+)/
  ];
  for (const re of patterns) {
    const m = raw.match(re);
    if (m) return m[1];
  }
  return raw;
}
