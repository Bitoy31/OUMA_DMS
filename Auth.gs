// ============================================
// AUTHENTICATION & USER MANAGEMENT
// ============================================

const SPREADSHEET_ID = '1VpOgnYBNMwqW58jClIRZQGB0rnR1u0Zx-VBpFSzTOf4'; // Replace with your actual Sheet ID
const AUTH_SHEET_NAME = 'Auth'; // Your auth sheet name
const CACHE_DURATION = 43200; // 12 hours in seconds

/**
 * Authenticates user against the Auth sheet
 * @param {string} email - User email
 * @param {string} password - User password
 * @returns {Object} { success, user, role, division, message }
 */
function authenticateUser(email, password) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(AUTH_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Auth sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    // Assuming columns: [Full Name, Role, Division, Email, Password]
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [fullName, role, division, sheetEmail, sheetPassword] = row;
      
      // Normalize email comparison (case-insensitive)
      if (sheetEmail && String(sheetEmail).toLowerCase().trim() === String(email).toLowerCase().trim()) {
        // Check password
        if (String(sheetPassword).trim() === String(password).trim()) {
          // Authentication successful
          return {
            success: true,
            user: {
              name: fullName,
              email: sheetEmail,
              role: normalizeRole(role),
              division: division || 'DIV1',
              rowIndex: i + 1 // For reference
            },
            message: 'Login successful'
          };
        } else {
          // Email found but password wrong
          return { success: false, message: 'Invalid email or password' };
        }
      }
    }
    
    // Email not found
    return { success: false, message: 'Invalid email or password' };
    
  } catch (error) {
    Logger.log('Auth Error: ' + error.toString());
    return { success: false, message: 'Authentication error: ' + error.toString() };
  }
}

/**
 * Normalizes role names to standard format
 * @param {string} roleText - Raw role from sheet
 * @returns {string} Normalized role code
 */
function normalizeRole(roleText) {
  if (!roleText) return 'DO';
  
  const text = String(roleText).toLowerCase().trim();
  
  const roleMap = {
    'administrator': 'ADMIN',
    'admin': 'ADMIN',
    'secretary': 'SECRETARY',
    'sec': 'SECRETARY',
    'under secretary': 'USEC',
    'usec': 'USEC',
    'assistant secretary': 'ASSEC',
    'assec': 'ASSEC',
    'deputy assistant secretary': 'DAS',
    'das': 'DAS',
    'director': 'DIRECTOR',
    'communication officer': 'COMMS',
    'comms': 'COMMS',
    'desk officer': 'DO',
    'do': 'DO'
  };
  
  for (const [key, value] of Object.entries(roleMap)) {
    if (text.includes(key)) return value;
  }
  
  return 'DO'; // Default to Desk Officer
}

/**
 * Normalizes division names to standard format
 * @param {string} divisionText - Raw division from sheet
 * @returns {string} Normalized division code
 */
function normalizeDivision(divisionText) {
  if (!divisionText) return 'DIV1';
  
  const text = String(divisionText).toLowerCase().trim();
  
  if (text.includes('i') && !text.includes('ii') && !text.includes('iii') && !text.includes('iv')) return 'DIV1';
  if (text.includes('ii') && !text.includes('iii') && !text.includes('iv')) return 'DIV2';
  if (text.includes('iii')) return 'DIV3';
  if (text.includes('iv')) return 'DIV4';
  if (text.includes('admin')) return 'ADMIN';
  
  return 'DIV1';
}

/**
 * Gets role-specific configuration
 * @param {string} role - User role code
 * @returns {Object} Role configuration with permissions
 */
function getRoleConfig(role) {
  const configs = {
    'ADMIN': {
      displayName: 'Administrator',
      color: '#0f2346',
      permissions: ['all'],
      visibleTabs: ['dashboard', 'newsubmission', 'mysubmissions', 'actionrequired', 'alldocuments', 'esignature', 'printing', 'outgoing', 'completed', 'archive', 'settings'],
      canApprove: true,
      canSign: false,
      canRelease: true,
      canViewAll: true
    },
    'SECRETARY': {
      displayName: 'Secretary',
      color: '#7c3aed',
      permissions: ['sign', 'release', 'view_all', 'view_actions'],
      visibleTabs: ['dashboard', 'actionrequired', 'alldocuments', 'esignature', 'outgoing', 'completed', 'settings'],
      canApprove: false,
      canSign: true,
      canRelease: true,
      canViewAll: true
    },
    'USEC': {
      displayName: 'Under Secretary',
      color: '#dc2626',
      permissions: ['approve', 'view_all', 'view_actions'],
      visibleTabs: ['dashboard', 'actionrequired', 'alldocuments', 'outgoing', 'completed', 'settings'],
      canApprove: true,
      canSign: false,
      canRelease: false,
      canViewAll: true
    },
    'ASSEC': {
      displayName: 'Assistant Secretary',
      color: '#ea580c',
      permissions: ['approve', 'view_all', 'view_actions'],
      visibleTabs: ['dashboard', 'actionrequired', 'alldocuments', 'outgoing', 'completed', 'settings'],
      canApprove: true,
      canSign: false,
      canRelease: false,
      canViewAll: true
    },
    'DAS': {
      displayName: 'Deputy Assistant Secretary',
      color: '#f59e0b',
      permissions: ['approve', 'view_all', 'view_actions'],
      visibleTabs: ['dashboard', 'actionrequired', 'alldocuments', 'outgoing', 'completed', 'settings'],
      canApprove: true,
      canSign: false,
      canRelease: false,
      canViewAll: true
    },
    'DIRECTOR': {
      displayName: 'Division Director',
      color: '#2563eb',
      permissions: ['approve', 'view_division', 'view_actions'],
      visibleTabs: ['dashboard', 'newsubmission', 'mysubmissions', 'actionrequired', 'alldocuments', 'completed', 'settings'],
      canApprove: true,
      canSign: false,
      canRelease: false,
      canViewAll: false
    },
    'COMMS': {
      displayName: 'Communication Officer',
      color: '#16a34a',
      permissions: ['release', 'view_all'],
      visibleTabs: ['dashboard', 'alldocuments', 'outgoing', 'completed', 'settings'],
      canApprove: false,
      canSign: false,
      canRelease: true,
      canViewAll: true
    },
    'DO': {
      displayName: 'Desk Officer',
      color: '#06b6d4',
      permissions: ['create', 'view_own', 'view_actions'],
      visibleTabs: ['dashboard', 'newsubmission', 'mysubmissions', 'actionrequired', 'settings'],
      canApprove: false,
      canSign: false,
      canRelease: false,
      canViewAll: false
    }
  };
  
  return configs[role] || configs['DO'];
}

/**
 * Gets current user session
 * @returns {Object|null} User object if logged in, null otherwise
 */
function getCurrentUser() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const userJson = userProperties.getProperty('currentUser');
    
    if (!userJson) return null;
    
    const user = JSON.parse(userJson);
    const loginTime = Number(userProperties.getProperty('loginTime')) || 0;
    const now = new Date().getTime();
    
    // Check if session expired (12 hours)
    if (now - loginTime > CACHE_DURATION * 1000) {
      logout(); // Clear expired session
      return null;
    }
    
    return user;
  } catch (error) {
    return null;
  }
}

/**
 * Creates user session
 * @param {Object} user - User object from authentication
 */
function setCurrentUser(user) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('currentUser', JSON.stringify(user));
    userProperties.setProperty('loginTime', String(new Date().getTime()));
  } catch (error) {
    Logger.log('Session Error: ' + error.toString());
  }
}

/**
 * Logs out user
 */
function logout() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('currentUser');
    userProperties.deleteProperty('loginTime');
  } catch (error) {
    Logger.log('Logout Error: ' + error.toString());
  }
}

/**
 * Server-side authentication check (called from frontend)
 * @param {string} email - User email
 * @param {string} password - User password
 * @returns {Object} Authentication result
 */
function serverAuth(email, password) {
  const result = authenticateUser(email, password);
  
  if (result.success) {
    setCurrentUser(result.user);
  }
  
  return {
    success: result.success,
    message: result.message,
    user: result.success ? {
      name: result.user.name,
      email: result.user.email,
      role: result.user.role,
      division: result.user.division
    } : null
  };
}

/**
 * Gets current user info (called from frontend)
 * @returns {Object} Current user or null
 */
function getCurrentUserInfo() {
  return getCurrentUser();
}

/**
 * Server logout (called from frontend)
 */
function serverLogout() {
  logout();
  return { success: true, message: 'Logged out successfully' };
}
