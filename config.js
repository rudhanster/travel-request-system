// Configuration file for Travel Request System - MIT Manipal
// IMPORTANT: Fill in the values marked with YOUR_XXX_HERE

const CONFIG = {
    // ========================================
    // STEP 1: Azure AD Configuration
    // ========================================
    // Get these from Azure Portal → App Registrations → Your App
    
    auth: {
        // Application (client) ID - looks like: 12345678-1234-1234-1234-123456789abc
        clientId: 'ef99f120-ead7-4335-81a6-da9998797ab5',
        
        // Directory (tenant) ID - looks like: 87654321-4321-4321-4321-cba987654321
        // Format: https://login.microsoftonline.com/YOUR_TENANT_ID
        authority: 'b528bfe7-6f84-447e-aa6c-8ca39e485594',
        
        // Your application URL - ALREADY FILLED FOR YOU!
        redirectUri: 'https://learnermanipal.sharepoint.com/sites/TravelManagement/Travel%20web/travel-request-app.html'
    },
    
    // ========================================
    // STEP 2: SharePoint Configuration
    // ========================================
    // ALREADY FILLED FOR YOU based on your screenshot!
    
    sharepoint: {
        // Your SharePoint site URL
        siteUrl: 'https://learnermanipal.sharepoint.com/sites/TravelManagement',
        
        // Your SharePoint list name (must match exactly)
        listName: 'TravelRequests'
    },
    
    // ========================================
    // STEP 3: Admin Configuration
    // ========================================
    // Add email addresses of people who should have admin access
    // IMPORTANT: Use your actual institutional emails
    
    admins: [
        'anirudhan.c@manipal.edu'         // Add more as needed
        // Add more admin emails below
    ],
    
    // ========================================
    // STEP 4: Transport Email Configuration
    // ========================================
    // Email address where approved requests will be sent
    
    transport: {
        email: 'anirudhan.c@manipal.edu',      // Replace with actual transport email
        subject: 'Travel Requests for '               // Subject prefix (date will be appended)
    },
    
    // ========================================
    // Technical Configuration (Don't Change)
    // ========================================
    
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
    },
    
    scopes: {
        sharepoint: ['https://graph.microsoft.com/.default'],
        graph: ['User.Read', 'Mail.ReadWrite', 'Sites.ReadWrite.All']
    }
};

// ========================================
// MSAL Request Configuration (Don't Change)
// ========================================

const loginRequest = {
    scopes: CONFIG.scopes.graph
};

const tokenRequest = {
    scopes: CONFIG.scopes.sharepoint,
    forceRefresh: false
};

// ========================================
// Configuration Validation
// ========================================
// This will help you identify missing configuration

console.log('=== Travel Request System Configuration ===');
console.log('SharePoint Site:', CONFIG.sharepoint.siteUrl);
console.log('Redirect URI:', CONFIG.auth.redirectUri);

if (CONFIG.auth.clientId === 'YOUR_CLIENT_ID_HERE') {
    console.warn('⚠️ WARNING: Client ID not configured! Get it from Azure Portal.');
}

if (CONFIG.auth.authority.includes('YOUR_TENANT_ID_HERE')) {
    console.warn('⚠️ WARNING: Tenant ID not configured! Get it from Azure Portal.');
}

if (CONFIG.admins.some(email => email.includes('your.email'))) {
    console.warn('⚠️ WARNING: Admin emails not configured! Update with actual email addresses.');
}

console.log('==========================================');
