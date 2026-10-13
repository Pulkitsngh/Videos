/*
 * Navjyoti Tracker – settings for LNB Group Microsoft 365.
 * Copy this file to config.js (next to it), fill in the three values, then run build_desktop_app.py.
 * These values are not secrets: they only say which tenant, app registration and SharePoint site to use.
 * Access is controlled by Microsoft 365 sign-in and SharePoint permissions.
 */
window.NAVJYOTI_CONFIG = {
  // Microsoft Entra admin center → Overview → Tenant ID
  tenantId: 'PASTE-TENANT-ID-HERE',
  // Entra admin center → App registrations → Navjyoti Tracker → Application (client) ID
  clientId: 'PASTE-CLIENT-ID-HERE',
  // The SharePoint site that holds the Navjyoti lists (created by Create-NavjyotiLists.ps1)
  siteUrl: 'https://YOUR-TENANT.sharepoint.com/sites/NavjyotiProduction',
  // How often each PC checks SharePoint for other people's changes, in seconds (minimum 10)
  pollSeconds: 30,
  // Optional. Microsoft Graph permissions to ask for. Leave as is unless IT grants the app one site only
  // (then use ['https://graph.microsoft.com/Sites.Selected', 'User.Read'], see the guide's security section).
  // scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All', 'User.Read'],
};
