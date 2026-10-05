// Settings for the standalone (Netlify) version. Inside Claude the page ignores this file.
// Both values come from the Azure app registration described in docs/SETUP.md. They are public
// identifiers, not secrets: a single-page app has no client secret, and sign-in is limited to the
// Cludo tenant and the registered redirect URIs.
window.OUTREACH_CONFIG = {
  clientId: '',   // Application (client) ID
  tenantId: '',   // Directory (tenant) ID
  scopes: ['User.Read', 'Mail.Read', 'Files.ReadWrite.AppFolder'],
};
