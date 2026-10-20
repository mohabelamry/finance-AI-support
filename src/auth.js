import {PublicClientApplication} from '@azure/msal-browser';

// Page Microsoft returns to after sign-in. It hands the result back to the
// dashboard window and closes the popup (see auth/index.html).
export const redirectUri = () => `${window.location.origin}/auth/`;

// Optional site-wide default, set as a Cloudflare build variable, so people
// only need to click "Sign in with Microsoft".
const defaultClientId = (import.meta.env.VITE_MS_CLIENT_ID || '').trim();
const defaultTenant = (import.meta.env.VITE_MS_TENANT_ID || '').trim();

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DOMAIN = /^[a-z0-9-]+(\.[a-z0-9-]+)+$/i;
const clients = new Map();

function resolve(config) {
  const clientId = (config.clientId || '').trim() || defaultClientId;
  const tenant = (config.tenant || '').trim() || defaultTenant || 'common';
  if (!GUID.test(clientId)) throw Error('Enter the Application (client) ID from your Microsoft Entra app registration.');
  if (!(GUID.test(tenant) || DOMAIN.test(tenant) || ['common', 'organizations', 'consumers'].includes(tenant))) {
    throw Error('Tenant must be a Directory (tenant) ID, a domain such as company.com, or left empty.');
  }
  return {clientId, tenant};
}

async function client(config) {
  const {clientId, tenant} = resolve(config);
  const key = `${tenant}|${clientId}`;
  if (!clients.has(key)) {
    const app = new PublicClientApplication({
      auth: {clientId, authority: `https://login.microsoftonline.com/${tenant}`, redirectUri: redirectUri()},
      cache: {cacheLocation: 'localStorage'},
    });
    await app.initialize();
    clients.set(key, app);
  }
  return clients.get(key);
}

// Turns Microsoft's technical errors into plain instructions.
function explain(e) {
  const text = `${e?.errorCode || ''} ${e?.message || ''}`;
  if (/popup_window_error|empty_window_error/i.test(text)) return Error('Your browser blocked the Microsoft sign-in window. Allow pop-ups for this site and try again.');
  if (/user_cancelled/i.test(text)) return Error('Sign-in was cancelled.');
  if (/AADSTS50011|redirect_uri/i.test(text)) return Error(`Microsoft rejected this site's address. In the app registration, add ${redirectUri()} as a Single-page application redirect URI.`);
  if (/AADSTS700016/i.test(text)) return Error('Microsoft could not find this Application (client) ID in that tenant. Check both IDs.');
  if (/AADSTS50194/i.test(text)) return Error('This app registration is single-tenant. Enter its Directory (tenant) ID in Settings.');
  if (/AADSTS9002326|AADSTS9002327/i.test(text)) return Error('The redirect URI is registered under "Web" instead of "Single-page application". Move it to the Single-page application platform.');
  if (/AADSTS65001|consent_required|AADSTS90094/i.test(text)) return Error('Your Microsoft administrator needs to approve Mail.Send for this app (Grant admin consent).');
  if (/timed_out|timeout/i.test(text)) return Error(`The sign-in window did not report back. Check that ${redirectUri()} is registered as a Single-page application redirect URI.`);
  return e instanceof Error ? e : Error(String(e));
}

const findAccount = (app, username) => app.getAllAccounts().find(a => a.username?.toLowerCase() === username?.toLowerCase());

export async function signIn(config) {
  try {
    const app = await client(config);
    const result = await app.loginPopup({scopes: ['User.Read', 'Mail.Send'], prompt: 'select_account'});
    return result.account.username;
  } catch (e) { throw explain(e); }
}

// Forgets the account on this browser only; no Microsoft sign-out window.
export async function disconnect(config) {
  try {
    const app = await client(config);
    const account = findAccount(app, config.username);
    if (account) await app.clearCache({account});
  } catch { /* nothing stored to remove */ }
}

export async function sendMessage(config, to, subject, body) {
  let app, result;
  try {
    app = await client(config);
    const account = findAccount(app, config.username);
    if (!account) throw Error('Sign in to the selected Microsoft 365 account first.');
    try {
      result = await app.acquireTokenSilent({account, scopes: ['Mail.Send']});
    } catch (e) {
      if (e.name === 'InteractionRequiredAuthError' || ['interaction_required', 'consent_required', 'login_required'].includes(e.errorCode)) {
        result = await app.acquireTokenPopup({account, scopes: ['Mail.Send']});
      } else throw e;
    }
  } catch (e) { throw explain(e); }
  const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: {Authorization: `Bearer ${result.accessToken}`, 'Content-Type': 'application/json'},
    body: JSON.stringify({message: {subject, body: {contentType: 'Text', content: body}, toRecipients: [{emailAddress: {address: to}}]}, saveToSentItems: true}),
  });
  if (!response.ok) throw Error(`Microsoft 365 rejected the send request (${response.status}). Check that the account has an Exchange Online mailbox and Mail.Send consent.`);
}
