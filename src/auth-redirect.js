// Runs in the Microsoft sign-in popup: passes the result back to the dashboard.
import {broadcastResponseToMainFrame} from '@azure/msal-browser/redirect-bridge';
broadcastResponseToMainFrame().catch(error => {
  document.getElementById('msg').textContent = 'Sign-in could not be completed. Close this window and try again.';
  console.error(error);
});
