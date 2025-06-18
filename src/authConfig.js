import { PublicClientApplication } from '@azure/msal-browser';

export const msalConfig = {
  auth: {
    clientId: '4d661288-c66e-45a8-a8a3-b76b795be7f9', // App (client) ID
    authority: 'https://login.microsoftonline.com/a00de4ec-48a8-43a6-be74-e31274e2060d', // Tenant ID
    redirectUri: 'https://finalfinalpwc40.netlify.app'
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true
  }
};

export const loginRequest = {
  scopes: [
    'User.Read',
    'Files.ReadWrite.All',
    'Sites.ReadWrite.All'
  ],
  prompt: 'login'
};

export const msalInstance = new PublicClientApplication(msalConfig);
