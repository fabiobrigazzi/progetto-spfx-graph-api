import { Configuration, PublicClientApplication } from '@azure/msal-browser';

export const msalConfig: Configuration = {
  auth: {
    clientId: '90d86df7-caac-428c-a83c-8d7a32042224', 
    authority: 'https://login.microsoftonline.com/93f7330a-2956-4a39-b10b-054fa00a1cde',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin
  },
  cache: {
    cacheLocation: 'sessionStorage' // o 'localStorage'
    //storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: ['User.Read', 'Mail.Read'] // Definisci gli scope necessari
};

// Inizializza MSAL instance
export const msalInstance = new PublicClientApplication(msalConfig);