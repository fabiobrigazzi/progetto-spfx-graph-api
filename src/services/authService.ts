import { msalInstance, loginRequest } from './authConfig';
import { AccountInfo, AuthenticationResult } from '@azure/msal-browser';

export class AuthService {
  private static instance: AuthService;

  private constructor() {
    // Inizializzazione MSAL
    //void msalInstance.initialize(); //1a soluzione
  }

  public async initialize(): Promise<void>{ //2a soluzione
    await msalInstance.initialize();
  }

  public static getInstance(): AuthService {
    if (!AuthService.instance) {
      AuthService.instance = new AuthService();
    }
    return AuthService.instance;
  }

  // Login interattivo
  public async login(): Promise<AuthenticationResult> {
    try {
      const response = await msalInstance.loginPopup(loginRequest);
      return response;
    } catch (error) {
      console.error('Login failed:', error);
      throw error;
    }
  }

  // Acquisizione token silente
  public async getToken(scopes: string[]): Promise<string> {
    const account = this.getAccount();
    
    if (!account) {
      throw new Error('No active account');
    }

    try {
      // Prova acquisizione silente
      const response = await msalInstance.acquireTokenSilent({
        scopes,
        account
      });
      return response.accessToken;
    } catch (error) {
      // Se fallisce, richiedi login interattivo
      const response = await msalInstance.acquireTokenPopup({
        scopes
      });
      return response.accessToken;
    }
  }

  // Ottieni l'account corrente
  private getAccount(): AccountInfo | null {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  // Logout
  public async logout(): Promise<void> {
    const account = this.getAccount();
    if (account) {
      await msalInstance.logoutPopup({ account });
    }
  }

  // Verifica se l'utente è autenticato
  public isAuthenticated(): boolean {
    return msalInstance.getAllAccounts().length > 0;
  }
}