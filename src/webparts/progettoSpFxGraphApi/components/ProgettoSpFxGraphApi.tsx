import * as React from 'react';
import styles from './ProgettoSpFxGraphApi.module.scss';
import type { IProgettoSpFxGraphApiProps } from './IProgettoSpFxGraphApiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { PrimaryButton, Stack, Spinner, SpinnerSize } from '@fluentui/react';

// Interfacce per i dati
interface IUser {
  id: string;
  displayName: string;
  mail: string;
  jobTitle: string;
  officeLocation: string;
}

interface IEmail {
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: string;
}

interface IEvent {
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  location: {
    displayName: string;
  };
}

interface IProgettoSpFxGraphApiState {
  users: IUser[];
  emails: IEmail[];
  events: IEvent[];
  loading: boolean;
  error: string;
  currentUser: IUser | null;
}



export default class ProgettoSpFxGraphApi extends React.Component<IProgettoSpFxGraphApiProps, IProgettoSpFxGraphApiState> {
  
  constructor(props: IProgettoSpFxGraphApiProps) {
    super(props);
    
    this.state = {
      users: [],
      emails: [],
      events: [],
      loading: false,
      error: '',
      currentUser: null
    };
  }

  // Ottieni utente corrente
  private async getCurrentUser(): Promise<void> {
    this.setState({ loading: true, error: '' });
    
    try {
      const response = await this.props.graphClient
        .api('/me')
        .get();
      
      this.setState({ 
        currentUser: response,
        loading: false 
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore nel recupero utente: ${errore.message}`,
        loading: false 
      });
    }
  }

  // Ottieni lista utenti
  private async getUsers(): Promise<void> {
    this.setState({ loading: true, error: '' });
    
    try {
      const response = await this.props.graphClient
        .api('/users')
        .top(10) // Prendi solo i primi 10
        .select('id,displayName,mail,jobTitle,officeLocation') // Seleziona solo campi specifici
        .get();
      
      this.setState({ 
        users: response.value,
        loading: false 
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore nel recupero utenti: ${errore.message}`,
        loading: false 
      });
    }
  }

  // Ottieni email
  private async getEmails(): Promise<void> {
    this.setState({ loading: true, error: '' });
    
    try {
      const response = await this.props.graphClient
        .api('/me/messages')
        .top(5)
        .select('subject,from,receivedDateTime')
        .orderby('receivedDateTime DESC')
        .get();
      
      this.setState({ 
        emails: response.value,
        loading: false 
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore nel recupero email: ${errore.message}`,
        loading: false 
      });
    }
  }

  // Ottieni eventi calendario
  private async getCalendarEvents(): Promise<void> {
    this.setState({ loading: true, error: '' });
    
    try {
      const response = await this.props.graphClient
        .api('/me/events')
        .top(5)
        .select('subject,start,location')
        .orderby('start/dateTime')
        .get();
      
      this.setState({ 
        events: response.value,
        loading: false 
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore nel recupero eventi: ${errore.message}`,
        loading: false 
      });
    }
  }

  // Cerca utenti con filtro
  private async searchUsers(searchText: string): Promise<void> {
    if (!searchText) return;
    
    this.setState({ loading: true, error: '' });
    
    try {
      const response = await this.props.graphClient
        .api('/users')
        .filter(`startswith(displayName,'${searchText}') or startswith(mail,'${searchText}')`)
        .top(10)
        .select('id,displayName,mail,jobTitle')
        .get();
      
      this.setState({ 
        users: response.value,
        loading: false 
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore nella ricerca: ${errore.message}`,
        loading: false 
      });
    }
  }

  // Batch request (multiple chiamate in una)
  private async getBatchData(): Promise<void> {
    this.setState({ loading: true, error: '' });
    
    try {
      const batch = {
        requests: [
          {
            id: '1',
            method: 'GET',
            url: '/me'
          },
          {
            id: '2',
            method: 'GET',
            url: '/me/messages?$top=5'
          },
          {
            id: '3',
            method: 'GET',
            url: '/me/events?$top=5'
          }
        ]
      };

      const response = await this.props.graphClient
        .api('/$batch')
        .post(batch);

      // Processa le risposte
      const responses = response.responses;
      const userResponse = responses.find((r: any) => r.id === '1');
      const emailsResponse = responses.find((r: any) => r.id === '2');
      const eventsResponse = responses.find((r: any) => r.id === '3');

      this.setState({
        currentUser: userResponse?.body,
        emails: emailsResponse?.body?.value || [],
        events: eventsResponse?.body?.value || [],
        loading: false
      });
    } catch (error) {
      const errore = error as Error;
      this.setState({ 
        error: `Errore batch request: ${errore.message}`,
        loading: false 
      });
    }
  }
  
  public render(): React.ReactElement<IProgettoSpFxGraphApiProps> {
    const { users, emails, events, loading, error, currentUser } = this.state;

    return (
      <div className={styles.graphApiStyles}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Microsoft Graph API Examples</span>
              <p className={styles.subTitle}>SPFx + React + TypeScript</p>
              
              <Stack tokens={{ childrenGap: 10 }}>
                <PrimaryButton 
                  text="Get Current User" 
                  onClick={() => this.getCurrentUser()} 
                  disabled={loading}
                />
                <PrimaryButton 
                  text="Get Users" 
                  onClick={() => this.getUsers()} 
                  disabled={loading}
                />
                <PrimaryButton 
                  text="Get My Emails" 
                  onClick={() => this.getEmails()} 
                  disabled={loading}
                />
                <PrimaryButton 
                  text="Get Calendar Events" 
                  onClick={() => this.getCalendarEvents()} 
                  disabled={loading}
                />
                <PrimaryButton 
                  text="Get All Data (Batch)" 
                  onClick={() => this.getBatchData()} 
                  disabled={loading}
                />
              </Stack>

              {loading && (
                <div className={styles.loading}>
                  <Spinner size={SpinnerSize.large} label="Loading..." />
                </div>
              )}

              {error && (
                <div className={styles.error}>
                  {error}
                </div>
              )}

              {/* Current User */}
              {currentUser && (
                <div className={styles.section}>
                  <h3>Current User</h3>
                  <div className={styles.card}>
                    <strong>{currentUser.displayName}</strong>
                    <div>{currentUser.mail}</div>
                    <div>{currentUser.jobTitle}</div>
                    <div>{currentUser.officeLocation}</div>
                  </div>
                </div>
              )}

              {/* Users List */}
              {users.length > 0 && (
                <div className={styles.section}>
                  <h3>Users ({users.length})</h3>
                  {users.map((user) => (
                    <div key={user.id} className={styles.card}>
                      <strong>{user.displayName}</strong>
                      <div>{user.mail}</div>
                      <div>{user.jobTitle}</div>
                    </div>
                  ))}
                </div>
              )}

              {/* Emails */}
              {emails.length > 0 && (
                <div className={styles.section}>
                  <h3>Recent Emails ({emails.length})</h3>
                  {emails.map((email, index) => (
                    <div key={index} className={styles.card}>
                      <strong>{email.subject}</strong>
                      <div>From: {email.from.emailAddress.name}</div>
                      <div>{new Date(email.receivedDateTime).toLocaleString()}</div>
                    </div>
                  ))}
                </div>
              )}

              {/* Calendar Events */}
              {events.length > 0 && (
                <div className={styles.section}>
                  <h3>Upcoming Events ({events.length})</h3>
                  {events.map((event, index) => (
                    <div key={index} className={styles.card}>
                      <strong>{event.subject}</strong>
                      <div>Location: {event.location.displayName}</div>
                      <div>{new Date(event.start.dateTime).toLocaleString()}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
