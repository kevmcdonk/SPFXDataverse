import * as React from 'react';
import styles from './SpfxDataverse.module.scss';
import { ISpfxDataverseProps } from './ISpfxDataverseProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export default class SpfxDataverse extends React.Component<ISpfxDataverseProps, {}> {
  private dataverseClient: AadHttpClient;
  private appRegistrationId: string = "49b1f515-aaab-426b-852a-0b6dff70e4dd";

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        .getClient(this.appRegistrationId)
        .then((client: AadHttpClient): void => {
          this.dataverseClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): React.ReactElement<ISpfxDataverseProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    this.dataverseClient
      .get('https://contoso-api-dp20191109.azurewebsites.net/api/Orders', AadHttpClient.configurations.v1)
      .then((res: HttpClientResponse): Promise<React.ReactElement<ISpfxDataverseProps>> => {
        return res.json();
      })
      .then((orders: any): React.ReactElement<ISpfxDataverseProps> => {
    

        return (
          <section className={`${styles.spfxDataverse} ${hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.welcome}>
              <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
              <h2>Well done, {escape(userDisplayName)}!</h2>
              <div>{environmentMessage}</div>
              <div>Web part property value: <strong>{escape(description)}</strong></div>
            </div>
            <div>
              <h3>Welcome to SharePoint Framework!</h3>
              <p>
                The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
              </p>
              <h4>Learn more about SPFx development:</h4>
              <ul className={styles.links}>
                <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
                <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
                <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
              </ul>
            </div>
          </section>
        );
      })
      .catch((error) => {
        return (
        <section className={`${styles.spfxDataverse} ${hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.welcome}>
              <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
              <h2>There was an error!</h2>
              <div>{error}</div>
            </div>
          </section> 
        );
      });

      return null;
  }
}
