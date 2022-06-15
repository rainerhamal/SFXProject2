import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import { DisplayMode } from '@microsoft/sp-core-library';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { Log } from '@microsoft/sp-core-library';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {

    
    // this.domElement.innerHTML = `
    // <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
    //   <div class="${styles.welcome}">
    //     <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
    //     <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
    //     <div>${this._environmentMessage}</div>
    //     <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
    //   </div>
    //   <div>
    //     <h3>Welcome to SharePoint Framework!</h3>
    //     <p>
    //     The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
    //     </p>
    //     <h4>Learn more about SPFx development:</h4>
    //       <ul class="${styles.links}">
    //         <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
    //         <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
    //         <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
    //       </ul>
    //   </div>
    // </section>`;
    const siteTitle : string = this.context.pageContext.web.title;
    const siteURL : string = this.context.pageContext.web.absoluteUrl;
    const siteRURL : string = this.context.pageContext.web.serverRelativeUrl;
    const currentUserName : string = this.context.pageContext.user.loginName;

    const pageMode: string = (this.displayMode === DisplayMode.Edit)
      ? 'You are in edit mode'
      : 'You are in read mode' ? 'You are in a test mode' : '';

    const environmentType : string = (Environment.type === EnvironmentType.ClassicSharePoint) ? 'You are running in a classic page' : 'You are running in a modern page';

    // this.context.statusRenderer.renderError(this.domElement, error);
    // this.context.statusRenderer.clearError(this.domElement);

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Web Part...");
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.domElement.innerHTML = `
      <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
          <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>

          <div>Site title: <strong>${escape(siteTitle)}</strong></div>
          <div>Page Mode: <strong>${escape(pageMode)}</strong></div>
          <div>Site Absolute URL: <strong>${escape(siteURL)}</strong></div>
          <div>Site Server-relative URL: <strong>${escape(siteRURL)}</strong></div>
          <div>Current User Sign-In Name: <strong>${escape(currentUserName)}</strong></div>
          <div>Environment: <strong>${escape(environmentType)}</strong></div>
          
        </div>

        
        <div>
          <h3>Welcome to SharePoint Framework!</h3>

          <p><b>This is your second web part!</b></p>
          <p>
          The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <button type="button">Show welcome message</button>
        </div>
      </section>`;

        this.domElement.getElementsByTagName("button")[0]
      .addEventListener('click', (event: any) => {
        event.preventDefault();
        alert('Welcome to the SharePoint Framework!');
      });
    }, 5000);

    Log.info('HelloWorld', 'message', this.context.serviceScope);
    Log.warn('HelloWorld', 'WARNING message', this.context.serviceScope);
    Log.error('HelloWorld', new Error('Error message'), this.context.serviceScope);
    Log.verbose('HelloWorld', 'VERBOSE message', this.context.serviceScope);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
