import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { IPpeFormWebPartProps } from "./components/IPpeFormProps";
import PpeFormHost from "./components/PpeFormHost";
import { SPCrudOperations } from "../../Classes/SPCrudOperations";
import MicrosoftGraphService from '../../Classes/MicrosoftGraphServices';
import { IUser } from "../../Interfaces/IUser";

export default class PpeFormWebPart extends BaseClientSideWebPart<IPpeFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  public spCrudRef: SPCrudOperations;
  private graphService: MicrosoftGraphService;


  protected async onInit(): Promise<void> {
    await super.onInit();
    this.graphService = new MicrosoftGraphService(this.context);

    return new Promise<void>(async (resolve, reject) => {
      this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();
      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
      // try {
      //   await this._fetchSharePointSiteListsGUIDs();
      // } catch (e) {
      //   // swallow fetch error and continue
      //   this._listsGUIDs = new Map();
      // }
      resolve();
    });
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public async getGroupMembers(groupEmail: string): Promise<IUser[]> {
    const members = await this.graphService.getGroupMembersByEmail(groupEmail);
    const users: IUser[] = members.map(member => ({
      id: member.id,
      displayName: member.displayName,
      mail: member.mail,
      userPrincipalName: member.userPrincipalName,
      jobTitle: member.jobTitle,
      mobilePhone: member.mobilePhone,
      officeLocation: member.officeLocation,
      businessPhones: member.businessPhones, 
      department: member.department,  
      manager: member.manager? member.manager.displayName : undefined,
      photo: member.photo ? `data:${member.photo['@odata.mediaContentType']};base64,${member.photo['@odata.mediaContent']} ` : undefined
    }));
    console.log("Group Members:", users);
    return users;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  public render(): void {

    // Read formId from the page URL so the form can be deep-linked
    const getQueryNumber = (name: string): number | undefined => {
      try {
        const href = (window.top?.location?.href) || window.location.href;
        const v = new URL(href).searchParams.get(name) || undefined;
        const n = v != null ? Number(v) : NaN;
        return Number.isFinite(n) ? n : undefined;
      } catch { return undefined; }
    };
    const formId = getQueryNumber('formId');
    // this.showGroupMembers("testdevteam@softflow.com.lb");
    const element = React.createElement(PpeFormHost, {
      context: this.context,
      ThemeColor: this._themeVariant?.palette?.themePrimary,
      IsDarkTheme: this._isDarkTheme,
      HasTeamsContext: !!this.context.sdks.microsoftTeams,
      formId,
    });
    ReactDom.render(element, this.domElement);
  }
}


