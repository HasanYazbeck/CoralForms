import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IPropertyFieldGroupOrPerson, PrincipalType, PropertyFieldPeoplePicker } from "@pnp/spfx-property-controls";
// import { IPTWFormProps } from './components/IPtwFormProps';
import { SPHelpers } from '../../Classes/SPHelpers';
import PTWFormHost from './components/PTWFormHost';
import { IPTWFormProps } from './components/IPTWFormProps';

export interface IPTWFormWebPartProps {
  useTargetAudience: boolean;
  targetAudience: IPropertyFieldGroupOrPerson[];
}

export default class PtwFormWebPart extends BaseClientSideWebPart<IPTWFormWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;;
  private _currentUserGroups: string[] = [];
  private _spHelpers: SPHelpers = new SPHelpers();
  // private _formStructure: IPTWForm = {}

  protected async onInit(): Promise<void> {
    await super.onInit();
    return new Promise<void>(async (resolve, reject) => {
      this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();
      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
      this._isDarkTheme = !!this._themeVariant?.isInverted;
      if (this.properties.useTargetAudience) {
        await this._loadCurrentUserGroups();
      }
      resolve();
    });
  }

  private async _loadCurrentUserGroups(): Promise<void> {
    try {
      const graphClient: MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient("3");

      // Get user's groups
      const userGroupsResponse = await graphClient
        .api("/me/memberOf/microsoft.graph.group")
        .version("v1.0")
        .select("id")
        .get();

      this._currentUserGroups = userGroupsResponse.value.map(
        (group: { id: string }) => group.id
      );

      // Also include the user's ID for direct user targeting
      const currentUserResponse = await graphClient
        .api("/me")
        .version("v1.0")
        .select("id")
        .get();

      this._currentUserGroups.push(currentUserResponse.id);
    } catch (error) {
      console.error("Error loading user groups:", error);
    }
  }

  private checkUserHasAccess(): boolean {
    if (
      !this.properties.targetAudience ||
      this.properties.targetAudience.length === 0
    ) {
      return true;
    }

    const currentUserId = this.context.pageContext.legacyPageContext.userId;
    const currentUserLogin =
      this.context.pageContext.user.loginName.toLowerCase();

    // Check if the current user is directly targeted by ID or login
    const isUserDirectlyTargeted = this.properties.targetAudience.some(
      (target) => {
        const targetId = target.id?.toLowerCase();
        const targetLogin = target.login?.toLowerCase();
        return targetId === currentUserId || targetLogin === currentUserLogin;
      }
    );

    if (isUserDirectlyTargeted) {
      return true;
    }

    // Check if the user is in any of the targeted groups
    return this.properties.targetAudience.some((target) => {
      return target.id ? this._currentUserGroups.includes(target.id) : false;
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public render(): void {
    const formId = this._spHelpers.getQueryNumber('formId');
    if (this.properties.useTargetAudience && this.properties.targetAudience?.length > 0 && !this.checkUserHasAccess()) {
      ReactDom.unmountComponentAtNode(this.domElement);
      return;
    }
    const element: React.ReactElement<IPTWFormProps> = React.createElement(
      PTWFormHost,
      {
        context: this.context,
        ThemeColor: this._themeVariant?.palette?.themePrimary,
        IsDarkTheme: this._isDarkTheme,
        useTargetAudience: this.properties.useTargetAudience,
        targetAudience: this.properties.targetAudience,
        formId,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (propertyPath === "targetAudience" && newValue) {
      const updatedAudience = (newValue as IPropertyFieldGroupOrPerson[]).map(
        (person) => ({
          id: person.id,
          login: person.login ?? person.email ?? person.id ?? "", // Ensure login is always a string
          fullName: person.fullName || "Unknown",
          email: person.email || "",
          imageUrl: person.imageUrl || "",
        })
      );

      this.properties.targetAudience = updatedAudience;

      // Refresh the property pane to reflect changes
      if (this.context.propertyPane.isPropertyPaneOpen()) {
        this.context.propertyPane.refresh();
      }

      this.render();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              // groupName: "Audience Targeting",
              groupFields: [
                PropertyPaneToggle("useTargetAudience", {
                  label: "Enable Target Audience",
                  checked: this.properties.useTargetAudience || false,
                }),
                PropertyFieldPeoplePicker("targetAudience", {
                  label: "Select Target Audience (Users or Groups)",
                  initialData: this.properties.targetAudience || [],
                  allowDuplicate: false,
                  principalType: [
                    PrincipalType.Users,
                    PrincipalType.SharePoint,
                    PrincipalType.Security,
                  ],
                  onPropertyChange: (
                    propertyPath: string,
                    newValue: IPropertyFieldGroupOrPerson[]
                  ) => {
                    // Create a new array reference to force change detection
                    [...newValue] = this.properties.targetAudience;

                    // Force property update
                    super.onPropertyPaneFieldChanged(
                      propertyPath,
                      [],
                      newValue
                    );

                    // Refresh the property pane if open
                    if (this.context.propertyPane.isPropertyPaneOpen()) {
                      this.context.propertyPane.refresh();
                    }

                    // Force a complete re-render
                    this.render();
                  },
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: (value: IPropertyFieldGroupOrPerson[]) => {
                    try {
                      // Validate your selections here
                      if (value && value.length > 150) {
                        return Promise.resolve(
                          "Maximum 150 users/groups allowed"
                        );
                      }

                      // Check for specific validation rules
                      const invalidEntries = value.filter(
                        (person) => !person.id
                      );
                      if (invalidEntries.length > 0) {
                        return Promise.resolve("Some entries are invalid");
                      }

                      // If no errors
                      return Promise.resolve("");
                    } catch (error) {
                      console.error("People Picker error:", error);
                      return Promise.resolve(
                        "An error occurred while validating selections"
                      );
                    }
                  },
                  deferredValidationTime: 0,
                  key: "targetAudiencePicker",
                  disabled: !this.properties.useTargetAudience,
                  multiSelect: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
