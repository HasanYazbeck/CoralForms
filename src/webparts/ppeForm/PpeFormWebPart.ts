import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';

import PpeForm from "./components/PpeForm";
import { IPpeFormWebPartProps } from "./components/IPpeFormProps";
import { IGraphResponse, IGraphUserResponse } from "../../Interfaces/ICommon";

import { IUser } from "../../Interfaces/IUser";
import { SPHelpers } from '../../Classes/SPHelpers';
import { IPPEItem } from "../../Interfaces/IPPEItem";
import { IPPEItemDetails } from "../../Interfaces/IPPEItemDetails";
import { ICoralFormsList } from "../../Interfaces/ICoralFormsList";
import { SPCrudOperations } from "../../Classes/SPCrudOperations";

export default class PpeFormWebPart extends BaseClientSideWebPart<IPpeFormWebPartProps> {

  private spCrudOperations: SPCrudOperations;
  private spHelpers: SPHelpers = new SPHelpers();
  private _isDarkTheme: boolean = false;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  private _users: IUser[] = [];
  private _ppeItems: IPPEItem[] = [];
  private _ppeItemsDetails: IPPEItemDetails[] = [];
  private _coralFormsList: ICoralFormsList = { Id: "" };
  private _hasFetchedUsers: boolean = false;
  private _isLoading: boolean = true;


  protected async onInit(): Promise<void> {
    await super.onInit();
    if (!this._hasFetchedUsers) {

      this._isLoading = true;
      this.render(); // show spinner before data loads

      await this._getUsers();
      // await this.getPPEItems();
      await this.getCoralFormsList();
      await this.getPPEItemsDetails();
      
      // await this._getBatchUserImages();
      this._hasFetchedUsers = true;

      this._isLoading = false;
      this.render(); // re-render with data after loading
    }

    return new Promise<void>(async (resolve, reject) => {
      this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();
      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
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

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  private async _getUsers(): Promise<void> {
    let users: IUser[] = [];
    let endpoint: string | null =
      "/users?$select=id,displayName,mail,department,jobTitle,mobilePhone,officeLocation&$expand=manager($select=id,displayName)";

    try {
      do {
        const client: MSGraphClientV3 =
          await this.context.msGraphClientFactory.getClient("3");
        const response: IGraphResponse = await client.api(endpoint).get();

        if (response?.value && Array.isArray(response.value)) {
          const seenIds = new Set<string>();
          const mappedUsers = response.value
            .filter((u: IGraphUserResponse) => u.mail)
            .filter(
              (user) =>
                user.mail &&
                !user.mail?.toLowerCase().includes("healthmailbox") &&
                !user.mail?.toLowerCase().includes("softflow-intl.com") &&
                !user.mail?.toLowerCase().includes("sync")
            )
            .filter((user) => {
              if (seenIds.has(user.id)) return false; // Duplicate found
              seenIds.add(user.id); // Add new Id to Set
              return true;
            })
            .map(
              (user: IGraphUserResponse) =>
              ({
                id: user.id,
                displayName: user.displayName,
                email: user.mail,
                jobTitle: user.jobTitle,
                department: user.department,
                officeLocation: user.officeLocation,
                mobilePhone: user.mobilePhone,
                profileImageUrl: undefined, // will load later
                isSelected: false,
                manager: user.manager
                  ? {
                    id: user.manager.id,
                    displayName: user.manager.displayName,
                  }
                  : undefined,
              } as IUser)
            );

          users.push(...mappedUsers);
          endpoint = response["@odata.nextLink"] || null;
        } else {
          break;
        }
      } while (endpoint);
      this._users = users;

      // const uniqueDepartments = Array.from(
      //   new Set(
      //     users
      //       .map((user: any) => user.department)
      //       .filter((dept) => dept && dept.trim() !== "")
      //       .map((dept) => (dept ? this._helpers.CamelString(dept) : "")) // format in CamelCase
      //   )
      // );
      // // Assign to _departments (or push if _departments already has items)
      // this._departments = uniqueDepartments.sort((a, b) => a.localeCompare(b));

      // const uniqueJobTitles = Array.from(
      //   new Set(
      //     users
      //       .map((user: any) => user.jobTitle)
      //       .filter((jobTitle) => jobTitle && jobTitle.trim() !== "")
      //       .map((jobTitle) =>
      //         jobTitle ? this._helpers.CamelString(jobTitle) : ""
      //       ) // format in CamelCase
      //   )
      // );
      // // Assign to _departments (or push if _departments already has items)
      // this._jobTitles = uniqueJobTitles.sort((a, b) => a.localeCompare(b));
    } catch (error) {
      console.error("Error fetching users:", error);
      this._users = [];
    }
  }

  // public async _getBatchUserImages(): Promise<void> {
  //   if (!this._users || this._users.length === 0) return;

  //   const client: MSGraphClientV3 =
  //     await this.context.msGraphClientFactory.getClient("3");
  //   const batchSize = 20;

  //   for (let i = 0; i < this._users.length; i += batchSize) {
  //     const batch = this._users.slice(i, i + batchSize);
  //     const batchRequests = batch.map((user, index) => ({
  //       id: `${i + index}`, // Keep real index for mapping back
  //       method: "GET",
  //       url: `/users/${user.id}/photo/$value`,
  //     }));

  //     try {
  //       const batchResponse = await client
  //         .api("/$batch")
  //         .post({ requests: batchRequests });

  //       batchResponse.responses.forEach((resp: any) => {
  //         if (resp.status === 200) {
  //           const userIndex = parseInt(resp.id);
  //           // Convert base64 or binary body depending on Graph response
  //           const imageBytes = resp.body;
  //           if (imageBytes) {
  //             const byteArray = new Uint8Array(imageBytes);
  //             const blob = new Blob([byteArray], { type: "image/jpeg" });
  //             this._users[userIndex].profileImageUrl =
  //               URL.createObjectURL(blob);
  //           }
  //         }
  //       });
  //     } catch (err) {
  //       console.error("Batch photo fetch failed:", err);
  //     }
  //   }
  // }

  public getCoralFormsList = async (): Promise<void> => {
    let result: ICoralFormsList = { Id: "" };
    try {
      const searchFormName = "PERSONAL PROTECTIVE EQUIPMENT";
      const searchEscaped = searchFormName.replace(/'/g, "''");
      const query: string = `?$select=Id,Title,hasInstructionForUse,hasWorkflow,Created` +
        `&$filter=substringof('${searchEscaped}', Title)`;
      this.spCrudOperations = new SPCrudOperations(this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl, 'CoralFormsList', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this._users !== undefined && this._users.length > 0 ? this._users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }

              result = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                hasInstructionForUse: obj.hasInstructionForUse !== undefined ? obj.hasInstructionForUse : undefined,
                hasWorkflow: obj.hasWorkflow !== undefined ? obj.hasWorkflow : undefined,
                Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
              }
            }
          });
          this._coralFormsList = result;
          // this.setState({ CoralFormsList: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

  public getPPEItems = async (): Promise<void> => {
    const result: IPPEItem[] = [];
    try {
      const query: string = `?$select=Id,Title,Required,hasInstructionForUse,hasWorkflow,Created`;
      // `PPEDetails/Id,PPEDetails/Title&$expand=PPEDetails`;
      this.spCrudOperations = new SPCrudOperations(this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl, 'PPEItems', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this._users !== undefined && this._users.length > 0 ? this._users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {// Convert string to Date first
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }
              const temp: IPPEItem = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                // hasInstructionForUse: obj.hasInstructionForUse !== undefined ? obj.hasInstructionForUse : undefined,
                // hasWorkflow: obj.hasWorkflow !== undefined ? obj.hasWorkflow : undefined,
                Title: obj.Title !== undefined && obj.Title !== null ? obj.Title : undefined,
                Required: obj.Required !== undefined ? obj.Required : undefined,
                // PPEDetails: obj.PPEDetails !== undefined ? obj.PPEDetails : undefined,
              }


              console.log(temp);
              // Get PPEDetails for each one (Types, Sizez )
              result.push(temp);
            }
          });
          this._ppeItems = result;
          // this.setState({ PPEItems: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

  public getPPEItemsDetails = async (): Promise<void> => {
    const result: IPPEItemDetails[] = [];
    try {
      const query: string = `?$select=Id,Title,PPEItem,Types,Sizes,Created,` +
        `PPEItem/Id,PPEItem/Title,Types/Id,Types/Title&$expand=PPEItem,Types`;
      this.spCrudOperations = new SPCrudOperations(this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl, 'PPEItemsDetails', query);
      await this.spCrudOperations._getItemsWithQuery()
        .then((data) => {
          data.map((obj) => {
            if (obj !== undefined) {
              const createdBy: IUser | undefined = this._users !== undefined && this._users.length > 0 ? this._users.filter(user => user.id.toString() === obj.AuthorId.toString())[0] : undefined;
              let created: Date | undefined;
              if (obj.Created !== undefined) {// Convert string to Date first
                created = new Date(this.spHelpers.adjustDateForGMTOffset(obj.Created));
              }
              const temp: IPPEItemDetails = {
                Id: obj.Id !== undefined && obj.Id !== null ? obj.Id : undefined,
                CreatedBy: createdBy !== undefined ? createdBy : undefined,
                Created: created !== undefined ? created : undefined,
                PPEItem: obj.PPEItem !== undefined ? {
                  Id: obj.PPEItem.Id !== undefined && obj.PPEItem.Id !== null ? obj.PPEItem.Id : undefined,
                  Title: obj.PPEItem.Title !== undefined && obj.PPEItem.Title !== null ? obj.PPEItem.Title : undefined,

                  // Required: obj.PPEItem.Required !== undefined ? obj.PPEItem.Required : undefined,
                  // Brands: obj.PPEItem.Brands !== undefined && obj.PPEItem.Brands !== null ? obj.PPEItem.Brands.split(",") : undefined,

                } : undefined,
                Types: obj.Types !== undefined && obj.Types !== null ? obj.Types : undefined,
                Sizes: obj.Sizes !== undefined && obj.Sizes !== null ? obj.Sizes.split(",") : undefined,
              };
              // console.log(temp);
              // Get PPEDetails for each one (Types, Sizez )
              result.push(temp);
            }
          });
          this._ppeItemsDetails = result;
          // this.setState({ PPEItems: result });
        })
        .catch(error => {
          console.error('An error has occurred while retrieving items!', error);
        });
    } catch (error) {
      console.error('An error has occurred!', error);
    }
  }

    public render(): void {

    const element: React.ReactElement<IPpeFormWebPartProps> =
      React.createElement(PpeForm, {
        context: this.context,
        Users: this._users,
        IsLoading: this._isLoading,
        ThemeColor: this._themeVariant?.palette?.themePrimary,
        IsDarkTheme: this._isDarkTheme,
        HasTeamsContext: !!this.context.sdks.microsoftTeams,
        PPEItems: this._ppeItems,
        CoralFormsList: this._coralFormsList,
        PPEItemDetails: this._ppeItemsDetails,
      });

    ReactDom.render(element, this.domElement);
  }
  
}
