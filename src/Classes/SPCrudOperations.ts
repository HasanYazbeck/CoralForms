import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { ISPItem } from "../Interfaces/Common/ISPItem";
import { IUser } from "../Interfaces/Common/IUser";
import { FieldTypeKind } from "../Enums/enums";
import { IPersonaProps } from "@fluentui/react";

export class SPCrudOperations {
  private listName: string;
  private siteUrl: string;
  private spHttpClient: SPHttpClient;
  private query?: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string, listName: string, query?: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
    this.listName = listName;
    this.query = query;
  }

  private escapeOData(value: string): string {
    return value.replace(/'/g, "''");
  }

  private isGuid(val?: string): boolean {
    if (!val) return false;
    const s = val.replace(/[{}]/g, '');
    return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$/.test(s);
  }

  private getListLocator(): string {
    if (!this.listName) throw new Error('List identifier not set (GUID or Title).');
    // If listGUID looks like a GUID, use getbyid; otherwise treat it as a Title
    if (this.isGuid(this.listName)) {
      const clean = this.listName.replace(/[{}]/g, '');
      return `getbyid('${clean}')`;
    }
    const title = this.escapeOData(this.listName);
    return `getbytitle('${title}')`;
  }

  private getListBaseUrl(): string {
    return `${this.siteUrl}/_api/web/lists/${this.getListLocator()}`;
  }

  public async _getSharePointListGUID(): Promise<string | undefined> {
    try {
      // Escape single quotes in title if provided
      const safeTitle = this.listName ? this.listName.replace(/'/g, "''") : undefined;
      const baseUrl: string = `${this.siteUrl}/_api/web/lists?$select=Id,Title,Fields`;
      const filter = safeTitle ? `&$filter=Title eq '${safeTitle}'` : '';
      const listUrl: string = `${baseUrl}${filter}`;
      // Resolve the list GUID by title first
      const response = await this.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);

      if (!response.ok) {
        return;
      }

      if (response.status === 200) {
        const responseData: any = await response.json();
        const data: any[] = responseData["value"] || [];
        if (data.length == 1) {
          return data[0].Id.toString();
        }
      } else {
        const responseError: any = await response.json();
        throw new Error(`Error retrieving items. Status: ${responseError.status}`);
      }
    } catch (e) {
      return;
    }
  }

  // Create List
  public async _createList(listDescription: string): Promise<void> {
    const listUrl: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')`;
    try {
      await this.spHttpClient
        .get(listUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 200) {
            alert("A List already exists with this name.");
            return;
          }
          if (response.status === 404) {
            const url: string = `${this.siteUrl}/_api/web/lists`;
            const listDefinition: any = {
              Title: this.listName,
              Description: listDescription,
              AllowContentTypes: true,
              BaseTemplate: 100,
              ContentTypesEnabled: true,
            };

            const spHttpClientOptions: ISPHttpClientOptions = {
              body: JSON.stringify(listDefinition),
            };
            this.spHttpClient
              .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
              .then((response: SPHttpClientResponse) => {
                if (response.status === 201) {
                  alert("A new List has been created successfully.");
                } else {
                  response.json().then((responseJson: JSON) => {
                    alert(
                      "Error Message" +
                      response.status +
                      " - " +
                      JSON.stringify(responseJson)
                    );
                  });
                }
              });
          } else {
            response.json().then((responseJson: JSON) => {
              alert(
                "Error Message" +
                response.status +
                " - " +
                JSON.stringify(responseJson)
              );
            });
          }
        });
    } catch (error) {
      console.error("Error creating item:", error);
      throw error;
    }
  }

  // Add columns to List
  public async _addColumnToList(columnName: string, columnType: FieldTypeKind): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/fields`;
    const columnDefinition: any = {
      Title: columnName,
      FieldTypeKind: columnType, // Change based on the type of column
      Required: false,
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(columnDefinition),
    };

    try {
      const response = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );
      if (response.status === 201) {
        alert("Column added successfully.");
      } else {
        const responseJson = await response.json();
        alert("Error adding column: " + JSON.stringify(responseJson));
      }
    } catch (error) {
      console.error("Error adding column:", error);
      throw error;
    }
  }

  // Delete columns from List
  public async _deleteColumnFromList(columnName: string): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/fields/getByTitle('${columnName}')`;

    try {
      const response = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*",
          },
        }
      );
      if (response.status === 204) {
        alert("Column deleted successfully.");
      } else {
        const responseJson = await response.json();
        alert("Error deleting column: " + JSON.stringify(responseJson));
      }
    } catch (error) {
      // console.error("Error deleting column:", error);
      throw error;
    }
  }

  // Insert item List
  public async _insertItem(item: any): Promise<number> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(item),
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        'OData-Version': '3.0',
      },
    };

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions);
      if (response.status === 201) {
        const created = await response.json();
        return created?.Id as number;
      } else {
        const errorText = await response.text();
        let parsed: any;
        try { parsed = JSON.parse(errorText); } catch { /* non-JSON */ }
        const message = parsed?.error?.message?.value || parsed?.error?.message ||
          errorText || `HTTP ${response.status}`;
        throw new Error(`Create Item failed: ${response.status} - ${message}`);
      }
    } catch (error) {
      throw error;
    }
  }

  // Add a new method that returns the created item (or at least the Id)
  public async _insertItemReturn<T = any>(item: any): Promise<T> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      },
      body: JSON.stringify(item),
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      spHttpClientOptions
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Insert failed: ${text}`);
    }
    return response.json() as Promise<T>;
  }

  // Get Items List
  public async _getItems(): Promise<any[]> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items`;

    try {
      const response = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (response.status === 200) {
        const responseData: any = await response.json();
        // console.log('Items retrieved successfully:', responseData);
        return responseData["value"];
      } else {
        const responseError: any = await response.json();
        // console.log(`Error retrieving items. Status: ${responseError.status}`, responseError);
        // alert("Error Message" + JSON.stringify(responseError));
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    } catch (error) {
      // console.error("Error Retreiving Items", error);
      throw error;
    }
  }

  // Get Items with query (ensure leading ?)
  public async _getItemsByListNameOrGuid(): Promise<any[]> {
    const qs = this.query ? (this.query.startsWith('?') ? this.query : `?${this.query}`) : '';
    const url: string = `${this.getListBaseUrl()}/items${qs}`;
    try {
      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.status === 200) {
        const responseData: any = await response.json();
        // console.log('Items retrieved successfully:', responseData.value);
        return responseData.value;
      } else {
        const responseError: any = await response.json();
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    }
    catch (error) {
      throw error;
    }
  }

  // Get Items List
  public async _getItemsWithQuery(): Promise<any[]> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items/${this.query}`;

    try {
      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.status === 200) {
        const responseData: any = await response.json();
        // console.log('Items retrieved successfully:', responseData.value);
        return responseData.value;
      } else {
        const responseError: any = await response.json();
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    } catch (error) {
      // console.error("Error Retreiving Items", error);
      throw error;
    }
  }

  // Get Item List By Id or Title
  public async _getItemById(id: string): Promise<ISPItem> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/items?filter=Id eq ${id}`;

    try {
      return this.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then((itemsList: any) => {
          const tempItem: any = itemsList.value[0];
          const listItem: ISPItem = tempItem as ISPItem; // Cast as interface ISPItem
          return listItem;
        }) as Promise<ISPItem>;
    } catch (error) {
      // console.error("Error Retreiving Item", error);
      throw error;
    }
  }

  // Update Item
  public async _updateItem(itemId: string, item: any): Promise<SPHttpClientResponse> {
    const url: string = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${itemId})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify(item),
    };
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );
      if (response.ok) {
        console.log("Item updated successfully");
        return response;
      } else {
        const errorResponse: any = await response.json();
        console.error(`Error updating item. Status: ${response.status}`, errorResponse);
        throw new Error(`Error updating item. Status: ${response.status}`);
      }
    } catch (error) {
      // console.error("Error updating item:", error);
      throw error;
    }
  }

  // Delete Item
  public async _deleteItem(itemId: number): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${itemId})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    };

    try {
      await this.spHttpClient.post(`${url}/undoCheckout`, SPHttpClient.configurations.v1, {});


      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );
      if (response.ok) {
        console.log("Item deleted successfully");
      } else {
        // const errorResponse: any = await response.json();
        // console.error(`Error deleting item. Status: ${response.status}`, errorResponse);
        throw new Error(`Error deleting item. Status: ${response.status}`);
      }
    } catch (error) {
      // console.error("Error deleting item:", error);
      throw error;
    }
  }

  // Bulk Delete Items related to a Lookup Field
  public async _deleteLookUPItems(lookupValueId: number, lookupFieldName: string): Promise<void> {
    const lookupFieldInternal = `${lookupFieldName}Id`;

    // Step 1: Get all items that match the lookup value
    const getUrl = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/items?$filter=${lookupFieldInternal} eq ${lookupValueId}&$select=Id`;

    try {
      const getResponse: SPHttpClientResponse = await this.spHttpClient.get(getUrl, SPHttpClient.configurations.v1);
      if (!getResponse.ok) throw new Error(`Error fetching items: ${getResponse.status}`);

      const items = await getResponse.json();
      const totalItems = items.value.length;

      if (totalItems > 0) {

        // Step 2: Loop and delete each item
        for (const item of items.value) {
          const deleteUrl = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${item.Id})`;

          const spHttpClientOptions: ISPHttpClientOptions = {
            headers: {
              "X-HTTP-Method": "DELETE",
              "IF-MATCH": "*",
            },
          };

          const deleteResponse: SPHttpClientResponse = await this.spHttpClient.post(
            deleteUrl,
            SPHttpClient.configurations.v1,
            spHttpClientOptions
          );

          if (deleteResponse.ok) {
            console.log(`Deleted item ID ${item.Id}`);
          } else {
            console.error(`Failed to delete item ID ${item.Id}: ${deleteResponse.status}`);
          }
        }
        console.log(`✅ Successfully deleted ${totalItems} items linked to ${lookupFieldInternal} = ${lookupValueId}`);
      }

    } catch (error) {
      console.error("Error during bulk delete:", error);
      throw error;
    }
  }

  // Get userPermission
  public async hasPermission(): Promise<boolean> {
    const url: string = `${this.siteUrl}/_api/web/effectiveBasePermissions`;

    try {
      const response: any = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      const permissions: any = await response.json();
      const hasPermission: boolean =
        permissions && permissions.High && permissions.Low;
      return hasPermission;
    } catch (ex) {
      // console.error("Error getting user permission", ex);
      return false;
    }
  }

  // Checks if logged in user is a site administrator
  public async IsCurrentUserSiteAdmin(): Promise<boolean> {
    const url: string = `${this.siteUrl}/_api/web/currentuser/isSiteAdmin`;

    try {
      const response: any = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      const isAdmin: any = await response.json();
      return isAdmin.value;
    } catch (ex) {
      // console.error("Error getting site admin permission", ex);
      return false;
    }
  }

  // Checks if logged in user is a site administrator
  public async GetSPUsers(): Promise<IUser[]> {
    const url: string = `${this.siteUrl}/_api/web/siteusers`;
    let result: IUser[] = [];
    try {
      const response: any = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (response.status === 200) {
        const responseData: any = await response.json();
        result = responseData.value.map((user: IUser) => {
          return user;
        });
        // console.log('Items retrieved successfully:', responseData.value);
        return result;
      } else {
        const responseError: any = await response.json();
        console.log(
          `Error retrieving users. Status: ${responseError.status}`,
          responseError
        );
        // alert('Error Message' + JSON.stringify(responseError));
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    } catch (ex) {
      // console.error("Error getting site users", ex);
      throw new Error("Error getting site users");
    }
  }

  // Update Choices Field within a list
  public async _updateChoicesField(fieldColumnName: string, itemId: string, item: any): Promise<SPHttpClientResponse> {
    const url: string = `${this.siteUrl}/_api/web/lists/GetByTitle('${this.listName}')/fields/GetByTitle(${fieldColumnName})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": "<form_digest_value>",
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
      body: JSON.stringify(item),
    };
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );
      if (response.ok) {
        console.log("Item updated successfully");
        return response;
      } else {
        // const errorResponse: any = await response.json();
        // console.error(`Error updating item. Status: ${response.status}`, errorResponse);
        throw new Error(`Error updating item. Status: ${response.status}`);
      }
    } catch (error) {
      // console.error("Error updating item:", error);
      throw error;
    }
  }

  public async _IsSPGroup(groupName: string): Promise<boolean | undefined> {

    try {
      const endpoint: string = `${this.siteUrl}/_api/web/currentuser/groups`;
      const response: SPHttpClientResponse = await this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        const groups = data.value.map((g: any) => g.Title.toLowerCase());
        return groups.includes(groupName.toLowerCase());
      }
    }
    catch (e) {
      return;
    }
  }

  public async _IsUserInSPGroup(groupName: string, userEmail: string): Promise<boolean> {
    try {
      // Get group by name → get its users
      const endpoint: string = `${this.siteUrl}/_api/web/sitegroups/getbyname('${groupName}')/users?$select=Email`;
      const response: SPHttpClientResponse = await this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

      if (!response.ok) return false;

      const data = await response.json();
      const emails: string[] = (data.value || []).map((u: any) => (u.Email || '').toLowerCase());

      return emails.includes(userEmail.toLowerCase());
    } catch (e) {
      console.error("Error checking if user is in group", e);
      return false;
    }
  }

  // Get SharePoint Group Members
  public async _getSharePointGroupMembers(goupName: string): Promise<IPersonaProps[]> {
    const members: IPersonaProps[] = [];
    if (!goupName) return members;
    const name = String(goupName).trim();
    const esc = (s: string) => s.replace(/'/g, "''");

    try {
      const url = `${this.siteUrl}/_api/web/sitegroups/getbyname('${esc(name)}')/users?$select=Id,Title,Email,LoginName`;
      const resp: any = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp || resp.status !== 200) {
        members;
      }
      const json = await resp.json();
      const personas: IPersonaProps[] = Array.isArray(json?.value) ? json.value.map((u: any) => ({
        text: u?.Title || u?.Email || u?.LoginName || '',
        secondaryText: u?.Email || '',
        id: (u?.Id != null ? String(u.Id) : (u?.LoginName || u?.Title || '')),
      } as IPersonaProps)) : [];
      return personas;
    }
    catch (ex) {
      return members;
    }
  };

  // Resolve a SharePoint user by email/login and return numeric user Id
  public async ensureUserId(loginOrEmail?: string): Promise<number | undefined> {
    if (!loginOrEmail) return undefined;

    const url = `${this.siteUrl}/_api/web/ensureuser`;
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
      },
      body: JSON.stringify({ logonName: `i:0#.f|membership|${loginOrEmail}` })
    };

    const res: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!res.ok) {
      const t = await res.text();
      throw new Error(`ensureUser failed for ${loginOrEmail}: ${t}`);
    }
    const u = await res.json();
    return u?.Id;
  }

}
