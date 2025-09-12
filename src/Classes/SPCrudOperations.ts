import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { ISPItem } from "../Interfaces/ISPItem";
import { IUser } from "../Interfaces/IUser";
import { FieldTypeKind } from "../Enums/enums";

export class SPCrudOperations {
  private listName: string;
  private siteUrl: string;
  private spHttpClient: SPHttpClient;
  private query?: string;

  constructor(
    spHttpClient: SPHttpClient,
    siteUrl: string,
    listName: string,
    query?: string
  ) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
    this.listName = listName;
    this.query = query;
  }

  // Create List
  public async _createList(
    listName: string,
    listDescription: string
  ): Promise<void> {
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
              Title: listName,
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
  public async _addColumnToList(
    columnName: string,
    columnType: FieldTypeKind
  ): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/fields`;
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
    const url: string = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/fields/getByTitle('${columnName}')`;

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
      console.error("Error deleting column:", error);
      throw error;
    }
  }

  // Insert item List
  public async _insertItem(item: any): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/getByTitle('${this.listName}')/items`;
    console.log("Item to insert:", JSON.stringify(item));
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(item),
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    };

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );

      if (response.status === 201) {
        // the item is created for response code 201
        alert("A new Item inserted successfully.");
      } else {
        const responseJson = await response.json();
        alert(
          "Error Message: " +
            response.status +
            " - " +
            JSON.stringify(responseJson)
        );
      }
    } catch (error) {
      console.error("Error creating item:", error);
      throw error;
    }
  }

  // Get Items List
  public async _getItems(): Promise<any[]> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;

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
        console.log(
          `Error retrieving items. Status: ${responseError.status}`,
          responseError
        );
        alert("Error Message" + JSON.stringify(responseError));
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    } catch (error) {
      console.error("Error Retreiving Items", error);
      throw error;
    }
  }

  // Get Items List
  public async _getItemsWithQuery(): Promise<any[]> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items/${this.query}`;

    try {
      const response = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (response.status === 200) {
        const responseData: any = await response.json();
        // console.log('Items retrieved successfully:', responseData.value);
        return responseData.value;
      } else {
        const responseError: any = await response.json();
        console.log(
          `Error retrieving items. Status: ${responseError.status}`,
          responseError
        );
        // alert('Error Message' + JSON.stringify(responseError));
        throw new Error(
          `Error retrieving items. Status: ${responseError.status}`
        );
      }
    } catch (error) {
      console.error("Error Retreiving Items", error);
      throw error;
    }
  }

  // Get Item List By Id or Title
  public async _getItemById(id: string): Promise<ISPItem> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle(${this.listName})/items?filter=Id eq ${id}`;

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
      console.error("Error Retreiving Item", error);
      throw error;
    }
  }

  // Update Item
  public async _updateItem(
    itemId: string,
    item: any
  ): Promise<SPHttpClientResponse> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;
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
        console.error(
          `Error updating item. Status: ${response.status}`,
          errorResponse
        );
        throw new Error(`Error updating item. Status: ${response.status}`);
      }
    } catch (error) {
      console.error("Error updating item:", error);
      throw error;
    }
  }

  // Delete Item
  public async _deleteItem(itemId: number): Promise<void> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${itemId})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        "X-HTTP-Method": "DELETE",
        "IF-MATCH": "*",
      },
    };

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        spHttpClientOptions
      );
      if (response.ok) {
        console.log("Item deleted successfully");
      } else {
        const errorResponse: any = await response.json();
        console.error(
          `Error deleting item. Status: ${response.status}`,
          errorResponse
        );
        throw new Error(`Error deleting item. Status: ${response.status}`);
      }
    } catch (error) {
      console.error("Error deleting item:", error);
      throw error;
    }
  }

  // Search Users by name or email
  // public async _getUserByTitleEmail(query: string): Promise<IUser> {
  //   const url: string = `${this.siteUrl}/_api/web/siteusers?$filter=Title eq '${query}' or Email eq '${query}'`;
  //   var result: IUser;
  //   try {
  //     const response: SPHttpClientResponse = await this.spHttpClient.get(
  //       url,
  //       SPHttpClient.configurations.v1
  //     );
  //     if (response.ok) {
  //       const itemsList = await response.json();

  //       if (Array.isArray(itemsList.value) && itemsList.value.length > 0) {
  //         const item = itemsList.value[0];
  //         result = {
  //           Id: item.Id,
  //           JobTitle: item.Title,
  //           Email: item.Email,
  //           Name: item.LoginName,
  //         };

  //         // itemsList.value.map((item: any) =>  {
  //         //   const user: IUser = {
  //         //     Id: item.Id,
  //         //     Title: item.Title,
  //         //     Email: item.Email,
  //         //     LoginName: item.LoginName
  //         //   }
  //         //   result = user;
  //         // });
  //       } else {
  //         console.error("Expected an array, but got:", itemsList.value[0]);
  //         throw new Error(`Error Retreiving item. Status: ${response.status}`);
  //       }
  //     } else {
  //       const errorResponse: any = await response.json();
  //       console.error(
  //         `Error Retreiving item. Status: ${response.status}`,
  //         errorResponse
  //       );
  //       throw new Error(`Error Retreiving item. Status: ${response.status}`);
  //     }
  //     return result;
  //   } catch (error) {
  //     console.error("Error Retreiving Item", error);
  //     throw new Error("Error Retreiving Item " + error);
  //   }
  // }

  // Search Users by name or email
  // public async _searchUsersByTitleEmail(query: string): Promise<IUser[]> {
  //   const url: string = `${this.siteUrl}/_api/web/siteusers?$filter=substringof('${query}',Title) or substringof('${query}',Email)`;
  //   let result: IUser[] = [];
  //   try {
  //     const response: SPHttpClientResponse = await this.spHttpClient.get(
  //       url,
  //       SPHttpClient.configurations.v1
  //     );
  //     if (response.ok) {
  //       const itemsList = await response.json();

  //       if (Array.isArray(itemsList.value)) {
  //         itemsList.value.map((item: any) => {
  //           const user: IUser = {
  //             Id: item.Id,
  //             Title: item.Title,
  //             Email: item.Email,
  //             LoginName: item.LoginName,
  //           };
  //           result.push(user);
  //         });
  //         return result;
  //       } else {
  //         console.error("Expected an array, but got:", itemsList.value[0]);
  //         throw new Error("Expected an array, but got:" + itemsList.value[0]);
  //       }
  //     } else {
  //       const errorResponse: any = await response.json();
  //       console.error(
  //         `Error Retreiving item. Status: ${response.status}`,
  //         errorResponse
  //       );
  //       throw new Error(`Error Retreiving item. Status: ${response.status}`);
  //     }
  //   } catch (error) {
  //     console.error("Error Retreiving Item", error);
  //     throw error("Error Retreiving Item");
  //   }
  // }

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
      console.error("Error getting user permission", ex);
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
      console.error("Error getting site admin permission", ex);
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
      console.error("Error getting site users", ex);
      throw new Error("Error getting site users");
    }
  }

  // Update Choices Field within a list
  public async _updateChoicesField(
    fieldColumnName: string,
    itemId: string,
    item: any
  ): Promise<SPHttpClientResponse> {
    const url: string = `${this.siteUrl}/_api/web/lists/getbytitle('${this.listName}')/fields/getbytitle(${fieldColumnName})`;
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
        const errorResponse: any = await response.json();
        console.error(
          `Error updating item. Status: ${response.status}`,
          errorResponse
        );
        throw new Error(`Error updating item. Status: ${response.status}`);
      }
    } catch (error) {
      console.error("Error updating item:", error);
      throw error;
    }
  }
}
