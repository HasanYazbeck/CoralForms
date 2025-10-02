import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUser {
    id: string;
    displayName: string;
    mail: string;
    userPrincipalName: string;
}

export default class MicrosoftGraphService {
    private context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;
    }

    // How to use getAllUsers()
    /*  const graphService = new MicrosoftGraphService(this.context);
        const members = await graphService.getGroupMembersByEmail("mygroup@domain.com");
        console.log("Group Members:", members);
        // Get all users in Azure AD
        const allUsers = await graphService.getAllUsers();
        console.log("All Users:", allUsers);
    */

    // Get members of a group by its email address
    //@param groupEmail - Email address of the Mail Enabled Security Group in Azure AD
    public async getGroupMembersByEmail(groupEmail: string): Promise<any[]> {
        try {
            const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // 1. Get the group by email
            const groupResponse = await client
                .api(`/groups?$filter=mail eq '${groupEmail}'`)
                .get();

            if (!groupResponse.value || groupResponse.value.length === 0) {
                throw new Error(`Group with email ${groupEmail} not found`);
            }

            const groupId = groupResponse.value[0].id;

            // 2. Get members of the group
            const membersResponse = await client
                .api(`/groups/${groupId}/members`)
                .get();

            return membersResponse.value || [];
        } catch (error) {
            console.error("Error getting group members:", error);
            return [];
        }
    }

    // Get all users in the organization
    public async getAllUsers(): Promise<IUser[]> {
        try {
            const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient("3");

            let users: IUser[] = [];
            let response = await client
                .api("/users")
                .select("id,displayName,mail,userPrincipalName")
                .top(999) // max allowed per request
                .get();

            users = users.concat(response.value);

            // Handle paging if @odata.nextLink exists
            while (response["@odata.nextLink"]) {
                response = await client.api(response["@odata.nextLink"]).get();
                users = users.concat(response.value);
            }

            return users;
        } catch (error) {
            console.error("Error getting all users:", error);
            return [];
        }
    }

    // @param groupEmail - Email address of the Mail Enabled Security Group in Azure AD
    // @param userEmail  - User Email address to check membership for within the Security Group
    public async isMemberofGroup(groupEmail: string, userEmail: string): Promise<boolean> {
        try {
            const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient("3");

            // 1. Get group by email
            const groupResponse = await client
                .api(`/groups`)
                .filter(`mail eq '${groupEmail}'`)
                .get();

            if (!groupResponse.value || groupResponse.value.length === 0) {
                throw new Error(`Group ${groupEmail} not found`);
            }
            const groupId = groupResponse.value[0].id;

            // 2. Get user by email
            const userResponse = await client
                .api(`/users/${userEmail}`)
                .select("id")
                .get();

            const userId = userResponse.id;

            // 3. Check if user is member of group
            const checkResponse = await client
                .api(`/users/${userId}/checkMemberGroups`)
                .post({
                    groupIds: [groupId]
                });

            // checkMemberGroups returns array of groupIds the user is in
            return checkResponse.length > 0;
        } catch (error) {
            console.error("Error checking membership:", error);
            return false;
        }
    }
}