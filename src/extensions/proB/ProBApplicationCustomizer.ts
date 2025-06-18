import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { MSGraphClientV3 } from '@microsoft/sp-http';

//import * as strings from 'ProBApplicationCustomizerStrings';

export interface IProBApplicationCustomizerProperties {
  cacheTime: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ProBApplicationCustomizer extends BaseApplicationCustomizer<IProBApplicationCustomizerProperties> {
  private _graphClient: MSGraphClientV3;

  public async onInit(): Promise<void> {

    if (window.location.pathname.indexOf('teams/b')) {
      console.log('pro b site');
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');

      try {
        const siteDetails = await this._graphClient.api(`/sites/${this.context.pageContext.site.id}`).select('displayName').get();
        console.log('siteDetails', siteDetails);

        if (siteDetails.displayName) {
          const groupsResponse = await this._graphClient
          .api(`/groups`)
          .filter(`groupTypes/any(c:c eq 'Unified') and startswith(displayName, '${siteDetails.displayName}')`)
          .select('id,displayName,description,mail,visibility') 
          .header('ConsistencyLevel', 'eventual') 
          .get();
          console.log('groupsResponse', groupsResponse);

          const group = groupsResponse.value[0];
          console.log('group', group);

          if (group && group.visibility === 'Public') {
            const groupOwners = await this._graphClient.api(`/groups/${group.id}/owners`).select('id,displayName,userPrincipalName').get();
            const groupMembers = await this._graphClient.api(`/groups/${group.id}/members`).select('id,displayName,userPrincipalName').get();
            console.log('groupOwners', groupOwners);
            console.log('groupMembers', groupMembers);

            const owners = groupOwners.value;
            const members = groupMembers.value;

            const currentUserEmail = this.context.pageContext.user.email;
            const currentUserId = this.context.pageContext.aadInfo.userId.toString();

            const isOwner = owners.some((o: { id: any; userPrincipalName: string; }) => o.id === currentUserId || o.userPrincipalName === currentUserEmail);
            const isMember = members.some((m: { id: any; userPrincipalName: string; }) => m.id === currentUserId || m.userPrincipalName === currentUserEmail);
            console.log('isOwner', isOwner);
            console.log('isMember', isMember);

            if (!isOwner && !isMember) {
              window.location.href = window.location.origin;
            }
          }
        }
      } catch (e) {
        console.error(e);
      }
    }

    //const cacheTme = parseInt(this.properties.cacheTime) ?? 300000;

    return Promise.resolve();
  }
}
