import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { OverlayLock } from './OverlayLock';
import { SessionCache } from './SessionCache';

export interface IProBApplicationCustomizerProperties {
  cacheTime: number;
  debug: boolean
  unlockOnError: boolean
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ProBApplicationCustomizer extends BaseApplicationCustomizer<IProBApplicationCustomizerProperties> {
  private _graphClient: MSGraphClientV3;
  private overlayLock: OverlayLock;

  public async onInit(): Promise<void> {
    this.overlayLock = new OverlayLock();

    if (window.location.href.indexOf(`${window.location.origin}/teams/b`) === 0) {
      if(this.properties.debug)
        console.log('Pro B site detected...')

      const cacheVal = SessionCache.get(this.context.pageContext.site.id.toString());
      if (cacheVal === this.context.pageContext.aadInfo.userId.toString()) {
        if (this.properties.debug)
          console.log('Access confirmed via cache.')

        this.overlayLock.destroy();
        return Promise.resolve();
      }

      this.overlayLock.lock();
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

          if(this.properties.debug)
            console.log('groupsResponse', groupsResponse);

          const group = groupsResponse.value[0];

          if(this.properties.debug)
            console.log('group', group);

          if (group && group.visibility === 'Public') {
            const groupOwners = await this._graphClient.api(`/groups/${group.id}/owners`).select('id,displayName,userPrincipalName').get();
            const groupMembers = await this._graphClient.api(`/groups/${group.id}/members`).select('id,displayName,userPrincipalName').get();

            if(this.properties.debug) {
              console.log('groupOwners', groupOwners);
              console.log('groupMembers', groupMembers);
            }

            const owners = groupOwners.value;
            const members = groupMembers.value;

            const currentUserEmail = this.context.pageContext.user.email;
            const currentUserId = this.context.pageContext.aadInfo.userId.toString();

            const isOwner = owners.some((o: { id: unknown; userPrincipalName: string; }) => o.id === currentUserId || o.userPrincipalName === currentUserEmail);
            const isMember = members.some((m: { id: unknown; userPrincipalName: string; }) => m.id === currentUserId || m.userPrincipalName === currentUserEmail);

            if(this.properties.debug) {
              console.log('isOwner', isOwner);
              console.log('isMember', isMember);
            }

            if (!isOwner && !isMember) {
              if(this.properties.debug)
                console.log('Access revoked.')

              this.overlayLock.unlock();
              window.location.href = window.location.origin;
            }
            else {
              this.overlayLock.destroy();
              SessionCache.set(this.context.pageContext.site.id.toString(), this.context.pageContext.aadInfo.userId.toString(), this.properties.cacheTime ?? 300000);

              if(this.properties.debug)
                console.log('Access confirmed.')
            }
          }
        }
      } catch (e) {
        console.error(e);

        if (this.properties.unlockOnError)
          this.overlayLock.unlock();
      }
    }

    return Promise.resolve();
  }
}
