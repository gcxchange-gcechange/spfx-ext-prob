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

const REVOKE_FLAG = 0;
const GRANT_FLAG = 1;

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ProBApplicationCustomizer extends BaseApplicationCustomizer<IProBApplicationCustomizerProperties> {
  private _graphClient: MSGraphClientV3;
  private overlayLock: OverlayLock;

  public async onInit(): Promise<void> {
    this.overlayLock = new OverlayLock();

    if (window.location.href.indexOf(`${window.location.origin}/teams/b`) === 0) {
        this.debugLog('Pro B site detected...')

      const cacheVal = SessionCache.get(this.context.pageContext.site.id.toString());
      if (cacheVal === `${this.context.pageContext.aadInfo.userId.toString()}${GRANT_FLAG}`) {
        this.debugLog('Access confirmed via cache.');

        this.overlayLock.destroy();
        return Promise.resolve();
      } else if (cacheVal === `${this.context.pageContext.aadInfo.userId.toString()}${REVOKE_FLAG}`) {
        this.debugLog('Access revoked via cache.');

        this.overlayLock.destroy();
        window.location.href = window.location.origin;
      }

      this.overlayLock.lock();
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');

      try {
        const siteDetails = await this._graphClient.api(`/sites/${this.context.pageContext.site.id}`).select('name').get();
        this.debugLog('siteDetails', siteDetails);

        if (siteDetails.name) {
          const groupsResponse = await this._graphClient
          .api(`/groups`)
          .filter(`groupTypes/any(c:c eq 'Unified') and startswith(mail, '${siteDetails.name}')`)
          .select('id,displayName,description,mail,visibility')
          .header('ConsistencyLevel', 'eventual')
          .get();

          this.debugLog('groupsResponse', groupsResponse);

          const group = groupsResponse.value[0];
          this.debugLog('group', group);

          if (group && group.visibility === 'Public') {
            const groupOwners = await this._graphClient.api(`/groups/${group.id}/owners`).select('id,displayName,userPrincipalName').get();
            const groupMembers = await this._graphClient.api(`/groups/${group.id}/members`).select('id,displayName,userPrincipalName').get();

            this.debugLog('groupOwners', groupOwners);
            this.debugLog('groupMembers', groupMembers);

            const owners = groupOwners.value;
            const members = groupMembers.value;

            const currentUserEmail = this.context.pageContext.user.email;
            const currentUserId = this.context.pageContext.aadInfo.userId.toString();

            const isOwner = owners.some((o: { id: unknown; userPrincipalName: string; }) => o.id === currentUserId || o.userPrincipalName === currentUserEmail);
            const isMember = members.some((m: { id: unknown; userPrincipalName: string; }) => m.id === currentUserId || m.userPrincipalName === currentUserEmail);

            this.debugLog('isOwner', isOwner);
            this.debugLog('isMember', isMember);

            if (!isOwner && !isMember) {
              this.debugLog('Access revoked.')

              SessionCache.set(this.context.pageContext.site.id.toString(), `${this.context.pageContext.aadInfo.userId.toString()}${REVOKE_FLAG}`, this.properties.cacheTime ?? 300000);
              this.overlayLock.unlock();
              window.location.href = window.location.origin;
            }
            else {
              this.overlayLock.destroy();
              SessionCache.set(this.context.pageContext.site.id.toString(), `${this.context.pageContext.aadInfo.userId.toString()}${GRANT_FLAG}`, this.properties.cacheTime ?? 300000);

              this.debugLog('Access confirmed.');
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
  
  private debugLog(text: string, obj?: unknown): void {
    if (this.properties.debug) {
      if (text && obj)
        console.log(text, obj);
      else
        console.log(text);
    }
  }
}
