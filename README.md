# prob

## Summary

This extension is intended to lock users out of accessing public Protected B sites.
Users who aren't in the Owner or Members list for the group associated with that site will re-route users to the home page.
Once a user is validated for that site they will be cached to local storage for a set amount of time to reduce calls.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## API Permissions

- Microsoft Graph - Group.Read.All
- Microsoft Graph - User.Read
- Microsoft Graph - Sites.Read.All

## Extension Settings

- `cacheTime` is a `number` in milliseconds that sets how long a validated user will be cached to local storage.
- `debug` is a `boolean` that enables console logs for the extension.
- `unlockOnError` is a `boolean` that will unlock the overlay if the extension fails.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**
