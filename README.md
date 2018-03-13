# teams-tab-msal-auth
Example demonstrating the simplest way to do MSAL based authentication inside Teams Tab

For ADAL please see this sample by [Richard diZerega](https://github.com/richdizz): https://github.com/richdizz/Microsoft-Teams-Tab-Auth

Few things to note:

- I have a sample [manifest](https://github.com/paragppp/teams-tab-msal-auth/tree/master/manifest) available to be loaded into Teams, please edit the manifest.json inside the test_manifest folder and update it with information inside the [ ] brackets. Once update, you will have to re-create the zip file and load it inside teams
- You will have to create an app registration at [https://apps.dev.micosoft.com](https://apps.dev.micosoft.com) and use the app id from there inside the code [auth.html](https://github.com/paragppp/teams-tab-msal-auth/blob/master/msalAuth/auth.html). This is the app the user will consent to, for providing access to reading/writing user information
- Please add access scopes, the code is currently set for user.read and user.readwrite