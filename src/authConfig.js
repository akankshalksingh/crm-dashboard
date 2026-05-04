export const msalConfig = {
  auth: {
    clientId: process.env.REACT_APP_AZURE_CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: process.env.REACT_APP_REDIRECT_URI || window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Mail.Read",
    "Mail.Send",
    "Calendars.Read",
  ],
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
  graphSentEndpoint: "https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages",
  graphCalendarEndpoint: "https://graph.microsoft.com/v1.0/me/calendarView",
};
