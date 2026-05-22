import { Configuration, LogLevel } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: import.meta.env.VITE_REDIRECT_URI || window.location.origin,
    postLogoutRedirectUri: "/",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            break;
          case LogLevel.Warning:
            console.warn(message);
            break;
        }
      },
    },
  },
};

// Delegated permission scopes
export const loginRequest = {
  scopes: [
    "openid",
    "profile",
    "User.Read",
    "User.ReadBasic.All",    // Org kullanıcı arama
    "Files.ReadWrite",        // SharePoint Excel
    "ChannelMessage.Send",    // Teams kanal mesajı
    "Chat.ReadWrite",         // Direkt mesaj
  ],
};

export const graphScopes = {
  scopes: ["https://graph.microsoft.com/.default"],
};
