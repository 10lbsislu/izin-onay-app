import { useState, useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";

interface TeamsContext {
  isTeams: boolean;
  isLoading: boolean;
  userObjectId: string | null;
  userPrincipalName: string | null;
  teamId: string | null;
  channelId: string | null;
  theme: string;
  teamsToken: string | null;
  tokenError: string | null;
}

export function useTeamsContext(): TeamsContext {
  const [ctx, setCtx] = useState<TeamsContext>({
    isTeams: false,
    isLoading: true,
    userObjectId: null,
    userPrincipalName: null,
    teamId: null,
    channelId: null,
    theme: "default",
    teamsToken: null,
    tokenError: null,
  });

  useEffect(() => {
    const init = async () => {
      try {
        await microsoftTeams.app.initialize();
        const context = await microsoftTeams.app.getContext();

        // Teams SSO token al
        let token: string | null = null;
        let tokenErr: string | null = null;
        try {
          token = await microsoftTeams.authentication.getAuthToken({
            resources: [`api://${import.meta.env.VITE_CLIENT_ID}`],
            silent: true,
          });
        } catch (err) {
          tokenErr = String(err);
          console.warn("Teams SSO token alınamadı, fallback MSAL kullanılacak:", err);
        }

        setCtx({
          isTeams: true,
          isLoading: false,
          userObjectId: context.user?.id ?? null,
          userPrincipalName: context.user?.userPrincipalName ?? null,
          teamId: context.team?.internalId ?? null,
          channelId: context.channel?.id ?? null,
          theme: context.app.theme ?? "default",
          teamsToken: token,
          tokenError: tokenErr,
        });
      } catch {
        // Teams dışında çalışıyor (geliştirme/test ortamı)
        setCtx((prev) => ({ ...prev, isTeams: false, isLoading: false }));
      }
    };

    init();
  }, []);

  return ctx;
}
