/**
 * App.tsx
 * Uygulama kök bileşeni.
 * Teams SSO → admin/çalışan görünümü ayrımı.
 */

import React, { useEffect, useState } from "react";
import {
  FluentProvider,
  teamsDarkTheme,
  teamsLightTheme,
  teamsHighContrastTheme,
  TabList,
  Tab,
  Text,
  Spinner,
  Avatar,
  Button,
  makeStyles,
  tokens,
  Badge,
} from "@fluentui/react-components";
import {
  SendRegular,
  DocumentBulletListRegular,
  ShieldCheckmarkRegular,
} from "@fluentui/react-icons";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./auth/authConfig";
import { useTeamsContext } from "./hooks/useTeamsContext";
import { getApprovers } from "./services/requestService";
import { RequestForm } from "./components/RequestForm";
import { MyRequests } from "./components/MyRequests";
import { AdminPanel } from "./components/AdminPanel";
import type { OrgUser } from "./types";

const msalInstance = new PublicClientApplication(msalConfig);

// ─── Fluent tema seçici ──────────────────────────────────────────────────────
function getTheme(teamsTheme: string) {
  if (teamsTheme === "dark" || teamsTheme === "contrast") {
    return teamsTheme === "contrast" ? teamsHighContrastTheme : teamsDarkTheme;
  }
  return teamsLightTheme;
}

// ─── Stiller ─────────────────────────────────────────────────────────────────
const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
    padding: "0",
  },
  appBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 20px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    position: "sticky",
    top: 0,
    zIndex: 100,
  },
  appBarLeft: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  content: {
    padding: "20px",
    maxWidth: "900px",
    margin: "0 auto",
  },
  centerScreen: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    minHeight: "80vh",
    flexDirection: "column",
    gap: "16px",
  },
  tabBar: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: "0 20px",
  },
});

// ─── Ana içerik bileşeni ──────────────────────────────────────────────────────
type AppTab = "yeni-talep" | "taleplerim" | "yonetici";

function AppContent() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const teamsCtx = useTeamsContext();

  const [token, setToken] = useState<string | null>(null);
  const [currentUser, setCurrentUser] = useState<OrgUser | null>(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [activeTab, setActiveTab] = useState<AppTab>("yeni-talep");
  const [refreshKey, setRefreshKey] = useState(0);
  const [isInitializing, setIsInitializing] = useState(true);

  // ─── Auth & kullanıcı bilgisi ────────────────────────────────────────────
  useEffect(() => {
    const init = async () => {
      try {
        setIsInitializing(true);

        // Önce Graph'tan kullanıcı bilgisini al (token olmadan da dene)
        let accessToken: string | null = teamsCtx.teamsToken;

        // Teams token yoksa MSAL ile dene
        if (!accessToken) {
          if (isAuthenticated && accounts.length > 0) {
            try {
              const result = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
              });
              accessToken = result.accessToken;
            } catch {
              try {
                const result = await instance.acquireTokenPopup(loginRequest);
                accessToken = result.accessToken;
              } catch (err) {
                console.error("MSAL popup hatası:", err);
              }
            }
          } else {
            // Giriş yok, popup ile giriş yap
            try {
              const result = await instance.loginPopup(loginRequest);
              accessToken = result.accessToken;
            } catch (err) {
              console.error("Login hatası:", err);
              setIsInitializing(false);
              return;
            }
          }
        }

        if (!accessToken) {
          setIsInitializing(false);
          return;
        }

        setToken(accessToken);

        const meRes = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: { Authorization: `Bearer ${accessToken}` },
        });
        const me = await meRes.json();
        const user: OrgUser = {
          id: me.id,
          displayName: me.displayName,
          mail: me.mail || me.userPrincipalName,
          jobTitle: me.jobTitle,
          department: me.department,
        };
        setCurrentUser(user);

        try {
          const approvers = await getApprovers(accessToken);
          setIsAdmin(approvers.some((a) => a.id === me.id));
        } catch {
          setIsAdmin(false);
        }
      } catch (err) {
        console.error("Auth hatası:", err);
      } finally {
        setIsInitializing(false);
      }
    };

    if (!teamsCtx.isLoading) {
      init();
    }
  }, [teamsCtx, isAuthenticated, accounts, instance]);

  // ─── Login ───────────────────────────────────────────────────────────────
  const handleLogin = async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (err) {
      console.error("Login hatası:", err);
    }
  };

  // ─── Yüklenme ekranı ─────────────────────────────────────────────────────
  if (teamsCtx.isLoading || isInitializing) {
    return (
      <div className={styles.centerScreen}>
        <Spinner size="large" label="İzin Onay Sistemi yükleniyor..." />
      </div>
    );
  }

  // ─── Login ekranı ─────────────────────────────────────────────────────────
  if (!currentUser) {
    return (
      <div className={styles.centerScreen}>
        <ShieldCheckmarkRegular fontSize={64} color={tokens.colorBrandForeground1} />
        <Text size={600} weight="semibold">
          İzin Onay Sistemi
        </Text>
        <Text size={300} style={{ color: "var(--colorNeutralForeground3)" }}>
          Devam etmek için Microsoft hesabınızla giriş yapın.
        </Text>
        <Button appearance="primary" size="large" onClick={handleLogin}>
          Microsoft ile Giriş Yap
        </Button>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      {/* Üst bar */}
      <div className={styles.appBar}>
        <div className={styles.appBarLeft}>
          <ShieldCheckmarkRegular fontSize={28} color={tokens.colorBrandForeground1} />
          <div>
            <Text weight="semibold" size={400}>
              İzin Onay Sistemi
            </Text>
            {isAdmin && (
              <Badge appearance="filled" color="brand" size="small" style={{ marginLeft: "8px" }}>
                Yönetici
              </Badge>
            )}
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <Avatar
            name={currentUser.displayName}
            size={32}
            color="colorful"
          />
          <div>
            <Text size={200} weight="semibold">{currentUser.displayName}</Text>
            <br />
            <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
              {currentUser.jobTitle ?? currentUser.mail}
            </Text>
          </div>
        </div>
      </div>

      {/* Tab bar */}
      <div className={styles.tabBar}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, d) => setActiveTab(d.value as AppTab)}
        >
          <Tab value="yeni-talep" icon={<SendRegular />}>
            Yeni Talep
          </Tab>
          <Tab value="taleplerim" icon={<DocumentBulletListRegular />}>
            Taleplerim
          </Tab>
          {isAdmin && (
            <Tab value="yonetici" icon={<ShieldCheckmarkRegular />}>
              Yönetici Paneli
            </Tab>
          )}
        </TabList>
      </div>

      {/* İçerik */}
      <div className={styles.content}>
        {activeTab === "yeni-talep" && token && (
          <RequestForm
            currentUser={currentUser}
            token={token}
            onSuccess={() => {
              setRefreshKey((k) => k + 1);
              setActiveTab("taleplerim");
            }}
          />
        )}
        {activeTab === "taleplerim" && token && (
          <MyRequests
            currentUser={currentUser}
            token={token}
            refreshKey={refreshKey}
          />
        )}
        {activeTab === "yonetici" && isAdmin && token && (
          <AdminPanel currentUser={currentUser} token={token} />
        )}
      </div>
    </div>
  );
}

// ─── Root wrapper ─────────────────────────────────────────────────────────────
export default function App() {
  const { theme } = useTeamsContext();

  return (
    <MsalProvider instance={msalInstance}>
      <FluentProvider theme={getTheme(theme)}>
        <AppContent />
      </FluentProvider>
    </MsalProvider>
  );
}
