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
  CheckmarkCircleRegular,
} from "@fluentui/react-icons";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./auth/authConfig";
import { useTeamsContext } from "./hooks/useTeamsContext";
import { getHierarchy } from "./services/requestService";
import { RequestForm } from "./components/RequestForm";
import { MyRequests } from "./components/MyRequests";
import { AdminPanel } from "./components/AdminPanel";
import { ApprovedLeaves } from "./components/ApprovedLeaves";
import type { OrgUser } from "./types";

const msalInstance = new PublicClientApplication(msalConfig);

// JWT payload'ı imzayı doğrulamadan decode et — sadece UI'da kullanıcı bilgisi göstermek için
function decodeJwt(token: string): Record<string, any> | null {
  try {
    const payload = token.split(".")[1];
    if (!payload) return null;
    const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
    return JSON.parse(decodeURIComponent(escape(decoded)));
  } catch {
    return null;
  }
}

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
    padding: "14px 24px",
    background: `linear-gradient(135deg, ${tokens.colorNeutralBackground1} 0%, ${tokens.colorBrandBackground2} 200%)`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow4,
    position: "sticky",
    top: 0,
    zIndex: 100,
  },
  appBarLeft: {
    display: "flex",
    alignItems: "center",
    gap: "14px",
  },
  brandIcon: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "40px",
    height: "40px",
    borderRadius: tokens.borderRadiusLarge,
    background: `linear-gradient(135deg, ${tokens.colorBrandBackground} 0%, ${tokens.colorBrandBackgroundHover} 100%)`,
    color: tokens.colorNeutralForegroundOnBrand,
    boxShadow: tokens.shadow4,
  },
  userInfoBox: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "6px 12px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground3,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  content: {
    padding: "24px",
    maxWidth: "960px",
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
    padding: "4px 20px 0",
    boxShadow: tokens.shadow2,
  },
});

// ─── Ana içerik bileşeni ──────────────────────────────────────────────────────
type AppTab = "yeni-talep" | "taleplerim" | "onaylanan" | "yonetici";

function AppContent() {
  const styles = useStyles();
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const teamsCtx = useTeamsContext();

  const [token, setToken] = useState<string | null>(null);
  const [currentUser, setCurrentUser] = useState<OrgUser | null>(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [isTreeAdmin, setIsTreeAdmin] = useState(false);
  const [isApprovalViewer, setIsApprovalViewer] = useState(false);
  const [hasDirectReports, setHasDirectReports] = useState(false);
  const [isRoot, setIsRoot] = useState(false);
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

        // Kullanıcı bilgisini al:
        //  - Teams SSO token'ı Graph audience değil, Graph /me 401 döner
        //  - Bu durumda Teams context + JWT decode ile kullanıcıyı kur
        //  - MSAL flow'unda token zaten Graph audience, /me çağrısı çalışır
        let user: OrgUser;
        const usingTeamsSso = !!teamsCtx.teamsToken;

        if (usingTeamsSso) {
          const claims = decodeJwt(accessToken);
          user = {
            id: teamsCtx.userObjectId || claims?.oid || "",
            displayName: claims?.name || teamsCtx.userPrincipalName || "Kullanıcı",
            mail: teamsCtx.userPrincipalName || claims?.preferred_username || claims?.upn || "",
          };
        } else {
          const meRes = await fetch("https://graph.microsoft.com/v1.0/me", {
            headers: { Authorization: `Bearer ${accessToken}` },
          });
          if (!meRes.ok) {
            console.error("Graph /me hatası:", meRes.status);
            setIsInitializing(false);
            return;
          }
          const me = await meRes.json();
          user = {
            id: me.id,
            displayName: me.displayName,
            mail: me.mail || me.userPrincipalName,
            jobTitle: me.jobTitle,
            department: me.department,
          };
        }
        setCurrentUser(user);

        try {
          const hierarchy = await getHierarchy(accessToken);
          const myNode = hierarchy.find((n) => n.id === user.id);
          const directReports = hierarchy.filter((n) => n.managerId === user.id);
          const treeAdmin = !!myNode?.isTreeAdmin;
          const hasReports = directReports.length > 0;
          const root = !!myNode && !myNode.managerId;
          setIsTreeAdmin(treeAdmin);
          setIsApprovalViewer(!!myNode?.isApprovalViewer);
          setHasDirectReports(hasReports);
          setIsRoot(root);
          setIsAdmin(treeAdmin || hasReports);
        } catch {
          setIsTreeAdmin(false);
          setIsApprovalViewer(false);
          setHasDirectReports(false);
          setIsRoot(false);
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
          <div className={styles.brandIcon}>
            <ShieldCheckmarkRegular fontSize={22} />
          </div>
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <Text weight="semibold" size={400}>
                İzin Onay Sistemi
              </Text>
              {isAdmin && (
                <Badge appearance="filled" color="brand" size="small">
                  Yönetici
                </Badge>
              )}
            </div>
            <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
              Çalışan izin yönetimi
            </Text>
          </div>
        </div>
        <div className={styles.userInfoBox}>
          <Avatar
            name={currentUser.displayName}
            size={36}
            color="colorful"
          />
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2 }}>
            <Text size={200} weight="semibold">{currentUser.displayName}</Text>
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
          {(isApprovalViewer || isTreeAdmin) && (
            <Tab value="onaylanan" icon={<CheckmarkCircleRegular />}>
              Onaylanan İzinler
            </Tab>
          )}
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
            isRoot={isRoot}
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
        {activeTab === "onaylanan" && (isApprovalViewer || isTreeAdmin) && token && (
          <ApprovedLeaves token={token} />
        )}
        {activeTab === "yonetici" && isAdmin && token && (
          <AdminPanel
            currentUser={currentUser}
            token={token}
            isTreeAdmin={isTreeAdmin}
            hasDirectReports={hasDirectReports}
          />
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
