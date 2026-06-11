/**
 * HierarchyEditor.tsx
 * Hiyerarşi ağacını görüntüler ve düzenler.
 */

import React, { useState, useCallback, useEffect } from "react";
import {
  Button, Avatar, Text, Card, Dialog, DialogSurface, DialogTitle,
  DialogBody, DialogActions, DialogTrigger, Field, Select,
  Toast, ToastTitle, ToastBody, useToastController, useId, Toaster,
  Spinner, Badge, MessageBar, MessageBarBody, makeStyles, tokens, Divider,
} from "@fluentui/react-components";
import {
  PersonAddRegular, DeleteRegular, ArrowMoveRegular,
  ShieldCheckmarkRegular, ShieldDismissRegular, ArrowClockwiseRegular,
  ChevronDownRegular, ChevronRightRegular,
  EyeRegular, EyeOffRegular,
} from "@fluentui/react-icons";
import type { HierarchyNode, TreeNode, OrgUser } from "../types";
import {
  getHierarchy, setUserManager, removeHierarchyNode,
  grantTreeAdmin, revokeTreeAdmin, buildTree, getDescendantIds,
  grantApprovalViewer, revokeApprovalViewer,
} from "../services/requestService";
import { UserPicker } from "./UserPicker";

const useStyles = makeStyles({
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap", marginBottom: "12px" },
  treeContainer: { display: "flex", flexDirection: "column", gap: "4px" },
  node: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  nodeRoot: {
    borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
  },
  nodeInfo: { flex: 1, display: "flex", flexDirection: "column" },
  nodeActions: { display: "flex", gap: "4px", flexWrap: "wrap" },
  childrenIndent: {
    marginLeft: "28px",
    paddingLeft: "12px",
    borderLeft: `1px dashed ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    marginTop: "4px",
  },
  toggleBtn: {
    background: "none",
    border: "none",
    cursor: "pointer",
    padding: 0,
    display: "flex",
    alignItems: "center",
    color: tokens.colorNeutralForeground2,
  },
  toggleSpacer: { width: "16px", display: "inline-block" },
  emptyState: {
    padding: "32px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
});

interface HierarchyEditorProps {
  token: string;
  currentUserId: string;
}

type DialogMode =
  | { type: "add-root" }
  | { type: "add-child"; parent: HierarchyNode }
  | { type: "move"; node: HierarchyNode };

export const HierarchyEditor: React.FC<HierarchyEditorProps> = ({ token, currentUserId }) => {
  const styles = useStyles();
  const toasterId = useId("hier-toaster");
  const { dispatchToast } = useToastController(toasterId);

  const [nodes, setNodes] = useState<HierarchyNode[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [dialog, setDialog] = useState<DialogMode | null>(null);
  const [pickedUser, setPickedUser] = useState<OrgUser | null>(null);
  const [pickedParentId, setPickedParentId] = useState<string>("");
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());

  const load = useCallback(async () => {
    setIsLoading(true);
    try {
      const all = await getHierarchy(token);
      setNodes(all);
    } catch (err) {
      dispatchToast(
        <Toast><ToastTitle>Yükleme Hatası</ToastTitle><ToastBody>{String(err)}</ToastBody></Toast>,
        { intent: "error" }
      );
    } finally {
      setIsLoading(false);
    }
  }, [token, dispatchToast]);

  useEffect(() => { load(); }, [load]);

  const toast = (title: string, body: string, intent: "success" | "warning" | "error" = "success") => {
    dispatchToast(
      <Toast><ToastTitle>{title}</ToastTitle><ToastBody>{body}</ToastBody></Toast>,
      { intent }
    );
  };

  // ── Dialog kapatma ────────────────────────────────────────────────────────
  const closeDialog = () => {
    setDialog(null);
    setPickedUser(null);
    setPickedParentId("");
  };

  // ── Ekleme (köke veya parent altına) ──────────────────────────────────────
  const handleSaveAdd = async () => {
    if (!dialog || !pickedUser) return;
    const parentId = dialog.type === "add-child" ? dialog.parent.id : "";
    try {
      await setUserManager(token, pickedUser, parentId);
      toast("Eklendi", `${pickedUser.displayName} hiyerarşiye eklendi.`);
      closeDialog();
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  // ── Taşıma ────────────────────────────────────────────────────────────────
  const handleSaveMove = async () => {
    if (!dialog || dialog.type !== "move") return;
    const { node } = dialog;
    try {
      await setUserManager(
        token,
        { id: node.id, displayName: node.displayName, mail: node.mail },
        pickedParentId
      );
      toast("Taşındı", `${node.displayName} yeni amire taşındı.`);
      closeDialog();
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  // ── Silme ─────────────────────────────────────────────────────────────────
  const handleRemove = async (node: HierarchyNode) => {
    if (!confirm(`${node.displayName} hiyerarşiden kaldırılsın mı?`)) return;
    try {
      await removeHierarchyNode(token, node.id);
      toast("Kaldırıldı", `${node.displayName} hiyerarşiden çıkarıldı.`, "warning");
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  // ── Admin yetkisi ─────────────────────────────────────────────────────────
  const handleToggleAdmin = async (node: HierarchyNode) => {
    try {
      if (node.isTreeAdmin) {
        await revokeTreeAdmin(token, node.id);
        toast("Yetki Kaldırıldı", `${node.displayName} artık tree admin değil.`, "warning");
      } else {
        await grantTreeAdmin(token, node.id);
        toast("Yetki Verildi", `${node.displayName} tree admin oldu.`);
      }
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  // ── Onaylanan izin görüntüleyici yetkisi ──────────────────────────────────
  const handleToggleViewer = async (node: HierarchyNode) => {
    try {
      if (node.isApprovalViewer) {
        await revokeApprovalViewer(token, node.id);
        toast("Yetki Kaldırıldı", `${node.displayName} artık onaylanan izinleri göremez.`, "warning");
      } else {
        await grantApprovalViewer(token, node.id);
        toast("Yetki Verildi", `${node.displayName} tüm onaylanan izinleri görebilir.`);
      }
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  // ── Collapse/expand ───────────────────────────────────────────────────────
  const toggleCollapse = (id: string) => {
    setCollapsed((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  // ── Rekürsif tree render ──────────────────────────────────────────────────
  const renderNode = (node: TreeNode, depth: number): React.ReactNode => {
    const hasChildren = node.children.length > 0;
    const isCollapsed = collapsed.has(node.id);
    return (
      <div key={node.id}>
        <div className={`${styles.node} ${depth === 0 ? styles.nodeRoot : ""}`}>
          {hasChildren ? (
            <button
              className={styles.toggleBtn}
              onClick={() => toggleCollapse(node.id)}
              aria-label={isCollapsed ? "Aç" : "Kapat"}
            >
              {isCollapsed ? <ChevronRightRegular /> : <ChevronDownRegular />}
            </button>
          ) : (
            <span className={styles.toggleSpacer} />
          )}
          <Avatar name={node.displayName} size={32} color="colorful" />
          <div className={styles.nodeInfo}>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <Text weight="semibold" size={300}>{node.displayName}</Text>
              {node.isTreeAdmin && (
                <Badge appearance="filled" color="brand" size="small">Admin</Badge>
              )}
              {node.isApprovalViewer && (
                <Badge appearance="filled" color="success" size="small" icon={<EyeRegular />}>
                  Onaylı İzin Görür
                </Badge>
              )}
              {depth === 0 && (
                <Badge appearance="outline" size="small">Root</Badge>
              )}
            </div>
            <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
              {node.mail}
              {hasChildren && ` • ${node.children.length} ast`}
            </Text>
          </div>
          <div className={styles.nodeActions}>
            <Button
              size="small"
              icon={<PersonAddRegular />}
              onClick={() => setDialog({ type: "add-child", parent: node })}
              title="Alt çalışan ekle"
            >
              Alt Ekle
            </Button>
            <Button
              size="small"
              icon={<ArrowMoveRegular />}
              appearance="subtle"
              onClick={() => setDialog({ type: "move", node })}
              title="Başka amire taşı"
            >
              Taşı
            </Button>
            <Button
              size="small"
              icon={node.isTreeAdmin ? <ShieldDismissRegular /> : <ShieldCheckmarkRegular />}
              appearance="subtle"
              onClick={() => handleToggleAdmin(node)}
              title={node.isTreeAdmin ? "Admin yetkisini kaldır" : "Admin yap"}
            />
            <Button
              size="small"
              icon={node.isApprovalViewer ? <EyeOffRegular /> : <EyeRegular />}
              appearance="subtle"
              onClick={() => handleToggleViewer(node)}
              title={node.isApprovalViewer
                ? "Onaylanan izin görme yetkisini kaldır"
                : "Tüm onaylanan izinleri görme yetkisi ver"}
            />
            <Button
              size="small"
              icon={<DeleteRegular />}
              appearance="subtle"
              onClick={() => handleRemove(node)}
              disabled={hasChildren}
              title={hasChildren ? "Önce alt çalışanları taşıyın" : "Hiyerarşiden çıkar"}
            />
          </div>
        </div>
        {hasChildren && !isCollapsed && (
          <div className={styles.childrenIndent}>
            {node.children.map((c) => renderNode(c, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  const tree = buildTree(nodes);

  // ── Taşıma için olası parent listesi ──────────────────────────────────────
  const movableParents = (movingNode: HierarchyNode | null): HierarchyNode[] => {
    if (!movingNode) return [];
    const forbidden = getDescendantIds(nodes, movingNode.id);
    forbidden.add(movingNode.id);
    return nodes.filter((n) => !forbidden.has(n.id));
  };

  // ── UserPicker exclude listesi ────────────────────────────────────────────
  const existingNodeIds = nodes.map((n) => n.id);

  if (isLoading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", padding: "32px" }}>
        <Spinner label="Hiyerarşi yükleniyor..." />
      </div>
    );
  }

  return (
    <>
      <Toaster toasterId={toasterId} position="top-end" />

      <Card>
        <div className={styles.toolbar}>
          <Button
            appearance="primary"
            icon={<PersonAddRegular />}
            onClick={() => setDialog({ type: "add-root" })}
          >
            Köke Ekle
          </Button>
          <Button icon={<ArrowClockwiseRegular />} appearance="subtle" onClick={load}>
            Yenile
          </Button>
          <Text size={100} style={{ color: "var(--colorNeutralForeground3)", marginLeft: "auto" }}>
            Toplam: {nodes.length} kullanıcı
          </Text>
        </div>

        <Divider />

        {tree.length === 0 ? (
          <div className={styles.emptyState}>
            <Text size={300}>Hiyerarşi boş.</Text>
            <br />
            <Text size={200}>"Köke Ekle" butonuyla ilk kullanıcıyı ekleyin.</Text>
          </div>
        ) : (
          <div className={styles.treeContainer} style={{ marginTop: "12px" }}>
            {tree.map((root) => renderNode(root, 0))}
          </div>
        )}
      </Card>

      {/* ── Dialog: Köke veya alt ekle ─────────────────────────────────────── */}
      <Dialog open={dialog?.type === "add-root" || dialog?.type === "add-child"} onOpenChange={(_, d) => !d.open && closeDialog()}>
        <DialogSurface>
          <DialogTitle>
            {dialog?.type === "add-child"
              ? `${dialog.parent.displayName} altına ekle`
              : "Köke kullanıcı ekle"}
          </DialogTitle>
          <DialogBody>
            <div style={{ minWidth: "320px", minHeight: "360px" }}>
              <UserPicker
                label="Kullanıcı seçin"
                token={token}
                value={pickedUser}
                onChange={setPickedUser}
                exclude={existingNodeIds}
              />
              {pickedUser === null && existingNodeIds.length > 0 && (
                <Text size={100} style={{ color: "var(--colorNeutralForeground3)", marginTop: "8px" }}>
                  Hiyerarşide zaten olanlar listeden gizli.
                </Text>
              )}
            </div>
          </DialogBody>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary">İptal</Button>
            </DialogTrigger>
            <Button appearance="primary" disabled={!pickedUser} onClick={handleSaveAdd}>
              Ekle
            </Button>
          </DialogActions>
        </DialogSurface>
      </Dialog>

      {/* ── Dialog: Taşı ────────────────────────────────────────────────────── */}
      <Dialog open={dialog?.type === "move"} onOpenChange={(_, d) => !d.open && closeDialog()}>
        <DialogSurface>
          <DialogTitle>
            {dialog?.type === "move" ? `${dialog.node.displayName} taşı` : ""}
          </DialogTitle>
          <DialogBody>
            <div style={{ minWidth: "320px", minHeight: "200px", display: "flex", flexDirection: "column", gap: "12px" }}>
              <MessageBar intent="info">
                <MessageBarBody>
                  Bu kullanıcının altındaki çalışanlar onunla birlikte taşınır.
                </MessageBarBody>
              </MessageBar>
              <Field label="Yeni amir">
                <Select
                  value={pickedParentId}
                  onChange={(_, d) => setPickedParentId(d.value)}
                >
                  <option value="">(Köke taşı — amiri yok)</option>
                  {dialog?.type === "move" && movableParents(dialog.node).map((n) => (
                    <option key={n.id} value={n.id}>{n.displayName}</option>
                  ))}
                </Select>
              </Field>
            </div>
          </DialogBody>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary">İptal</Button>
            </DialogTrigger>
            <Button appearance="primary" onClick={handleSaveMove}>
              Taşı
            </Button>
          </DialogActions>
        </DialogSurface>
      </Dialog>
    </>
  );
};
