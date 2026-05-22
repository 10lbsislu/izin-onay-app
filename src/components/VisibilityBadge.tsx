/**
 * VisibilityBadge.tsx
 * Beklemedeki taleplerde "sadece siz ve onaylayıcılar görebilir" göstergesi.
 */

import React from "react";
import { Badge, Tooltip, makeStyles, tokens } from "@fluentui/react-components";
import { LockClosedRegular } from "@fluentui/react-icons";
import type { VisibilityMeta } from "../services/requestService";

const useStyles = makeStyles({
  lockBadge: {
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    padding: "2px 8px",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground3,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    cursor: "default",
    userSelect: "none",
  },
});

interface VisibilityBadgeProps {
  meta?: VisibilityMeta;
}

export const VisibilityBadge: React.FC<VisibilityBadgeProps> = ({ meta }) => {
  const styles = useStyles();

  if (!meta || meta.badge === "normal" || !meta.label) return null;

  return (
    <Tooltip content={meta.label} relationship="label">
      <span className={styles.lockBadge}>
        <LockClosedRegular fontSize={11} />
        {meta.label}
      </span>
    </Tooltip>
  );
};
