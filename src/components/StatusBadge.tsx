import React from "react";
import { Badge } from "@fluentui/react-components";
import type { LeaveStatus } from "../types";

interface StatusBadgeProps {
  status: LeaveStatus;
}

const CONFIG = {
  beklemede: { color: "warning" as const, label: "Beklemede" },
  onaylandi: { color: "success" as const, label: "Onaylandı" },
  reddedildi: { color: "danger" as const, label: "Reddedildi" },
};

export const StatusBadge: React.FC<StatusBadgeProps> = ({ status }) => {
  const { color, label } = CONFIG[status] ?? CONFIG.beklemede;
  return (
    <Badge appearance="filled" color={color}>
      {label}
    </Badge>
  );
};
