/**
 * UserPicker.tsx
 * Kuruluştan kullanıcı seçme bileşeni.
 * /api/searchUsers → Graph API /users?$search="displayName:..." ile arama yapar.
 */

import React, { useState, useCallback, useRef, useEffect } from "react";
import {
  Field,
  Input,
  Spinner,
  Text,
  Avatar,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { SearchRegular, DismissRegular } from "@fluentui/react-icons";
import type { OrgUser } from "../types";
import { searchOrgUsers } from "../services/requestService";

const useStyles = makeStyles({
  wrapper: {
    position: "relative",
  },
  dropdown: {
    position: "absolute",
    top: "calc(100% + 4px)",
    left: 0,
    right: 0,
    zIndex: 1000,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow16,
    maxHeight: "280px",
    overflowY: "auto",
  },
  item: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "8px 12px",
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  itemInfo: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  noResult: {
    padding: "12px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  selectedChip: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "6px 10px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    marginTop: "6px",
  },
  clearBtn: {
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    color: tokens.colorNeutralForeground3,
    "&:hover": { color: tokens.colorNeutralForeground1 },
    background: "none",
    border: "none",
    padding: 0,
  },
});

interface UserPickerProps {
  label: string;
  token: string;
  value: OrgUser | null;
  onChange: (user: OrgUser | null) => void;
  required?: boolean;
  exclude?: string[]; // Listede gösterilmeyecek user ID'leri
}

export const UserPicker: React.FC<UserPickerProps> = ({
  label,
  token,
  value,
  onChange,
  required,
  exclude = [],
}) => {
  const styles = useStyles();
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<OrgUser[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isOpen, setIsOpen] = useState(false);
  const debounceTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  // Dışarıya tıklanınca kapat
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const handleSearch = useCallback(
    (q: string) => {
      setQuery(q);
      if (debounceTimer.current) clearTimeout(debounceTimer.current);
      if (q.length < 2) {
        setResults([]);
        setIsOpen(false);
        return;
      }
      debounceTimer.current = setTimeout(async () => {
        setIsLoading(true);
        try {
          const users = await searchOrgUsers(token, q);
          setResults(users.filter((u) => !exclude.includes(u.id)));
          setIsOpen(true);
        } catch (err) {
          console.error("Kullanıcı arama hatası:", err);
        } finally {
          setIsLoading(false);
        }
      }, 350);
    },
    [token, exclude]
  );

  const handleSelect = (user: OrgUser) => {
    onChange(user);
    setQuery("");
    setResults([]);
    setIsOpen(false);
  };

  const handleClear = () => {
    onChange(null);
    setQuery("");
  };

  return (
    <div ref={containerRef} className={styles.wrapper}>
      <Field label={label} required={required}>
        {value ? (
          <div className={styles.selectedChip}>
            <Avatar
              name={value.displayName}
              size={24}
              color="colorful"
            />
            <div style={{ flex: 1 }}>
              <Text weight="semibold" size={200}>
                {value.displayName}
              </Text>
              <br />
              <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
                {value.mail}
              </Text>
            </div>
            <button className={styles.clearBtn} onClick={handleClear}>
              <DismissRegular fontSize={16} />
            </button>
          </div>
        ) : (
          <Input
            placeholder="İsim yazın..."
            value={query}
            onChange={(_, d) => handleSearch(d.value)}
            contentBefore={
              isLoading ? <Spinner size="tiny" /> : <SearchRegular />
            }
            autoComplete="off"
          />
        )}
      </Field>

      {isOpen && results.length > 0 && !value && (
        <div className={styles.dropdown}>
          {results.map((user) => (
            <div
              key={user.id}
              className={styles.item}
              onMouseDown={() => handleSelect(user)}
            >
              <Avatar name={user.displayName} size={32} color="colorful" />
              <div className={styles.itemInfo}>
                <Text weight="semibold" size={200}>
                  {user.displayName}
                </Text>
                <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
                  {user.jobTitle ? `${user.jobTitle} • ` : ""}
                  {user.mail}
                </Text>
              </div>
            </div>
          ))}
        </div>
      )}

      {isOpen && results.length === 0 && !isLoading && query.length >= 2 && !value && (
        <div className={styles.dropdown}>
          <Text className={styles.noResult} size={200}>
            "{query}" için sonuç bulunamadı
          </Text>
        </div>
      )}
    </div>
  );
};
