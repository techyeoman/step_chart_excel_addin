import {
  makeStyles,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Subtitle2,
  tokens,
} from "@fluentui/react-components";
import * as React from "react";
import { ValidationIssue } from "../../core/types";

export const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  row: { display: "flex", gap: "8px", alignItems: "flex-end", flexWrap: "wrap" },
  grow: { flexGrow: 1, flexBasis: "120px", minWidth: 0 },
  wrap: { overflowWrap: "anywhere" },
  muted: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  mono: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
  table: {
    borderCollapse: "collapse",
    width: "100%",
    fontSize: tokens.fontSizeBase200,
    "& th, & td": {
      textAlign: "right",
      padding: "2px 6px",
      borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
      fontVariantNumeric: "tabular-nums",
    },
    "& th": { fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2 },
    "& th:first-child, & td:first-child": { textAlign: "left" },
  },
  scroll: { maxHeight: "220px", overflowY: "auto" },
  card: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "8px 10px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
});

export function SectionTitle({ children }: { children: React.ReactNode }) {
  return <Subtitle2>{children}</Subtitle2>;
}

/** Locale-aware number formatting with trailing zeros trimmed. */
export function fmtNum(value: number, maxDigits = 6): string {
  if (!Number.isFinite(value)) return String(value);
  return value.toLocaleString(undefined, { maximumFractionDigits: maxDigits });
}

export function fmtPct(value: number, maxDigits = 2): string {
  return (value * 100).toLocaleString(undefined, { maximumFractionDigits: maxDigits }) + "%";
}

export function IssueList({ issues }: { issues: ValidationIssue[] }) {
  const styles = useStyles();
  if (issues.length === 0) return null;
  const order = { error: 0, warning: 1, info: 2 } as const;
  const sorted = [...issues].sort((a, b) => order[a.severity] - order[b.severity]);
  const shown = sorted.slice(0, 8);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      {shown.map((issue, i) => (
        <MessageBar key={i} intent={issue.severity === "info" ? "info" : issue.severity}>
          <MessageBarBody className={styles.wrap}>
            {issue.cell ? <strong>{issue.cell}: </strong> : null}
            {issue.message}
          </MessageBarBody>
        </MessageBar>
      ))}
      {sorted.length > shown.length ? (
        <MessageBar intent="info">
          <MessageBarBody>…and {sorted.length - shown.length} more.</MessageBarBody>
        </MessageBar>
      ) : null}
    </div>
  );
}

export interface StatusMessage {
  intent: "success" | "error" | "warning" | "info";
  title?: string;
  text: string;
  notes?: string[];
}

export function Status({ message }: { message: StatusMessage | null }) {
  const styles = useStyles();
  if (!message) return null;
  return (
    <MessageBar intent={message.intent} layout="multiline">
      <MessageBarBody className={styles.wrap}>
        {message.title ? <MessageBarTitle>{message.title}</MessageBarTitle> : null}
        {message.text}
        {message.notes?.length ? (
          <ul style={{ margin: "4px 0 0", paddingLeft: 18 }}>
            {message.notes.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>
        ) : null}
      </MessageBarBody>
    </MessageBar>
  );
}
