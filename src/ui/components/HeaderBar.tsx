import * as React from "react";
import { Button, Tooltip, Switch, makeStyles, Dropdown, Option, tokens } from "@fluentui/react-components";
import { Compose24Regular, History24Regular, SlideLayout24Regular } from "@fluentui/react-icons";

export type ModelType = string;

interface HeaderBarProps {
  onNewChat: () => void;
  onShowHistory: () => void;
  onShowTemplates: () => void;
  selectedModel: ModelType;
  onModelChange: (model: ModelType) => void;
  models: { key: string; label: string }[];
  debugEnabled: boolean;
  onDebugChange: (v: boolean) => void;
  activeTemplateName?: string | null;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "0px 0px",
    paddingRight: "44px",
    gap: "8px",
    minHeight: "40px",
  },
  leftSection: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    minWidth: 0,
    flex: 1,
  },
  debugRow: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  dropdown: {
    minWidth: "120px",
    opacity: 0.6,
    fontSize: "12px",
    borderBottom: "none",
    ":hover": {
      opacity: 1,
    },
  },
  buttonGroup: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flexShrink: 0,
  },
  iconButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
  },
  primaryIconButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
    backgroundColor: "#0078d4",
    color: "white",
    borderRadius: "10px",
    ":hover": {
      backgroundColor: "#106ebe",
      color: "white",
    },
  },
});

export const HeaderBar: React.FC<HeaderBarProps> = ({
  onNewChat,
  onShowHistory,
  onShowTemplates,
  selectedModel,
  onModelChange,
  models,
  debugEnabled,
  onDebugChange,
  activeTemplateName,
}) => {
  const styles = useStyles();
  const selectedLabel = models.find(m => m.key === selectedModel)?.label || selectedModel;

  return (
    <div className={styles.header}>
      <div className={styles.leftSection}>
        <Dropdown
          className={styles.dropdown}
          appearance="underline"
          value={selectedLabel}
          selectedOptions={[selectedModel]}
          onOptionSelect={(_, data) => {
            if (data.optionValue && data.optionValue !== selectedModel) {
              onModelChange(data.optionValue as ModelType);
            }
          }}
        >
          {models.map((model) => (
            <Option key={model.key} value={model.key}>
              {model.label}
            </Option>
          ))}
        </Dropdown>
        {/* Debug toggle — hidden by default, enable via localStorage: copilot-debug-visible=true */}
        {localStorage.getItem("copilot-debug-visible") === "true" && (
          <div className={styles.debugRow}>
            <Switch
              checked={debugEnabled}
              onChange={(_, data) => onDebugChange(data.checked)}
              label="Debug"
              style={{ fontSize: "11px" }}
            />
          </div>
        )}
      </div>
      <div className={styles.buttonGroup}>
        <Tooltip content={activeTemplateName ? `Template: ${activeTemplateName}` : "Template Library"} relationship="label">
          <Button
            icon={<SlideLayout24Regular />}
            appearance="subtle"
            onClick={onShowTemplates}
            aria-label="Template Library"
            className={styles.iconButton}
            style={activeTemplateName ? { color: tokens.colorBrandForeground1 } : undefined}
          />
        </Tooltip>
        <Tooltip content="History" relationship="label">
          <Button
            icon={<History24Regular />}
            appearance="subtle"
            onClick={onShowHistory}
            aria-label="History"
            className={styles.iconButton}
          />
        </Tooltip>
        <Tooltip content="New chat" relationship="label">
          <Button
            icon={<Compose24Regular />}
            appearance="subtle"
            onClick={onNewChat}
            aria-label="New chat"
            className={styles.primaryIconButton}
          />
        </Tooltip>
      </div>
    </div>
  );
};
