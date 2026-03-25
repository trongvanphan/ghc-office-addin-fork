import * as React from "react";
import { useState, useEffect, useCallback, useRef } from "react";
import {
  Button,
  Spinner,
  Text,
  Badge,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowLeft24Regular,
  Delete24Regular,
  Tag24Regular,
  Checkmark24Regular,
} from "@fluentui/react-icons";
import type { TemplateMetadata } from "../templateStorage";
import {
  fetchTemplates,
  uploadTemplate,
  deleteTemplate,
} from "../templateStorage";
import { TemplateTagging } from "./TemplateTagging";

interface TemplateManagerProps {
  activeTemplateId: string | null;
  onClose: () => void;
  onSelectTemplate: (template: TemplateMetadata | null) => void;
}

type View = "list" | "upload" | "tagging";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    backgroundColor: "var(--colorNeutralBackground2)",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "10px 12px",
    borderBottom: "1px solid var(--colorNeutralStroke2)",
    flexShrink: 0,
  },
  headerTitle: {
    fontWeight: "600",
    fontSize: "14px",
    flex: 1,
  },
  scrollArea: {
    flex: 1,
    overflowY: "auto",
    padding: "8px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "12px",
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
    padding: "24px",
    textAlign: "center",
  },
  templateCard: {
    padding: "10px 12px",
    borderRadius: "6px",
    backgroundColor: "var(--colorNeutralBackground1)",
    border: "1px solid var(--colorNeutralStroke2)",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  templateCardActive: {
    border: "1px solid var(--colorBrandStroke1)",
    backgroundColor: "var(--colorBrandBackground2)",
  },
  cardRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  templateName: {
    flex: 1,
    fontSize: "13px",
    fontWeight: "500",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  templateMeta: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  actionRow: {
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
  },
  footer: {
    padding: "10px 12px",
    borderTop: "1px solid var(--colorNeutralStroke2)",
    flexShrink: 0,
  },
  // Upload form
  uploadArea: {
    flex: 1,
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  dropZone: {
    border: `2px dashed var(--colorNeutralStroke1)`,
    borderRadius: "8px",
    padding: "24px",
    textAlign: "center",
    cursor: "pointer",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
    fontSize: "13px",
    color: tokens.colorNeutralForeground2,
    ":hover": {
      border: `2px solid var(--colorBrandStroke1)`,
      backgroundColor: "var(--colorBrandBackground2)",
    },
  },
  nameInput: {
    width: "100%",
    padding: "6px 8px",
    fontSize: "13px",
    borderRadius: "4px",
    border: `1px solid var(--colorNeutralStroke1)`,
    backgroundColor: "var(--colorNeutralBackground1)",
    color: "var(--colorNeutralForeground1)",
    outline: "none",
    boxSizing: "border-box",
    ":focus": {
      outline: `2px solid var(--colorBrandStroke1)`,
      outlineOffset: "-1px",
    },
  },
  errorText: {
    fontSize: "12px",
    color: tokens.colorPaletteRedForeground1,
  },
});

function formatDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  } catch {
    return "";
  }
}

function countTaggedSlides(template: TemplateMetadata): number {
  return template.slides.filter((s) => s.type !== "other").length;
}

// Convert a File to base64 string (strips data URL prefix)
function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      // Strip "data:...;base64," prefix
      const base64 = result.split(",")[1];
      resolve(base64);
    };
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

// Count slides in a PPTX base64 by scanning the zip manifest for slide entries.
// PPTX is a zip file; ppt/slides/slide*.xml entries reveal the slide count.
function countSlidesInBase64(base64: string): number {
  try {
    const binary = atob(base64);
    // Scan for central directory file headers to find ppt/slides/slideN.xml entries
    const re = /ppt\/slides\/slide\d+\.xml/g;
    const matches = binary.match(re);
    return matches ? new Set(matches).size : 0;
  } catch {
    return 0;
  }
}

export const TemplateManager: React.FC<TemplateManagerProps> = ({
  activeTemplateId,
  onClose,
  onSelectTemplate,
}) => {
  const styles = useStyles();
  const [view, setView] = useState<View>("list");
  const [templates, setTemplates] = useState<TemplateMetadata[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [taggingTemplate, setTaggingTemplate] = useState<TemplateMetadata | null>(null);

  // Upload form state
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadName, setUploadName] = useState("");
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState("");

  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadTemplates = useCallback(async () => {
    setIsLoading(true);
    try {
      setTemplates(await fetchTemplates());
    } catch {
      // Silently ignore — server may still be starting
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    loadTemplates();
  }, [loadTemplates]);

  const handleDelete = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    try {
      await deleteTemplate(id);
      if (activeTemplateId === id) onSelectTemplate(null);
      setTemplates((prev) => prev.filter((t) => t.id !== id));
    } catch (err: any) {
      console.error("Delete template failed:", err);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadFile(file);
    if (!uploadName) setUploadName(file.name.replace(/\.pptx$/i, ""));
    setUploadError("");
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file?.name.endsWith(".pptx")) {
      setUploadFile(file);
      if (!uploadName) setUploadName(file.name.replace(/\.pptx$/i, ""));
      setUploadError("");
    } else {
      setUploadError("Chỉ chấp nhận file .pptx");
    }
  };

  const handleUpload = async () => {
    if (!uploadFile || !uploadName.trim()) return;
    setIsUploading(true);
    setUploadError("");
    try {
      const base64 = await fileToBase64(uploadFile);
      const slideCount = countSlidesInBase64(base64);
      if (slideCount === 0) {
        throw new Error(
          "Không đọc được số slide. Hãy chắc chắn file là .pptx hợp lệ.",
        );
      }
      const created = await uploadTemplate(uploadName.trim(), base64, slideCount);
      setTemplates((prev) => [created, ...prev]);
      setUploadFile(null);
      setUploadName("");
      // Immediately go to tagging
      setTaggingTemplate(created);
      setView("tagging");
    } catch (err: any) {
      setUploadError(err.message);
    } finally {
      setIsUploading(false);
    }
  };

  // Tagging completed
  const handleTaggingSaved = (updated: TemplateMetadata) => {
    setTemplates((prev) => prev.map((t) => (t.id === updated.id ? updated : t)));
    setTaggingTemplate(null);
    setView("list");
  };

  // ---- Tagging view ----
  if (view === "tagging" && taggingTemplate) {
    return (
      <TemplateTagging
        template={taggingTemplate}
        onBack={() => setView("list")}
        onSaved={handleTaggingSaved}
      />
    );
  }

  // ---- Upload view ----
  if (view === "upload") {
    return (
      <div className={styles.container}>
        <div className={styles.header}>
          <Button
            appearance="subtle"
            icon={<ArrowLeft24Regular />}
            size="small"
            onClick={() => { setView("list"); setUploadFile(null); setUploadName(""); setUploadError(""); }}
            aria-label="Back"
          />
          <Text className={styles.headerTitle}>Upload Template mới</Text>
        </div>

        <div className={styles.uploadArea}>
          <div
            className={styles.dropZone}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(e) => e.preventDefault()}
            onDrop={handleDrop}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => e.key === "Enter" && fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".pptx"
              style={{ display: "none" }}
              onChange={handleFileChange}
            />
            {uploadFile ? (
              <>
                <Text style={{ fontWeight: 600 }}>{uploadFile.name}</Text>
                <Text style={{ fontSize: "11px", color: tokens.colorNeutralForeground3 }}>
                  {(uploadFile.size / 1024 / 1024).toFixed(1)} MB
                </Text>
              </>
            ) : (
              <>
                <Text>Kéo thả file .pptx vào đây</Text>
                <Text style={{ fontSize: "11px" }}>hoặc click để chọn file</Text>
              </>
            )}
          </div>

          <div>
            <Text style={{ fontSize: "12px", marginBottom: "4px", display: "block" }}>
              Tên template
            </Text>
            <input
              className={styles.nameInput}
              type="text"
              placeholder="Nhập tên template..."
              value={uploadName}
              onChange={(e) => setUploadName(e.target.value)}
              maxLength={100}
            />
          </div>

          {uploadError && <Text className={styles.errorText}>{uploadError}</Text>}

          <Button
            appearance="primary"
            onClick={handleUpload}
            disabled={!uploadFile || !uploadName.trim() || isUploading}
            icon={isUploading ? <Spinner size="extra-tiny" /> : undefined}
          >
            {isUploading ? "Đang upload..." : "Upload & Tag Slides"}
          </Button>
        </div>
      </div>
    );
  }

  // ---- List view (default) ----
  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          size="small"
          onClick={onClose}
          aria-label="Close"
        />
        <Text className={styles.headerTitle}>Template Library</Text>
      </div>

      {isLoading ? (
        <div className={styles.emptyState}>
          <Spinner size="small" label="Đang tải..." />
        </div>
      ) : templates.length === 0 ? (
        <div className={styles.emptyState}>
          <Text>Chưa có template nào.</Text>
          <Text style={{ fontSize: "12px" }}>
            Upload file .pptx mẫu để bắt đầu.
          </Text>
          <Button appearance="primary" size="small" onClick={() => setView("upload")}>
            Upload Template
          </Button>
        </div>
      ) : (
        <div className={styles.scrollArea}>
          {templates.map((t) => {
            const isActive = t.id === activeTemplateId;
            const tagged = countTaggedSlides(t);
            return (
              <div
                key={t.id}
                className={`${styles.templateCard} ${isActive ? styles.templateCardActive : ""}`}
              >
                <div className={styles.cardRow}>
                  <Text className={styles.templateName}>{t.name}</Text>
                  {isActive && (
                    <Badge appearance="filled" color="brand" size="small">
                      Active
                    </Badge>
                  )}
                </div>
                <div className={styles.cardRow}>
                  <Text className={styles.templateMeta}>
                    {t.slideCount} slides · {tagged}/{t.slideCount} tagged · {formatDate(t.createdAt)}
                  </Text>
                </div>
                <div className={styles.actionRow}>
                  {isActive ? (
                    <Button
                      appearance="secondary"
                      size="small"
                      onClick={() => onSelectTemplate(null)}
                    >
                      Bỏ chọn
                    </Button>
                  ) : (
                    <Button
                      appearance="primary"
                      size="small"
                      icon={<Checkmark24Regular />}
                      onClick={() => onSelectTemplate(t)}
                    >
                      Dùng template này
                    </Button>
                  )}
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<Tag24Regular />}
                    onClick={() => { setTaggingTemplate(t); setView("tagging"); }}
                  >
                    Tag Slides
                  </Button>
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<Delete24Regular />}
                    onClick={(e) => handleDelete(t.id, e)}
                    style={{ color: tokens.colorPaletteRedForeground1 }}
                    aria-label="Delete template"
                  />
                </div>
              </div>
            );
          })}
        </div>
      )}

      <div className={styles.footer}>
        <Button
          appearance="primary"
          size="small"
          style={{ width: "100%" }}
          onClick={() => setView("upload")}
        >
          + Upload Template mới
        </Button>
      </div>
    </div>
  );
};
