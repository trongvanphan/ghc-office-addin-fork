import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import {
  Button,
  Dropdown,
  Option,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ArrowLeft24Regular, Save24Regular } from "@fluentui/react-icons";
import type { SlideTag, SlideType, TemplateMetadata } from "../templateStorage";
import {
  SLIDE_TYPES,
  SLIDE_TYPE_LABELS,
  saveTemplateTags,
  slideThumbnailUrl,
  uploadThumbnails,
} from "../templateStorage";

interface TemplateTaggingProps {
  template: TemplateMetadata;
  onBack: () => void;
  onSaved: (updated: TemplateMetadata) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    overflow: "hidden",
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
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  scrollArea: {
    flex: 1,
    overflowY: "auto",
    padding: "8px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  slideRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "8px 10px",
    borderRadius: "6px",
    backgroundColor: "var(--colorNeutralBackground1)",
    border: "1px solid var(--colorNeutralStroke2)",
  },
  thumbnail: {
    width: "80px",
    height: "45px",
    objectFit: "cover",
    borderRadius: "3px",
    border: "1px solid var(--colorNeutralStroke2)",
    flexShrink: 0,
    backgroundColor: "var(--colorNeutralBackground3)",
  },
  thumbnailPlaceholder: {
    width: "80px",
    height: "45px",
    borderRadius: "3px",
    border: "1px solid var(--colorNeutralStroke2)",
    flexShrink: 0,
    backgroundColor: "var(--colorNeutralBackground3)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  slideInfo: {
    flex: 1,
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  slideLabel: {
    fontSize: "12px",
    fontWeight: "500",
    color: tokens.colorNeutralForeground2,
  },
  dropdown: {
    minWidth: "0px",
    width: "100%",
    fontSize: "12px",
  },
  footer: {
    padding: "10px 12px",
    borderTop: "1px solid var(--colorNeutralStroke2)",
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    flexShrink: 0,
  },
  captureNote: {
    padding: "8px 12px",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    borderBottom: "1px solid var(--colorNeutralStroke2)",
    flexShrink: 0,
  },
  errorText: {
    padding: "8px 12px",
    fontSize: "12px",
    color: tokens.colorPaletteRedForeground1,
  },
});

// Capture thumbnails for all template slides by temporarily inserting them
async function captureThumbnails(
  base64: string,
  slideCount: number,
  onProgress?: (msg: string) => void,
): Promise<{ slideIndex: number; imageData: string }[]> {
  const thumbnails: { slideIndex: number; imageData: string }[] = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const countBefore = slides.items.length;

    // Insert template slides at the end
    onProgress?.("Đang chèn template tạm để chụp ảnh...");
    const insertOptions: PowerPoint.InsertSlideOptions = {
      formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
    };
    if (countBefore > 0) {
      const lastSlide = slides.items[countBefore - 1];
      lastSlide.load("id");
      await context.sync();
      insertOptions.targetSlideId = lastSlide.id;
    }
    context.presentation.insertSlidesFromBase64(base64, insertOptions);
    await context.sync();

    slides.load("items");
    await context.sync();

    const countAfter = slides.items.length;
    const insertedCount = countAfter - countBefore;

    // Capture each inserted slide
    for (let i = 0; i < insertedCount; i++) {
      const slidePos = countBefore + i;
      if (slidePos >= slides.items.length) break;
      const slide = slides.items[slidePos];

      onProgress?.(`Chụp ảnh slide ${i + 1}/${insertedCount}...`);
      try {
        const imgResult = slide.getImageAsBase64({ width: 400 });
        await context.sync();
        thumbnails.push({ slideIndex: i, imageData: imgResult.value });
      } catch {
        // Thumbnail capture fails on some platforms — skip
      }
    }

    // Delete all inserted slides (in reverse to avoid index shifting)
    onProgress?.("Dọn dẹp slides tạm...");
    slides.load("items");
    await context.sync();
    for (let i = countAfter - 1; i >= countBefore; i--) {
      slides.items[i].delete();
    }
    await context.sync();
  });

  return thumbnails;
}

export const TemplateTagging: React.FC<TemplateTaggingProps> = ({
  template,
  onBack,
  onSaved,
}) => {
  const styles = useStyles();
  const [tags, setTags] = useState<SlideTag[]>(() =>
    template.slides.length > 0
      ? template.slides
      : Array.from({ length: template.slideCount }, (_, i) => ({
          index: i,
          type: "other" as SlideType,
          label: "",
        })),
  );
  const [thumbnails, setThumbnails] = useState<Record<number, string>>({});
  const [isCapturing, setIsCapturing] = useState(false);
  const [captureMsg, setCaptureMsg] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState("");

  // Try to load thumbnails already stored on server
  useEffect(() => {
    const map: Record<number, string> = {};
    let added = false;
    for (let i = 0; i < template.slideCount; i++) {
      const url = slideThumbnailUrl(template.id, i);
      // Probe whether thumbnail exists
      fetch(url, { method: "HEAD" })
        .then((r) => {
          if (r.ok) {
            setThumbnails((prev) => ({ ...prev, [i]: url }));
          }
        })
        .catch(() => {});
    }
    void map;
    void added;
  }, [template.id, template.slideCount]);

  const handleCapture = useCallback(async () => {
    setIsCapturing(true);
    setError("");
    try {
      // Need the template PPTX binary
      const res = await fetch(`/api/templates/${encodeURIComponent(template.id)}`);
      if (!res.ok) throw new Error("Không thể tải template để chụp ảnh");
      const data = await res.json();

      const captured = await captureThumbnails(data.data, template.slideCount, (msg) =>
        setCaptureMsg(msg),
      );

      if (captured.length > 0) {
        await uploadThumbnails(template.id, captured);
        const map: Record<number, string> = {};
        captured.forEach(({ slideIndex }) => {
          map[slideIndex] = `${slideThumbnailUrl(template.id, slideIndex)}?t=${Date.now()}`;
        });
        setThumbnails((prev) => ({ ...prev, ...map }));
      }
    } catch (e: any) {
      setError(`Lỗi khi chụp thumbnail: ${e.message}`);
    } finally {
      setIsCapturing(false);
      setCaptureMsg("");
    }
  }, [template.id, template.slideCount]);

  const handleTypeChange = (slideIndex: number, newType: SlideType) => {
    setTags((prev) =>
      prev.map((t) => (t.index === slideIndex ? { ...t, type: newType } : t)),
    );
  };

  const handleSave = async () => {
    setIsSaving(true);
    setError("");
    try {
      const updated = await saveTemplateTags(template.id, tags);
      onSaved(updated);
    } catch (e: any) {
      setError(`Không thể lưu: ${e.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onBack}
          size="small"
          aria-label="Back"
        />
        <Text className={styles.headerTitle}>Tag Slides — {template.name}</Text>
      </div>

      <div className={styles.captureNote}>
        Gán loại cho từng slide trong template.{" "}
        <Button
          appearance="transparent"
          size="small"
          onClick={handleCapture}
          disabled={isCapturing}
          style={{ padding: "0 4px", fontSize: "11px" }}
        >
          {isCapturing ? captureMsg || "Đang chụp..." : "Chụp thumbnail từ file đang mở"}
        </Button>
        {isCapturing && <Spinner size="extra-tiny" style={{ marginLeft: 4 }} />}
      </div>

      {error && <div className={styles.errorText}>{error}</div>}

      <div className={styles.scrollArea}>
        {tags.map((tag) => (
          <div key={tag.index} className={styles.slideRow}>
            {thumbnails[tag.index] ? (
              <img
                src={thumbnails[tag.index]}
                alt={`Slide ${tag.index + 1}`}
                className={styles.thumbnail}
              />
            ) : (
              <div className={styles.thumbnailPlaceholder}>
                {tag.index + 1}
              </div>
            )}
            <div className={styles.slideInfo}>
              <Text className={styles.slideLabel}>Slide {tag.index + 1}</Text>
              <Dropdown
                className={styles.dropdown}
                appearance="underline"
                value={SLIDE_TYPE_LABELS[tag.type as SlideType] ?? tag.type}
                selectedOptions={[tag.type]}
                onOptionSelect={(_, data) => {
                  if (data.optionValue) {
                    handleTypeChange(tag.index, data.optionValue as SlideType);
                  }
                }}
                size="small"
              >
                {SLIDE_TYPES.map((type) => (
                  <Option key={type} value={type}>
                    {SLIDE_TYPE_LABELS[type]}
                  </Option>
                ))}
              </Dropdown>
            </div>
          </div>
        ))}
      </div>

      <div className={styles.footer}>
        <Button appearance="secondary" size="small" onClick={onBack}>
          Hủy
        </Button>
        <Button
          appearance="primary"
          size="small"
          icon={<Save24Regular />}
          onClick={handleSave}
          disabled={isSaving}
        >
          {isSaving ? "Đang lưu..." : "Lưu"}
        </Button>
      </div>
    </div>
  );
};
