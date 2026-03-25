// Types, constants and browser-side API helpers for the Template Library feature.

export const SLIDE_TYPES = [
  "intro",
  "agenda",
  "content",
  "two_column",
  "image_text",
  "chart",
  "table",
  "qa",
  "thank_you",
  "other",
] as const;

export type SlideType = (typeof SLIDE_TYPES)[number];

export const SLIDE_TYPE_LABELS: Record<SlideType, string> = {
  intro: "Trang giới thiệu (Intro)",
  agenda: "Agenda / Mục lục",
  content: "Nội dung (Content)",
  two_column: "Hai cột (Two Column)",
  image_text: "Hình + Văn bản (Image & Text)",
  chart: "Biểu đồ (Chart)",
  table: "Bảng (Table)",
  qa: "Q&A / Hỏi đáp",
  thank_you: "Trang kết (Thank You)",
  other: "Khác (Other)",
};

export interface SlideTag {
  index: number;
  type: SlideType;
  label: string;
}

export interface TemplateMetadata {
  id: string;
  name: string;
  slideCount: number;
  slides: SlideTag[];
  createdAt: string;
}

export interface TemplateWithData extends TemplateMetadata {
  /** Base64-encoded PPTX binary */
  data: string;
}

// ---- Browser API client -----------------------------------------------------

export async function fetchTemplates(): Promise<TemplateMetadata[]> {
  const res = await fetch("/api/templates");
  if (!res.ok) throw new Error(`Failed to fetch templates: ${res.statusText}`);
  return res.json();
}

export async function fetchTemplate(id: string): Promise<TemplateWithData> {
  const res = await fetch(`/api/templates/${encodeURIComponent(id)}`);
  if (!res.ok) throw new Error(`Failed to fetch template: ${res.statusText}`);
  return res.json();
}

export async function uploadTemplate(
  name: string,
  base64Data: string,
  slideCount: number,
): Promise<TemplateMetadata> {
  const res = await fetch("/api/templates/upload", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, data: base64Data, slideCount }),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(body.error || res.statusText);
  }
  return res.json();
}

export async function saveTemplateTags(
  id: string,
  slides: SlideTag[],
): Promise<TemplateMetadata> {
  const res = await fetch(`/api/templates/${encodeURIComponent(id)}/tags`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ slides }),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(body.error || res.statusText);
  }
  return res.json();
}

export async function deleteTemplate(id: string): Promise<void> {
  const res = await fetch(`/api/templates/${encodeURIComponent(id)}`, {
    method: "DELETE",
  });
  if (!res.ok) throw new Error(`Failed to delete template: ${res.statusText}`);
}

export async function uploadThumbnails(
  id: string,
  thumbnails: { slideIndex: number; imageData: string }[],
): Promise<void> {
  const res = await fetch(`/api/templates/${encodeURIComponent(id)}/thumbnails`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ thumbnails }),
  });
  if (!res.ok) throw new Error(`Failed to upload thumbnails: ${res.statusText}`);
}

export function slideThumbnailUrl(id: string, slideIndex: number): string {
  return `/api/templates/${encodeURIComponent(id)}/slide-thumbnail/${slideIndex}`;
}

// ---- localStorage active-template cache -------------------------------------

const ACTIVE_TEMPLATE_KEY = "copilot-active-template-id";

export function getActiveTemplateId(): string | null {
  try {
    return localStorage.getItem(ACTIVE_TEMPLATE_KEY);
  } catch {
    return null;
  }
}

export function setActiveTemplateId(id: string | null): void {
  try {
    if (id === null) {
      localStorage.removeItem(ACTIVE_TEMPLATE_KEY);
    } else {
      localStorage.setItem(ACTIVE_TEMPLATE_KEY, id);
    }
  } catch {}
}
