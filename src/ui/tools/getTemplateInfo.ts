import type { Tool } from "@github/copilot-sdk";

export const getTemplateInfo: Tool = {
  name: "get_template_info",
  description: `Get information about the currently active PowerPoint slide template.

Returns the template name, total slide count, and a list of tagged slides — each slide has an index, a type (intro, agenda, content, two_column, image_text, chart, table, qa, thank_you, other), and an optional label.

Call this tool first when the user wants to create a presentation using the active template, so you know which slide types are available and at which indices they sit in the template file.

The returned slideIndex for each slide is used when calling insert_template_slide.`,
  parameters: {
    type: "object",
    properties: {
      templateId: {
        type: "string",
        description: "The ID of the template to query. Provided in the system prompt when a template is active.",
      },
    },
    required: ["templateId"],
  },
  handler: async (args) => {
    const { templateId } = args as { templateId: string };

    try {
      const res = await fetch(`/api/templates/${encodeURIComponent(templateId)}`);
      if (!res.ok) {
        return {
          textResultForLlm: `Template not found (id: ${templateId}). Ask the user to select a template from the Template Library.`,
          resultType: "failure",
          error: "Template not found",
          toolTelemetry: {},
        };
      }

      // Strip large base64 data — we only need metadata here
      const { data: _data, ...meta } = await res.json();

      const slideList = (meta.slides ?? [])
        .map((s: { index: number; type: string; label: string }) =>
          `  - slideIndex ${s.index}: type="${s.type}"${s.label ? ` (${s.label})` : ""}`,
        )
        .join("\n");

      const summary = [
        `Template: "${meta.name}"`,
        `Total slides in template: ${meta.slideCount}`,
        `Tagged slides:\n${slideList || "  (none tagged yet)"}`,
        "",
        "Use insert_template_slide to insert a slide from this template by type.",
        "If multiple slides share the same type, the first one is used by default.",
      ].join("\n");

      return summary;
    } catch (e: any) {
      return {
        textResultForLlm: `Failed to fetch template info: ${e.message}`,
        resultType: "failure",
        error: e.message,
        toolTelemetry: {},
      };
    }
  },
};
