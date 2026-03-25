import type { Tool } from "@github/copilot-sdk";

export const insertTemplateSlide: Tool = {
  name: "insert_template_slide",
  description: `Insert a slide from the active template into the current PowerPoint presentation.

The slide is copied from the template file while preserving its visual layout, colors, and design.
After insertion you can update text via update_slide_shape for targeted replacements.

Workflow:
1. Call get_template_info to discover available slide types and their slideIndex in the template.
2. Call insert_template_slide with the desired slideType (and optional replacements).
3. If placeholders need updating, call update_slide_shape on the newly inserted slide.

Slide types: intro | agenda | content | two_column | image_text | chart | table | qa | thank_you | other

replacements: A key→value map where each key is placeholder text to find (case-insensitive substring)
and value is the new text to set on that shape. Shapes whose text matches the key take on the new value.
Example: { "{{TITLE}}": "Q3 2024 Results", "{{SUBTITLE}}": "Finance Team" }`,
  parameters: {
    type: "object",
    properties: {
      templateId: {
        type: "string",
        description: "The ID of the template. Available from the system prompt when a template is active.",
      },
      slideType: {
        type: "string",
        description:
          "Type of slide to insert: intro | agenda | content | two_column | image_text | chart | table | qa | thank_you | other",
      },
      targetIndex: {
        type: "number",
        description:
          "0-based index in the presentation where the slide should be placed. Defaults to appending at the end.",
      },
      replacements: {
        type: "object",
        description:
          "Optional key→value map. Each key is placeholder text (case-insensitive substring match). Matching shapes will have their entire text replaced with the value.",
        additionalProperties: { type: "string" },
      },
    },
    required: ["templateId", "slideType"],
  },
  handler: async (args) => {
    const { templateId, slideType, targetIndex, replacements } = args as {
      templateId: string;
      slideType: string;
      targetIndex?: number;
      replacements?: Record<string, string>;
    };

    // 1. Fetch template from server
    let base64: string;
    let templateSlideIndex: number;

    try {
      const res = await fetch(`/api/templates/${encodeURIComponent(templateId)}`);
      if (!res.ok) {
        return {
          textResultForLlm: `Template not found (id: ${templateId}). Ask the user to select or upload a template.`,
          resultType: "failure",
          error: "Template not found",
          toolTelemetry: {},
        };
      }

      const template = await res.json();
      base64 = template.data;

      // Find the first slide matching slideType
      const match = (template.slides ?? []).find(
        (s: { type: string }) => s.type === slideType,
      );
      if (!match) {
        return {
          textResultForLlm: `No slide tagged as "${slideType}" found in the template. Available types: ${[...(template.slides ?? []).map((s: { type: string }) => s.type)].join(", ")}. Use get_template_info to check.`,
          resultType: "failure",
          error: "Slide type not found in template",
          toolTelemetry: {},
        };
      }
      templateSlideIndex = match.index;
    } catch (e: any) {
      return {
        textResultForLlm: `Failed to load template: ${e.message}`,
        resultType: "failure",
        error: e.message,
        toolTelemetry: {},
      };
    }

    // 2. Insert slide via Office.js
    try {
      const resultMessage = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const countBefore = slides.items.length;

        // Determine insertion point
        const insertAfterIndex =
          targetIndex !== undefined
            ? Math.max(-1, Math.min(targetIndex - 1, countBefore - 1))
            : countBefore - 1;

        const insertOptions: PowerPoint.InsertSlideOptions = {
          formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
        };

        if (insertAfterIndex >= 0) {
          const anchorSlide = slides.items[insertAfterIndex];
          anchorSlide.load("id");
          await context.sync();
          insertOptions.targetSlideId = anchorSlide.id;
        }

        // Insert ALL slides from template (Office.js limitation)
        context.presentation.insertSlidesFromBase64(base64, insertOptions);
        await context.sync();

        // Reload to count inserted slides
        slides.load("items");
        await context.sync();

        const countAfter = slides.items.length;
        const insertedCount = countAfter - countBefore;

        if (insertedCount === 0) {
          throw new Error("insertSlidesFromBase64 inserted 0 slides — check that the PPTX is valid");
        }

        // The inserted block starts at (insertAfterIndex + 1) in 0-based terms
        const insertionStart = insertAfterIndex + 1;
        const desiredPosition = insertionStart + templateSlideIndex;

        // Validate bounds
        if (desiredPosition >= countAfter) {
          throw new Error(
            `templateSlideIndex ${templateSlideIndex} out of range — template only has ${insertedCount} slide(s)`,
          );
        }

        // Delete all inserted slides EXCEPT the desired one
        // Iterate backwards to avoid index shifting
        for (let i = insertionStart + insertedCount - 1; i >= insertionStart; i--) {
          if (i !== desiredPosition) {
            slides.items[i].delete();
          }
        }
        await context.sync();

        // Final slide position after deletions
        // Deleted slides before desiredPosition shift index down
        const deletedBefore = desiredPosition - insertionStart;
        const finalIndex = insertionStart + (desiredPosition - insertionStart - deletedBefore);
        // Simplification: after deleting all others the kept slide lands at insertionStart
        const keptSlideIndex = insertionStart;

        // 3. Apply text replacements if provided
        if (replacements && Object.keys(replacements).length > 0) {
          slides.load("items");
          await context.sync();

          const keptSlide = slides.items[keptSlideIndex];
          keptSlide.shapes.load("items");
          await context.sync();

          for (const shape of keptSlide.shapes.items) {
            try {
              shape.textFrame.textRange.load("text");
            } catch {}
          }
          await context.sync();

          for (const shape of keptSlide.shapes.items) {
            try {
              const currentText = shape.textFrame?.textRange?.text ?? "";
              for (const [placeholder, newText] of Object.entries(replacements)) {
                if (currentText.toLowerCase().includes(placeholder.toLowerCase())) {
                  shape.textFrame.textRange.text = newText;
                  break;
                }
              }
            } catch {}
          }
          await context.sync();
        }

        return `Successfully inserted "${slideType}" slide at position ${keptSlideIndex + 1} (1-based) in the presentation.`;
      });

      return resultMessage;
    } catch (e: any) {
      return {
        textResultForLlm: `Failed to insert template slide: ${e.message}`,
        resultType: "failure",
        error: e.message,
        toolTelemetry: {},
      };
    }
  },
};
