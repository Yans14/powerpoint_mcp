import path from "node:path";

import { PythonBridgeClient } from "../bridge/client.js";

interface ResourceEntry {
  uri: string;
  name: string;
  title: string;
  description: string;
  mimeType: string;
}

interface ResourceTemplateEntry {
  uriTemplate: string;
  name: string;
  title: string;
  description: string;
  mimeType: string;
}

export type ResourceContent =
  | { uri: string; mimeType: string; text: string }
  | { uri: string; mimeType: string; blob: string };

/** The URI does not name a resource this server can serve. */
export class ResourceNotFoundError extends Error {
  public constructor(public readonly uri: string) {
    super(`Resource not found: ${uri}`);
    this.name = "ResourceNotFoundError";
  }
}

const JSON_MIME = "application/json";
const ROOT = "pptx://presentations";
const PRESENTATION = `${ROOT}/{presentation_id}`;

/**
 * Everything below a presentation is addressed through templates rather than listed: a listing
 * would need one bridge call per presentation and one entry per slide, and go stale on every edit.
 */
export const RESOURCE_TEMPLATES: ResourceTemplateEntry[] = [
  {
    uriTemplate: PRESENTATION,
    name: "presentation",
    title: "Presentation",
    description: "Slide count and per-slide summary (title, layout, shape count) of an open presentation.",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/slides`,
    name: "presentation-slides",
    title: "Slides",
    description: "Slide index of an open presentation.",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/layouts`,
    name: "presentation-layouts",
    title: "Layouts",
    description: "Layout names and placeholder types available in the presentation.",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/masters`,
    name: "presentation-masters",
    title: "Slide masters",
    description: "Slide masters of the presentation.",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/theme`,
    name: "presentation-theme",
    title: "Theme",
    description: "Theme colors and font scheme of the presentation.",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/slides/{slide_index}`,
    name: "slide",
    title: "Slide",
    description: "Shapes, placeholders and notes of one slide (1-based index).",
    mimeType: JSON_MIME,
  },
  {
    uriTemplate: `${PRESENTATION}/slides/{slide_index}/snapshot`,
    name: "slide-snapshot",
    title: "Slide snapshot",
    description: "JPEG rendering of one slide. Requires LibreOffice and poppler in OOXML mode.",
    mimeType: "image/jpeg",
  },
];

/** Lists the open presentations with a single bridge call; deeper resources use the templates. */
export async function listResources(bridge: PythonBridgeClient): Promise<ResourceEntry[]> {
  const list = await bridge.call<{
    presentations?: Array<{ presentation_id: string; original_path?: string }>;
  }>("pptx_list_open_presentations", {});

  const resources: ResourceEntry[] = [
    {
      uri: ROOT,
      name: "open-presentations",
      title: "Open presentations",
      description: "All active presentation sessions.",
      mimeType: JSON_MIME,
    },
  ];

  for (const presentation of list.presentations ?? []) {
    const label = presentation.original_path ? path.basename(presentation.original_path) : "Untitled presentation";
    resources.push({
      uri: `${ROOT}/${presentation.presentation_id}`,
      name: `presentation-${presentation.presentation_id}`,
      title: label,
      description: `Slide summary of ${label}. Slides, layouts, theme and snapshots are available under this URI.`,
      mimeType: JSON_MIME,
    });
  }

  return resources;
}

function jsonContent(uri: string, data: unknown): ResourceContent {
  return { uri, mimeType: JSON_MIME, text: JSON.stringify(data) };
}

function parseSlideIndex(value: string, uri: string): number {
  if (!/^[1-9]\d*$/.test(value)) {
    throw new ResourceNotFoundError(uri);
  }
  return Number(value);
}

/** Throws ResourceNotFoundError for an unknown URI; bridge errors propagate as BridgeClientError. */
export async function readResource(uri: string, bridge: PythonBridgeClient): Promise<ResourceContent> {
  const prefix = "pptx://";
  const parts = uri.startsWith(prefix) ? uri.slice(prefix.length).split("/") : [];

  if (parts[0] !== "presentations" || parts.some((part) => part === "")) {
    throw new ResourceNotFoundError(uri);
  }

  if (parts.length === 1) {
    return jsonContent(uri, await bridge.call("pptx_list_open_presentations", {}));
  }

  const presentationId = parts[1];
  const [, , section, slide, leaf] = parts;

  if (parts.length === 2) {
    return jsonContent(uri, await bridge.call("pptx_get_presentation_state", { presentation_id: presentationId }));
  }

  if (parts.length === 3) {
    switch (section) {
      case "slides": {
        const state = await bridge.call<{ slides: unknown[] }>("pptx_get_presentation_state", {
          presentation_id: presentationId,
        });
        return jsonContent(uri, { presentation_id: presentationId, slides: state.slides });
      }
      case "layouts":
        return jsonContent(uri, await bridge.call("pptx_get_layouts", { presentation_id: presentationId }));
      case "masters":
        return jsonContent(uri, await bridge.call("pptx_get_masters", { presentation_id: presentationId }));
      case "theme":
        return jsonContent(uri, await bridge.call("pptx_get_theme", { presentation_id: presentationId }));
      default:
        throw new ResourceNotFoundError(uri);
    }
  }

  if (section === "slides" && parts.length === 4) {
    const slideIndex = parseSlideIndex(slide, uri);
    return jsonContent(
      uri,
      await bridge.call("pptx_get_slide", { presentation_id: presentationId, slide_index: slideIndex }),
    );
  }

  if (section === "slides" && parts.length === 5 && leaf === "snapshot") {
    const slideIndex = parseSlideIndex(slide, uri);
    const snapshot = await bridge.call<{ mime_type?: string; snapshot_base64: string }>("pptx_get_slide_snapshot", {
      presentation_id: presentationId,
      slide_index: slideIndex,
    });
    return { uri, mimeType: snapshot.mime_type ?? "image/jpeg", blob: snapshot.snapshot_base64 };
  }

  throw new ResourceNotFoundError(uri);
}
