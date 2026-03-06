import { PythonBridgeClient } from "../bridge/client.js";

interface ResourceEntry {
  uri: string;
  name: string;
  description: string;
  mimeType: string;
}

function asResource(uri: string, name: string, description: string, mimeType = "application/json"): ResourceEntry {
  return { uri, name, description, mimeType };
}

export async function listResources(bridge: PythonBridgeClient): Promise<ResourceEntry[]> {
  const resources: ResourceEntry[] = [
    asResource("pptx://presentations", "Open Presentations", "All active presentation sessions."),
  ];

  const list = (await bridge.call("pptx_list_open_presentations", {})) as {
    presentations?: Array<{ presentation_id: string }>;
  };

  const presentations = list.presentations ?? [];

  for (const presentation of presentations) {
    const id = presentation.presentation_id;
    resources.push(
      asResource(`pptx://presentations/${id}`, `Presentation ${id}`, "Presentation metadata and slide summary."),
      asResource(`pptx://presentations/${id}/slides`, `Presentation ${id} Slides`, "Slide index for presentation."),
      asResource(`pptx://presentations/${id}/layouts`, `Presentation ${id} Layouts`, "Available layout names and placeholder types."),
      asResource(`pptx://presentations/${id}/masters`, `Presentation ${id} Masters`, "Slide masters list."),
      asResource(`pptx://presentations/${id}/theme`, `Presentation ${id} Theme`, "Theme colors and font scheme."),
    );

    const state = (await bridge.call("pptx_get_presentation_state", { presentation_id: id })) as {
      slides?: Array<{ index: number }>;
    };

    for (const slide of state.slides ?? []) {
      resources.push(
        asResource(
          `pptx://presentations/${id}/slides/${slide.index}`,
          `Presentation ${id} Slide ${slide.index}`,
          "Single slide details including shapes and placeholders.",
        ),
        asResource(
          `pptx://presentations/${id}/slides/${slide.index}/snapshot`,
          `Presentation ${id} Slide ${slide.index} Snapshot`,
          "Base64 JPEG preview for the slide.",
        ),
      );
    }
  }

  return resources;
}

function parsePresentationUri(uri: string): string[] {
  const prefix = "pptx://";
  if (!uri.startsWith(prefix)) {
    return [];
  }

  const rest = uri.slice(prefix.length);
  if (!rest) {
    return [];
  }

  return rest.split("/").filter(Boolean);
}

export async function readResource(uri: string, bridge: PythonBridgeClient): Promise<unknown> {
  const parts = parsePresentationUri(uri);

  if (parts.length === 1 && parts[0] === "presentations") {
    return bridge.call("pptx_list_open_presentations", {});
  }

  if (parts.length < 2 || parts[0] !== "presentations") {
    throw new Error(`Unsupported resource URI: ${uri}`);
  }

  const presentationId = parts[1];

  if (parts.length === 2) {
    return bridge.call("pptx_get_presentation_state", { presentation_id: presentationId });
  }

  if (parts.length === 3 && parts[2] === "slides") {
    const state = (await bridge.call("pptx_get_presentation_state", {
      presentation_id: presentationId,
    })) as { slides: unknown[] };
    return { presentation_id: presentationId, slides: state.slides };
  }

  if (parts.length === 3 && parts[2] === "layouts") {
    return bridge.call("pptx_get_layouts", { presentation_id: presentationId });
  }

  if (parts.length === 3 && parts[2] === "masters") {
    return bridge.call("pptx_get_masters", { presentation_id: presentationId });
  }

  if (parts.length === 3 && parts[2] === "theme") {
    return bridge.call("pptx_get_theme", { presentation_id: presentationId });
  }

  if (parts.length === 4 && parts[2] === "slides") {
    const slideIndex = Number(parts[3]);
    return bridge.call("pptx_get_slide", {
      presentation_id: presentationId,
      slide_index: slideIndex,
    });
  }

  if (parts.length === 5 && parts[2] === "slides" && parts[4] === "snapshot") {
    const slideIndex = Number(parts[3]);
    return bridge.call("pptx_get_slide_snapshot", {
      presentation_id: presentationId,
      slide_index: slideIndex,
    });
  }

  throw new Error(`Unsupported resource URI: ${uri}`);
}
