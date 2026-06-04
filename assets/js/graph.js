import { state } from "./state.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

class GraphRequestError extends Error {
  constructor(message, status) {
    super(message);
    this.name = "GraphRequestError";
    this.status = status;
  }
}

function authHeaders() {
  if (!state.graphToken) {
    throw new Error("Microsoft no esta conectado.");
  }
  return { Authorization: `Bearer ${state.graphToken}` };
}

async function graphJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...authHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    throw new GraphRequestError(await response.text(), response.status);
  }
  return response.status === 204 ? {} : response.json();
}

export async function resolveDriveId() {
  if (state.driveId) return state.driveId;

  const { spHostname, spSitePath, spDriveName } = state.config;
  if (!spHostname || !spSitePath) {
    throw new Error("Falta SP_HOSTNAME o SP_SITE_PATH.");
  }

  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("No se pudo resolver el drive de SharePoint.");
  }
  state.driveId = drive.id;
  return drive.id;
}

export async function uploadSharePointFile(path, content, contentType = "application/pdf") {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  return graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`, {
    method: "PUT",
    headers: { "Content-Type": contentType },
    body: content,
  });
}
