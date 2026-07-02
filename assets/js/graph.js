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

async function graphText(url) {
  const response = await fetch(url, {
    headers: authHeaders(),
  });
  if (!response.ok) {
    throw new GraphRequestError(await response.text(), response.status);
  }
  return response.text();
}

function isNotFound(error) {
  return error instanceof GraphRequestError
    && (error.status === 404 || /itemNotFound|could not be found|not found/i.test(error.message || ""));
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

export async function downloadSharePointTextFile(path) {
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  const fileName = String(path || "").split("/").pop();

  try {
    return await graphText(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`);
  } catch (error) {
    if (!isNotFound(error)) throw error;
  }

  const oneDriveMatch = await findMyDriveFileByName(fileName);
  if (oneDriveMatch) {
    return graphText(`${GRAPH_BASE}/drives/${oneDriveMatch.parentReference.driveId}/items/${oneDriveMatch.id}/content`);
  }

  const searchMatch = await findFileWithMicrosoftSearch(fileName);
  if (searchMatch) {
    return graphText(`${GRAPH_BASE}/drives/${searchMatch.parentReference.driveId}/items/${searchMatch.id}/content`);
  }

  const accessibleDriveMatch = await findAccessibleDriveFile(path, fileName);
  if (accessibleDriveMatch) {
    return graphText(`${GRAPH_BASE}/drives/${accessibleDriveMatch.parentReference.driveId}/items/${accessibleDriveMatch.id}/content`);
  }

  let driveId = "";
  try {
    driveId = await resolveDriveId();
  } catch (error) {
    if (!isNotFound(error)) throw error;
  }

  if (driveId) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`);
    } catch (error) {
      if (!isNotFound(error)) throw error;
    }

    const sharePointMatch = await findDriveFileByName(driveId, fileName);
    if (sharePointMatch) {
      return graphText(`${GRAPH_BASE}/drives/${sharePointMatch.parentReference.driveId}/items/${sharePointMatch.id}/content`);
    }
  }

  throw new GraphRequestError(`No se encontro ${fileName} en SharePoint ni en OneDrive.`, 404);
}

async function findAccessibleDriveFile(path, fileName) {
  let drives = [];
  try {
    drives = (await graphJson(`${GRAPH_BASE}/me/drives?$select=id,name,driveType`)).value || [];
  } catch (error) {
    if (isNotFound(error)) return null;
    throw error;
  }

  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  for (const drive of drives) {
    try {
      const item = await graphJson(`${GRAPH_BASE}/drives/${drive.id}/root:/${encodedPath}?$select=id,name,file,parentReference`);
      if (item.file && String(item.name || "").toLowerCase() === String(fileName || "").toLowerCase()) {
        return item;
      }
    } catch (error) {
      if (!isNotFound(error)) throw error;
    }

    const found = await findDriveFileByName(drive.id, fileName);
    if (found) return found;
  }
  return null;
}

async function findDriveFileByName(driveId, fileName) {
  const query = encodeURIComponent(String(fileName || "").replaceAll("'", "''"));
  try {
    const result = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root/search(q='${query}')?$select=id,name,file,parentReference`);
    return exactFileMatch(result.value, fileName);
  } catch (error) {
    if (isNotFound(error)) return null;
    throw error;
  }
}

async function findMyDriveFileByName(fileName) {
  const query = encodeURIComponent(String(fileName || "").replaceAll("'", "''"));
  try {
    const result = await graphJson(`${GRAPH_BASE}/me/drive/root/search(q='${query}')?$select=id,name,file,parentReference`);
    return exactFileMatch(result.value, fileName);
  } catch (error) {
    if (isNotFound(error)) return null;
    throw error;
  }
}

async function findFileWithMicrosoftSearch(fileName) {
  try {
    const result = await graphJson(`${GRAPH_BASE}/search/query`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: [{
          entityTypes: ["driveItem"],
          query: { queryString: `"${String(fileName || "").replaceAll('"', '\\"')}"` },
          fields: ["id", "name", "file", "parentReference"],
          from: 0,
          size: 25,
        }],
      }),
    });
    const hits = result.value?.flatMap((entry) => (
      (entry.hitsContainers || []).flatMap((container) => container.hits || [])
    )) || [];
    return exactFileMatch(hits.map((hit) => hit.resource), fileName);
  } catch (error) {
    if (isNotFound(error)) return null;
    throw error;
  }
}

function exactFileMatch(items, fileName) {
  return (items || []).find((item) => (
    item.file
    && item.parentReference?.driveId
    && String(item.name || "").toLowerCase() === String(fileName || "").toLowerCase()
  ));
}
