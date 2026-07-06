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
    throw new Error("Microsoft is not connected.");
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

async function graphArrayBuffer(url) {
  const response = await fetch(url, {
    headers: authHeaders(),
  });
  if (!response.ok) {
    throw new GraphRequestError(await response.text(), response.status);
  }
  return response.arrayBuffer();
}

function isNotFound(error) {
  return error instanceof GraphRequestError
    && (error.status === 404 || /itemNotFound|could not be found|not found/i.test(error.message || ""));
}

function isAccessDenied(error) {
  return error instanceof GraphRequestError
    && (error.status === 403 || /accessDenied|access denied/i.test(error.message || ""));
}

function isLookupMiss(error) {
  return isNotFound(error) || isAccessDenied(error);
}

export async function resolveDriveId() {
  if (state.driveId) return state.driveId;

  const { spHostname, spSitePath, spDriveName } = state.config;
  if (!spHostname || !spSitePath) {
    throw new Error("Missing SP_HOSTNAME or SP_SITE_PATH.");
  }

  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("Could not resolve the SharePoint drive.");
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

export async function assertSharePointFolderPath(path) {
  const driveId = await resolveDriveId();
  const parts = String(path || "").split("/").filter(Boolean);
  let currentPath = "";

  for (const part of parts) {
    const nextPath = currentPath ? `${currentPath}/${part}` : part;
    const encodedPath = encodeURIComponent(nextPath).replaceAll("%2F", "/");
    try {
      const item = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}?$select=id,name,folder`);
      if (!item.folder) {
        throw new Error(`Expected a folder but found a file at: ${nextPath}`);
      }
      currentPath = nextPath;
    } catch (error) {
      if (isNotFound(error)) {
        const stoppedAt = currentPath || "(drive root)";
        throw new Error(`Missing SharePoint folder: ${nextPath}. Expected full folder path: ${path}. Last existing location: ${stoppedAt}. Create the missing folder manually before uploading.`);
      }
      throw error;
    }
  }
}

async function createSharePointFolder(driveId, parentPath, folderName) {
  const encodedParentPath = encodeURIComponent(parentPath).replaceAll("%2F", "/");
  const url = parentPath
    ? `${GRAPH_BASE}/drives/${driveId}/root:/${encodedParentPath}:/children`
    : `${GRAPH_BASE}/drives/${driveId}/root/children`;

  return graphJson(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    }),
  });
}

export async function ensureSharePointFolderPath(path, options = {}) {
  const driveId = await resolveDriveId();
  const parts = String(path || "").split("/").filter(Boolean);
  let currentPath = "";

  for (let index = 0; index < parts.length; index += 1) {
    const part = parts[index];
    const nextPath = currentPath ? `${currentPath}/${part}` : part;
    const encodedPath = encodeURIComponent(nextPath).replaceAll("%2F", "/");

    try {
      const item = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}?$select=id,name,folder`);
      if (!item.folder) {
        throw new Error(`Expected a folder but found a file at: ${nextPath}`);
      }
      currentPath = nextPath;
    } catch (error) {
      if (!isNotFound(error)) throw error;

      const missingParts = parts.slice(index);
      const lastExistingPath = currentPath || "(drive root)";
      const confirmed = await options.confirmCreate?.({
        fullPath: path,
        missingPath: nextPath,
        missingParts,
        lastExistingPath,
      });

      if (!confirmed) {
        throw new Error(`Folder creation cancelled. Missing SharePoint folder: ${nextPath}. Expected full folder path: ${path}. Last existing location: ${lastExistingPath}.`);
      }

      let parentPath = currentPath;
      for (const missingPart of missingParts) {
        await createSharePointFolder(driveId, parentPath, missingPart);
        parentPath = parentPath ? `${parentPath}/${missingPart}` : missingPart;
      }
      return;
    }
  }
}

export async function downloadSharePointTextFile(path) {
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  const fileName = String(path || "").split("/").pop();

  const driveId = await resolveDriveId();

  try {
    return await graphText(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`);
  } catch (error) {
    if (!isLookupMiss(error)) throw error;
  }

  const sharePointMatch = await findDriveFileByName(driveId, fileName);
  if (sharePointMatch) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${sharePointMatch.parentReference.driveId}/items/${sharePointMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const folderMatch = await findFileInDriveFolderByName(driveId, parentPath(path), fileName);
  if (folderMatch) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${folderMatch.parentReference.driveId}/items/${folderMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const accessibleDriveMatch = await findAccessibleDriveFile(path, fileName);
  if (accessibleDriveMatch) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${accessibleDriveMatch.parentReference.driveId}/items/${accessibleDriveMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  try {
    return await graphText(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`);
  } catch (error) {
    if (!isLookupMiss(error)) throw error;
  }

  const oneDriveMatch = await findMyDriveFileByName(fileName);
  if (oneDriveMatch) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${oneDriveMatch.parentReference.driveId}/items/${oneDriveMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const searchMatch = await findFileWithMicrosoftSearch(fileName);
  if (searchMatch) {
    try {
      return await graphText(`${GRAPH_BASE}/drives/${searchMatch.parentReference.driveId}/items/${searchMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  throw new GraphRequestError(await fileAccessMessage(driveId, path, fileName), 403);
}

export async function downloadSharePointFile(path) {
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  const fileName = String(path || "").split("/").pop();

  const driveId = await resolveDriveId();

  try {
    return await graphArrayBuffer(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`);
  } catch (error) {
    if (!isLookupMiss(error)) throw error;
  }

  const sharePointMatch = await findDriveFileByName(driveId, fileName);
  if (sharePointMatch) {
    try {
      return await graphArrayBuffer(`${GRAPH_BASE}/drives/${sharePointMatch.parentReference.driveId}/items/${sharePointMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const folderMatch = await findFileInDriveFolderByName(driveId, parentPath(path), fileName);
  if (folderMatch) {
    try {
      return await graphArrayBuffer(`${GRAPH_BASE}/drives/${folderMatch.parentReference.driveId}/items/${folderMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const accessibleDriveMatch = await findAccessibleDriveFile(path, fileName);
  if (accessibleDriveMatch) {
    try {
      return await graphArrayBuffer(`${GRAPH_BASE}/drives/${accessibleDriveMatch.parentReference.driveId}/items/${accessibleDriveMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  try {
    return await graphArrayBuffer(`${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`);
  } catch (error) {
    if (!isLookupMiss(error)) throw error;
  }

  const oneDriveMatch = await findMyDriveFileByName(fileName);
  if (oneDriveMatch) {
    try {
      return await graphArrayBuffer(`${GRAPH_BASE}/drives/${oneDriveMatch.parentReference.driveId}/items/${oneDriveMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  const searchMatch = await findFileWithMicrosoftSearch(fileName);
  if (searchMatch) {
    try {
      return await graphArrayBuffer(`${GRAPH_BASE}/drives/${searchMatch.parentReference.driveId}/items/${searchMatch.id}/content`);
    } catch (error) {
      if (!isLookupMiss(error)) throw error;
    }
  }

  throw new GraphRequestError(await fileAccessMessage(driveId, path, fileName), 403);
}

async function findAccessibleDriveFile(path, fileName) {
  let drives = [];
  try {
    drives = (await graphJson(`${GRAPH_BASE}/me/drives?$select=id,name,driveType`)).value || [];
  } catch (error) {
    if (isLookupMiss(error)) return null;
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
      if (!isLookupMiss(error)) throw error;
    }

    const found = await findDriveFileByName(drive.id, fileName);
    if (found) return found;
  }
  return null;
}

async function findFileInDriveFolderByName(driveId, folderPath, fileName) {
  if (!folderPath) return null;
  const encodedPath = encodeURIComponent(folderPath).replaceAll("%2F", "/");
  try {
    const result = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/children?$select=id,name,file,parentReference`);
    return exactFileMatch(result.value, fileName);
  } catch (error) {
    if (isLookupMiss(error)) return null;
    throw error;
  }
}

async function fileAccessMessage(driveId, path, fileName) {
  const folder = parentPath(path) || "root";
  const names = await listDriveFolderFileNames(driveId, parentPath(path));
  if (!names.length) {
    return `Could not read ${fileName}. No visible files could be listed in ${folder}. Check permissions and location.`;
  }
  const visibleWorkbooks = names.filter((name) => /\.xlsx$/i.test(name));
  const visibleText = visibleWorkbooks.length ? visibleWorkbooks.join(", ") : names.slice(0, 12).join(", ");
  return `Could not read ${fileName}. Visible files in ${folder}: ${visibleText || "none"}.`;
}

async function listDriveFolderFileNames(driveId, folderPath) {
  if (!folderPath) return [];
  const encodedPath = encodeURIComponent(folderPath).replaceAll("%2F", "/");
  try {
    const result = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/children?$select=name,file`);
    return (result.value || [])
      .filter((item) => item.file)
      .map((item) => item.name)
      .filter(Boolean)
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  } catch (error) {
    if (isLookupMiss(error)) return [];
    throw error;
  }
}

async function findDriveFileByName(driveId, fileName) {
  const query = encodeURIComponent(String(fileName || "").replaceAll("'", "''"));
  try {
    const result = await graphJson(`${GRAPH_BASE}/drives/${driveId}/root/search(q='${query}')?$select=id,name,file,parentReference`);
    return exactFileMatch(result.value, fileName);
  } catch (error) {
    if (isLookupMiss(error)) return null;
    throw error;
  }
}

async function findMyDriveFileByName(fileName) {
  const query = encodeURIComponent(String(fileName || "").replaceAll("'", "''"));
  try {
    const result = await graphJson(`${GRAPH_BASE}/me/drive/root/search(q='${query}')?$select=id,name,file,parentReference`);
    return exactFileMatch(result.value, fileName);
  } catch (error) {
    if (isLookupMiss(error)) return null;
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
    if (isLookupMiss(error)) return null;
    throw error;
  }
}

function parentPath(path) {
  const parts = String(path || "").split("/").filter(Boolean);
  parts.pop();
  return parts.join("/");
}

function exactFileMatch(items, fileName) {
  return (items || []).find((item) => (
    item.file
    && item.parentReference?.driveId
    && String(item.name || "").toLowerCase() === String(fileName || "").toLowerCase()
  ));
}
