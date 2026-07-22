function parseJsonDetail(text) {
  const source = String(text || "").trim();
  const jsonStart = source.indexOf("{");
  if (jsonStart < 0) return "";
  try {
    const payload = JSON.parse(source.slice(jsonStart));
    return String(payload?.error?.innerError?.message || payload?.error?.message || payload?.message || payload?.error_description || "").trim();
  } catch {
    return "";
  }
}

function cleanDetail(value) {
  const raw = String(value || "").trim();
  return (parseJsonDetail(raw) || raw)
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 700);
}

function statusAdvice(status) {
  if (status === 400) return "Review the selected values and try again.";
  if (status === 401) return "Your Microsoft session may have expired. Disconnect, reconnect, and try again.";
  if (status === 403) return "Your Microsoft account does not have permission for this SharePoint location.";
  if (status === 404) return "The requested SharePoint file or folder could not be found. Refresh the folders and verify the destination.";
  if (status === 409) return "A file or folder with conflicting information already exists. Refresh and try again.";
  if (status === 413) return "The file is too large for this operation.";
  if (status === 429) return "The service is temporarily limiting requests. Wait a moment and try again.";
  if (status >= 500) return "The remote service is temporarily unavailable. Try again shortly.";
  return "";
}

export function describeHttpError(operation, status, responseText) {
  const detail = cleanDetail(responseText);
  const advice = statusAdvice(Number(status));
  return [
    `${operation} failed${status ? ` (HTTP ${status})` : ""}.`,
    detail && detail !== "[object Object]" ? detail : "The service did not provide an error description.",
    advice,
  ].filter(Boolean).join(" ");
}

export function describeError(error) {
  const status = Number(error?.status) || 0;
  const rawMessage = error instanceof Error ? error.message : error;
  const detail = cleanDetail(rawMessage);
  if (/failed to fetch|networkerror|network request failed/i.test(detail)) {
    return "The network request failed. Check your connection and confirm that Microsoft is still connected, then try again.";
  }
  if (/aborterror|timed?\s*out|timeout/i.test(detail)) {
    return "The operation took too long and was stopped. Try again; use fewer or smaller files if the problem continues.";
  }
  const advice = statusAdvice(status);
  if (!detail || detail === "[object Object]") {
    return advice || "An unexpected error occurred. Try again; if it continues, reconnect Microsoft and review the selected files.";
  }
  return advice && !detail.includes(advice) ? `${detail} ${advice}` : detail;
}
