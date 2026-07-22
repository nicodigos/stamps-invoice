export async function sha256Hex(bytes) {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", input);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export async function assertMatchingSha256(expectedBytes, actualBytes, label = "Saved file") {
  const [expectedHash, actualHash] = await Promise.all([
    sha256Hex(expectedBytes),
    sha256Hex(actualBytes),
  ]);
  if (expectedHash !== actualHash) {
    throw new Error(`${label} content mismatch after upload.`);
  }
}
