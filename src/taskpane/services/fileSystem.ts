import { get, set } from "idb-keyval";

const ROOT_HANDLE_KEY = "certgen-root-dir-handle";
const LAST_BATCH_ID_KEY = "certgen-last-batch-id";

/** Prompts the user once to choose the root folder under which all Batches/ are created. */
export async function pickRootDirectory(): Promise<FileSystemDirectoryHandle> {
  const handle = await showDirectoryPicker({ id: "certgen-root", mode: "readwrite" });
  await set(ROOT_HANDLE_KEY, handle);
  return handle;
}

/** Returns the previously chosen folder handle, if any, without prompting for permission. */
export async function getPersistedRootHandle(): Promise<FileSystemDirectoryHandle | null> {
  return (await get<FileSystemDirectoryHandle>(ROOT_HANDLE_KEY)) ?? null;
}

/** Checks read/write permission on a handle without prompting - safe to call outside a user gesture. */
export async function queryPermission(
  handle: FileSystemHandle,
  mode: FileSystemPermissionMode = "readwrite"
): Promise<PermissionState> {
  return handle.queryPermission({ mode });
}

/** Prompts for permission on a handle. Must be called from within a user-gesture (click) handler, or the browser throws. */
export async function requestPermission(
  handle: FileSystemHandle,
  mode: FileSystemPermissionMode = "readwrite"
): Promise<boolean> {
  return (await handle.requestPermission({ mode })) === "granted";
}

/** Returns the previously chosen root folder only if permission is already granted (never prompts). */
export async function getRootDirectory(): Promise<FileSystemDirectoryHandle | null> {
  const handle = await getPersistedRootHandle();
  if (!handle) return null;
  return (await queryPermission(handle)) === "granted" ? handle : null;
}

export function sanitizeFolderName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "-").trim();
}


/** Remembers which batch was last opened so the add-in can resume on it next time. */
export async function setLastBatchId(batchId: string): Promise<void> {
  await set(LAST_BATCH_ID_KEY, batchId);
}

export async function getLastBatchId(): Promise<string | null> {
  return (await get<string>(LAST_BATCH_ID_KEY)) ?? null;
}

async function getBatchesDirectory(root: FileSystemDirectoryHandle): Promise<FileSystemDirectoryHandle> {
  return root.getDirectoryHandle("Batches", { create: true });
}

export async function createBatchFolder(
  root: FileSystemDirectoryHandle,
  batchId: string
): Promise<{ batchDir: FileSystemDirectoryHandle; certificatesDir: FileSystemDirectoryHandle }> {
  const batchesDir = await getBatchesDirectory(root);
  const batchDir = await batchesDir.getDirectoryHandle(sanitizeFolderName(batchId), { create: true });
  const certificatesDir = await batchDir.getDirectoryHandle("Certificates", { create: true });
  return { batchDir, certificatesDir };
}

export async function openBatchFolder(
  root: FileSystemDirectoryHandle,
  batchId: string
): Promise<{ batchDir: FileSystemDirectoryHandle; certificatesDir: FileSystemDirectoryHandle }> {
  const batchesDir = await getBatchesDirectory(root);
  const batchDir = await batchesDir.getDirectoryHandle(sanitizeFolderName(batchId));
  const certificatesDir = await batchDir.getDirectoryHandle("Certificates", { create: true });
  return { batchDir, certificatesDir };
}

export async function listBatchIds(root: FileSystemDirectoryHandle): Promise<string[]> {
  const batchesDir = await getBatchesDirectory(root);
  const ids: string[] = [];
  for await (const [name, handle] of batchesDir.entries()) {
    if (handle.kind === "directory") ids.push(name);
  }
  return ids;
}

export async function writeFile(
  dir: FileSystemDirectoryHandle,
  fileName: string,
  data: Uint8Array | Blob | string
): Promise<void> {
  const fileHandle = await dir.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  // Cast needed: lib.dom's BufferSource union doesn't line up with TS's generic TypedArray types.
  await writable.write(data as FileSystemWriteChunkType);
  await writable.close();
}

export async function readTextFile(dir: FileSystemDirectoryHandle, fileName: string): Promise<string | null> {
  try {
    const fileHandle = await dir.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    return await file.text();
  } catch {
    return null;
  }
}
