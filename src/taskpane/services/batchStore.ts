import { Batch } from "./types";
import * as fs from "./fileSystem";

export async function saveBatch(root: FileSystemDirectoryHandle, batch: Batch): Promise<void> {
  const { batchDir } = await fs.openBatchFolder(root, batch.id);
  await fs.writeFile(batchDir, "batch.json", JSON.stringify(batch, null, 2));
}

export async function loadBatch(root: FileSystemDirectoryHandle, batchId: string): Promise<Batch | null> {
  const { batchDir } = await fs.openBatchFolder(root, batchId);
  const text = await fs.readTextFile(batchDir, "batch.json");
  return text ? (JSON.parse(text) as Batch) : null;
}

export async function listBatches(root: FileSystemDirectoryHandle): Promise<Batch[]> {
  const ids = await fs.listBatchIds(root);
  const batches: Batch[] = [];
  for (const id of ids) {
    const batch = await loadBatch(root, id);
    if (batch) batches.push(batch);
  }
  return batches.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export function buildBatchId(trainingLevel: string, startDate: string, endDate: string): string {
  return fs.sanitizeFolderName(`${trainingLevel}_${startDate}_to_${endDate}`);
}
