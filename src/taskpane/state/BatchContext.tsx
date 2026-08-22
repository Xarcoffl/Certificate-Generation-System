import * as React from "react";
import { createContext, useCallback, useContext, useEffect, useMemo, useState } from "react";
import { Batch, Student } from "../services/types";
import * as fs from "../services/fileSystem";
import * as batchStore from "../services/batchStore";
import { captureTemplateFields } from "../services/mergeEngine";
import { getActiveDocumentAsPptx } from "../services/documentFile";

interface BatchContextValue {
  rootDir: FileSystemDirectoryHandle | null;
  needsReconnect: boolean;
  batches: Batch[];
  currentBatch: Batch | null;
  loading: boolean;
  error: string | null;
  chooseRootFolder: () => Promise<void>;
  reconnectRootFolder: () => Promise<void>;
  refreshBatches: () => Promise<void>;
  createBatch: (trainingLevel: string, startDate: string, endDate: string) => Promise<void>;
  selectBatch: (batchId: string) => Promise<void>;
  updateStudents: (students: Student[]) => Promise<void>;
  rescanTemplateFields: () => Promise<number>;
}

const BatchContext = createContext<BatchContextValue | undefined>(undefined);

export const BatchProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [rootDir, setRootDir] = useState<FileSystemDirectoryHandle | null>(null);
  const [pendingHandle, setPendingHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [batches, setBatches] = useState<Batch[]>([]);
  const [currentBatch, setCurrentBatch] = useState<Batch | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refreshBatches = useCallback(async (dirOverride?: FileSystemDirectoryHandle | null) => {
    const dir = dirOverride ?? rootDir;
    if (!dir) return;
    setLoading(true);
    try {
      setBatches(await batchStore.listBatches(dir));
    } catch (err) {
      setError(String(err));
    } finally {
      setLoading(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rootDir]);

  useEffect(() => {
    (async () => {
      const handle = await fs.getPersistedRootHandle();
      if (!handle) return;

      const granted = (await fs.queryPermission(handle)) === "granted";
      if (!granted) {
        // Can't silently request permission outside a user gesture - surface a reconnect prompt instead.
        setPendingHandle(handle);
        return;
      }

      setRootDir(handle);
      await refreshBatches(handle);

      const lastBatchId = await fs.getLastBatchId();
      if (lastBatchId) {
        const batch = await batchStore.loadBatch(handle, lastBatchId).catch(() => null);
        if (batch) setCurrentBatch(batch);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const chooseRootFolder = useCallback(async () => {
    setError(null);
    try {
      const dir = await fs.pickRootDirectory();
      setRootDir(dir);
      await refreshBatches(dir);
    } catch (err) {
      setError(String(err));
    }
  }, [refreshBatches]);

  const createBatch = useCallback(
    async (trainingLevel: string, startDate: string, endDate: string) => {
      if (!rootDir) throw new Error("Choose a root folder first.");
      setLoading(true);
      setError(null);
      try {
        const id = batchStore.buildBatchId(trainingLevel, startDate, endDate);
        const { batchDir } = await fs.createBatchFolder(rootDir, id);

        const templateFileName = "Certificate_Template.pptx";
        const templateBytes = await getActiveDocumentAsPptx();
        await fs.writeFile(batchDir, templateFileName, templateBytes);

        const fields = await captureTemplateFields();

        const batch: Batch = {
          id,
          trainingLevel,
          startDate,
          endDate,
          templateFileName,
          fields,
          students: [],
          createdAt: new Date().toISOString(),
        };
        await batchStore.saveBatch(rootDir, batch);
        setCurrentBatch(batch);
        await fs.setLastBatchId(batch.id);
        await refreshBatches(rootDir);
      } catch (err) {
        setError(String(err));
        throw err;
      } finally {
        setLoading(false);
      }
    },
    [rootDir, refreshBatches]
  );

  const selectBatch = useCallback(
    async (batchId: string) => {
      if (!rootDir) return;
      setLoading(true);
      setError(null);
      try {
        const batch = await batchStore.loadBatch(rootDir, batchId);
        setCurrentBatch(batch);
        if (batch) await fs.setLastBatchId(batch.id);
      } catch (err) {
        setError(String(err));
      } finally {
        setLoading(false);
      }
    },
    [rootDir]
  );

  const updateStudents = useCallback(
    async (students: Student[]) => {
      if (!rootDir || !currentBatch) throw new Error("No batch selected.");
      const updated: Batch = { ...currentBatch, students };
      await batchStore.saveBatch(rootDir, updated);
      setCurrentBatch(updated);
      await refreshBatches(rootDir);
    },
    [rootDir, currentBatch, refreshBatches]
  );

  /** Re-scans the currently open document for {{tokens}} and updates the current batch's field map. */
  const rescanTemplateFields = useCallback(async (): Promise<number> => {
    if (!rootDir || !currentBatch) throw new Error("No batch selected.");
    const fields = await captureTemplateFields();
    const updated: Batch = { ...currentBatch, fields };
    await batchStore.saveBatch(rootDir, updated);
    setCurrentBatch(updated);
    return fields.length;
  }, [rootDir, currentBatch]);

  /** Re-requests permission for the previously chosen folder. Must run inside a click handler (user gesture required). */
  const reconnectRootFolder = useCallback(async () => {
    if (!pendingHandle) return;
    setError(null);
    try {
      const granted = await fs.requestPermission(pendingHandle);
      if (!granted) {
        setError("Folder access was not granted.");
        return;
      }
      setRootDir(pendingHandle);
      setPendingHandle(null);
      await refreshBatches(pendingHandle);

      const lastBatchId = await fs.getLastBatchId();
      if (lastBatchId) {
        const batch = await batchStore.loadBatch(pendingHandle, lastBatchId).catch(() => null);
        if (batch) setCurrentBatch(batch);
      }
    } catch (err) {
      setError(String(err));
    }
  }, [pendingHandle, refreshBatches]);

  const value = useMemo<BatchContextValue>(
    () => ({
      rootDir,
      needsReconnect: pendingHandle !== null,
      batches,
      currentBatch,
      loading,
      error,
      chooseRootFolder,
      reconnectRootFolder,
      refreshBatches: () => refreshBatches(),
      createBatch,
      selectBatch,
      updateStudents,
      rescanTemplateFields,
    }),
    [
      rootDir,
      pendingHandle,
      batches,
      currentBatch,
      loading,
      error,
      chooseRootFolder,
      reconnectRootFolder,
      refreshBatches,
      createBatch,
      selectBatch,
      updateStudents,
      rescanTemplateFields,
    ]
  );

  return <BatchContext.Provider value={value}>{children}</BatchContext.Provider>;
};

export function useBatch(): BatchContextValue {
  const ctx = useContext(BatchContext);
  if (!ctx) throw new Error("useBatch must be used within a BatchProvider");
  return ctx;
}
