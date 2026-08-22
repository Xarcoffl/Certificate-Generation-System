import * as React from "react";
import { useState } from "react";
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Stack,
  Text,
} from "@fluentui/react";
import { useBatch } from "../state/BatchContext";
import * as fs from "../services/fileSystem";
import { applyValuesToOpenDocument, buildFieldValues, restoreTemplateFields } from "../services/mergeEngine";
import { getActiveDocumentAsPdf } from "../services/documentFile";
import { GenerateResult, Student } from "../services/types";

function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "-").trim();
}

/** Merges new outcomes into the prior result list by register number, preserving each student's position. */
function mergeOutcomes(prev: GenerateResult[], updated: GenerateResult[]): GenerateResult[] {
  const byRegNo = new Map(prev.map((r) => [r.registerNumber, r]));
  for (const outcome of updated) {
    byRegNo.set(outcome.registerNumber, outcome);
  }
  return Array.from(byRegNo.values());
}

export const GenerateExport: React.FC = () => {
  const { rootDir, currentBatch } = useBatch();
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [results, setResults] = useState<GenerateResult[]>([]);

  if (!currentBatch) return null;

  const runGeneration = async (students: Student[]) => {
    if (!rootDir) return;
    setRunning(true);
    setProgress(0);

    const { certificatesDir } = await fs.openBatchFolder(rootDir, currentBatch.id);
    const outcomes: GenerateResult[] = [];

    for (let i = 0; i < students.length; i++) {
      const student = students[i];
      try {
        const values = buildFieldValues(currentBatch, student);
        await applyValuesToOpenDocument(currentBatch.fields, values);
        const pdfBytes = await getActiveDocumentAsPdf();
        const fileName = `${sanitizeFileName(student.studentName)}_${sanitizeFileName(student.registerNumber)}.pdf`;
        await fs.writeFile(certificatesDir, fileName, pdfBytes);
        outcomes.push({ registerNumber: student.registerNumber, studentName: student.studentName, success: true });
      } catch (err) {
        outcomes.push({
          registerNumber: student.registerNumber,
          studentName: student.studentName,
          success: false,
          error: String(err),
        });
      }
      setProgress((i + 1) / students.length);
      setResults((prev) => mergeOutcomes(prev, outcomes));
    }

    try {
      await restoreTemplateFields(currentBatch.fields);
    } catch {
      // Non-fatal: the last student's data may remain visible on the open slide(s).
    }

    setRunning(false);
  };

  const handleGenerateAll = () => {
    setResults([]);
    void runGeneration(currentBatch.students);
  };

  const handleRetryFailed = () => {
    const failedRegNos = new Set(results.filter((r) => !r.success).map((r) => r.registerNumber));
    const failedStudents = currentBatch.students.filter((s) => failedRegNos.has(s.registerNumber));
    void runGeneration(failedStudents);
  };

  const failureCount = results.filter((r) => !r.success).length;

  const columns = [
    { key: "name", name: "Student", fieldName: "studentName", minWidth: 120 },
    { key: "reg", name: "Register No.", fieldName: "registerNumber", minWidth: 100 },
    {
      key: "status",
      name: "Status",
      minWidth: 140,
      onRender: (item: GenerateResult) => (item.success ? "Saved" : `Failed: ${item.error}`),
    },
  ];

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Text>
        Generates one PDF per student into Certificates/ inside this batch's folder, named "StudentName_RegisterNo.pdf".
      </Text>
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text={`Generate All Certificates (${currentBatch.students.length})`}
          onClick={handleGenerateAll}
          disabled={running || currentBatch.students.length === 0}
        />
        <DefaultButton text={`Retry Failed (${failureCount})`} onClick={handleRetryFailed} disabled={running || failureCount === 0} />
      </Stack>
      {running && <ProgressIndicator percentComplete={progress} label="Generating certificates…" />}
      {results.length > 0 && (
        <>
          <MessageBar messageBarType={failureCount > 0 ? MessageBarType.warning : MessageBarType.success}>
            {results.length - failureCount} succeeded, {failureCount} failed.
          </MessageBar>
          <DetailsList items={results} columns={columns} layoutMode={DetailsListLayoutMode.justified} />
        </>
      )}
    </Stack>
  );
};
