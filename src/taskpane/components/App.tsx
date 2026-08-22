import * as React from "react";
import { useEffect, useState } from "react";
import { MessageBar, MessageBarType, Pivot, PivotItem, Stack, Text } from "@fluentui/react";
import { BatchProvider, useBatch } from "../state/BatchContext";
import { formatDateOrdinal } from "../services/dateFormat";
import { BatchList } from "./BatchList";
import { NewBatch } from "./NewBatch";
import { ImportStudents } from "./ImportStudents";
import { MergePreview } from "./MergePreview";
import { GenerateExport } from "./GenerateExport";

/* global Office */

const AppShell: React.FC = () => {
  const { currentBatch, error } = useBatch();

  return (
    <Stack tokens={{ childrenGap: 12, padding: 12 }}>
      <Text variant="xLarge">Certificate Generation System</Text>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {currentBatch && (
        <MessageBar messageBarType={MessageBarType.info}>
          Current batch: <b>{currentBatch.trainingLevel}</b> ({formatDateOrdinal(currentBatch.startDate)} to{" "}
          {formatDateOrdinal(currentBatch.endDate)}) — {currentBatch.students.length} student(s)
        </MessageBar>
      )}
      <Pivot>
        <PivotItem headerText="Batches">
          <BatchList />
        </PivotItem>
        <PivotItem headerText="New Batch">
          <NewBatch />
        </PivotItem>
        <PivotItem headerText="Students">
          {currentBatch ? <ImportStudents /> : <MessageBar>Select or create a batch first.</MessageBar>}
        </PivotItem>
        <PivotItem headerText="Merge & Preview">
          {currentBatch ? <MergePreview /> : <MessageBar>Select or create a batch first.</MessageBar>}
        </PivotItem>
        <PivotItem headerText="Generate & Export">
          {currentBatch ? <GenerateExport /> : <MessageBar>Select or create a batch first.</MessageBar>}
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export const App: React.FC = () => {
  const [isReady, setIsReady] = useState(false);
  const [hostOk, setHostOk] = useState(false);

  useEffect(() => {
    Office.onReady((info) => {
      setIsReady(true);
      setHostOk(info.host === Office.HostType.PowerPoint);
    });
  }, []);

  if (!isReady) return <Text>Loading…</Text>;
  if (!hostOk) {
    return <MessageBar messageBarType={MessageBarType.error}>This add-in only runs in PowerPoint.</MessageBar>;
  }

  return (
    <BatchProvider>
      <AppShell />
    </BatchProvider>
  );
};
