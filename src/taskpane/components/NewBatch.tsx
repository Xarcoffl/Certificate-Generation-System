import * as React from "react";
import { useState } from "react";
import { MessageBar, MessageBarType, PrimaryButton, Stack, Text, TextField } from "@fluentui/react";
import { useBatch } from "../state/BatchContext";

export const NewBatch: React.FC = () => {
  const { rootDir, createBatch, loading } = useBatch();
  const [trainingLevel, setTrainingLevel] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [status, setStatus] = useState<{ type: MessageBarType; text: string } | null>(null);

  const canSubmit = !!rootDir && !!trainingLevel && !!startDate && !!endDate;

  const handleCreate = async () => {
    setStatus(null);
    try {
      await createBatch(trainingLevel, startDate, endDate);
      setStatus({ type: MessageBarType.success, text: "Batch created. A copy of the open template was saved into its folder." });
      setTrainingLevel("");
      setStartDate("");
      setEndDate("");
    } catch (err) {
      setStatus({ type: MessageBarType.error, text: `Failed: ${String(err)}` });
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Text>
        Open the certificate template in PowerPoint, then fill in the batch details below. A copy of the currently open
        template is saved into the new batch's folder.
      </Text>
      {!rootDir && (
        <MessageBar messageBarType={MessageBarType.warning}>Choose a root folder in the "Batches" tab first.</MessageBar>
      )}
      <TextField
        label="Training Level"
        value={trainingLevel}
        onChange={(_, v) => setTrainingLevel(v ?? "")}
        placeholder="Level 1 - Unity for Virtual Reality"
      />
      <TextField label="Start Date" type="date" value={startDate} onChange={(_, v) => setStartDate(v ?? "")} />
      <TextField label="End Date" type="date" value={endDate} onChange={(_, v) => setEndDate(v ?? "")} />
      <PrimaryButton text="Create Batch" onClick={handleCreate} disabled={!canSubmit || loading} />
      {status && <MessageBar messageBarType={status.type}>{status.text}</MessageBar>}
    </Stack>
  );
};
