import * as React from "react";
import { useState } from "react";
import { DefaultButton, Dropdown, IDropdownOption, MessageBar, MessageBarType, PrimaryButton, Stack, Text } from "@fluentui/react";
import { useBatch } from "../state/BatchContext";
import { applyValuesToOpenDocument, buildFieldValues, restoreTemplateFields } from "../services/mergeEngine";

export const MergePreview: React.FC = () => {
  const { currentBatch, rescanTemplateFields } = useBatch();
  const [selectedReg, setSelectedReg] = useState<string | undefined>(undefined);
  const [status, setStatus] = useState<{ type: MessageBarType; text: string } | null>(null);

  if (!currentBatch) return null;

  const options: IDropdownOption[] = currentBatch.students.map((s) => ({
    key: s.registerNumber,
    text: `${s.studentName} (${s.registerNumber})`,
  }));

  const handlePreview = async () => {
    const student = currentBatch.students.find((s) => s.registerNumber === selectedReg);
    if (!student) return;
    try {
      const values = buildFieldValues(currentBatch, student);
      await applyValuesToOpenDocument(currentBatch.fields, values);
      setStatus({ type: MessageBarType.success, text: `Previewing certificate for ${student.studentName}.` });
    } catch (err) {
      setStatus({ type: MessageBarType.error, text: String(err) });
    }
  };

  const handleRestore = async () => {
    try {
      await restoreTemplateFields(currentBatch.fields);
      setStatus({ type: MessageBarType.success, text: "Template placeholders restored." });
    } catch (err) {
      setStatus({ type: MessageBarType.error, text: String(err) });
    }
  };

  const handleRescan = async () => {
    try {
      const count = await rescanTemplateFields();
      setStatus({
        type: count > 0 ? MessageBarType.success : MessageBarType.warning,
        text: `Found ${count} placeholder field(s) in the open template.`,
      });
    } catch (err) {
      setStatus({ type: MessageBarType.error, text: String(err) });
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      {currentBatch.fields.length === 0 && (
        <MessageBar messageBarType={MessageBarType.warning}>
          No placeholder tokens were found in the open template when this batch was created. Make sure the batch's
          template is open in PowerPoint (with text like {"{{StudentName}}"} added, even inside grouped shapes), then
          click "Rescan Template" below.
        </MessageBar>
      )}
      {currentBatch.students.length === 0 && <MessageBar>Import students in the "Students" tab first.</MessageBar>}
      <Dropdown
        label="Student"
        options={options}
        selectedKey={selectedReg}
        onChange={(_, opt) => setSelectedReg(opt?.key as string)}
        placeholder="Select a student"
      />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Preview on Open Slide(s)" onClick={handlePreview} disabled={!selectedReg} />
        <DefaultButton text="Restore Template" onClick={handleRestore} />
        <DefaultButton text="Rescan Template" onClick={handleRescan} />
      </Stack>
      {status && <MessageBar messageBarType={status.type}>{status.text}</MessageBar>}
    </Stack>
  );
};
