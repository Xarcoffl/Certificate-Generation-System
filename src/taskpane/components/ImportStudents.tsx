import * as React from "react";
import { useEffect, useState } from "react";
import { DetailsList, DetailsListLayoutMode, MessageBar, MessageBarType, PrimaryButton, Stack, Text } from "@fluentui/react";
import { useBatch } from "../state/BatchContext";
import { parseStudentsFile } from "../services/xlsxParser";
import { Student } from "../services/types";

export const ImportStudents: React.FC = () => {
  const { currentBatch, updateStudents } = useBatch();
  const [pending, setPending] = useState<Student[]>(currentBatch?.students ?? []);
  const [errors, setErrors] = useState<string[]>([]);
  const [status, setStatus] = useState<string | null>(null);

  useEffect(() => {
    setPending(currentBatch?.students ?? []);
    setErrors([]);
    setStatus(null);
  }, [currentBatch?.id]);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setStatus(null);
    const result = await parseStudentsFile(file);
    setPending(result.students);
    setErrors(result.errors);
    e.target.value = "";
  };

  const handleSave = async () => {
    await updateStudents(pending);
    setStatus(`Saved ${pending.length} student(s) to batch "${currentBatch?.id}".`);
  };

  const columns = [
    { key: "name", name: "Student Name", fieldName: "studentName", minWidth: 120 },
    { key: "reg", name: "Register No.", fieldName: "registerNumber", minWidth: 100 },
    { key: "grade", name: "Grade", fieldName: "grade", minWidth: 50 },
    { key: "class", name: "Classification", fieldName: "classification", minWidth: 110 },
    { key: "cert", name: "Certificate No.", fieldName: "certificateNo", minWidth: 100 },
  ];

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Text>
        Import an Excel/CSV file with columns: Student Name, Register Number, Grade, Certificate No. Classification is
        derived automatically from Grade.
      </Text>
      <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} />
      {errors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning}>
          <Stack>
            {errors.map((err, i) => (
              <Text key={i}>{err}</Text>
            ))}
          </Stack>
        </MessageBar>
      )}
      {pending.length > 0 && (
        <>
          <DetailsList items={pending} columns={columns} layoutMode={DetailsListLayoutMode.justified} />
          <PrimaryButton text={`Save ${pending.length} Student(s) to Batch`} onClick={handleSave} />
        </>
      )}
      {status && <MessageBar messageBarType={MessageBarType.success}>{status}</MessageBar>}
    </Stack>
  );
};
