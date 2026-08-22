import * as XLSX from "xlsx";
import { Student } from "./types";
import { lookupGrade } from "./gradeTable";

export interface ParseResult {
  students: Student[];
  errors: string[];
}

const HEADER_ALIASES: Record<string, string> = {
  studentname: "studentName",
  "student name": "studentName",
  name: "studentName",
  registernumber: "registerNumber",
  "register number": "registerNumber",
  regno: "registerNumber",
  "register no": "registerNumber",
  "register no.": "registerNumber",
  grade: "grade",
  certificateno: "certificateNo",
  "certificate no": "certificateNo",
  "certificate no.": "certificateNo",
  "certificate number": "certificateNo",
};

function normalizeHeader(header: string): string | null {
  return HEADER_ALIASES[header.trim().toLowerCase()] ?? null;
}

export async function parseStudentsFile(file: File): Promise<ParseResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows: Record<string, unknown>[] = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const errors: string[] = [];
  const students: Student[] = [];
  const seenRegNos = new Set<string>();

  rows.forEach((row, index) => {
    const rowNumber = index + 2; // account for the header row
    const mapped: Record<string, string> = {};
    for (const [header, value] of Object.entries(row)) {
      const field = normalizeHeader(header);
      if (field) mapped[field] = String(value).trim();
    }

    const { studentName, registerNumber, grade, certificateNo } = mapped;

    if (!studentName || !registerNumber || !grade || !certificateNo) {
      errors.push(`Row ${rowNumber}: missing required field(s) (Student Name, Register Number, Grade, Certificate No.).`);
      return;
    }
    if (seenRegNos.has(registerNumber)) {
      errors.push(`Row ${rowNumber}: duplicate register number "${registerNumber}".`);
      return;
    }
    const gradeInfo = lookupGrade(grade);
    if (!gradeInfo) {
      errors.push(`Row ${rowNumber}: unknown grade "${grade}".`);
      return;
    }

    seenRegNos.add(registerNumber);
    students.push({
      studentName,
      registerNumber,
      grade: grade.trim().toUpperCase(),
      certificateNo,
      classification: gradeInfo.classification,
      gradePoint: gradeInfo.gradePoint,
    });
  });

  return { students, errors };
}
