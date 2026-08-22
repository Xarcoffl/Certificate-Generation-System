export interface GradeInfo {
  gradePoint: number;
  classification: string;
}

export const GRADE_ORDER = ["O", "A+", "A", "B+", "B", "C"] as const;

export const GRADE_TABLE: Record<string, GradeInfo> = {
  O: { gradePoint: 10, classification: "Outstanding" },
  "A+": { gradePoint: 9, classification: "Excellent" },
  A: { gradePoint: 8, classification: "Very Good" },
  "B+": { gradePoint: 7, classification: "Good" },
  B: { gradePoint: 6, classification: "Above Average" },
  C: { gradePoint: 5, classification: "Satisfactory" },
};

export function lookupGrade(grade: string): GradeInfo | undefined {
  return GRADE_TABLE[grade.trim().toUpperCase()];
}
