export interface Student {
  studentName: string;
  registerNumber: string;
  grade: string;
  certificateNo: string;
  classification: string;
  gradePoint: number;
}

/** A shape/slide pair whose text contains {{tokens}}, captured from the open template. shapePath drills through nested groups. */
export interface TemplateField {
  slideId: string;
  shapePath: string[];
  template: string;
}

export interface Batch {
  id: string;
  trainingLevel: string;
  startDate: string;
  endDate: string;
  templateFileName: string;
  fields: TemplateField[];
  students: Student[];
  createdAt: string;
}

export interface GenerateResult {
  registerNumber: string;
  studentName: string;
  success: boolean;
  error?: string;
}
