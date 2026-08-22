/* global PowerPoint */
import { Batch, Student, TemplateField } from "./types";
import { formatDateOrdinal } from "./dateFormat";

/** Recursively scans a shape collection, drilling into groups, collecting shapes whose text contains {{tokens}}. */
async function scanShapes(
  context: PowerPoint.RequestContext,
  slideId: string,
  shapes: PowerPoint.ShapeCollection | PowerPoint.ShapeScopedCollection,
  path: string[],
  fields: TemplateField[]
): Promise<void> {
  shapes.load("items/id,items/type");
  await context.sync();

  for (const shape of shapes.items) {
    const currentPath = [...path, shape.id];

    if (shape.type === PowerPoint.ShapeType.group) {
      await scanShapes(context, slideId, shape.group.shapes, currentPath, fields);
      continue;
    }

    try {
      shape.textFrame.load("hasText");
      await context.sync();
      if (!shape.textFrame.hasText) continue;

      shape.textFrame.textRange.load("text");
      await context.sync();
      const text = shape.textFrame.textRange.text;
      if (text && text.includes("{{")) {
        fields.push({ slideId, shapePath: currentPath, template: text });
      }
    } catch {
      // Shape type doesn't support a text frame (e.g. picture) - skip it.
    }
  }
}

/** Scans every slide/shape (including nested groups) in the open presentation for {{tokens}}. */
export async function captureTemplateFields(): Promise<TemplateField[]> {
  const fields: TemplateField[] = [];
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    for (const slide of slides.items) {
      await scanShapes(context, slide.id, slide.shapes, [], fields);
    }
  });
  return fields;
}

/** Resolves a TemplateField's shape, drilling through nested groups via its shapePath. */
function resolveShape(context: PowerPoint.RequestContext, field: TemplateField): PowerPoint.Shape {
  const slide = context.presentation.slides.getItem(field.slideId);
  let shape = slide.shapes.getItem(field.shapePath[0]);
  for (let i = 1; i < field.shapePath.length; i++) {
    shape = shape.group.shapes.getItem(field.shapePath[i]);
  }
  return shape;
}

export function buildFieldValues(
  batch: Pick<Batch, "trainingLevel" | "startDate" | "endDate">,
  student: Student
): Record<string, string> {
  return {
    StudentName: student.studentName,
    RegisterNumber: student.registerNumber,
    TrainingLevel: batch.trainingLevel,
    StartDate: formatDateOrdinal(batch.startDate),
    EndDate: formatDateOrdinal(batch.endDate),
    Grade: student.grade,
    Classification: student.classification,
    CertificateNo: student.certificateNo,
  };
}

function renderTemplate(template: string, values: Record<string, string>): string {
  return template.replace(/{{\s*(\w+)\s*}}/g, (match, key) => (key in values ? values[key] : match));
}

/** Writes merged text for one student into the corresponding shapes of the open presentation. */
export async function applyValuesToOpenDocument(fields: TemplateField[], values: Record<string, string>): Promise<void> {
  await PowerPoint.run(async (context) => {
    for (const field of fields) {
      const shape = resolveShape(context, field);
      shape.textFrame.textRange.text = renderTemplate(field.template, values);
    }
    await context.sync();
  });
}

/** Restores the original {{token}} text, undoing whatever student was last merged in. */
export async function restoreTemplateFields(fields: TemplateField[]): Promise<void> {
  await PowerPoint.run(async (context) => {
    for (const field of fields) {
      const shape = resolveShape(context, field);
      shape.textFrame.textRange.text = field.template;
    }
    await context.sync();
  });
}
