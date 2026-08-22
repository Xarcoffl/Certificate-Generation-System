/* global Office */

function concatSlices(slices: Uint8Array[]): Uint8Array {
  const total = slices.reduce((sum, slice) => sum + slice.length, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  for (const slice of slices) {
    merged.set(slice, offset);
    offset += slice.length;
  }
  return merged;
}

function getFileSlices(fileType: Office.FileType): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }
      const file = result.value;
      const sliceCount = file.sliceCount;
      if (sliceCount === 0) {
        file.closeAsync(() => resolve(new Uint8Array(0)));
        return;
      }
      const slices: Uint8Array[] = new Array(sliceCount);
      let received = 0;

      const requestSlice = (index: number) => {
        file.getSliceAsync(index, (sliceResult) => {
          if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
            file.closeAsync(() => {});
            reject(sliceResult.error);
            return;
          }
          slices[index] = new Uint8Array(sliceResult.value.data);
          received++;
          if (received === sliceCount) {
            file.closeAsync(() => resolve(concatSlices(slices)));
          } else if (index + 1 < sliceCount) {
            requestSlice(index + 1);
          }
        });
      };
      requestSlice(0);
    });
  });
}

/** Bytes of the currently open presentation in OOXML (.pptx) format. */
export async function getActiveDocumentAsPptx(): Promise<Uint8Array> {
  return getFileSlices(Office.FileType.Compressed);
}

/** Bytes of the currently open presentation rendered to PDF, reflecting whatever text is on the slides right now. */
export async function getActiveDocumentAsPdf(): Promise<Uint8Array> {
  return getFileSlices(Office.FileType.Pdf);
}
