/** Requirement-set checks and error helpers for the Excel layer. */

/** Minimum Excel API set declared in the manifests. */
export const MIN_EXCEL_API = "1.7";

export function excelApi(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", version);
  } catch {
    return false;
  }
}

/** An error whose message is meant to be shown to the user as-is. */
export class UserError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UserError";
  }
}

/** Human-readable text for any error thrown by Office.js or our code. */
export function describeError(error: unknown): string {
  if (error instanceof UserError) return error.message;
  if (typeof OfficeExtension !== "undefined" && error instanceof OfficeExtension.Error) {
    const where = error.debugInfo?.errorLocation ? ` (at ${error.debugInfo.errorLocation})` : "";
    return `Excel reported: ${error.message}${where}`;
  }
  if (error instanceof Error) return error.message;
  return String(error);
}
