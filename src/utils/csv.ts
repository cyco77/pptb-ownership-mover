const CSV_CONTENT_TYPE = "text/csv;charset=utf-8;";
const CSV_BOM = "\uFEFF";

const escapeCsvValue = (value: unknown): string =>
  `"${String(value ?? "").replace(/"/g, '""')}"`;

export const toCsvLine = (values: unknown[]): string =>
  values.map(escapeCsvValue).join(",");

export const downloadCsv = (lines: string[], fileName: string): void => {
  if (lines.length === 0) {
    return;
  }

  const blob = new Blob([CSV_BOM + lines.join("\r\n")], {
    type: CSV_CONTENT_TYPE,
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

export const createCsvTimestamp = (): string =>
  new Date().toISOString().replace(/[:.]/g, "-");
