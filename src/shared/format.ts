export const formatFileSize = (size: number) => `${(size / 1024).toFixed(1)} KB`;

export const formatDateTime = () =>
  new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date());

export const normalizeDatasetKey = (value: string) =>
  value
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\p{L}\p{N}_-]/gu, "")
    .toLowerCase();

export const cellValueToJson = (value: unknown) => {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
};
