/** Посада в рапортах — перша літера велика, решта як була. */
export const capitalizeReportPosition = (value: string) => {
  const text = String(value ?? "").trim();
  if (!text) return "";
  return text.charAt(0).toLocaleUpperCase("uk-UA") + text.slice(1);
};
