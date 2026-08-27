export type AppPage =
  | "overview"
  | "import"
  | "analytics"
  | "ejournal"
  | "bchs"
  | "bchsLab"
  | "excelFill"
  | "questionnaireParser"
  | "anketaData"
  | "socPassport"
  | "personnel"
  | "documents"
  | "documentSettings"
  | "usersAccess"
  | "profile";
export type BchsAnalyticsTab = "overview" | "comparison" | "combat" | "supplement";

/** Pages available to non-admin users (USER). Admins see all pages. */
export const USER_ALLOWED_PAGES: readonly AppPage[] = [
  "overview",
  "personnel",
  "bchs",
  "anketaData",
  "documents",
  "profile",
] as const;

export const isUserAllowedPage = (page: AppPage) =>
  (USER_ALLOWED_PAGES as readonly string[]).includes(page);

/** Write permission required to mutate data on a page (null = no page-level writes). */
export const writeAreaForPage = (
  page: AppPage,
): "personnel" | "bchs" | "anketaData" | "documents" | null => {
  switch (page) {
    case "personnel":
      return "personnel";
    case "bchs":
      return "bchs";
    case "anketaData":
      return "anketaData";
    case "documents":
      return "documents";
    default:
      return null;
  }
};

export const pagePaths: Record<AppPage, string> = {
  overview: "/overview",
  import: "/import",
  analytics: "/analytics",
  ejournal: "/ejournal",
  bchs: "/bchs",
  bchsLab: "/bchs-lab",
  excelFill: "/excel-fill",
  questionnaireParser: "/questionnaire-parser",
  anketaData: "/anketa-data",
  socPassport: "/soc-passport",
  personnel: "/personnel",
  documents: "/documents",
  documentSettings: "/document-settings",
  usersAccess: "/users-access",
  profile: "/profile",
};

const pathPages = Object.fromEntries(
  Object.entries(pagePaths).map(([page, path]) => [path, page]),
) as Record<string, AppPage>;

const getPathname = (path: string) => {
  try {
    return new URL(path, window.location.origin).pathname;
  } catch {
    return path.split("?")[0] || path;
  }
};

export const getPageFromPath = (path: string): AppPage => {
  const pathname = getPathname(path);

  return pathname.startsWith(`${pagePaths.documents}/`)
    ? "documents"
    : pathname.startsWith(`${pagePaths.personnel}/`)
      ? "personnel"
      : pathPages[pathname] ?? (pathname === "/" ? "overview" : "overview");
};

export const getInitialBchsAnalyticsTab = (): BchsAnalyticsTab => {
  const tab = new URLSearchParams(window.location.search).get("bchsTab");

  return tab === "comparison" ||
    tab === "combat" ||
    tab === "overview" ||
    tab === "supplement"
    ? tab
    : "overview";
};

export type EjoosWorkspaceTab =
  | "import"
  | "shpo"
  | "oos"
  | "excluded"
  | "tempArrivals"
  | "tempAbsents"
  | "timesheet"
  | "irrevocableLosses";

const EJOOS_WORKSPACE_TABS: readonly EjoosWorkspaceTab[] = [
  "import",
  "shpo",
  "oos",
  "excluded",
  "tempArrivals",
  "tempAbsents",
  "timesheet",
  "irrevocableLosses",
] as const;

export const isEjoosWorkspaceTab = (
  value: string | null | undefined,
): value is EjoosWorkspaceTab =>
  !!value &&
  (EJOOS_WORKSPACE_TABS as readonly string[]).includes(value);

export const getInitialEjoosTab = (): EjoosWorkspaceTab => {
  const tab = new URLSearchParams(window.location.search).get("ejoosTab");
  return isEjoosWorkspaceTab(tab) ? tab : "import";
};

/** Persist ЕЖООС sub-tab in the URL so reload keeps the same screen. */
export const replaceEjoosTabInUrl = (tab: EjoosWorkspaceTab) => {
  if (typeof window === "undefined") return;
  if (getPageFromPath(window.location.pathname) !== "ejournal") return;

  const url = new URL(window.location.href);
  if (tab === "import") {
    url.searchParams.delete("ejoosTab");
  } else {
    url.searchParams.set("ejoosTab", tab);
  }

  const next = `${url.pathname}${url.search}${url.hash}`;
  if (next === `${window.location.pathname}${window.location.search}${window.location.hash}`) {
    return;
  }

  window.history.replaceState(
    { ...(window.history.state as object | null), page: "ejournal" },
    "",
    next,
  );
};

export const getCurrentRouteKey = () =>
  `${window.location.pathname}${window.location.search}`;

export const pushAppRoute = (path: string, page = getPageFromPath(path)) => {
  window.history.pushState({ page }, "", path);
  return {
    page,
    routeKey: getCurrentRouteKey(),
  };
};

export const navigateToPage = (page: AppPage) => {
  return pushAppRoute(pagePaths[page], page);
};

export const buildDocumentRoute = ({
  personExternalId,
  rowId,
  documentId,
  type,
}: {
  personExternalId?: string;
  rowId?: string;
  documentId?: string;
  type?:
    | "salaryPowerAttorney"
    | "ubdReport"
    | "ubdRestoreReport"
    | "form6Report"
    | "form12Report"
    | "temporaryMilitaryId"
    | "default"
    | string;
}) => {
  const params = new URLSearchParams();
  if (rowId) params.set("rowId", rowId);
  if (documentId) params.set("documentId", documentId);
  if (type === "salaryPowerAttorney") params.set("type", "salary-power-attorney");
  if (type === "ubdReport") params.set("type", "ubd-report");
  if (type === "ubdRestoreReport") params.set("type", "ubd-restore-report");
  if (type === "form6Report") params.set("type", "form6-report");
  if (type === "form12Report") params.set("type", "form12-report");
  if (type === "serviceCharacteristic")
    params.set("type", "service-characteristic");
  if (type === "zhbdCertificate") params.set("type", "zhbd-certificate");
  if (type === "temporaryMilitaryId")
    params.set("type", "temporary-military-id");

  const path = `${pagePaths.documents}${
    personExternalId ? `/${encodeURIComponent(personExternalId)}` : ""
  }`;
  const query = params.toString();
  return query ? `${path}?${query}` : path;
};

export const buildPersonnelRoute = ({
  rowId,
  externalId,
}: {
  rowId?: string;
  externalId?: string;
}) => {
  const params = new URLSearchParams();
  if (rowId) params.set("rowId", rowId);
  if (externalId) params.set("externalId", externalId);

  const query = params.toString();
  return query ? `${pagePaths.personnel}?${query}` : pagePaths.personnel;
};

export const openPersonnelInNewTab = ({
  rowId,
  externalId,
}: {
  rowId?: string;
  externalId?: string;
}) => {
  const url = buildPersonnelRoute({ rowId, externalId });
  window.open(url, "_blank", "noopener,noreferrer");
};
