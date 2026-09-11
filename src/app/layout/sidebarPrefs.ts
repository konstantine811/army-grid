export const SIDEBAR_COLLAPSED_STORAGE_KEY = "army-grid.sidebar-collapsed";

export const readSidebarCollapsed = (): boolean => {
  try {
    return window.localStorage.getItem(SIDEBAR_COLLAPSED_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
};

export const writeSidebarCollapsed = (collapsed: boolean) => {
  try {
    window.localStorage.setItem(
      SIDEBAR_COLLAPSED_STORAGE_KEY,
      collapsed ? "1" : "0",
    );
  } catch {
    // Ignore storage write failures (private mode, etc.).
  }
};
