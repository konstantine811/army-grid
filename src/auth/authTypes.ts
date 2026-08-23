const TOKEN_KEY = "army-grid.auth-token";

export type AppUserRole = "ADMIN" | "USER";

export type AuthUser = {
  id: string;
  email: string;
  role: AppUserRole;
  accessGranted: boolean;
  displayName: string | null;
};

export type AuthSession = {
  accessToken: string;
  user: AuthUser;
};

export type RegisteredUser = AuthUser & {
  createdAt: string;
  updatedAt: string;
};

export const getAuthToken = () => {
  try {
    return window.localStorage.getItem(TOKEN_KEY) ?? "";
  } catch {
    return "";
  }
};

export const setAuthToken = (token: string | null) => {
  try {
    if (!token) window.localStorage.removeItem(TOKEN_KEY);
    else window.localStorage.setItem(TOKEN_KEY, token);
  } catch {
    // Ignore storage errors.
  }
};

export const clearAuthToken = () => setAuthToken(null);

export const AUTH_LOGOUT_EVENT = "army-grid:auth-logout";

export const emitAuthLogout = () => {
  window.dispatchEvent(new Event(AUTH_LOGOUT_EVENT));
};

export const isAdminUser = (user: AuthUser | null | undefined) =>
  user?.role === "ADMIN";

export const canViewApp = (user: AuthUser | null | undefined) =>
  Boolean(user && (user.role === "ADMIN" || user.accessGranted));

export const canEditApp = (user: AuthUser | null | undefined) =>
  user?.role === "ADMIN";
