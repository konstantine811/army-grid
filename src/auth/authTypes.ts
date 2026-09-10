const TOKEN_KEY = "army-grid.auth-token";
const USER_KEY = "army-grid.auth-user";
const LAST_EMAIL_KEY = "army-grid.auth-last-email";

export type AppUserRole = "ADMIN" | "USER";

export const WRITE_PERMISSIONS = [
  "personnel",
  "bchs",
  "anketaData",
  "documents",
] as const;

export type WritePermission = (typeof WRITE_PERMISSIONS)[number];

export const WRITE_PERMISSION_LABELS: Record<WritePermission, string> = {
  personnel: "Особовий склад",
  bchs: "БЧС",
  anketaData: "Анкетні дані",
  documents: "Документи",
};

export type AuthUser = {
  id: string;
  email: string;
  role: AppUserRole;
  accessGranted: boolean;
  writePermissions: WritePermission[];
  displayName: string | null;
  nickname: string | null;
  photoData: string | null;
  linkedPersonExternalId: string | null;
  linkedPersonFullName: string | null;
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

export const getCachedAuthUser = (): AuthUser | null => {
  try {
    const raw = window.localStorage.getItem(USER_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as AuthUser;
  } catch {
    return null;
  }
};

/** Keep sidebar avatar when session refresh omits photoData. */
export const mergeAuthUserWithCachedPhoto = (user: AuthUser): AuthUser => {
  if (user.photoData) return user;
  const cached = getCachedAuthUser();
  if (cached?.id === user.id && cached.photoData) {
    return { ...user, photoData: cached.photoData };
  }
  return user;
};

export const setCachedAuthUser = (user: AuthUser | null) => {
  try {
    if (!user) window.localStorage.removeItem(USER_KEY);
    else {
      window.localStorage.setItem(
        USER_KEY,
        JSON.stringify(mergeAuthUserWithCachedPhoto(user)),
      );
    }
  } catch {
    // Ignore storage errors.
  }
};

export const getLastAuthEmail = () => {
  try {
    return window.localStorage.getItem(LAST_EMAIL_KEY) ?? "";
  } catch {
    return "";
  }
};

export const setLastAuthEmail = (email: string) => {
  try {
    const trimmed = email.trim();
    if (!trimmed) window.localStorage.removeItem(LAST_EMAIL_KEY);
    else window.localStorage.setItem(LAST_EMAIL_KEY, trimmed);
  } catch {
    // Ignore storage errors.
  }
};

export const AUTH_LOGOUT_EVENT = "army-grid:auth-logout";

export const emitAuthLogout = () => {
  window.dispatchEvent(new Event(AUTH_LOGOUT_EVENT));
};

export const isAdminUser = (user: AuthUser | null | undefined) =>
  user?.role === "ADMIN";

export const canViewApp = (user: AuthUser | null | undefined) =>
  Boolean(user && (user.role === "ADMIN" || user.accessGranted));

export const normalizeWritePermissions = (
  value: unknown,
): WritePermission[] => {
  if (!Array.isArray(value)) return [];
  const allowed = new Set<string>(WRITE_PERMISSIONS);
  const unique = new Set<WritePermission>();
  for (const item of value) {
    if (typeof item === "string" && allowed.has(item)) {
      unique.add(item as WritePermission);
    }
  }
  return [...unique];
};

export const canEditArea = (
  user: AuthUser | null | undefined,
  area: WritePermission,
) => {
  if (!user) return false;
  if (user.role === "ADMIN") return true;
  return normalizeWritePermissions(user.writePermissions).includes(area);
};

/** Admin or any granted write area. */
export const canEditApp = (user: AuthUser | null | undefined) => {
  if (!user) return false;
  if (user.role === "ADMIN") return true;
  return normalizeWritePermissions(user.writePermissions).length > 0;
};
