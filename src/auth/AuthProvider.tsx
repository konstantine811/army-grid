import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import { api, isApiHttpError } from "../api";
import {
  AUTH_LOGOUT_EVENT,
  canEditApp,
  canEditArea as canEditAreaForUser,
  canViewApp,
  clearAuthToken,
  getAuthToken,
  getCachedAuthUser,
  isAdminUser,
  mergeAuthUserWithCachedPhoto,
  setAuthToken,
  setCachedAuthUser,
  setLastAuthEmail,
  type AuthUser,
  type WritePermission,
} from "./authTypes";
import { bootstrapPersonnelAppShell } from "../data/personnelAppWarm";

type AuthContextValue = {
  user: AuthUser | null;
  loading: boolean;
  canView: boolean;
  canEdit: boolean;
  isAdmin: boolean;
  canEditArea: (area: WritePermission) => boolean;
  login: (email: string, password: string) => Promise<AuthUser>;
  register: (
    email: string,
    password: string,
    displayName: string,
  ) => Promise<AuthUser>;
  logout: () => void;
  refreshUser: () => Promise<AuthUser | null>;
};

const AuthContext = createContext<AuthContextValue | null>(null);

const scheduleAuthPhotoRefresh = (task: () => void) => {
  if (typeof requestIdleCallback === "function") {
    requestIdleCallback(task, { timeout: 12_000 });
    return;
  }
  window.setTimeout(task, 3_000);
};

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(null);
  const [loading, setLoading] = useState(true);

  const logout = useCallback(() => {
    clearAuthToken();
    setCachedAuthUser(null);
    setUser(null);
  }, []);

  const applySession = useCallback((accessToken: string, nextUser: AuthUser) => {
    const merged = mergeAuthUserWithCachedPhoto(nextUser);
    setAuthToken(accessToken);
    setCachedAuthUser(merged);
    setLastAuthEmail(merged.email);
    setUser(merged);
    bootstrapPersonnelAppShell();
  }, []);

  const handleAuthFailure = useCallback(
    (error: unknown) => {
      if (isApiHttpError(error) && error.status === 401) {
        clearAuthToken();
        setCachedAuthUser(null);
        setUser(null);
        return;
      }
      const cached = getCachedAuthUser();
      if (cached) setUser(cached);
    },
    [],
  );

  const refreshAuthPhoto = useCallback(async () => {
    const token = getAuthToken();
    if (!token) return;
    try {
      const me = await api.getAuthMe();
      if (!me.photoData) return;
      const merged = mergeAuthUserWithCachedPhoto(me);
      setCachedAuthUser(merged);
      setUser((current) =>
        current?.id === merged.id ? merged : current,
      );
    } catch {
      /* avatar is optional */
    }
  }, []);

  const refreshUser = useCallback(async () => {
    const token = getAuthToken();
    if (!token) {
      setUser(null);
      return null;
    }
    try {
      const session = await api.refreshAuth({ omitPhoto: true });
      applySession(session.accessToken, session.user);
      scheduleAuthPhotoRefresh(() => {
        void refreshAuthPhoto();
      });
      return mergeAuthUserWithCachedPhoto(session.user);
    } catch (error) {
      try {
        const me = await api.getAuthMe({ omitPhoto: true });
        const merged = mergeAuthUserWithCachedPhoto(me);
        setCachedAuthUser(merged);
        setUser(merged);
        scheduleAuthPhotoRefresh(() => {
          void refreshAuthPhoto();
        });
        return merged;
      } catch (innerError) {
        handleAuthFailure(innerError);
        return getCachedAuthUser();
      }
    }
  }, [applySession, handleAuthFailure, refreshAuthPhoto]);

  useEffect(() => {
    let cancelled = false;
    const boot = async () => {
      const token = getAuthToken();
      if (!token) {
        if (!cancelled) {
          setUser(null);
          setLoading(false);
        }
        return;
      }
      const cached = getCachedAuthUser();
      if (cached && !cancelled) {
        setUser(cached);
        bootstrapPersonnelAppShell();
      }
      try {
        const me = await api.getAuthMe({ omitPhoto: true });
        if (!cancelled) {
          const merged = mergeAuthUserWithCachedPhoto(me);
          setCachedAuthUser(merged);
          setUser(merged);
          if (!merged.photoData) {
            scheduleAuthPhotoRefresh(() => {
              if (!cancelled) void refreshAuthPhoto();
            });
          }
        }
      } catch (error) {
        if (!cancelled) handleAuthFailure(error);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    void boot();
    return () => {
      cancelled = true;
    };
  }, [handleAuthFailure, refreshAuthPhoto]);

  useEffect(() => {
    const onLogout = () => logout();
    window.addEventListener(AUTH_LOGOUT_EVENT, onLogout);
    return () => window.removeEventListener(AUTH_LOGOUT_EVENT, onLogout);
  }, [logout]);

  const login = useCallback(async (email: string, password: string) => {
    const session = await api.login({ email, password });
    applySession(session.accessToken, session.user);
    return mergeAuthUserWithCachedPhoto(session.user);
  }, [applySession]);

  const register = useCallback(
    async (email: string, password: string, displayName: string) => {
      const session = await api.register({ email, password, displayName });
      applySession(session.accessToken, session.user);
      return mergeAuthUserWithCachedPhoto(session.user);
    },
    [applySession],
  );

  const value = useMemo<AuthContextValue>(
    () => ({
      user,
      loading,
      canView: canViewApp(user),
      canEdit: canEditApp(user),
      isAdmin: isAdminUser(user),
      canEditArea: (area: WritePermission) => canEditAreaForUser(user, area),
      login,
      register,
      logout,
      refreshUser,
    }),
    [user, loading, login, register, logout, refreshUser],
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error("useAuth must be used within AuthProvider");
  }
  return ctx;
}
