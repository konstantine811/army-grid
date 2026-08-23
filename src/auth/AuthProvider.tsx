import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";
import { api } from "../api";
import {
  AUTH_LOGOUT_EVENT,
  canEditApp,
  canViewApp,
  clearAuthToken,
  getAuthToken,
  isAdminUser,
  setAuthToken,
  type AuthUser,
} from "./authTypes";

type AuthContextValue = {
  user: AuthUser | null;
  loading: boolean;
  canView: boolean;
  canEdit: boolean;
  isAdmin: boolean;
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

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<AuthUser | null>(null);
  const [loading, setLoading] = useState(true);

  const logout = useCallback(() => {
    clearAuthToken();
    setUser(null);
  }, []);

  const refreshUser = useCallback(async () => {
    const token = getAuthToken();
    if (!token) {
      setUser(null);
      return null;
    }
    try {
      const session = await api.refreshAuth();
      setAuthToken(session.accessToken);
      setUser(session.user);
      return session.user;
    } catch {
      try {
        const me = await api.getAuthMe();
        setUser(me);
        return me;
      } catch {
        clearAuthToken();
        setUser(null);
        return null;
      }
    }
  }, []);

  useEffect(() => {
    let cancelled = false;
    const boot = async () => {
      if (!getAuthToken()) {
        if (!cancelled) {
          setUser(null);
          setLoading(false);
        }
        return;
      }
      try {
        const me = await api.getAuthMe();
        if (!cancelled) setUser(me);
      } catch {
        clearAuthToken();
        if (!cancelled) setUser(null);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    void boot();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const onLogout = () => logout();
    window.addEventListener(AUTH_LOGOUT_EVENT, onLogout);
    return () => window.removeEventListener(AUTH_LOGOUT_EVENT, onLogout);
  }, [logout]);

  const login = useCallback(async (email: string, password: string) => {
    const session = await api.login({ email, password });
    setAuthToken(session.accessToken);
    setUser(session.user);
    return session.user;
  }, []);

  const register = useCallback(
    async (email: string, password: string, displayName: string) => {
      const session = await api.register({ email, password, displayName });
      setAuthToken(session.accessToken);
      setUser(session.user);
      return session.user;
    },
    [],
  );

  const value = useMemo<AuthContextValue>(
    () => ({
      user,
      loading,
      canView: canViewApp(user),
      canEdit: canEditApp(user),
      isAdmin: isAdminUser(user),
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
