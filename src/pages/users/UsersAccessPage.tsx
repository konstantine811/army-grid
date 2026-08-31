import { useCallback, useEffect, useState } from "react";
import {
  Alert,
  Button,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { api } from "../../api";
import {
  WRITE_PERMISSION_LABELS,
  WRITE_PERMISSIONS,
  normalizeWritePermissions,
  type RegisteredUser,
  type WritePermission,
} from "../../auth/authTypes";

export function UsersAccessPage() {
  const [users, setUsers] = useState<RegisteredUser[]>([]);
  const [error, setError] = useState("");
  const [busyId, setBusyId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setError("");
    setLoading(true);
    try {
      const items = await api.listAuthUsers();
      setUsers(
        items.map((item) => ({
          ...item,
          writePermissions: normalizeWritePermissions(item.writePermissions),
        })),
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося завантажити");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const patchUser = (updated: RegisteredUser) => {
    setUsers((current) =>
      current.map((item) =>
        item.id === updated.id
          ? {
              ...item,
              ...updated,
              writePermissions: normalizeWritePermissions(
                updated.writePermissions,
              ),
            }
          : item,
      ),
    );
  };

  const toggleAccess = async (user: RegisteredUser) => {
    if (user.role === "ADMIN") return;
    setBusyId(user.id);
    setError("");
    try {
      const updated = await api.setUserAccess(user.id, !user.accessGranted);
      patchUser(updated);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося змінити доступ");
    } finally {
      setBusyId(null);
    }
  };

  const togglePermission = async (
    user: RegisteredUser,
    permission: WritePermission,
  ) => {
    if (user.role === "ADMIN") return;
    const current = normalizeWritePermissions(user.writePermissions);
    const next = current.includes(permission)
      ? current.filter((item) => item !== permission)
      : [...current, permission];

    setBusyId(user.id);
    setError("");
    try {
      const updated = await api.setUserPermissions(user.id, next);
      patchUser(updated);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося змінити права",
      );
    } finally {
      setBusyId(null);
    }
  };

  return (
    <main className="main-panel">
      <header className="topbar">
        <div>
          <Typography component="h1" variant="h5">
            Зареєстровані користувачі
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Надайте доступ до перегляду та окремі права на зміни в розділах.
          </Typography>
        </div>
        <Button variant="outlined" onClick={() => void load()} disabled={loading}>
          Оновити
        </Button>
      </header>

      {error ? <Alert severity="error">{error}</Alert> : null}

      <div className="users-access-table-wrap">
        <table className="users-access-table">
          <thead>
            <tr>
              <th>Ім’я</th>
              <th>Пошта</th>
              <th>Роль</th>
              <th>Перегляд</th>
              <th>Права змін</th>
              <th>Реєстрація</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {users.map((user) => {
              const permissions = normalizeWritePermissions(
                user.writePermissions,
              );
              return (
                <tr key={user.id}>
                  <td>{user.displayName || "—"}</td>
                  <td>{user.email}</td>
                  <td>{user.role === "ADMIN" ? "Адмін" : "Користувач"}</td>
                  <td>
                    {user.role === "ADMIN" || user.accessGranted ? "Так" : "Ні"}
                  </td>
                  <td>
                    {user.role === "ADMIN" ? (
                      <Typography variant="caption" color="text.secondary">
                        Усі розділи
                      </Typography>
                    ) : (
                      <div className="users-access-permissions">
                        {WRITE_PERMISSIONS.map((permission) => {
                          const checked = permissions.includes(permission);
                          return (
                            <label
                              key={permission}
                              className="users-access-permission"
                            >
                              <input
                                type="checkbox"
                                checked={checked}
                                disabled={busyId === user.id}
                                onChange={() =>
                                  void togglePermission(user, permission)
                                }
                              />
                              <span>{WRITE_PERMISSION_LABELS[permission]}</span>
                            </label>
                          );
                        })}
                      </div>
                    )}
                  </td>
                  <td>
                    {new Date(user.createdAt).toLocaleString("uk-UA", {
                      dateStyle: "short",
                      timeStyle: "short",
                    })}
                  </td>
                  <td>
                    {user.role === "ADMIN" ? (
                      <Typography variant="caption" color="text.secondary">
                        —
                      </Typography>
                    ) : (
                      <Button
                        size="small"
                        variant={user.accessGranted ? "outlined" : "contained"}
                        disabled={busyId === user.id}
                        onClick={() => void toggleAccess(user)}
                      >
                        {user.accessGranted
                          ? "Забрати перегляд"
                          : "Надати перегляд"}
                      </Button>
                    )}
                  </td>
                </tr>
              );
            })}
            {!loading && users.length === 0 ? (
              <tr>
                <td colSpan={7}>
                  <Typography variant="body2" color="text.secondary">
                    Поки немає зареєстрованих користувачів.
                  </Typography>
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>

      <Stack sx={{ mt: 2 }} spacing={0.5}>
        <Typography variant="caption" color="text.secondary">
          Перегляд — Огляд, Особовий склад, БЧС, Анкетні дані, Документи.
        </Typography>
        <Typography variant="caption" color="text.secondary">
          Права змін — окремо: редагування відповідного розділу (додавання,
          імпорт у межах розділу тощо). Без права розділ лишається лише для
          перегляду.
        </Typography>
      </Stack>
    </main>
  );
}
