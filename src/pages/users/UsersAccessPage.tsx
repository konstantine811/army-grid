import { useCallback, useEffect, useState } from "react";
import {
  Alert,
  Button,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { api } from "../../api";
import type { RegisteredUser } from "../../auth/authTypes";

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
      setUsers(items);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося завантажити");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const toggleAccess = async (user: RegisteredUser) => {
    if (user.role === "ADMIN") return;
    setBusyId(user.id);
    setError("");
    try {
      const updated = await api.setUserAccess(user.id, !user.accessGranted);
      setUsers((current) =>
        current.map((item) => (item.id === updated.id ? { ...item, ...updated } : item)),
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося змінити доступ");
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
            Надайте доступ до перегляду. Редагувати дані може лише адміністратор.
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
              <th>Доступ</th>
              <th>Реєстрація</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {users.map((user) => (
              <tr key={user.id}>
                <td>{user.displayName || "—"}</td>
                <td>{user.email}</td>
                <td>{user.role === "ADMIN" ? "Адмін" : "Користувач"}</td>
                <td>
                  {user.role === "ADMIN" || user.accessGranted
                    ? "Так"
                    : "Ні"}
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
                      {user.accessGranted ? "Забрати доступ" : "Надати доступ"}
                    </Button>
                  )}
                </td>
              </tr>
            ))}
            {!loading && users.length === 0 ? (
              <tr>
                <td colSpan={6}>
                  <Typography variant="body2" color="text.secondary">
                    Поки немає зареєстрованих користувачів.
                  </Typography>
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>

      <Stack sx={{ mt: 2 }}>
        <Typography variant="caption" color="text.secondary">
          Користувачі з доступом бачать усі розділи, але не можуть змінювати,
          імпортувати чи створювати дані.
        </Typography>
      </Stack>
    </main>
  );
}
