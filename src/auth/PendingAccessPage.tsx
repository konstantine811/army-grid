import { Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { useAuth } from "./AuthProvider";

export function PendingAccessPage() {
  const { user, logout, refreshUser } = useAuth();

  return (
    <div className="auth-shell">
      <div className="auth-card">
        <Typography variant="overline" color="text.secondary">
          Army Grid
        </Typography>
        <Typography component="h1" variant="h5">
          Очікування доступу
        </Typography>
        <Typography variant="body2" color="text.secondary">
          Акаунт <strong>{user?.email}</strong> зареєстровано. Адміністратор ще
          не надав доступ до перегляду.
        </Typography>
        <Stack direction="row" spacing={1} sx={{ mt: 2 }}>
          <Button
            variant="outlined"
            onClick={() => void refreshUser()}
          >
            Перевірити знову
          </Button>
          <Button variant="text" onClick={logout}>
            Вийти
          </Button>
        </Stack>
      </div>
    </div>
  );
}
