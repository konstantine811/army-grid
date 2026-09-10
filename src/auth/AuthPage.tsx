import { useState, type FormEvent } from "react";
import { Alert, Button, Stack, TextField, Typography } from "@/components/sci/SciPrimitives";
import {
  VisibilityOffOutlinedIcon,
  VisibilityOutlinedIcon,
} from "@/components/sci/icons";
import { useAuth } from "./AuthProvider";
import { getLastAuthEmail } from "./authTypes";

type Mode = "login" | "register";

export function AuthPage() {
  const { login, register } = useAuth();
  const [mode, setMode] = useState<Mode>("login");
  const [email, setEmail] = useState(() => getLastAuthEmail());
  const [password, setPassword] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const onSubmit = async (event: FormEvent) => {
    event.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (mode === "login") {
        await login(email.trim(), password);
      } else {
        await register(email.trim(), password, displayName.trim());
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Помилка автентифікації");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="auth-shell">
      <form
        className="auth-card"
        autoComplete="on"
        onSubmit={(event) => void onSubmit(event)}
      >
        <Typography variant="overline" color="text.secondary">
          Army Grid
        </Typography>
        <Typography component="h1" variant="h5">
          {mode === "login" ? "Вхід" : "Реєстрація"}
        </Typography>
        <Typography variant="body2" color="text.secondary">
          {mode === "login"
            ? "Увійдіть, щоб переглянути дані. Зміни доступні лише адміністратору."
            : "Після реєстрації адміністратор надасть доступ до перегляду."}
        </Typography>

        <Stack spacing={1.5} sx={{ mt: 2 }}>
          {mode === "register" ? (
            <TextField
              label="Ім’я"
              value={displayName}
              onChange={(event) => setDisplayName(event.target.value)}
              required
              autoComplete="name"
            />
          ) : null}
          <TextField
            label="Пошта"
            name="email"
            id="auth-email"
            type="email"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
            required
            autoComplete="username email"
          />
          <TextField
            label="Пароль"
            name="password"
            id="auth-password"
            type={showPassword ? "text" : "password"}
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            required
            autoComplete={mode === "login" ? "current-password" : "new-password"}
            minLength={6}
            suffix={
              <button
                type="button"
                className="auth-password-toggle"
                aria-label={showPassword ? "Сховати пароль" : "Показати пароль"}
                aria-pressed={showPassword}
                tabIndex={0}
                onClick={() => setShowPassword((value) => !value)}
              >
                {showPassword ? (
                  <VisibilityOffOutlinedIcon fontSize="small" />
                ) : (
                  <VisibilityOutlinedIcon fontSize="small" />
                )}
              </button>
            }
          />
          {error ? <Alert severity="error">{error}</Alert> : null}
          <Button type="submit" variant="contained" disabled={busy} fullWidth>
            {busy
              ? "Зачекайте…"
              : mode === "login"
                ? "Увійти"
                : "Зареєструватися"}
          </Button>
          <Button
            type="button"
            variant="text"
            disabled={busy}
            onClick={() => {
              setMode((current) => (current === "login" ? "register" : "login"));
              setError("");
            }}
          >
            {mode === "login"
              ? "Немає акаунта? Зареєструватися"
              : "Вже є акаунт? Увійти"}
          </Button>
        </Stack>
      </form>
    </div>
  );
}
