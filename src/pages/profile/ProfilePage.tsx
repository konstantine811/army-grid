import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import { api } from "../../api";
import { useAuth } from "../../auth/AuthProvider";
import {
  WRITE_PERMISSION_LABELS,
  WRITE_PERMISSIONS,
  normalizeWritePermissions,
  type WritePermission,
} from "../../auth/authTypes";
import {
  buildPersonSummary,
  getPersonDisplayName,
} from "../personnel/personnelUtils";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

const fileToDataUrl = (file: File) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Не вдалося прочитати файл"));
    reader.readAsDataURL(file);
  });

const normalizeSearch = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

export function ProfilePage() {
  const { user, refreshUser, isAdmin } = useAuth();
  const [displayName, setDisplayName] = useState(user?.displayName || "");
  const [nickname, setNickname] = useState(user?.nickname || "");
  const [photoData, setPhotoData] = useState(user?.photoData || "");
  const [linkedPersonExternalId, setLinkedPersonExternalId] = useState(
    user?.linkedPersonExternalId || "",
  );
  const [linkedPersonFullName, setLinkedPersonFullName] = useState(
    user?.linkedPersonFullName || "",
  );
  const [rosterRows, setRosterRows] = useState<EjournalPreviewRow[]>([]);
  const [personQuery, setPersonQuery] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [loadingRoster, setLoadingRoster] = useState(false);

  useEffect(() => {
    if (!user) return;
    setDisplayName(user.displayName || "");
    setNickname(user.nickname || "");
    setPhotoData(user.photoData || "");
    setLinkedPersonExternalId(user.linkedPersonExternalId || "");
    setLinkedPersonFullName(user.linkedPersonFullName || "");
  }, [user]);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoadingRoster(true);
      try {
        const latest = await api.getLatestPersonnelRoster();
        if (cancelled) return;
        const rows = (latest?.rows ?? []).map(
          (row) =>
            ({
              ...row.values,
              __dbRowId: row.id,
            }) as EjournalPreviewRow,
        );
        setRosterRows(rows);
      } catch {
        if (!cancelled) setRosterRows([]);
      } finally {
        if (!cancelled) setLoadingRoster(false);
      }
    };
    void load();
    return () => {
      cancelled = true;
    };
  }, []);

  const permissions = useMemo(
    () => normalizeWritePermissions(user?.writePermissions),
    [user?.writePermissions],
  );

  const personMatches = useMemo(() => {
    const query = normalizeSearch(personQuery);
    if (query.length < 2) return [];
    const matches: Array<{
      externalId: string;
      fullName: string;
      row: EjournalPreviewRow;
    }> = [];
    for (const row of rosterRows) {
      const summary = buildPersonSummary(row);
      const fullName = getPersonDisplayName(row) || summary.fullName || "";
      const externalId = String(summary.externalId || "").trim();
      if (!externalId || !fullName) continue;
      const haystack = normalizeSearch(`${fullName} ${externalId}`);
      if (!haystack.includes(query)) continue;
      matches.push({ externalId, fullName, row });
      if (matches.length >= 12) break;
    }
    return matches;
  }, [personQuery, rosterRows]);

  const saveProfile = async () => {
    if (!user) return;
    setBusy(true);
    setError("");
    setMessage("");
    try {
      await api.updateOwnProfile({
        displayName: displayName.trim(),
        nickname: nickname.trim(),
        photoData: photoData || "",
        linkedPersonExternalId: linkedPersonExternalId.trim() || null,
        linkedPersonFullName: linkedPersonFullName.trim() || null,
      });
      await refreshUser();
      setMessage("Профіль збережено.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося зберегти профіль");
    } finally {
      setBusy(false);
    }
  };

  const onPickPhoto = async (file: File | undefined) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setError("Оберіть файл зображення.");
      return;
    }
    if (file.size > 4_500_000) {
      setError("Фото завелике (до ~4.5 МБ).");
      return;
    }
    try {
      const dataUrl = await fileToDataUrl(file);
      setPhotoData(dataUrl);
      setError("");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Помилка читання фото");
    }
  };

  const linkPerson = async (match: {
    externalId: string;
    fullName: string;
  }) => {
    setLinkedPersonExternalId(match.externalId);
    setLinkedPersonFullName(match.fullName);
    if (!nickname.trim()) {
      const short =
        match.fullName.split(/\s+/).filter(Boolean).slice(0, 2).join(" ") ||
        match.fullName;
      setNickname(short);
    }
    if (!displayName.trim()) setDisplayName(match.fullName);
    try {
      const photo = await api.getPersonPhoto(match.externalId);
      if (photo?.photoData) setPhotoData(photo.photoData);
    } catch {
      /* photo optional */
    }
    setPersonQuery("");
    setMessage(`Привʼязано до особового складу: ${match.fullName}`);
  };

  if (!user) return null;

  return (
    <main className="main-panel profile-page">
      <header className="topbar">
        <div>
          <Typography component="h1" variant="h5">
            Мій профіль
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Доступи, нікнейм, фото та привʼязка до особового складу / анкети.
          </Typography>
        </div>
        <Button variant="contained" disabled={busy} onClick={() => void saveProfile()}>
          {busy ? "Збереження…" : "Зберегти"}
        </Button>
      </header>

      {error ? <Alert severity="error">{error}</Alert> : null}
      {message ? <Alert severity="success">{message}</Alert> : null}

      <section className="analytics-panel profile-access-panel">
        <div className="panel-heading">Доступи</div>
        <Stack spacing={1} sx={{ p: 2 }}>
          <Typography variant="body2">
            Пошта: <strong>{user.email}</strong>
          </Typography>
          <Typography variant="body2">
            Роль: <strong>{isAdmin ? "Адміністратор" : "Користувач"}</strong>
          </Typography>
          <Typography variant="body2">
            Перегляд застосунку:{" "}
            <strong>{user.accessGranted || isAdmin ? "так" : "ні"}</strong>
          </Typography>
          <div className="profile-permissions">
            <Typography variant="body2" sx={{ mb: 1 }}>
              Права змін:
            </Typography>
            {WRITE_PERMISSIONS.map((permission: WritePermission) => {
              const on = isAdmin || permissions.includes(permission);
              return (
                <div
                  key={permission}
                  className={
                    on
                      ? "profile-permission-chip is-on"
                      : "profile-permission-chip"
                  }
                >
                  {WRITE_PERMISSION_LABELS[permission]}
                  {on ? " · так" : " · ні"}
                </div>
              );
            })}
            {!isAdmin && permissions.length === 0 ? (
              <Typography variant="caption" color="text.secondary">
                Права на зміни ще не надано адміністратором — доступні сторінки
                лише для перегляду.
              </Typography>
            ) : null}
          </div>
        </Stack>
      </section>

      <section className="analytics-panel profile-edit-panel">
        <div className="panel-heading">Про себе</div>
        <div className="profile-edit-grid">
          <div className="profile-photo-block">
            {photoData ? (
              <img
                className="profile-photo"
                src={photoData}
                alt={nickname || displayName || "Фото"}
              />
            ) : (
              <div className="profile-photo is-empty">Немає фото</div>
            )}
            <Stack direction="row" spacing={1} sx={{ mt: 1 }}>
              <Button component="label" variant="outlined" size="small">
                Завантажити
                <input
                  hidden
                  type="file"
                  accept="image/*"
                  onChange={(event) => {
                    void onPickPhoto(event.target.files?.[0]);
                    event.currentTarget.value = "";
                  }}
                />
              </Button>
              {photoData ? (
                <Button
                  size="small"
                  variant="outlined"
                  onClick={() => setPhotoData("")}
                >
                  Прибрати
                </Button>
              ) : null}
            </Stack>
          </div>
          <Stack spacing={1.5} sx={{ minWidth: 0 }}>
            <TextField
              size="small"
              label="Імʼя для відображення"
              value={displayName}
              onChange={(event) => setDisplayName(event.target.value)}
              fullWidth
            />
            <TextField
              size="small"
              label="Нікнейм"
              value={nickname}
              onChange={(event) => setNickname(event.target.value)}
              fullWidth
              helperText="Коротке імʼя для журналів і підписів"
            />
            <TextField
              size="small"
              label="Пошук у особовому складі"
              value={personQuery}
              onChange={(event) => setPersonQuery(event.target.value)}
              fullWidth
              placeholder="ПІБ або ID"
              helperText={
                loadingRoster
                  ? "Завантажую список…"
                  : linkedPersonFullName
                    ? `Привʼязано: ${linkedPersonFullName}${
                        linkedPersonExternalId
                          ? ` · ID ${linkedPersonExternalId}`
                          : ""
                      }`
                    : "Можна підтягнути ПІБ і фото з картки службовця"
              }
            />
            {personMatches.length ? (
              <div className="profile-person-matches">
                {personMatches.map((match) => (
                  <button
                    key={match.externalId}
                    type="button"
                    className="profile-person-match"
                    onClick={() => void linkPerson(match)}
                  >
                    <strong>{match.fullName}</strong>
                    <span>ID {match.externalId}</span>
                  </button>
                ))}
              </div>
            ) : null}
            {linkedPersonExternalId ? (
              <Button
                size="small"
                variant="outlined"
                onClick={() => {
                  setLinkedPersonExternalId("");
                  setLinkedPersonFullName("");
                }}
              >
                Відвʼязати особовий склад
              </Button>
            ) : null}
          </Stack>
        </div>
      </section>
    </main>
  );
}
