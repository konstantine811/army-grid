import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  Chip,
  IconButton,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  CheckCircleOutlineIcon,
  DeleteOutlineOutlinedIcon,
  EditOutlinedIcon,
  FormatListBulletedOutlinedIcon,
} from "@/components/sci/icons";
import {
  api,
  type BackendWorkTask,
  type WorkTaskCategory,
  type WorkTaskStatus,
} from "../../api";
import { useAuth } from "../../auth/AuthProvider";

const CATEGORIES: Array<{ value: WorkTaskCategory; label: string }> = [
  { value: "ubd_status", label: "Запит статусу УБД (інша частина)" },
  { value: "ubd_send", label: "Відправити УБД" },
  { value: "tvk", label: "ТВК на іншу локацію" },
  { value: "other", label: "Інше" },
];

const STATUS_FILTERS: Array<{ value: WorkTaskStatus | "all"; label: string }> = [
  { value: "open", label: "Відкриті" },
  { value: "done", label: "Зроблені" },
  { value: "irrelevant", label: "Неактуальні" },
  { value: "all", label: "Усі" },
];

const STATUS_LABELS: Record<WorkTaskStatus, string> = {
  open: "Відкрите",
  done: "Зроблено",
  irrelevant: "Неактуально",
};

const categoryLabel = (value: string) =>
  CATEGORIES.find((item) => item.value === value)?.label || "Інше";

const formatWhen = (iso: string) => {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return iso;
  return date.toLocaleString("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
};

export function WorkTasksPage() {
  const { user } = useAuth();
  const [tasks, setTasks] = useState<BackendWorkTask[]>([]);
  const [statusFilter, setStatusFilter] = useState<WorkTaskStatus | "all">(
    "open",
  );
  const [queryInput, setQueryInput] = useState("");
  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  const [title, setTitle] = useState("");
  const [personName, setPersonName] = useState("");
  const [category, setCategory] = useState<WorkTaskCategory>("ubd_status");
  const [location, setLocation] = useState("");
  const [firstComment, setFirstComment] = useState("");
  const [commentDraft, setCommentDraft] = useState("");
  const [showCreate, setShowCreate] = useState(false);
  const [editing, setEditing] = useState(false);
  const [editTitle, setEditTitle] = useState("");
  const [editPersonName, setEditPersonName] = useState("");
  const [editCategory, setEditCategory] = useState<WorkTaskCategory>("ubd_status");
  const [editLocation, setEditLocation] = useState("");

  const load = useCallback(async () => {
    setError("");
    setLoading(true);
    try {
      const items = await api.listWorkTasks(
        statusFilter,
        query.trim() || undefined,
      );
      setTasks(items);
      setSelectedId((current) =>
        current && items.some((item) => item.id === current)
          ? current
          : items[0]?.id ?? null,
      );
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося завантажити завдання",
      );
    } finally {
      setLoading(false);
    }
  }, [statusFilter, query]);

  useEffect(() => {
    void load();
  }, [load]);

  useEffect(() => {
    if (!error) return;
    const retry = () => {
      void load();
    };
    const onVisible = () => {
      if (document.visibilityState === "visible") retry();
    };
    window.addEventListener("focus", retry);
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      window.removeEventListener("focus", retry);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, [error, load]);

  useEffect(() => {
    const timer = window.setTimeout(() => setQuery(queryInput), 350);
    return () => window.clearTimeout(timer);
  }, [queryInput]);

  const selected = useMemo(
    () => tasks.find((item) => item.id === selectedId) ?? null,
    [tasks, selectedId],
  );

  useEffect(() => {
    setEditing(false);
  }, [selected?.id]);

  const counts = { total: tasks.length };

  const replaceTask = (next: BackendWorkTask) => {
    setTasks((current) =>
      current
        .map((item) => (item.id === next.id ? next : item))
        .filter((item) =>
          statusFilter === "all" ? true : item.status === statusFilter,
        ),
    );
    setSelectedId(next.id);
  };

  const createTask = async () => {
    if (!title.trim()) {
      setError("Вкажіть, що треба зробити");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const created = await api.createWorkTask({
        title: title.trim(),
        personName: personName.trim() || undefined,
        category,
        location: location.trim() || undefined,
        firstComment: firstComment.trim() || undefined,
      });
      setTitle("");
      setPersonName("");
      setLocation("");
      setFirstComment("");
      setCategory("ubd_status");
      setShowCreate(false);
      setMessage("Завдання створено");
      if (statusFilter !== "open" && statusFilter !== "all") {
        setStatusFilter("open");
      }
      setTasks((current) => [created, ...current]);
      setSelectedId(created.id);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося створити");
    } finally {
      setBusy(false);
    }
  };

  const startEdit = () => {
    if (!selected) return;
    const knownCategory = CATEGORIES.some((item) => item.value === selected.category)
      ? (selected.category as WorkTaskCategory)
      : "other";
    setEditTitle(selected.title);
    setEditPersonName(selected.personName || "");
    setEditCategory(knownCategory);
    setEditLocation(selected.location || "");
    setEditing(true);
  };

  const saveEdit = async () => {
    if (!selected) return;
    if (!editTitle.trim()) {
      setError("Вкажіть, що треба зробити");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const updated = await api.updateWorkTask(selected.id, {
        title: editTitle.trim(),
        personName: editPersonName.trim(),
        category: editCategory,
        location: editLocation.trim(),
      });
      replaceTask(updated);
      setEditing(false);
      setMessage("Зміни збережено");
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Не вдалося зберегти зміни",
      );
    } finally {
      setBusy(false);
    }
  };

  const setStatus = async (status: WorkTaskStatus) => {
    if (!selected) return;
    setBusy(true);
    setError("");
    try {
      const next = await api.updateWorkTask(selected.id, { status });
      replaceTask(next);
      setMessage(
        status === "done"
          ? "Позначено як зроблене"
          : status === "irrelevant"
            ? "Позначено як неактуальне"
            : "Завдання знову відкрите",
      );
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося змінити статус");
    } finally {
      setBusy(false);
    }
  };

  const addComment = async () => {
    if (!selected || !commentDraft.trim()) return;
    setBusy(true);
    setError("");
    try {
      const next = await api.addWorkTaskComment(
        selected.id,
        commentDraft.trim(),
      );
      setCommentDraft("");
      replaceTask(next);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося додати коментар");
    } finally {
      setBusy(false);
    }
  };

  const removeTask = async () => {
    if (!selected) return;
    if (!window.confirm("Видалити це завдання назавжди?")) return;
    setBusy(true);
    setError("");
    try {
      await api.deleteWorkTask(selected.id);
      setTasks((current) => current.filter((item) => item.id !== selected.id));
      setSelectedId(null);
      setMessage("Завдання видалено");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не вдалося видалити");
    } finally {
      setBusy(false);
    }
  };

  return (
    <main className="main-panel work-tasks-page">
      <header className="topbar">
        <div>
          <Typography component="h1" variant="h5">
            Мої завдання
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Особисті кроки по УБД, ТВК і запитах. Бачите лише ви —{" "}
            {user?.email || "пошта акаунта"}.
          </Typography>
        </div>
        <Button
          variant="contained"
          startIcon={<FormatListBulletedOutlinedIcon />}
          onClick={() => setShowCreate((value) => !value)}
        >
          {showCreate ? "Сховати форму" : "Нове завдання"}
        </Button>
      </header>

      {error ? (
        <Alert
          severity="error"
          action={
            <Button size="small" variant="outlined" onClick={() => void load()}>
              Повторити
            </Button>
          }
        >
          {error}
        </Alert>
      ) : null}
      {message ? <Alert severity="success">{message}</Alert> : null}

      <section className="work-tasks-toolbar">
        <div className="work-tasks-stats">
          <span>
            У списку: <strong>{counts.total}</strong>
          </span>
        </div>
        <div className="work-tasks-filters">
          {STATUS_FILTERS.map((item) => (
            <button
              key={item.value}
              type="button"
              className={
                statusFilter === item.value
                  ? "work-tasks-filter is-on"
                  : "work-tasks-filter"
              }
              onClick={() => setStatusFilter(item.value)}
            >
              {item.label}
            </button>
          ))}
        </div>
        <TextField
          size="small"
          placeholder="Пошук: ПІБ, дія, локація…"
          value={queryInput}
          onChange={(event) => setQueryInput(event.target.value)}
        />
      </section>

      <div className="work-tasks-layout">
        <div className="work-tasks-main">
          {showCreate ? (
            <section className="analytics-panel work-tasks-create">
              <div className="panel-heading">Нове завдання</div>
              <div className="work-tasks-create-grid">
                <TextField
                  label="Що зробити"
                  value={title}
                  onChange={(event) => setTitle(event.target.value)}
                  placeholder="Запитати статус УБД / відправити УБД / ТВК…"
                />
                <TextField
                  label="По кому"
                  value={personName}
                  onChange={(event) => setPersonName(event.target.value)}
                  placeholder="ПІБ"
                />
                <TextField
                  select
                  label="Тип"
                  value={category}
                  onChange={(event) =>
                    setCategory(event.target.value as WorkTaskCategory)
                  }
                >
                  {CATEGORIES.map((item) => (
                    <MenuItem key={item.value} value={item.value}>
                      {item.label}
                    </MenuItem>
                  ))}
                </TextField>
                <TextField
                  label="Локація / частина"
                  value={location}
                  onChange={(event) => setLocation(event.target.value)}
                  placeholder="Куди / звідки"
                />
                <TextField
                  className="work-tasks-textarea"
                  label="Перший коментар (що вже зроблено / що відбувається)"
                  value={firstComment}
                  onChange={(event) => setFirstComment(event.target.value)}
                  multiline
                  rows={4}
                  style={{ resize: "vertical" }}
                />
                <Button
                  variant="contained"
                  disabled={busy}
                  onClick={() => void createTask()}
                >
                  Створити
                </Button>
              </div>
            </section>
          ) : null}

          <section className="work-tasks-list">
          {loading ? (
            <Typography variant="body2" color="text.secondary">
              Завантажую…
            </Typography>
          ) : tasks.length === 0 ? (
            <Typography variant="body2" color="text.secondary">
              Немає завдань у цьому фільтрі. Створіть перше — щоб не губити, по
              кому що робили.
            </Typography>
          ) : (
            tasks.map((task) => (
              <button
                key={task.id}
                type="button"
                className={
                  task.id === selectedId
                    ? `work-task-card is-selected status-${task.status}`
                    : `work-task-card status-${task.status}`
                }
                onClick={() => setSelectedId(task.id)}
              >
                <div className="work-task-card-head">
                  <strong>{task.title}</strong>
                  <Chip
                    size="small"
                    label={STATUS_LABELS[task.status as WorkTaskStatus] || task.status}
                  />
                </div>
                <span>
                  {task.personName || "без ПІБ"}
                  {task.location ? ` · ${task.location}` : ""}
                </span>
                <span className="work-task-card-meta">
                  {categoryLabel(task.category)} · кроків{" "}
                  {task.comments.length} · {formatWhen(task.updatedAt)}
                </span>
              </button>
            ))
          )}
        </section>
        </div>

        <section className="analytics-panel work-tasks-detail">
          {!selected ? (
            <Typography variant="body2" color="text.secondary" sx={{ p: 2 }}>
              Оберіть завдання зліва або створіть нове.
            </Typography>
          ) : (
            <>
              <div className="panel-heading work-tasks-detail-heading">
                <span>{selected.title}</span>
                {!editing ? (
                  <IconButton
                    aria-label="Редагувати завдання"
                    size="small"
                    onClick={startEdit}
                  >
                    <EditOutlinedIcon fontSize="small" />
                  </IconButton>
                ) : null}
              </div>
              <div className="work-tasks-detail-body">
                {editing ? (
                  <div className="work-tasks-create-grid work-tasks-edit-grid">
                    <TextField
                      label="Що зробити"
                      value={editTitle}
                      onChange={(event) => setEditTitle(event.target.value)}
                    />
                    <TextField
                      label="По кому"
                      value={editPersonName}
                      onChange={(event) => setEditPersonName(event.target.value)}
                      placeholder="ПІБ"
                    />
                    <TextField
                      select
                      label="Тип"
                      value={editCategory}
                      onChange={(event) =>
                        setEditCategory(event.target.value as WorkTaskCategory)
                      }
                    >
                      {CATEGORIES.map((item) => (
                        <MenuItem key={item.value} value={item.value}>
                          {item.label}
                        </MenuItem>
                      ))}
                    </TextField>
                    <TextField
                      className="work-tasks-textarea"
                      label="Локація / частина"
                      value={editLocation}
                      onChange={(event) => setEditLocation(event.target.value)}
                      multiline
                      rows={3}
                      placeholder="Куди / звідки"
                      style={{ resize: "vertical" }}
                    />
                    <Stack direction="row" spacing={1}>
                      <Button
                        variant="contained"
                        disabled={busy}
                        onClick={() => void saveEdit()}
                      >
                        Зберегти
                      </Button>
                      <Button
                        variant="outlined"
                        disabled={busy}
                        onClick={() => setEditing(false)}
                      >
                        Скасувати
                      </Button>
                    </Stack>
                  </div>
                ) : (
                  <>
                    <Typography variant="body2">
                      По кому: <strong>{selected.personName || "—"}</strong>
                    </Typography>
                    <Typography variant="body2">
                      Тип: <strong>{categoryLabel(selected.category)}</strong>
                    </Typography>
                    <Typography variant="body2">
                      Локація: <strong>{selected.location || "—"}</strong>
                    </Typography>
                  </>
                )}
                <Typography variant="body2">
                  Статус:{" "}
                  <strong>
                    {STATUS_LABELS[selected.status as WorkTaskStatus] ||
                      selected.status}
                  </strong>
                </Typography>
                <Typography variant="caption" color="text.secondary">
                  Створено {formatWhen(selected.createdAt)} · оновлено{" "}
                  {formatWhen(selected.updatedAt)}
                </Typography>

                <Stack direction="row" spacing={1} sx={{ mt: 1, flexWrap: "wrap" }}>
                  {selected.status !== "done" ? (
                    <Button
                      size="small"
                      variant="contained"
                      startIcon={<CheckCircleOutlineIcon />}
                      disabled={busy}
                      onClick={() => void setStatus("done")}
                    >
                      Зроблено
                    </Button>
                  ) : null}
                  {selected.status !== "irrelevant" ? (
                    <Button
                      size="small"
                      variant="outlined"
                      disabled={busy}
                      onClick={() => void setStatus("irrelevant")}
                    >
                      Неактуально
                    </Button>
                  ) : null}
                  {selected.status !== "open" ? (
                    <Button
                      size="small"
                      variant="outlined"
                      disabled={busy}
                      onClick={() => void setStatus("open")}
                    >
                      Відкрити знову
                    </Button>
                  ) : null}
                  <Button
                    size="small"
                    variant="outlined"
                    startIcon={<DeleteOutlineOutlinedIcon />}
                    disabled={busy}
                    onClick={() => void removeTask()}
                  >
                    Видалити
                  </Button>
                </Stack>

                <Typography variant="subtitle2" sx={{ mt: 2 }}>
                  Хід роботи
                </Typography>
                {selected.comments.length === 0 ? (
                  <Typography variant="body2" color="text.secondary">
                    Коментарів ще немає. Допишіть, що запитали / куди відправили.
                  </Typography>
                ) : (
                  <ol className="work-task-comments">
                    {selected.comments.map((comment) => (
                      <li key={comment.id}>
                        <time>{formatWhen(comment.createdAt)}</time>
                        <p>{comment.body}</p>
                      </li>
                    ))}
                  </ol>
                )}

                <div className="work-tasks-comment-compose">
                  <TextField
                    className="work-tasks-textarea"
                    label="Новий коментар / крок"
                    value={commentDraft}
                    onChange={(event) => setCommentDraft(event.target.value)}
                    multiline
                    rows={4}
                    placeholder="Що відбувається зараз…"
                    style={{ resize: "vertical" }}
                  />
                  <Button
                    variant="contained"
                    disabled={busy || !commentDraft.trim()}
                    onClick={() => void addComment()}
                  >
                    Додати коментар
                  </Button>
                </div>
              </div>
            </>
          )}
        </section>
      </div>
    </main>
  );
}
