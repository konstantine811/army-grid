export type AppToastVariant = "STATUS" | "WARNING" | "CRITICAL" | "INFO";

export type AppToastData = {
  id: string;
  title: string;
  description?: string;
  variant?: AppToastVariant;
  duration?: number;
};

type Listener = (toasts: AppToastData[]) => void;

let toasts: AppToastData[] = [];
let toastCounter = 0;
const listeners = new Set<Listener>();

const emit = () => {
  for (const listener of listeners) listener(toasts);
};

export const subscribeAppToasts = (listener: Listener) => {
  listeners.add(listener);
  listener(toasts);
  return () => {
    listeners.delete(listener);
  };
};

export const dismissAppToast = (id: string) => {
  toasts = toasts.filter((item) => item.id !== id);
  emit();
};

export const showAppToast = (input: Omit<AppToastData, "id">) => {
  const id = `toast-${++toastCounter}`;
  const duration = input.duration ?? 5200;
  const next: AppToastData = { ...input, id, variant: input.variant ?? "INFO" };
  toasts = [...toasts, next].slice(-5);
  emit();
  window.setTimeout(() => dismissAppToast(id), duration);
  return id;
};

/** Backend blocked a write / access (403). */
export const showBackendBlockedToast = (message: string) =>
  showAppToast({
    title: "Доступ заборонено",
    description: message.trim() || "Бекенд відхилив цю дію.",
    variant: "WARNING",
    duration: 6500,
  });
