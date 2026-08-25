import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import {
  Toast,
  ToastDescription,
  ToastProvider,
  ToastTitle,
  ToastViewport,
} from "@/components/ui/toast/toast";
import {
  dismissAppToast,
  subscribeAppToasts,
  type AppToastData,
} from "../shared/appToast";

export function AppToastHost() {
  const [toasts, setToasts] = useState<AppToastData[]>([]);
  const [mounted, setMounted] = useState(false);

  useEffect(() => subscribeAppToasts(setToasts), []);
  useEffect(() => {
    setMounted(true);
  }, []);

  if (!mounted || typeof document === "undefined" || toasts.length === 0) {
    return null;
  }

  return createPortal(
    <ToastProvider swipeDirection="right" duration={5200}>
      {toasts.map((item) => (
        <Toast
          key={item.id}
          variant={item.variant ?? "INFO"}
          open
          onOpenChange={(open) => {
            if (!open) dismissAppToast(item.id);
          }}
          duration={item.duration ?? 5200}
        >
          <ToastTitle>{item.title}</ToastTitle>
          {item.description ? (
            <ToastDescription>{item.description}</ToastDescription>
          ) : null}
        </Toast>
      ))}
      <ToastViewport className="app-toast-viewport" />
    </ToastProvider>,
    document.body,
  );
}
