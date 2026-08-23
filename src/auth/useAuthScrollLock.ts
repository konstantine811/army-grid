import { useEffect } from "react";

/** Lock document scroll while auth screens are shown (mobile rubber-band / double scroll). */
export function useAuthScrollLock(locked: boolean) {
  useEffect(() => {
    if (!locked) return;
    const root = document.documentElement;
    root.classList.add("auth-lock");
    return () => {
      root.classList.remove("auth-lock");
    };
  }, [locked]);
}
