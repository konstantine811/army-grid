import { useEffect, useRef } from "react";

const AUDIO_STORAGE_KEY = "army-grid:sci-live-audio";
const INTERACTIVE_SELECTOR =
  "button, a, [role='button'], .sci-button, input, select, textarea, label, .panel-heading, .nav-item, [data-sci-live]";

const prefersReducedMotion = () =>
  typeof window !== "undefined" &&
  window.matchMedia("(prefers-reduced-motion: reduce)").matches;

const isAudioEnabled = () => {
  try {
    return window.localStorage.getItem(AUDIO_STORAGE_KEY) !== "0";
  } catch {
    return true;
  }
};

/** Soft terminal tick via Web Audio (no asset files). */
const playClickTick = (kind: "tap" | "action" = "tap") => {
  if (prefersReducedMotion() || !isAudioEnabled()) return;
  try {
    const AudioCtx =
      window.AudioContext ||
      (window as unknown as { webkitAudioContext: typeof AudioContext })
        .webkitAudioContext;
    const ctx = new AudioCtx();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.value = kind === "action" ? 720 : 520;
    gain.gain.value = kind === "action" ? 0.045 : 0.028;
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.09);
    window.setTimeout(() => {
      try {
        osc.stop();
        void ctx.close();
      } catch {
        /* ignore */
      }
    }, 120);
  } catch {
    /* ignore */
  }
};

const spawnRadarPing = (clientX: number, clientY: number) => {
  const ping = document.createElement("div");
  ping.className = "sci-radar-ping";
  ping.style.left = `${clientX}px`;
  ping.style.top = `${clientY}px`;
  document.body.appendChild(ping);
  window.setTimeout(() => ping.remove(), 900);
};

const pulseTarget = (target: EventTarget | null) => {
  if (!(target instanceof Element)) return;
  const el = target.closest(INTERACTIVE_SELECTOR);
  if (!(el instanceof HTMLElement)) return;
  el.classList.add("sci-pulse-once");
  window.setTimeout(() => el.classList.remove("sci-pulse-once"), 420);
};

/**
 * Global sci-fi live feedback inspired by CodePen gridghost/GggZdBL:
 * radar click rings, soft UI ticks, button pulse, ambient particles + scan line.
 */
export function SciLiveFeedback() {
  const layerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (prefersReducedMotion()) return;

    const layer = layerRef.current;
    if (!layer) return;

    // Ambient particles
    const field = document.createElement("div");
    field.className = "sci-particle-field";
    const count = 18;
    for (let i = 0; i < count; i += 1) {
      const particle = document.createElement("span");
      particle.className = "sci-particle";
      particle.style.left = `${Math.random() * 100}%`;
      particle.style.animationDelay = `${Math.random() * 18}s`;
      particle.style.animationDuration = `${14 + Math.random() * 16}s`;
      particle.style.opacity = String(0.12 + Math.random() * 0.28);
      field.appendChild(particle);
    }
    layer.appendChild(field);

    const scanner = document.createElement("div");
    scanner.className = "sci-ambient-scanner";
    layer.appendChild(scanner);

    const onPointerDown = (event: PointerEvent) => {
      if (event.button !== 0) return;
      spawnRadarPing(event.clientX, event.clientY);
      pulseTarget(event.target);
      const interactive = event.target instanceof Element
        ? event.target.closest(INTERACTIVE_SELECTOR)
        : null;
      playClickTick(interactive ? "action" : "tap");
    };

    document.addEventListener("pointerdown", onPointerDown, { passive: true });

    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      field.remove();
      scanner.remove();
    };
  }, []);

  return (
    <div
      ref={layerRef}
      className="sci-live-feedback-layer"
      aria-hidden="true"
    />
  );
}
