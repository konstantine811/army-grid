import { useEffect, useState } from "react";
import { useProgress } from "@react-three/drei";
import { Spinner } from "@/components/ui/spinner/spinner";

export function CombatModelsPreloader({
  sceneReady,
}: {
  sceneReady: boolean;
}) {
  const { active, progress, loaded, total } = useProgress();
  const [visible, setVisible] = useState(true);

  useEffect(() => {
    if (active || !sceneReady) {
      setVisible(true);
      return;
    }

    const timer = window.setTimeout(() => setVisible(false), 220);
    return () => window.clearTimeout(timer);
  }, [active, sceneReady]);

  if (!visible) return null;

  const percent = Math.min(
    100,
    Math.round(sceneReady && !active ? 100 : progress),
  );

  return (
    <div className="combat-models-preloader" aria-live="polite">
      <div className="combat-models-preloader-card">
        <Spinner size="LG" label="LOADING" />
        <strong>Завантаження 3D-моделей</strong>
        <span>
          {total > 0
            ? `${loaded} / ${total} · ${percent}%`
            : "soldier.glb · dron.glb…"}
        </span>
      </div>
    </div>
  );
}

/** Mount only after Suspense has resolved model loads. */
export function CombatSceneReadySignal({
  onReady,
}: {
  onReady: () => void;
}) {
  useEffect(() => {
    const frame = window.requestAnimationFrame(() => onReady());
    return () => window.cancelAnimationFrame(frame);
  }, [onReady]);

  return null;
}
