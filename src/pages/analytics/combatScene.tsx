import { useEffect, useRef } from "react";
import { useFrame, useThree } from "@react-three/fiber";
import { OrbitControls } from "@react-three/drei";
import type { Group } from "three";
import { CombatSoldierCrowd } from "./combatSoldiers";
import { CombatDroneFleet } from "./combatDrones";
import { CombatSceneReadySignal } from "./combatModelsPreloader";

export type CombatBlock = {
  label: string;
  staff: number;
  actual: number;
  x: number;
  z: number;
};

export const COMBAT_CAMERA_POSITION = [1.8, 3.8, 8.4] as const;
export const COMBAT_CAMERA_TARGET = [0, -1.35, 0] as const;

type OrbitControlsHandle = {
  target: { set: (x: number, y: number, z: number) => void };
  update: () => void;
};

export function CombatCameraRig({
  orbitEnabled,
  resetKey,
}: {
  orbitEnabled: boolean;
  resetKey: number;
}) {
  const { camera } = useThree();
  const controlsRef = useRef<OrbitControlsHandle | null>(null);

  useEffect(() => {
    camera.position.set(...COMBAT_CAMERA_POSITION);
    camera.lookAt(...COMBAT_CAMERA_TARGET);
    camera.updateProjectionMatrix();

    const frame = requestAnimationFrame(() => {
      const controls = controlsRef.current;
      if (!controls) return;
      controls.target.set(...COMBAT_CAMERA_TARGET);
      controls.update();
    });

    return () => cancelAnimationFrame(frame);
  }, [camera, orbitEnabled, resetKey]);

  if (!orbitEnabled) return null;

  return (
    <OrbitControls
      ref={controlsRef as never}
      enableDamping
      dampingFactor={0.08}
      target={[...COMBAT_CAMERA_TARGET]}
      minDistance={3}
      maxDistance={18}
      maxPolarAngle={Math.PI * 0.48}
      enablePan
    />
  );
}

export function CombatStructureBlock({
  block,
  selected,
}: {
  block: CombatBlock;
  selected?: boolean;
}) {
  const baseSize = Math.max(0.7, Math.min(1.35, Math.sqrt(block.staff / 80)));

  return (
    <group position={[block.x, 0, block.z]}>
      <mesh position={[0, -0.04, 0]} rotation={[-Math.PI / 2, 0, 0]}>
        <boxGeometry args={[baseSize + 0.55, baseSize + 0.42, 0.08]} />
        <meshStandardMaterial
          color="#24261b"
          metalness={0.2}
          roughness={0.75}
        />
      </mesh>
      <mesh position={[0, 0.82, 0]}>
        <boxGeometry args={[baseSize, 1.65, baseSize]} />
        <meshBasicMaterial
          color={selected ? "#d7d785" : "#ded9c9"}
          wireframe
          transparent
          opacity={selected ? 0.7 : 0.32}
        />
      </mesh>
    </group>
  );
}

export function CombatStructureScene({
  blocks,
  selectedLabel,
  orbitEnabled = false,
  cameraResetKey = 0,
  onReady,
}: {
  blocks: CombatBlock[];
  selectedLabel?: string;
  orbitEnabled?: boolean;
  cameraResetKey?: number;
  onReady?: () => void;
}) {
  const groupRef = useRef<Group>(null);

  useFrame(({ clock }) => {
    if (!groupRef.current || orbitEnabled) return;
    groupRef.current.rotation.y =
      -0.2 + Math.sin(clock.elapsedTime * 0.28) * 0.025;
  });

  useEffect(() => {
    if (!groupRef.current) return;
    if (orbitEnabled) groupRef.current.rotation.y = -0.2;
  }, [orbitEnabled]);

  return (
    <>
      <CombatCameraRig
        orbitEnabled={orbitEnabled}
        resetKey={cameraResetKey}
      />
      <ambientLight intensity={0.78} />
      <directionalLight position={[4, 6, 3]} intensity={1.55} />
      <pointLight position={[-3, 3, -4]} color="#d7d785" intensity={0.75} />
      <group ref={groupRef} position={[0, -0.75, 0]} rotation={[0, -0.2, 0]}>
        <gridHelper
          args={[7, 14, "#3e4425", "#22251b"]}
          position={[0, -0.08, 0]}
        />
        {blocks.map((block) => (
          <CombatStructureBlock
            block={block}
            key={block.label}
            selected={block.label === selectedLabel}
          />
        ))}
        <CombatSoldierCrowd blocks={blocks} selectedLabel={selectedLabel} />
        <CombatDroneFleet blocks={blocks} />
      </group>
      {onReady ? <CombatSceneReadySignal onReady={onReady} /> : null}
    </>
  );
}
