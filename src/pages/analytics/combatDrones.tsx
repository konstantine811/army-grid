import { useMemo, useRef } from "react";
import { useFrame } from "@react-three/fiber";
import { useGLTF } from "@react-three/drei";
import type { Group, Object3D } from "three";
import type { CombatBlock } from "./combatScene";

const DRONE_URL = "/3d_models/dron.glb";
const RBPAK_LABEL = "РБпАК";
const DRONE_COUNT = 5;

useGLTF.preload(DRONE_URL, true);

type DroneSpec = {
  radius: number;
  speed: number;
  phase: number;
  height: number;
  bob: number;
  scaleMul: number;
};

function modelSpan(scene: Object3D) {
  let maxAxis = 1;
  scene.updateMatrixWorld(true);
  scene.traverse((object) => {
    const mesh = object as Object3D & {
      isMesh?: boolean;
      geometry?: {
        computeBoundingBox: () => void;
        boundingBox: {
          max: { x: number; y: number; z: number };
          min: { x: number; y: number; z: number };
        } | null;
      };
    };
    if (!mesh.isMesh || !mesh.geometry) return;
    mesh.geometry.computeBoundingBox();
    const box = mesh.geometry.boundingBox;
    if (!box) return;
    maxAxis = Math.max(
      maxAxis,
      box.max.x - box.min.x,
      box.max.y - box.min.y,
      box.max.z - box.min.z,
    );
  });
  return maxAxis;
}

function buildSpecs(count: number): DroneSpec[] {
  return Array.from({ length: count }, (_, index) => ({
    radius: 0.28 + (index % 3) * 0.12,
    speed: 0.55 + (index % 4) * 0.12,
    phase: (index / count) * Math.PI * 2,
    height: 0.95 + (index % 3) * 0.18,
    bob: 0.04 + (index % 2) * 0.02,
    scaleMul: 0.9 + (index % 3) * 0.12,
  }));
}

function prepareDrone(scene: Object3D) {
  const root = scene.clone(true);
  root.traverse((object) => {
    object.frustumCulled = false;
    const mesh = object as Object3D & {
      isMesh?: boolean;
      castShadow?: boolean;
      receiveShadow?: boolean;
    };
    if (mesh.isMesh) {
      mesh.castShadow = false;
      mesh.receiveShadow = false;
    }
  });
  return root;
}

function CombatDroneFleetInner({
  origin,
}: {
  origin: { x: number; z: number };
}) {
  const gltf = useGLTF(DRONE_URL, true);
  const droneRefs = useRef<Array<Group | null>>([]);
  const specs = useMemo(() => buildSpecs(DRONE_COUNT), []);

  const clones = useMemo(
    () => specs.map(() => prepareDrone(gltf.scene)),
    [gltf.scene, specs],
  );

  const baseScale = useMemo(() => {
    const span = modelSpan(gltf.scene);
    return Math.min(0.045, 0.55 / Math.max(1, span));
  }, [gltf.scene]);

  useFrame(({ clock }) => {
    const t = clock.elapsedTime;
    specs.forEach((spec, index) => {
      const drone = droneRefs.current[index];
      if (!drone) return;

      const angle = t * spec.speed + spec.phase;
      const x = Math.cos(angle) * spec.radius;
      const z = Math.sin(angle) * spec.radius * 0.72;
      const y =
        spec.height + Math.sin(t * (1.1 + index * 0.17) + spec.phase) * spec.bob;

      const yaw = -angle + Math.PI / 2;
      const roll = Math.sin(angle) * 0.22;
      const pitch = Math.sin(t * 0.9 + spec.phase) * 0.08;

      drone.position.set(x, y, z);
      drone.rotation.set(pitch, yaw, roll);
      drone.scale.setScalar(baseScale * spec.scaleMul);
    });
  });

  return (
    <group position={[origin.x, 0, origin.z]}>
      {clones.map((clone, index) => (
        <group
          key={`drone-${index}`}
          ref={(node) => {
            droneRefs.current[index] = node;
          }}
        >
          <primitive object={clone} />
        </group>
      ))}
    </group>
  );
}

export function CombatDroneFleet({ blocks }: { blocks: CombatBlock[] }) {
  const rbpak = useMemo(
    () => blocks.find((block) => block.label === RBPAK_LABEL),
    [blocks],
  );

  if (!rbpak) return null;

  return <CombatDroneFleetInner origin={{ x: rbpak.x, z: rbpak.z }} />;
}
