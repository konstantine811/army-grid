import { useEffect, useMemo, useRef } from "react";
import { useFrame } from "@react-three/fiber";
import { useGLTF } from "@react-three/drei";
import {
  AnimationMixer,
  type AnimationAction,
  type AnimationClip,
  type Group,
  type Object3D,
  type SkinnedMesh,
} from "three";
import { clone as cloneSkinned } from "three/examples/jsm/utils/SkeletonUtils.js";
import type { CombatBlock } from "./combatScene";

const SOLDIER_URL = "/3d_models/soldier.glb";
/** Soft visual budget — keeps relative size, labels still show real actual. */
export const MAX_VISUAL_SOLDIERS = 160;
const DETAIL_MESH_RE = /eyes|gloves|masks|default/i;

useGLTF.preload(SOLDIER_URL, true);

function findClip(animations: AnimationClip[], candidates: string[]) {
  const lowered = candidates.map((name) => name.toLowerCase());
  return (
    animations.find((clip) =>
      lowered.some(
        (name) =>
          clip.name.toLowerCase() === name ||
          clip.name.toLowerCase().includes(name),
      ),
    ) ?? null
  );
}

/** Exact headcount from combat statistics (cards / labels). */
export function soldierCountForBlock(actual: number) {
  if (!Number.isFinite(actual) || actual <= 0) return 0;
  return Math.round(actual);
}

/** Rendered figures — proportional to stats, capped for FPS. */
export function visualSoldierCountForBlock(
  actual: number,
  totalActual: number,
) {
  const exact = soldierCountForBlock(actual);
  if (exact <= 0) return 0;
  if (totalActual <= MAX_VISUAL_SOLDIERS) return exact;
  return Math.max(1, Math.round((exact / totalActual) * MAX_VISUAL_SOLDIERS));
}

function padSizeForStaff(staff: number) {
  return Math.max(0.7, Math.min(1.35, Math.sqrt(staff / 80)));
}

function layoutInPad(count: number, padSize: number) {
  if (count <= 0) return [] as Array<{ x: number; z: number; rot: number }>;

  const cols = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / cols);
  const usable = padSize * 0.78;
  const stepX = usable / Math.max(1, cols);
  const stepZ = usable / Math.max(1, rows);
  const items: Array<{ x: number; z: number; rot: number }> = [];

  for (let index = 0; index < count; index += 1) {
    const col = index % cols;
    const row = Math.floor(index / cols);
    items.push({
      x: (col - (cols - 1) / 2) * stepX,
      z: (row - (rows - 1) / 2) * stepZ,
      rot: ((index % 7) - 3) * 0.05,
    });
  }

  return items;
}

type SoldierSlot = {
  worldX: number;
  worldZ: number;
  rot: number;
  active: boolean;
  scale: number;
};

type PooledSoldier = {
  root: Object3D;
  mixer: AnimationMixer;
  idle: AnimationAction | null;
  walk: AnimationAction | null;
  active: boolean;
};

function modelHeight(scene: Object3D) {
  let height = 1.8;
  scene.updateMatrixWorld(true);
  scene.traverse((object) => {
    const mesh = object as SkinnedMesh;
    if (!mesh.isMesh || !mesh.geometry) return;
    mesh.geometry.computeBoundingBox();
    const box = mesh.geometry.boundingBox;
    if (!box) return;
    height = Math.max(height, box.max.y - box.min.y);
  });
  return height;
}

function buildSlots(
  blocks: CombatBlock[],
  selectedLabel: string | undefined,
  baseScale: number,
): SoldierSlot[] {
  const totalActual = blocks.reduce(
    (sum, block) => sum + soldierCountForBlock(block.actual),
    0,
  );
  const slots: SoldierSlot[] = [];

  for (const block of blocks) {
    const count = visualSoldierCountForBlock(block.actual, totalActual);
    const pad = padSizeForStaff(block.staff);
    const layout = layoutInPad(count, pad);
    const cols = Math.max(1, Math.ceil(Math.sqrt(count)));
    const cell = (pad * 0.78) / cols;
    // Prefer readable size even if figures slightly overlap in dense pads.
    const scale = Math.max(0.14, Math.min(baseScale, cell * 2.45));
    const active = block.label === selectedLabel;

    layout.forEach((item) => {
      slots.push({
        worldX: block.x + item.x,
        worldZ: block.z + item.z,
        rot: item.rot,
        active,
        scale,
      });
    });
  }

  return slots;
}

function prepareClone(template: Object3D) {
  const root = cloneSkinned(template);
  root.traverse((object) => {
    object.frustumCulled = true;
    const mesh = object as SkinnedMesh;
    if (!mesh.isMesh) return;
    mesh.castShadow = false;
    mesh.receiveShadow = false;
    if (DETAIL_MESH_RE.test(mesh.name)) mesh.visible = false;
  });
  return root;
}

function playIdle(soldier: PooledSoldier) {
  soldier.walk?.stop();
  if (!soldier.idle) return;
  soldier.idle.reset().setEffectiveWeight(1).play();
  // Freeze on a stable idle frame for inactive crowd.
  soldier.mixer.setTime(0.35);
  soldier.mixer.update(0);
  soldier.idle.paused = true;
}

function playWalk(soldier: PooledSoldier) {
  soldier.idle?.stop();
  if (!soldier.walk) return;
  soldier.walk.paused = false;
  soldier.walk.reset().setEffectiveWeight(1).play();
}

function CombatSkinnedSoldiers({
  blocks,
  selectedLabel,
}: {
  blocks: CombatBlock[];
  selectedLabel?: string;
}) {
  const gltf = useGLTF(SOLDIER_URL, true);
  const scene = gltf.scene;
  const animations = gltf.animations as AnimationClip[];
  const groupRef = useRef<Group>(null);
  const poolRef = useRef<PooledSoldier[]>([]);
  const slotsRef = useRef<SoldierSlot[]>([]);

  const idleClip = useMemo(
    () => findClip(animations, ["idle", "stand", "breathing"]),
    [animations],
  );
  const walkClip = useMemo(
    () => findClip(animations, ["walk", "run", "move"]),
    [animations],
  );

  const baseScale = useMemo(() => {
    const height = modelHeight(scene);
    return Math.min(0.32, 0.48 / Math.max(0.5, height));
  }, [scene]);

  const slots = useMemo(
    () => buildSlots(blocks, selectedLabel, baseScale),
    [baseScale, blocks, selectedLabel],
  );
  slotsRef.current = slots;

  useEffect(() => {
    const group = groupRef.current;
    if (!group) return;

    const pool = poolRef.current;
    while (pool.length < slots.length) {
      const root = prepareClone(scene);
      const mixer = new AnimationMixer(root);
      const soldier: PooledSoldier = {
        root,
        mixer,
        idle: idleClip ? mixer.clipAction(idleClip) : null,
        walk: walkClip ? mixer.clipAction(walkClip) : null,
        active: false,
      };
      playIdle(soldier);
      group.add(root);
      pool.push(soldier);
    }

    for (let index = 0; index < pool.length; index += 1) {
      const soldier = pool[index];
      const slot = slots[index];
      if (!soldier) continue;

      if (!slot) {
        soldier.root.visible = false;
        continue;
      }

      soldier.root.visible = true;
      soldier.root.position.set(slot.worldX, 0.02, slot.worldZ);
      soldier.root.rotation.set(0, slot.rot + (slot.active ? 0.35 : 0), 0);
      soldier.root.scale.setScalar(slot.scale);

      if (slot.active !== soldier.active) {
        if (slot.active) playWalk(soldier);
        else playIdle(soldier);
        soldier.active = slot.active;
      }
    }
  }, [idleClip, scene, slots, walkClip]);

  useFrame((_, delta) => {
    const pool = poolRef.current;
    const currentSlots = slotsRef.current;
    const count = currentSlots.length;

    for (let index = 0; index < count; index += 1) {
      const soldier = pool[index];
      const slot = currentSlots[index];
      if (!soldier || !slot?.active) continue;
      soldier.mixer.update(delta);
    }
  });

  return <group ref={groupRef} />;
}

export function CombatSoldierCrowd({
  blocks,
  selectedLabel,
}: {
  blocks: CombatBlock[];
  selectedLabel?: string;
}) {
  const total = useMemo(
    () =>
      blocks.reduce(
        (sum, block) => sum + soldierCountForBlock(block.actual),
        0,
      ),
    [blocks],
  );

  if (total <= 0) return null;

  return (
    <CombatSkinnedSoldiers blocks={blocks} selectedLabel={selectedLabel} />
  );
}
