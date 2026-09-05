import { createStore } from "zustand/vanilla";

export type SharedDataEntry = {
  key: string;
  value: unknown;
  savedAt: number;
  formatVersion?: number;
};

type SharedDataState = {
  entries: Map<string, SharedDataEntry>;
  put: (entry: SharedDataEntry) => void;
  remove: (key: string) => void;
  clear: () => void;
};

/**
 * First-level application cache. It survives page unmounts without copying
 * large row arrays; IndexedDB remains the persistent second level.
 */
export const sharedDataStore = createStore<SharedDataState>((set) => ({
  entries: new Map(),
  put: (entry) =>
    set((state) => {
      const entries = new Map(state.entries);
      entries.set(entry.key, entry);
      return { entries };
    }),
  remove: (key) =>
    set((state) => {
      if (!state.entries.has(key)) return state;
      const entries = new Map(state.entries);
      entries.delete(key);
      return { entries };
    }),
  clear: () => set({ entries: new Map() }),
}));

export const getSharedDataEntry = (key: string) =>
  sharedDataStore.getState().entries.get(key);

export const setSharedDataEntry = (entry: SharedDataEntry) =>
  sharedDataStore.getState().put(entry);

export const deleteSharedDataEntry = (key: string) =>
  sharedDataStore.getState().remove(key);

export const getSharedDataKeys = () =>
  sharedDataStore.getState().entries.keys();

export const clearSharedDataStore = () =>
  sharedDataStore.getState().clear();
