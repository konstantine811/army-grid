const DB_NAME = "army-grid-documents";
const STORE_NAME = "templates";
const UBD_KEY = "ubd-report";

const openDb = () =>
  new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

export const defaultUbdTemplateUrl = () =>
  `${import.meta.env.BASE_URL}templates/ubd-report.docx`;

export const loadDefaultUbdTemplate = async () => {
  const response = await fetch(defaultUbdTemplateUrl());
  if (!response.ok) {
    throw new Error("Не знайшов стандартний шаблон рапорту УБД.");
  }
  return response.arrayBuffer();
};

export const loadCustomUbdTemplate = async () => {
  const db = await openDb();
  try {
    return await new Promise<ArrayBuffer | null>((resolve, reject) => {
      const request = db.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).get(UBD_KEY);
      request.onsuccess = () => {
        const value = request.result;
        resolve(value instanceof ArrayBuffer ? value : null);
      };
      request.onerror = () => reject(request.error);
    });
  } finally {
    db.close();
  }
};

export const loadUbdTemplate = async () =>
  (await loadCustomUbdTemplate()) ?? (await loadDefaultUbdTemplate());

export const saveCustomUbdTemplate = async (buffer: ArrayBuffer) => {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const request = db
        .transaction(STORE_NAME, "readwrite")
        .objectStore(STORE_NAME)
        .put(buffer, UBD_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } finally {
    db.close();
  }
};

export const clearCustomUbdTemplate = async () => {
  const db = await openDb();
  try {
    await new Promise<void>((resolve, reject) => {
      const request = db
        .transaction(STORE_NAME, "readwrite")
        .objectStore(STORE_NAME)
        .delete(UBD_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } finally {
    db.close();
  }
};

export const hasCustomUbdTemplate = async () =>
  Boolean(await loadCustomUbdTemplate());
