export type PersonnelAttachmentKind = "questionnaire" | "photo";

export type PersonnelAttachmentChange = {
  externalId: string;
  kind: PersonnelAttachmentKind;
  at: number;
};

const CHANNEL_NAME = "army-grid:personnel-attachments";
const STORAGE_KEY = "army-grid:personnel-attachments:v1";

let channel: BroadcastChannel | null = null;

const getChannel = () => {
  if (typeof BroadcastChannel === "undefined") return null;
  channel ??= new BroadcastChannel(CHANNEL_NAME);
  return channel;
};

/** Повідомити інші вкладки (і цю), що фото/анкету оновлено в БД. */
export const notifyPersonnelAttachmentChanged = (
  externalId: string,
  kind: PersonnelAttachmentKind,
) => {
  const trimmed = externalId.trim();
  if (!trimmed) return;

  const payload: PersonnelAttachmentChange = {
    externalId: trimmed,
    kind,
    at: Date.now(),
  };

  getChannel()?.postMessage(payload);

  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch {
    /* ignore quota / private mode */
  }
};

export const subscribePersonnelAttachmentChanges = (
  listener: (change: PersonnelAttachmentChange) => void,
) => {
  const bc = getChannel();

  const onMessage = (event: MessageEvent<PersonnelAttachmentChange>) => {
    if (event.data?.externalId) listener(event.data);
  };

  const onStorage = (event: StorageEvent) => {
    if (event.key !== STORAGE_KEY || !event.newValue) return;
    try {
      const parsed = JSON.parse(event.newValue) as PersonnelAttachmentChange;
      if (parsed.externalId) listener(parsed);
    } catch {
      /* ignore */
    }
  };

  bc?.addEventListener("message", onMessage);
  window.addEventListener("storage", onStorage);

  return () => {
    bc?.removeEventListener("message", onMessage);
    window.removeEventListener("storage", onStorage);
  };
};
