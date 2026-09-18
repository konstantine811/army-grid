// Keep writes to each document ordered, including across page unmounts.
const pending = new Map<string, Promise<unknown>>();

export function enqueueDocumentSave<T>(id: string, save: () => Promise<T>): Promise<T> {
  const previous = pending.get(id) ?? Promise.resolve();
  const result = previous.catch(() => undefined).then(save);
  pending.set(id, result);
  const cleanup = () => {
    if (pending.get(id) === result) pending.delete(id);
  };
  void result.then(cleanup, cleanup);
  return result;
}

export async function waitForDocumentSaves(id: string): Promise<void> {
  while (pending.has(id)) await pending.get(id);
}
