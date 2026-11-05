import { useCallback, useEffect, useState } from "react";

export function useClipboard(active: boolean) {
  const [items, setItems] = useState<string[]>([]);
  const [error, setError] = useState("");
  const remember = useCallback((value: string) => {
    if (value)
      setItems((previous) =>
        [value, ...previous.filter((item) => item !== value)].slice(0, 24),
      );
  }, []);
  const read = useCallback(async () => {
    try {
      const text = await navigator.clipboard.readText();
      remember(text);
      setError(text ? "" : "The system clipboard has no text.");
      return text;
    } catch {
      setError(
        "Clipboard access was blocked. Click Read Clipboard or paste with Ctrl+V on the page.",
      );
      return null;
    }
  }, [remember]);
  useEffect(() => {
    const pasted = (event: ClipboardEvent) =>
      remember(event.clipboardData?.getData("text/plain") ?? "");
    const copied = () => remember(window.getSelection()?.toString() ?? "");
    document.addEventListener("paste", pasted);
    document.addEventListener("copy", copied);
    document.addEventListener("cut", copied);
    return () => {
      document.removeEventListener("paste", pasted);
      document.removeEventListener("copy", copied);
      document.removeEventListener("cut", copied);
    };
  }, [remember]);
  useEffect(() => {
    if (!active) return;
    let disposed = false;
    const sync = async () => {
      if (!document.hasFocus()) return;
      try {
        const permission = await navigator.permissions.query({
          name: "clipboard-read" as PermissionName,
        });
        if (!disposed && permission.state === "granted") void read();
      } catch {
        /* Read button supplies activation on browsers without this permission. */
      }
    };
    void sync();
    window.addEventListener("focus", sync);
    navigator.clipboard?.addEventListener("clipboardchange", sync);
    return () => {
      disposed = true;
      window.removeEventListener("focus", sync);
      navigator.clipboard?.removeEventListener("clipboardchange", sync);
    };
  }, [active, read]);
  return { items, setItems, read, error };
}
