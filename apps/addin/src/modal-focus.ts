import { useEffect, useRef } from "react";

/** Keep keyboard navigation inside the active approval/question and restore its origin on close. */
export function useModalFocus(active: boolean, dismiss: () => void) {
  const container = useRef<HTMLElement>(null);
  const onDismiss = useRef(dismiss);
  onDismiss.current = dismiss;
  useEffect(() => {
    if (!active || !container.current) return;
    const dialog = container.current;
    const previous = document.activeElement;
    const controls = () => [
      ...dialog.querySelectorAll<HTMLElement>(
        'button:not(:disabled), input:not(:disabled), textarea:not(:disabled), select:not(:disabled), a[href], [tabindex="0"]',
      ),
    ];
    controls()[0]?.focus();
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onDismiss.current();
        return;
      }
      if (event.key !== "Tab") return;
      const items = controls();
      const first = items[0];
      const last = items.at(-1);
      if (
        !dialog.contains(document.activeElement) ||
        (event.shiftKey ? document.activeElement === first : document.activeElement === last)
      ) {
        event.preventDefault();
        (event.shiftKey ? last : first)?.focus();
      }
    };
    document.addEventListener("keydown", onKey, true);
    return () => {
      document.removeEventListener("keydown", onKey, true);
      if (previous instanceof HTMLElement && previous.isConnected) previous.focus();
    };
  }, [active]);
  return container;
}
