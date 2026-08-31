import { useEffect } from "react";

const SCROLL_SELECTOR = [
  ".personnel-list",
  ".overview-table-body",
  ".person-card-scroll",
  ".sci-data-table-wrap",
  ".document-page-preview",
  ".document-fields",
  ".documents-journal-table-body",
  ".ubd-preview-viewport",
  ".word-document-preview",
  ".ejournal-event-list",
].join(",");

const LAYOUT_SELECTOR = ".app-shell, .sidebar, .personnel-layout, .personnel-list-panel";

type ScrollbarParts = {
  rail: HTMLDivElement;
  thumb: HTMLDivElement;
  onScroll: () => void;
  onPointerDown: (event: PointerEvent) => void;
  onRailPointerDown: (event: PointerEvent) => void;
};

const isLayoutTransitionTarget = (target: EventTarget | null) => {
  if (!(target instanceof HTMLElement)) return false;
  return (
    target.classList.contains("app-shell") ||
    target.classList.contains("sidebar")
  );
};

export function SciScrollbars() {
  useEffect(() => {
    const bars = new Map<HTMLElement, ScrollbarParts>();
    let frame = 0;
    let transitionFrame = 0;

    const scheduleUpdate = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(updateAll);
    };

    const updateBar = (element: HTMLElement, parts: ScrollbarParts) => {
      const rect = element.getBoundingClientRect();
      const canScroll = element.scrollHeight > element.clientHeight + 1;
      const visible =
        canScroll &&
        rect.width > 0 &&
        rect.height > 0 &&
        rect.bottom > 0 &&
        rect.top < window.innerHeight;

      parts.rail.style.display = visible ? "block" : "none";
      if (!visible) return;

      const railWidth = 10;
      const railInset = 1;
      const usableHeight = Math.max(1, rect.height);
      const thumbHeight = Math.max(
        28,
        Math.min(
          usableHeight - 8,
          Math.round((element.clientHeight / element.scrollHeight) * usableHeight),
        ),
      );
      const maxThumbTop = Math.max(0, usableHeight - thumbHeight);
      const maxScrollTop = Math.max(1, element.scrollHeight - element.clientHeight);
      const thumbTop = Math.round((element.scrollTop / maxScrollTop) * maxThumbTop);

      parts.rail.style.left = `${Math.round(rect.right - railWidth - railInset)}px`;
      parts.rail.style.top = `${Math.round(rect.top)}px`;
      parts.rail.style.height = `${Math.round(rect.height)}px`;
      parts.rail.style.width = `${railWidth}px`;
      parts.thumb.style.top = `${thumbTop}px`;
      parts.thumb.style.height = `${thumbHeight}px`;
    };

    function updateAll() {
      document.querySelectorAll<HTMLElement>(SCROLL_SELECTOR).forEach((element) => {
        if (!bars.has(element)) {
          mountScrollbar(element);
        }
      });
      document.querySelectorAll<HTMLElement>(LAYOUT_SELECTOR).forEach((element) => {
        resizeObserver.observe(element);
      });

      for (const [element, parts] of bars) {
        if (!document.body.contains(element)) {
          unmountScrollbar(element, parts);
          bars.delete(element);
          continue;
        }
        updateBar(element, parts);
      }
    }

    const resizeObserver = new ResizeObserver(scheduleUpdate);

    const mountScrollbar = (element: HTMLElement) => {
      element.classList.add("sci-custom-scroll-target");
      resizeObserver.observe(element);

      const rail = document.createElement("div");
      rail.className = "sci-custom-scrollbar";
      const thumb = document.createElement("div");
      thumb.className = "sci-custom-scrollbar-thumb";
      rail.appendChild(thumb);
      document.body.appendChild(rail);

      const onScroll = () => scheduleUpdate();
      const onPointerDown = (event: PointerEvent) => {
        event.preventDefault();
        event.stopPropagation();
        const startY = event.clientY;
        const startScrollTop = element.scrollTop;
        const startHeight = thumb.offsetHeight;
        const railHeight = Math.max(1, rail.offsetHeight - startHeight);
        const scrollableHeight = Math.max(1, element.scrollHeight - element.clientHeight);

        const onPointerMove = (moveEvent: PointerEvent) => {
          const delta = moveEvent.clientY - startY;
          element.scrollTop = startScrollTop + (delta / railHeight) * scrollableHeight;
        };
        const onPointerUp = () => {
          window.removeEventListener("pointermove", onPointerMove);
          window.removeEventListener("pointerup", onPointerUp);
        };

        window.addEventListener("pointermove", onPointerMove);
        window.addEventListener("pointerup", onPointerUp);
      };
      const onRailPointerDown = (event: PointerEvent) => {
        if (event.target === thumb) return;
        const railRect = rail.getBoundingClientRect();
        const targetRatio = (event.clientY - railRect.top) / Math.max(1, railRect.height);
        element.scrollTop = targetRatio * (element.scrollHeight - element.clientHeight);
      };

      element.addEventListener("scroll", onScroll, { passive: true });
      thumb.addEventListener("pointerdown", onPointerDown);
      rail.addEventListener("pointerdown", onRailPointerDown);

      bars.set(element, { rail, thumb, onScroll, onPointerDown, onRailPointerDown });
    };

    const unmountScrollbar = (element: HTMLElement, parts: ScrollbarParts) => {
      element.classList.remove("sci-custom-scroll-target");
      resizeObserver.unobserve(element);
      element.removeEventListener("scroll", parts.onScroll);
      parts.thumb.removeEventListener("pointerdown", parts.onPointerDown);
      parts.rail.removeEventListener("pointerdown", parts.onRailPointerDown);
      parts.rail.remove();
    };

    const stopTransitionLoop = () => {
      window.cancelAnimationFrame(transitionFrame);
      transitionFrame = 0;
    };

    const startTransitionLoop = () => {
      if (transitionFrame) return;
      const tick = () => {
        updateAll();
        transitionFrame = window.requestAnimationFrame(tick);
      };
      transitionFrame = window.requestAnimationFrame(tick);
    };

    const onLayoutTransition = (event: TransitionEvent) => {
      if (!isLayoutTransitionTarget(event.target)) return;
      if (event.type === "transitionend" || event.type === "transitioncancel") {
        stopTransitionLoop();
        scheduleUpdate();
        return;
      }
      startTransitionLoop();
    };

    const observer = new MutationObserver(scheduleUpdate);
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class"],
    });
    window.addEventListener("resize", scheduleUpdate);
    window.addEventListener("scroll", scheduleUpdate, true);
    window.addEventListener("transitionstart", onLayoutTransition, true);
    window.addEventListener("transitionend", onLayoutTransition, true);
    window.addEventListener("transitioncancel", onLayoutTransition, true);
    scheduleUpdate();

    return () => {
      window.cancelAnimationFrame(frame);
      stopTransitionLoop();
      observer.disconnect();
      resizeObserver.disconnect();
      window.removeEventListener("resize", scheduleUpdate);
      window.removeEventListener("scroll", scheduleUpdate, true);
      window.removeEventListener("transitionstart", onLayoutTransition, true);
      window.removeEventListener("transitionend", onLayoutTransition, true);
      window.removeEventListener("transitioncancel", onLayoutTransition, true);
      for (const [element, parts] of bars) {
        unmountScrollbar(element, parts);
      }
      bars.clear();
    };
  }, []);

  return null;
}
