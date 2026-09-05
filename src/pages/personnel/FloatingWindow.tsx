import { useEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { Button } from "@/components/sci/SciPrimitives";

export type FloatingPlacement = "center" | "left" | "right" | "top";

type FloatingWindowProps = {
  open: boolean;
  title: string;
  subtitle?: string;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  placement?: FloatingPlacement;
  defaultWidth?: number;
  defaultHeight?: number;
  minWidth?: number;
  minHeight?: number;
  className?: string;
  bodyClassName?: string;
};

let floatingZCounter = 10040;

export function claimFloatingZIndex() {
  floatingZCounter += 1;
  return floatingZCounter;
}

function resolvePlacement(
  placement: FloatingPlacement,
  width: number,
  height: number,
) {
  const maxX = Math.max(8, window.innerWidth - width - 8);
  if (placement === "left") {
    return { x: 8, y: 8 };
  }
  if (placement === "right") {
    return { x: maxX, y: 8 };
  }
  if (placement === "top") {
    return {
      x: Math.max(8, Math.round((window.innerWidth - width) / 2)),
      y: 8,
    };
  }
  return {
    x: Math.max(8, Math.round((window.innerWidth - width) / 2)),
    y: Math.max(8, Math.round((window.innerHeight - height) / 2)),
  };
}

export function FloatingWindow({
  open,
  title,
  subtitle = "перетягніть за шапку · розмір за кут · клік — наперед",
  onClose,
  children,
  footer,
  placement = "center",
  defaultWidth = 720,
  defaultHeight = 640,
  minWidth = 360,
  minHeight = 320,
  className = "",
  bodyClassName = "",
}: FloatingWindowProps) {
  const [pos, setPos] = useState({ x: 40, y: 40 });
  const [size, setSize] = useState({ w: defaultWidth, h: defaultHeight });
  const [zIndex, setZIndex] = useState(() => claimFloatingZIndex());
  const [interacting, setInteracting] = useState(false);
  const dragOffset = useRef({ x: 0, y: 0 });

  const bringToFront = () => setZIndex(claimFloatingZIndex());

  useEffect(() => {
    if (!open) return;
    const w = Math.min(
      defaultWidth,
      Math.max(minWidth, window.innerWidth - 32),
    );
    const h = Math.min(
      defaultHeight,
      Math.max(minHeight, window.innerHeight - 32),
    );
    setSize({ w, h });
    setPos(resolvePlacement(placement, w, h));
    setZIndex(claimFloatingZIndex());
  }, [open, placement, defaultWidth, defaultHeight, minWidth, minHeight]);

  if (!open || typeof document === "undefined") return null;

  return createPortal(
    <div
      className={`floating-window${interacting ? " is-interacting" : ""}${className ? ` ${className}` : ""}`}
      style={{
        left: pos.x,
        top: pos.y,
        width: size.w,
        height: size.h,
        zIndex,
      }}
      role="dialog"
      aria-label={title}
      onPointerDown={bringToFront}
    >
      <header
        className="floating-window-title"
        onPointerDown={(event) => {
          if (event.button !== 0) return;
          const target = event.target as HTMLElement;
          if (target.closest("button")) return;
          bringToFront();
          dragOffset.current = {
            x: event.clientX - pos.x,
            y: event.clientY - pos.y,
          };
          setInteracting(true);

          const onMove = (moveEvent: PointerEvent) => {
            const maxX = window.innerWidth - 120;
            const maxY = window.innerHeight - 48;
            setPos({
              x: Math.min(
                maxX,
                Math.max(0, moveEvent.clientX - dragOffset.current.x),
              ),
              y: Math.min(
                maxY,
                Math.max(0, moveEvent.clientY - dragOffset.current.y),
              ),
            });
          };
          const onUp = () => {
            setInteracting(false);
            window.removeEventListener("pointermove", onMove);
            window.removeEventListener("pointerup", onUp);
          };
          window.addEventListener("pointermove", onMove);
          window.addEventListener("pointerup", onUp);
        }}
      >
        <div className="floating-window-title-text">
          <strong>{title}</strong>
          {subtitle ? <span>{subtitle}</span> : null}
        </div>
        <Button size="small" variant="outlined" onClick={onClose}>
          ✕
        </Button>
      </header>

      <div className={`floating-window-body${bodyClassName ? ` ${bodyClassName}` : ""}`}>
        {children}
      </div>

      {footer ? <footer className="floating-window-footer">{footer}</footer> : null}

      <div
        className="floating-window-resize"
        onPointerDown={(event) => {
          if (event.button !== 0) return;
          event.preventDefault();
          event.stopPropagation();
          bringToFront();
          const startX = event.clientX;
          const startY = event.clientY;
          const startW = size.w;
          const startH = size.h;
          setInteracting(true);

          const onMove = (moveEvent: PointerEvent) => {
            const nextW = Math.min(
              window.innerWidth - pos.x - 8,
              Math.max(minWidth, startW + (moveEvent.clientX - startX)),
            );
            const nextH = Math.min(
              window.innerHeight - pos.y - 8,
              Math.max(minHeight, startH + (moveEvent.clientY - startY)),
            );
            setSize({ w: nextW, h: nextH });
          };
          const onUp = () => {
            setInteracting(false);
            window.removeEventListener("pointermove", onMove);
            window.removeEventListener("pointerup", onUp);
          };
          window.addEventListener("pointermove", onMove);
          window.addEventListener("pointerup", onUp);
        }}
      />
    </div>,
    document.body,
  );
}
