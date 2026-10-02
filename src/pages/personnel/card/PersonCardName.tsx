import { useLayoutEffect, useRef } from "react";

export function PersonCardName({ name }: { name: string }) {
  const ref = useRef<HTMLHeadingElement>(null);

  useLayoutEffect(() => {
    const el = ref.current;
    const parent = el?.parentElement;
    if (!el || !parent) return;

    const fit = () => {
      el.style.fontSize = "";
      const available = parent.clientWidth;
      if (available <= 0) return;
      const width = el.scrollWidth;
      if (width <= available) return;
      const current = parseFloat(getComputedStyle(el).fontSize);
      if (!current) return;
      el.style.fontSize = `${Math.max(13, (current * available) / width)}px`;
    };

    fit();
    const observer = new ResizeObserver(fit);
    observer.observe(parent);
    return () => observer.disconnect();
  }, [name]);

  return (
    <h2
      ref={ref}
      className="sci-text sci-text-h4 person-card-name"
      title={name}
      style={{
        ["--name-len" as string]: Math.max(name.trim().length, 8),
      }}
    >
      {name}
    </h2>
  );
}
