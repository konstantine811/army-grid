import type { ReactNode } from "react";
import { ArticleOutlinedIcon } from "@/components/sci/icons";
import { FormatListBulletedOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";

export type PersonnelMobilePane = "list" | "card" | "side";

export function PersonnelMobileTabs({
  pane,
  canOpenPerson,
  onChange,
}: {
  pane: PersonnelMobilePane;
  canOpenPerson: boolean;
  onChange: (pane: PersonnelMobilePane) => void;
}) {
  return (
    <div
      className="personnel-mobile-tabs"
      role="tablist"
      aria-label="Розділи особового складу"
    >
      <PersonnelMobileTab
        label="Список"
        active={pane === "list"}
        icon={<FormatListBulletedOutlinedIcon fontSize="small" />}
        onClick={() => onChange("list")}
      />
      <PersonnelMobileTab
        label="Картка"
        active={pane === "card"}
        disabled={!canOpenPerson}
        icon={<PersonOutlinedIcon fontSize="small" />}
        onClick={() => canOpenPerson && onChange("card")}
      />
      <PersonnelMobileTab
        label="Дії"
        active={pane === "side"}
        disabled={!canOpenPerson}
        icon={<ArticleOutlinedIcon fontSize="small" />}
        onClick={() => canOpenPerson && onChange("side")}
      />
    </div>
  );
}

function PersonnelMobileTab({
  label,
  active,
  disabled,
  icon,
  onClick,
}: {
  label: string;
  active: boolean;
  disabled?: boolean;
  icon: ReactNode;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      role="tab"
      aria-selected={active}
      className={active ? "is-active" : undefined}
      disabled={disabled}
      onClick={onClick}
    >
      {icon}
      {label}
    </button>
  );
}
