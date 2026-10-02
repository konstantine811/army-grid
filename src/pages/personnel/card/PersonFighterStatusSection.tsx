import { CalendarMonthOutlinedIcon } from "@/components/sci/icons";
import { InfoOutlinedIcon } from "@/components/sci/icons";
import { LoginOutlinedIcon } from "@/components/sci/icons";
import { LogoutOutlinedIcon } from "@/components/sci/icons";
import { PushPinOutlinedIcon } from "@/components/sci/icons";
import { WarningAmberOutlinedIcon } from "@/components/sci/icons";
import { getFighterStatusFieldTone } from "../fighterStatusImport";
import type { FighterStatusCardField } from "./personCardTypes";

const STATUS_ICON = {
  exit: LogoutOutlinedIcon,
  return: LoginOutlinedIcon,
  entry: CalendarMonthOutlinedIcon,
  days: InfoOutlinedIcon,
  status: WarningAmberOutlinedIcon,
  direction: PushPinOutlinedIcon,
} as const;

export function PersonFighterStatusSection({
  fields,
}: {
  fields: FighterStatusCardField[];
}) {
  if (!fields.length) return null;
  return (
    <div className="person-edit-section">
      <div className="panel-heading">Статус бійців</div>
      <div className="person-roster-grid">
        {fields.map((field) => {
          const tone = getFighterStatusFieldTone(field.key);
          const Icon = tone ? STATUS_ICON[tone] : null;
          return (
            <span
              key={field.key}
              className={tone ? `person-roster-tile is-${tone}` : "person-roster-tile"}
            >
              <strong>
                {Icon ? <Icon fontSize="small" aria-hidden /> : null}
                {field.label.replace(/^Статус бійців · /, "")}
              </strong>
              <em>{field.value}</em>
            </span>
          );
        })}
      </div>
    </div>
  );
}
