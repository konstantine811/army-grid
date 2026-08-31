import { ANKETA_MISSING_VALUE_PRESETS } from "./anketaGaps";

export function AnketaMissingPresets({
  onPick,
}: {
  onPick: (value: string) => void;
}) {
  return (
    <div className="anketa-missing-presets" role="listbox">
      <div className="anketa-missing-presets-title">Статус пропуску</div>
      <div className="anketa-missing-presets-list">
        {ANKETA_MISSING_VALUE_PRESETS.map((preset) => (
          <button
            key={preset}
            type="button"
            role="option"
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => onPick(preset)}
          >
            {preset}
          </button>
        ))}
      </div>
    </div>
  );
}
