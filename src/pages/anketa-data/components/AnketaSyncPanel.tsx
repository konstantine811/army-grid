import { Button, TextField } from "@/components/sci/SciPrimitives";
import { ANKETA_APPS_SCRIPT_TEMPLATE } from "../anketaGaps";
import { writeAnketaAppsScriptUrl } from "../anketaGaps";

type AnketaSyncPanelProps = {
  appsScriptUrl: string;
  onAppsScriptUrlChange: (value: string) => void;
  expanded: boolean;
  onToggleExpanded: () => void;
  showSyncHelp: boolean;
  onToggleSyncHelp: () => void;
};

export function AnketaSyncPanel({
  appsScriptUrl,
  onAppsScriptUrlChange,
  expanded,
  onToggleExpanded,
  showSyncHelp,
  onToggleSyncHelp,
}: AnketaSyncPanelProps) {
  return (
    <section
      className={`analytics-panel anketa-sync-panel${expanded ? " is-expanded" : " is-collapsed"}`}
    >
      <div className="panel-heading anketa-sync-heading">
        <button
          type="button"
          className="anketa-sync-toggle"
          aria-expanded={expanded}
          onClick={onToggleExpanded}
        >
          <span>Синхронізація з Google</span>
          <span className="anketa-sync-toggle-hint">
            {expanded ? "згорнути" : "розгорнути"}
          </span>
        </button>
        {expanded ? (
          <Button size="small" variant="text" onClick={onToggleSyncHelp}>
            {showSyncHelp ? "Сховати" : "Як підключити"}
          </Button>
        ) : null}
      </div>
      {expanded ? (
        <div className="anketa-sync-body">
          <TextField
            size="small"
            fullWidth
            label="URL Google Apps Script Web App"
            placeholder="https://script.google.com/macros/s/.../exec"
            value={appsScriptUrl}
            onChange={(event) => onAppsScriptUrlChange(event.target.value)}
            onBlur={() => writeAnketaAppsScriptUrl(appsScriptUrl)}
          />
          {!showSyncHelp ? (
            <p className="anketa-sync-hint">
              Правки в IndexedDB не зникають після «Оновити з Google». Apps
              Script — додатково в Google Sheet.
            </p>
          ) : (
            <pre className="anketa-apps-script-template">
              {ANKETA_APPS_SCRIPT_TEMPLATE}
            </pre>
          )}
        </div>
      ) : (
        <p className="anketa-sync-collapsed-hint">
          {appsScriptUrl.trim()
            ? "Apps Script підключено · натисніть, щоб змінити URL"
            : "Не обов’язково для локальної роботи · розгорніть за потреби"}
        </p>
      )}
    </section>
  );
}
