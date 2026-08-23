import { Button, TextField } from "@/components/sci/SciPrimitives";
import { ANKETA_APPS_SCRIPT_TEMPLATE } from "../anketaGaps";
import { writeAnketaAppsScriptUrl } from "../anketaGaps";

type AnketaSyncPanelProps = {
  appsScriptUrl: string;
  onAppsScriptUrlChange: (value: string) => void;
  showSyncHelp: boolean;
  onToggleSyncHelp: () => void;
};

export function AnketaSyncPanel({
  appsScriptUrl,
  onAppsScriptUrlChange,
  showSyncHelp,
  onToggleSyncHelp,
}: AnketaSyncPanelProps) {
  return (
    <section className="analytics-panel anketa-sync-panel">
      <div className="panel-heading anketa-sync-heading">
        Синхронізація з Google
        <Button size="small" variant="text" onClick={onToggleSyncHelp}>
          {showSyncHelp ? "Сховати" : "Як підключити"}
        </Button>
      </div>
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
            Правки в IndexedDB не зникають після «Оновити з Google». Apps Script —
            додатково в Google Sheet.
          </p>
        ) : (
          <pre className="anketa-apps-script-template">
            {ANKETA_APPS_SCRIPT_TEMPLATE}
          </pre>
        )}
      </div>
    </section>
  );
}
