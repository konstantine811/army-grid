import {
  Alert,
  Box,
  Chip,
  LinearProgress,
} from "@/components/sci/SciPrimitives";
import {
  EjoosWorkspaceProvider,
  useEjoosWorkspace,
  type EjoosWorkspaceTab,
} from "./EjoosWorkspaceContext";
import { EjoosImportPanel } from "./EjoosImportPanel";
import {
  EjoosSheetViewPanel,
  type EjoosSheetKind,
} from "./EjoosSheetViewPanel";

const TABS: { id: EjoosWorkspaceTab; label: string; group?: string }[] = [
  { id: "import", label: "Імпорт", group: "flow" },
  { id: "shpo", label: "ШПО", group: "view" },
  { id: "oos", label: "ООС", group: "view" },
  { id: "excluded", label: "Виключені", group: "view" },
  { id: "tempArrivals", label: "Тимчасові прибулі", group: "view" },
  { id: "tempAbsents", label: "Тимчасові відсутні", group: "view" },
  { id: "timesheet", label: "Табель", group: "view" },
  { id: "irrevocableLosses", label: "Безповоротні втрати", group: "view" },
];

const SHEET_TABS = new Set<EjoosWorkspaceTab>([
  "shpo",
  "oos",
  "excluded",
  "tempArrivals",
  "tempAbsents",
  "timesheet",
  "irrevocableLosses",
]);

function WorkspaceBody() {
  const { tab } = useEjoosWorkspace();
  if (tab === "import") return <EjoosImportPanel />;
  if (SHEET_TABS.has(tab)) {
    return <EjoosSheetViewPanel kind={tab as EjoosSheetKind} />;
  }
  return <EjoosImportPanel />;
}

function EjoosWorkspaceInner() {
  const { tab, setTab, live, message, error, isLoading } = useEjoosWorkspace();

  return (
    <Box className="ejoos-workspace">
      <nav className="ejoos-workspace-nav" aria-label="ЕЖООС">
        {TABS.map((item, index) => {
          const prev = TABS[index - 1];
          const showDivider = prev && prev.group !== item.group;
          return (
            <span key={item.id} className="ejoos-workspace-nav-item">
              {showDivider ? (
                <span className="ejoos-workspace-nav-divider" aria-hidden />
              ) : null}
              <button
                type="button"
                className={
                  tab === item.id
                    ? "ejoos-workspace-nav-btn is-active"
                    : "ejoos-workspace-nav-btn"
                }
                onClick={() => setTab(item.id)}
              >
                <span>{item.label}</span>
              </button>
            </span>
          );
        })}
        {!live?.current ? (
          <Chip
            size="small"
            color="warning"
            label="Немає ЕЖООС у БД"
            sx={{ ml: "auto" }}
          />
        ) : null}
      </nav>

      {isLoading ? <LinearProgress color="primary" /> : null}
      {message ? (
        <Alert severity="success" variant="outlined" sx={{ mb: 1.5 }}>
          {message}
        </Alert>
      ) : null}
      {error ? (
        <Alert severity="error" variant="outlined" sx={{ mb: 1.5 }}>
          {error}
        </Alert>
      ) : null}

      <WorkspaceBody />
    </Box>
  );
}

export function EjoosWorkspacePage() {
  return (
    <EjoosWorkspaceProvider>
      <EjoosWorkspaceInner />
    </EjoosWorkspaceProvider>
  );
}
