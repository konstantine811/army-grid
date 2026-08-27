import { Box, Typography } from "@/components/sci/SciPrimitives";
import { EjoosWorkspacePage } from "./EjoosWorkspacePage";

export function EjournalPage() {
  return (
    <main className="main-panel">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            ЕЖООС
          </Typography>
          <Typography variant="body2" className="ejoos-muted">
            Імпорт 1ПБ і ЕЖООС · перегляд аркушів журналу
          </Typography>
        </Box>
      </header>

      <EjoosWorkspacePage />
    </main>
  );
}
