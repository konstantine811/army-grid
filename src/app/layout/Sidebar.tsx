import { useEffect, useState, type ReactNode } from "react";
import { Box, IconButton, Stack, Typography } from "@/components/sci/SciPrimitives";
import { AnalyticsOutlinedIcon } from "@/components/sci/icons";
import { ArticleOutlinedIcon } from "@/components/sci/icons";
import { DashboardOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { GridViewOutlinedIcon } from "@/components/sci/icons";
import { HelpOutlineOutlinedIcon } from "@/components/sci/icons";
import { MenuOutlinedIcon } from "@/components/sci/icons";
import { MenuOpenOutlinedIcon } from "@/components/sci/icons";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { SettingsOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import { SyncAltOutlinedIcon } from "@/components/sci/icons";
import { TableChartOutlinedIcon } from "@/components/sci/icons";
import type { AppPage } from "../navigation";

const SIDEBAR_COLLAPSED_KEY = "army-grid.sidebar-collapsed";

const navItems: Array<{ label: string; page?: AppPage; icon: ReactNode }> = [
  { label: "Огляд", page: "overview", icon: <DashboardOutlinedIcon /> },
  { label: "Особовий склад", page: "personnel", icon: <PersonSearchOutlinedIcon /> },
  { label: "Статуси", icon: <GridViewOutlinedIcon /> },
  { label: "ЕЖООС", page: "ejournal", icon: <TableChartOutlinedIcon /> },
  { label: "Заповнення Excel", page: "excelFill", icon: <SyncAltOutlinedIcon /> },
  { label: "БЧС", page: "bchs", icon: <ShieldOutlinedIcon /> },
  { label: "Аналітика", page: "analytics", icon: <AnalyticsOutlinedIcon /> },
  { label: "Імпорт", page: "import", icon: <FileDownloadOutlinedIcon /> },
  { label: "Парсинг анкет", page: "questionnaireParser", icon: <ArticleOutlinedIcon /> },
  { label: "Документи", page: "documents", icon: <ArticleOutlinedIcon /> },
];

export function Sidebar({
  activePage,
  onPageChange,
}: {
  activePage: AppPage;
  onPageChange: (page: AppPage) => void;
}) {
  const [collapsed, setCollapsed] = useState(() => {
    try {
      return window.localStorage.getItem(SIDEBAR_COLLAPSED_KEY) === "1";
    } catch {
      return false;
    }
  });

  useEffect(() => {
    try {
      window.localStorage.setItem(SIDEBAR_COLLAPSED_KEY, collapsed ? "1" : "0");
    } catch {
      // Ignore storage write failures (private mode, etc.).
    }
  }, [collapsed]);

  return (
    <aside className={`sidebar${collapsed ? " collapsed" : ""}`}>
      <div className="brand-block">
        <Stack
          direction="row"
          sx={{
            alignItems: collapsed ? "center" : "flex-start",
            justifyContent: collapsed ? "center" : "space-between",
          }}
        >
          <Box className="brand-copy">
            <Typography variant="overline" color="text.secondary">
              Система управління
            </Typography>
            <Typography variant="body2">статусами</Typography>
          </Box>
          <IconButton
            aria-label={collapsed ? "Розгорнути меню" : "Згорнути меню"}
            aria-expanded={!collapsed}
            className="sidebar-toggle"
            size="small"
            onClick={() => setCollapsed((value) => !value)}
          >
            {collapsed ? (
              <MenuOutlinedIcon fontSize="small" />
            ) : (
              <MenuOpenOutlinedIcon fontSize="small" />
            )}
          </IconButton>
        </Stack>
        <Box className="brand-status" sx={{ mt: 2 }}>
          <Typography variant="caption" color="text.secondary">
            <span className="status-dot" />
            Оператор
          </Typography>
          <Typography
            variant="caption"
            color="text.secondary"
            sx={{ display: "block", ml: 2.2 }}
          >
            Онлайн
          </Typography>
        </Box>
      </div>

      <nav className="nav-list" aria-label="Головна навігація">
        {navItems.map((item) => (
          <button
            className={`nav-item${item.page === activePage ? " active" : ""}`}
            disabled={!item.page}
            key={item.label}
            title={item.label}
            aria-label={item.label}
            aria-current={item.page === activePage ? "page" : undefined}
            type="button"
            onClick={() => item.page && onPageChange(item.page)}
          >
            {item.icon}
            <Typography className="nav-item-label" variant="body2">
              {item.label}
            </Typography>
          </button>
        ))}
      </nav>

      <div className="sidebar-footer">
        <div className="nav-item" title="Налаштування" aria-label="Налаштування">
          <SettingsOutlinedIcon />
          <Typography className="nav-item-label" variant="body2">
            Налаштування
          </Typography>
        </div>
        <div className="nav-item" title="Довідка" aria-label="Довідка">
          <HelpOutlineOutlinedIcon />
          <Typography className="nav-item-label" variant="body2">
            Довідка
          </Typography>
        </div>
      </div>
    </aside>
  );
}
