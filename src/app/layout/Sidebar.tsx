import { type ReactNode } from "react";
import { Box, IconButton, Stack, Typography } from "@/components/sci/SciPrimitives";
import { AnalyticsOutlinedIcon } from "@/components/sci/icons";
import { ArticleOutlinedIcon } from "@/components/sci/icons";
import { DashboardOutlinedIcon } from "@/components/sci/icons";
import { FormatListBulletedOutlinedIcon } from "@/components/sci/icons";
import { GroupsOutlinedIcon } from "@/components/sci/icons";
import { LogoutOutlinedIcon } from "@/components/sci/icons";
import { MenuOutlinedIcon } from "@/components/sci/icons";
import { MenuOpenOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { SettingsOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import { SyncAltOutlinedIcon } from "@/components/sci/icons";
import { TableChartOutlinedIcon } from "@/components/sci/icons";
import { useAuth } from "../../auth/AuthProvider";
import { isUserAllowedPage, type AppPage } from "../navigation";

const navItems: Array<{ label: string; page?: AppPage; icon: ReactNode; adminOnly?: boolean }> = [
  { label: "Огляд", page: "overview", icon: <DashboardOutlinedIcon /> },
  { label: "Особовий склад", page: "personnel", icon: <PersonSearchOutlinedIcon /> },
  { label: "ЕЖООС", page: "ejournal", icon: <TableChartOutlinedIcon /> },
  { label: "Заповнення Excel", page: "excelFill", icon: <SyncAltOutlinedIcon /> },
  {
    label: "Excel Lab",
    page: "excelLab",
    icon: <TableChartOutlinedIcon />,
    adminOnly: true,
  },
  { label: "БЧС", page: "bchs", icon: <ShieldOutlinedIcon /> },
  { label: "Аналітика", page: "analytics", icon: <AnalyticsOutlinedIcon /> },
  { label: "Анкетні дані", page: "anketaData", icon: <TableChartOutlinedIcon /> },
  { label: "Анкета", page: "preAnketa", icon: <ArticleOutlinedIcon /> },
  { label: "Соц. паспорт", page: "socPassport", icon: <TableChartOutlinedIcon /> },
  { label: "Документи", page: "documents", icon: <ArticleOutlinedIcon /> },
  {
    label: "Мої завдання",
    page: "workTasks",
    icon: <FormatListBulletedOutlinedIcon />,
  },
  {
    label: "Записи документів",
    page: "documentSettings",
    icon: <SettingsOutlinedIcon />,
  },
  {
    label: "Мій профіль",
    page: "profile",
    icon: <PersonOutlinedIcon />,
  },
  {
    label: "Доступи",
    page: "usersAccess",
    icon: <GroupsOutlinedIcon />,
    adminOnly: true,
  },
];

export const APP_PAGE_LABELS: Record<AppPage, string> = Object.fromEntries(
  navItems
    .filter((item): item is { label: string; page: AppPage; icon: ReactNode; adminOnly?: boolean } =>
      Boolean(item.page),
    )
    .map((item) => [item.page, item.label]),
) as Record<AppPage, string>;

export function Sidebar({
  activePage,
  onPageChange,
  collapsed,
  onCollapsedChange,
  mobileOpen = false,
  onMobileClose,
}: {
  activePage: AppPage;
  onPageChange: (page: AppPage) => void;
  collapsed: boolean;
  onCollapsedChange: (collapsed: boolean) => void;
  mobileOpen?: boolean;
  onMobileClose?: () => void;
}) {
  const { user, isAdmin, canEdit, logout } = useAuth();

  const selectPage = (page: AppPage) => {
    onPageChange(page);
    onMobileClose?.();
  };

  const visibleNav = navItems.filter((item) => {
    if (item.adminOnly && !isAdmin) return false;
    if (isAdmin) return true;
    // Non-admins: only the allowlisted pages (hide placeholders without a page).
    return Boolean(item.page && isUserAllowedPage(item.page));
  });
  const identityLabel =
    user?.nickname || user?.displayName || user?.email || "Користувач";
  const roleLabel = isAdmin
    ? "Адміністратор"
    : canEdit
      ? "Редактор"
      : "Лише перегляд";
  const showRoleLabel =
    identityLabel.trim().toLocaleLowerCase("uk-UA") !==
    roleLabel.toLocaleLowerCase("uk-UA");

  return (
    <aside
      id="app-sidebar"
      className={`sidebar${collapsed ? " collapsed" : ""}${mobileOpen ? " mobile-open" : ""}`}
    >
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
            className="sidebar-toggle sidebar-toggle-desktop"
            size="small"
            onClick={() => onCollapsedChange(!collapsed)}
          >
            {collapsed ? (
              <MenuOutlinedIcon fontSize="small" />
            ) : (
              <MenuOpenOutlinedIcon fontSize="small" />
            )}
          </IconButton>
          <IconButton
            aria-label="Закрити меню"
            className="sidebar-toggle sidebar-toggle-mobile-close"
            size="small"
            onClick={() => onMobileClose?.()}
          >
            <MenuOpenOutlinedIcon fontSize="small" />
          </IconButton>
        </Stack>
        <Box className="brand-status" sx={{ mt: 2 }}>
          <span className="sidebar-user-avatar">
            {user?.photoData ? (
              <img src={user.photoData} alt="" />
            ) : (
              <PersonOutlinedIcon aria-hidden />
            )}
            <span className="status-dot" aria-hidden />
          </span>
          <span className="sidebar-user-copy">
            <Typography variant="caption" color="text.secondary">
              {identityLabel}
            </Typography>
            {showRoleLabel ? (
              <Typography
                variant="caption"
                color="text.secondary"
                sx={{ display: "block" }}
              >
                {roleLabel}
              </Typography>
            ) : null}
          </span>
        </Box>
      </div>

      <nav className="nav-list" aria-label="Головна навігація">
        {visibleNav.map((item) => (
          <button
            className={`nav-item${item.page === activePage ? " active" : ""}`}
            disabled={!item.page}
            key={item.label}
            title={item.label}
            aria-label={item.label}
            aria-current={item.page === activePage ? "page" : undefined}
            type="button"
            onClick={() => item.page && selectPage(item.page)}
          >
            {item.icon}
            <Typography className="nav-item-label" variant="body2">
              {item.label}
            </Typography>
          </button>
        ))}
      </nav>

      <div className="sidebar-footer">
        <button
          className="nav-item"
          type="button"
          title="Вийти"
          aria-label="Вийти"
          onClick={() => {
            logout();
            onMobileClose?.();
          }}
        >
          <LogoutOutlinedIcon />
          <Typography className="nav-item-label" variant="body2">
            Вийти
          </Typography>
        </button>
      </div>
    </aside>
  );
}
