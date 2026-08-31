import { useEffect, useState } from "react";
import { Button, Typography } from "@/components/sci/SciPrimitives";
import { MenuOutlinedIcon } from "@/components/sci/icons";
import "./App.css";
import {
  buildDocumentRoute,
  buildPersonnelRoute,
  getCurrentRouteKey,
  getPageFromPath,
  isRetiredPagePath,
  isUserAllowedPage,
  navigateToPage,
  pagePaths,
  pushAppRoute,
  writeAreaForPage,
  type AppPage,
} from "./app/navigation";
import { Sidebar, APP_PAGE_LABELS } from "./app/layout/Sidebar";
import { AnalyticsPage } from "./pages/analytics/AnalyticsPage";
import { PersonnelPage } from "./pages/personnel/PersonnelPage";
import { DocumentsPage } from "./pages/documents/DocumentsPage";
import { storeSelectedPersonFullPosition } from "./pages/documents/zhbdCertificateReport";
import {
  getPersonFullPositionTitle,
  pickFullPositionFromPersonRow,
} from "./pages/personnel/personnelUtils";
import { EjournalPage } from "./pages/ejournal/EjournalPage";
import type { EjournalPreviewRow } from "./pages/ejournal/ejournalTypes";
import { buildPersonSummary } from "./pages/personnel/personnelUtils";
import { BchsPage } from "./pages/bchs/BchsPage";
import { OverviewPage } from "./pages/overview/OverviewPage";
import { ExcelFillPage } from "./pages/excel-fill/ExcelFillPage";
import { AnketaDataPage } from "./pages/anketa-data/AnketaDataPage";
import { SocPassportPage } from "./pages/soc-passport/SocPassportPage";
import { DocumentSignatoriesSettingsPage } from "./pages/document-settings/DocumentSignatoriesSettingsPage";
import { SciScrollbars } from "./components/sci/SciScrollbars";
import { SciLiveFeedback } from "./components/sci/SciLiveFeedback";
import { AppToastHost } from "./components/AppToastHost";
import { useAuth } from "./auth/AuthProvider";
import { AuthPage } from "./auth/AuthPage";
import { PendingAccessPage } from "./auth/PendingAccessPage";
import { useAuthScrollLock } from "./auth/useAuthScrollLock";
import { UsersAccessPage } from "./pages/users/UsersAccessPage";
import { ProfilePage } from "./pages/profile/ProfilePage";
import { WorkTasksPage } from "./pages/work-tasks/WorkTasksPage";

function App() {
  const { loading, user, canView, canEdit, canEditArea, isAdmin } = useAuth();
  useAuthScrollLock(loading || !user || !canView);
  const [activePage, setActivePage] = useState<AppPage>(() =>
    getPageFromPath(window.location.pathname),
  );
  const [routeKey, setRouteKey] = useState(getCurrentRouteKey);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  useEffect(() => {
    const handlePopState = () => {
      setActivePage(getPageFromPath(window.location.pathname));
      setRouteKey(getCurrentRouteKey());
    };

    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, []);

  useEffect(() => {
    const path = window.location.pathname;
    if (path === "/" || isRetiredPagePath(path)) {
      window.history.replaceState(
        { page: activePage },
        "",
        pagePaths[activePage],
      );
    }
  }, [activePage]);

  useEffect(() => {
    if (!mobileNavOpen) return;
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setMobileNavOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => {
      document.body.style.overflow = previous;
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [mobileNavOpen]);

  useEffect(() => {
    const media = window.matchMedia("(max-width: 980px)");
    const onChange = () => {
      if (!media.matches) setMobileNavOpen(false);
    };
    onChange();
    media.addEventListener("change", onChange);
    return () => media.removeEventListener("change", onChange);
  }, []);

  const applyRoute = (route: { page: AppPage; routeKey: string }) => {
    setActivePage(route.page);
    setRouteKey(route.routeKey);
    setMobileNavOpen(false);
  };

  const changePage = (page: AppPage) => {
    applyRoute(navigateToPage(page));
  };

  useEffect(() => {
    // Wait until session is known — during boot isAdmin is false and would
    // wrongly bounce admins from /excel-fill (etc.) to /overview on every reload.
    if (loading || !user) return;
    if (isAdmin) return;
    if (!isUserAllowedPage(activePage)) {
      applyRoute(navigateToPage("overview"));
    }
  }, [loading, user, isAdmin, activePage]);

  const openPersonnelForPerson = (target: {
    rowId: string;
    externalId: string;
  }) => {
    window.localStorage.setItem(
      "army-grid:focus-personnel",
      JSON.stringify(target),
    );
    applyRoute(
      pushAppRoute(
        buildPersonnelRoute(target),
        "personnel",
      ),
    );
  };
  const openDocumentsForPerson = (
    row: EjournalPreviewRow,
    mode:
      | "default"
      | "salaryPowerAttorney"
      | "ubdReport"
      | "ubdRestoreReport"
      | "form6Report"
      | "form12Report"
      | "serviceCharacteristic"
      | "zhbdCertificate"
      | "temporaryMilitaryId"
      | "lostMilitaryId" = "default",
    meta?: { fullPosition?: string },
  ) => {
    const externalId = buildPersonSummary(row).externalId;
    const fullPosition =
      meta?.fullPosition?.trim() ||
      pickFullPositionFromPersonRow(row) ||
      getPersonFullPositionTitle(row);
    const rowForDocuments: EjournalPreviewRow = {
      ...row,
      ...(fullPosition
        ? {
            __zhbdFullPosition: fullPosition,
            roster__повна_посада: fullPosition,
            повна_посада: fullPosition,
          }
        : {}),
    };
    window.localStorage.setItem(
      "army-grid:selected-person",
      JSON.stringify(rowForDocuments),
    );
    window.localStorage.setItem("army-grid:selected-document-mode", mode);
    storeSelectedPersonFullPosition(fullPosition);
    applyRoute(
      pushAppRoute(
        buildDocumentRoute({
          personExternalId: externalId,
          rowId: row.__dbRowId,
          type: mode,
        }),
        "documents",
      ),
    );
  };

  if (loading) {
    return (
      <div className="auth-shell auth-shell--loading">
        <Typography variant="body2" color="text.secondary">
          Завантаження…
        </Typography>
      </div>
    );
  }

  if (!user) {
    return <AuthPage />;
  }

  if (!canView) {
    return <PendingAccessPage />;
  }

  const pageWriteArea = writeAreaForPage(activePage);
  const canEditActivePage = pageWriteArea
    ? canEditArea(pageWriteArea)
    : canEdit || isAdmin;
  const pageReadonly = Boolean(pageWriteArea) && !canEditActivePage;

  return (
    <>
    <div
      className={`app-shell${mobileNavOpen ? " mobile-nav-open" : ""}${pageReadonly ? " app-shell--readonly" : ""}`}
    >
      <SciScrollbars />
      <header className="mobile-topbar">
        <Button
          className="mobile-burger-btn"
          variant="outlined"
          size="small"
          aria-label="Відкрити меню"
          aria-expanded={mobileNavOpen}
          aria-controls="app-sidebar"
          onClick={() => setMobileNavOpen(true)}
        >
          <MenuOutlinedIcon fontSize="small" />
        </Button>
        <div className="mobile-topbar-copy">
          <Typography variant="overline" color="text.secondary">
            Army Grid
          </Typography>
          <Typography variant="body2">
            {APP_PAGE_LABELS[activePage] ?? activePage}
          </Typography>
        </div>
      </header>
      <button
        type="button"
        className="mobile-nav-backdrop"
        aria-label="Закрити меню"
        tabIndex={mobileNavOpen ? 0 : -1}
        onClick={() => setMobileNavOpen(false)}
      />
      <Sidebar
        activePage={activePage}
        onPageChange={changePage}
        mobileOpen={mobileNavOpen}
        onMobileClose={() => setMobileNavOpen(false)}
      />

      {activePage === "overview" ? (
        <OverviewPage onOpenPersonnel={openPersonnelForPerson} />
      ) : activePage === "analytics" ? (
        <AnalyticsPage />
      ) : activePage === "bchs" ? (
        <BchsPage />
      ) : activePage === "excelFill" ? (
        <ExcelFillPage />
      ) : activePage === "anketaData" ? (
        <AnketaDataPage />
      ) : activePage === "socPassport" ? (
        <SocPassportPage />
      ) : activePage === "ejournal" ? (
        <EjournalPage />
      ) : activePage === "personnel" ? (
        <PersonnelPage onOpenDocuments={openDocumentsForPerson} />
      ) : activePage === "documents" ? (
        <DocumentsPage
          key={routeKey}
          onNavigate={(path) => {
            applyRoute(pushAppRoute(path, getPageFromPath(path)));
          }}
        />
      ) : activePage === "documentSettings" ? (
        <DocumentSignatoriesSettingsPage />
      ) : activePage === "usersAccess" ? (
        isAdmin ? <UsersAccessPage /> : null
      ) : activePage === "profile" ? (
        <ProfilePage />
      ) : (
        <WorkTasksPage />
      )}
    </div>
    <SciLiveFeedback />
    <AppToastHost />
    </>
  );
}

export default App;
