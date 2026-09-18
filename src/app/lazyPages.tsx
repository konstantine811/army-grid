import {
  lazy,
  Suspense,
  type LazyExoticComponent,
  type ReactNode,
} from "react";
import { MagneticTapePreloader } from "@/components/sci/MagneticTapePreloader";

export const LazyOverviewPage = lazy(() =>
  import("../pages/overview/OverviewPage").then((m) => ({
    default: m.OverviewPage,
  })),
) as LazyExoticComponent<
  (props: { active?: boolean }) => React.ReactElement
>;
export const LazyAnalyticsPage = lazy(() =>
  import("../pages/analytics/AnalyticsPage").then((m) => ({
    default: m.AnalyticsPage,
  })),
);
export const LazyBchsPage = lazy(() =>
  import("../pages/bchs/BchsPage").then((m) => ({ default: m.BchsPage })),
) as LazyExoticComponent<(props: { active?: boolean }) => React.ReactElement>;
export const LazyExcelFillPage = lazy(() =>
  import("../pages/excel-fill/ExcelFillPage").then((m) => ({
    default: m.ExcelFillPage,
  })),
);
export const LazyAnketaDataPage = lazy(() =>
  import("../pages/anketa-data/AnketaDataPage").then((m) => ({
    default: m.AnketaDataPage,
  })),
);
export const LazyPreAnketaPage = lazy(() =>
  import("../pages/pre-anketa/PreAnketaPage").then((m) => ({
    default: m.PreAnketaPage,
  })),
);
export const LazySocPassportPage = lazy(() =>
  import("../pages/soc-passport/SocPassportPage").then((m) => ({
    default: m.SocPassportPage,
  })),
);
export const LazyEjournalPage = lazy(() =>
  import("../pages/ejournal/EjournalPage").then((m) => ({
    default: m.EjournalPage,
  })),
);
export const LazyPersonnelPage = lazy(() =>
  import("../pages/personnel/PersonnelPage").then((m) => ({
    default: m.PersonnelPage,
  })),
) as LazyExoticComponent<
  typeof import("../pages/personnel/PersonnelPage").PersonnelPage
>;
export const LazyDocumentsPage = lazy(() =>
  import("../pages/documents/DocumentsPage").then((m) => ({
    default: m.DocumentsPage,
  })),
) as LazyExoticComponent<
  typeof import("../pages/documents/DocumentsPage").DocumentsPage
>;

function PageLoadingFallback({
  status = "ЗАВАНТАЖЕННЯ МОДУЛЯ",
  hint,
}: {
  status?: string;
  hint?: string;
}) {
  return (
    <div className="app-page-slot--pending app-page-preloader">
      <MagneticTapePreloader status={status} hint={hint} />
    </div>
  );
}

export function LazyPageBoundary({
  children,
  status,
  hint,
}: {
  children: ReactNode;
  status?: string;
  hint?: string;
}) {
  return (
    <Suspense fallback={<PageLoadingFallback status={status} hint={hint} />}>
      {children}
    </Suspense>
  );
}
