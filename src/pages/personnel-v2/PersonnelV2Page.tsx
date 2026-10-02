import { useEffect, useMemo, useRef, useState } from "react";
import { Alert, Box, Typography } from "@/components/sci/SciPrimitives";
import { MagneticTapePreloader } from "@/components/sci/MagneticTapePreloader";
import { useVirtualizer } from "@tanstack/react-virtual";
import {
  loadPersonnelV2StaffSheet,
  type PersonnelV2StaffPerson,
  type PersonnelV2StaffSheet,
} from "./loadPersonnelV2StaffSheet";

const staffCell = (person: PersonnelV2StaffPerson, column: number) =>
  person.cells.find((cell) => cell.column === column)?.value ?? "";

export function PersonnelV2Page({ active = true }: { active?: boolean }) {
  const [staffSheet, setStaffSheet] = useState<PersonnelV2StaffSheet | null>(
    null,
  );
  const [isLoading, setIsLoading] = useState(true);
  const [message, setMessage] = useState("");
  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!active) return;
    const controller = new AbortController();
    setIsLoading(true);
    setMessage("");
    void loadPersonnelV2StaffSheet({ signal: controller.signal })
      .then((next) => {
        console.log("[personnel-v2] штатка", next);
        setStaffSheet(next);
        setMessage("");
      })
      .catch((error: unknown) => {
        if (controller.signal.aborted) return;
        setStaffSheet(null);
        setMessage(
          error instanceof Error
            ? error.message
            : "Не вдалося завантажити штатку.",
        );
      })
      .finally(() => {
        if (!controller.signal.aborted) setIsLoading(false);
      });
    return () => controller.abort();
  }, [active]);

  const people = useMemo(
    () => staffSheet?.people.filter((person) => !person.inArchive) ?? [],
    [staffSheet],
  );
  const rowVirtualizer = useVirtualizer({
    count: people.length,
    getScrollElement: () => listRef.current,
    estimateSize: () => 58,
    overscan: 12,
    gap: 8,
  });

  if (!active) return null;

  return (
    <main className="main-panel overview-page">
      <header className="topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Особовий склад
          </Typography>
          <Typography color="text.secondary" variant="body2">
            Список зі штатки
          </Typography>
        </Box>
      </header>

      {isLoading && !staffSheet ? (
        <div className="personnel-list-preloader">
          <MagneticTapePreloader
            status="ЗАВАНТАЖЕННЯ ШТАТКИ"
            hint="Читаю Загальний список і готую список осіб."
          />
        </div>
      ) : null}

      {message ? <Alert severity="warning">{message}</Alert> : null}

      {staffSheet ? (
        <section className="personnel-v2-staff">
          <div className="personnel-v2-staff-summary">
            <span>
              У штаті: <strong>{staffSheet.inStaff}</strong>
            </span>
            <span>
              Архів: <strong>{staffSheet.inArchive}</strong>
            </span>
            <span>{staffSheet.sourceFileName || staffSheet.sheetName}</span>
          </div>
          <div className="personnel-v2-staff-list" ref={listRef}>
            <div
              className="personnel-v2-staff-virtual"
              style={{ height: `${rowVirtualizer.getTotalSize()}px` }}
            >
              {rowVirtualizer.getVirtualItems().map((virtualRow) => {
                const person = people[virtualRow.index];
                if (!person) return null;
                const callSign = staffCell(person, 15);
                const unit = staffCell(person, 2);
                const status = staffCell(person, 21);
                return (
                  <div
                    className="personnel-v2-staff-row"
                    key={person.id}
                    style={{ transform: `translateY(${virtualRow.start}px)` }}
                  >
                    <strong>{person.name}</strong>
                    <span>
                      {[callSign, unit, status].filter(Boolean).join(" · ") ||
                        "—"}
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        </section>
      ) : null}
    </main>
  );
}
