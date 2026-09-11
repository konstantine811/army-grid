# ЕЖООС: план рефакторингу по PR

Документ доповнює [EJOOS_FUNCTIONAL_PLAN.md](./EJOOS_FUNCTIONAL_PLAN.md). Мета — зробити pipeline прозорим і безпечним для змін **без big-bang переписування**.

## Принципи

1. **Кожен PR — поведінка не змінюється** (лише переміщення коду + re-export зі старих шляхів).
2. **Перед розрізанням — safety net**: CI має проходити `ejoos*.test.ts` + august/hop fixtures.
3. **Новий код — pure functions + контракт-тести**; ZIP/Excel — окремий шар.
4. **Старі імпорти живуть** через `ejoosSyncPlan.ts` / `ejoosSyncApply.ts` як barrel re-export, поки не приберемо їх у PR 12.

## Цільова архітектура

```text
src/pages/ejournal/ejoos/
  types/           EjoosSyncOp, EjoosSyncPlan, PbMovement, …
  parse/
    pb/            parsePbShPeople, parsePbArchive, parsePbMovements
    workbook/      parseEjoosOos, parseEjoosShpo, parseEjoosAbsents, …
  plan/
    buildPlan.ts   buildEjoosSyncPlan (фасад)
    arrival.ts
    absent.ts
    absentClose.ts
    excludeTransfer.ts
    disposition.ts
    positionChange.ts
    rankChange.ts
    contract.ts
    timesheetDay.ts
    shpoOccupant.ts
  validate/
    opRequirements.ts   (вже є)
    planBlocks.ts       planBlocksWorkbookApply, workbookApplyBlockMessage
  apply/
    applyConfirmedOps.ts  (фасад)
    writePlan.ts          ZipCellWrite[] без JSZip
    handlers/             по kind → делегує в *Zip.ts
  presentation/    exportStyle, zipCellWrites, restyle*
  workspace/         actions.ts, state.ts (тонкий Context)
```

Поточні файли (`ejoosSyncPlan.ts`, `ejoosSyncApply.ts`) залишаються як **тонкі фасади** до повного видалення.

---

## PR 0 — Safety net (без змін коду)

**Мета:** зафіксувати baseline перед рефакторингом.

**Дії:**
- Переконатися, що проходять:
  ```bash
  npx vitest run src/pages/ejournal/ejoos*.test.ts src/pages/ejournal/EjoosChangesPanel.test.ts
  ```
- Додати npm-скрипт `test:ejoos` (опційно).
- Зафіксувати в README/CI, що PR з `ejoos*` обов'язково ганяють ці тести.

**Критерій приймання:** усі існуючі ejoos-тести зелені; august + hop + bulkApply — обов'язкові в CI.

---

## PR 1 — Спільні типи та контракт op

**Мета:** один source of truth для `EjoosSyncOp`, `EjoosOpKind`, `EjoosSyncPlan`.

**Дії:**
- Створити `ejoos/types/syncOp.ts`, `ejoos/types/pb.ts`, `ejoos/types/workbookRows.ts`.
- Перенести типи з `ejoosSyncPlan.ts` (рядки ~74–120, PbShPerson, PbMovement, …).
- У `ejoosSyncPlan.ts` — `export type { … } from "./ejoos/types/…"`.
- Додати `ejoosContract.test.ts` розширити: snapshot усіх `EjoosOpKind` (12 kinds), обов'язкові поля op.

**Файли:** ~3 нових, `ejoosSyncPlan.ts` −200 рядків типів.

**Критерій:** жодна зміна поведінки; імпорти `from "./ejoosSyncPlan"` працюють як раніше.

---

## PR 2 — Виділити парсинг 1ПБ

**Мета:** винести `parsePbShPeople`, `parsePbArchive`, `parsePbMovements`, `createMovementKey`, `collectProcessedMovementKeys` з 6533-рядкового файлу.

**Дії:**
- `ejoos/parse/pb/shPeople.ts`
- `ejoos/parse/pb/archive.ts`
- `ejoos/parse/pb/movements.ts`
- `ejoos/parse/pb/index.ts` — re-export
- **Нові unit-тести:** `ejoos/parse/pb/*.test.ts` на мінімальні snapshot-и (не лише через august).

**Критерій:** `buildEjoosSyncPlan` імпортує з `./parse/pb`; august fixtures без змін.

---

## PR 3 — Виділити парсинг аркушів ЕЖООС

**Мета:** `parseEjoosOos`, `parseEjoosShpo`, `parseEjoosAbsents`, `parseEjoosArrivals`, `parseEjoosExcluded`, `parseEjoosTimesheetDay`, `parseEjoosTimesheetPeople`, `findEjoosSheet`.

**Дії:**
- `ejoos/parse/workbook/*.ts`
- Об'єднати з дублікатами в `ejoosParsers.ts` / `ejoosLiveViews.ts` де можливо (лише re-export, без злиття логіки в цьому PR).
- Тести: по одному кейсу на кожен parse* (рядок з august fixture → очікуваний об'єкт).

**Критерій:** `parseEjoosOos.test.ts` вже частково покриває — перенести поруч із модулем.

---

## PR 4 — Journal day + plan guards

**Мета:** ізолувати дату журналу та блокування apply.

**Дії:**
- `ejoos/plan/journalDay.ts` — `parseTimesheetDayFromPbName`, `parseAsOfDateLabel`, `resolveJournalTimesheetDay`, `refreshPlanTimesheetHorizon`.
- `ejoos/validate/planBlocks.ts` — `planBlocksWorkbookApply`, `workbookApplyBlockMessage`, константи `SOURCE_DATE_UNKNOWN_MESSAGE`, …
- Тести: перенести/розширити `ejoosJournalHorizon.test.ts`, `ejoosAsOfDate.test.ts`.

**Критерій:** логіка блокування місяця / невідомої дати — окремий модуль з тестами.

---

## PR 5 — Розрізати `buildEjoosSyncPlan`: частина 1 (arrival + absent)

**Мета:** перші два op kind у окремих файлах.

**Дії:**
- `ejoos/plan/arrival.ts` — `planArrivalOps(ctx): EjoosSyncOp[]`
- `ejoos/plan/absent.ts` — `planAbsentUpsertOps(ctx): EjoosSyncOp[]`
- `ejoos/plan/context.ts` — спільний `PlanBuildContext` (shPeople, ejoosOos, movements, timesheetDay, …)
- `buildEjoosSyncPlan` збирає ops через `[...planArrivalOps(ctx), ...planAbsentOps(ctx), …]`

**Тести:**
- `ejoos/plan/arrival.test.ts` — 3–5 кейсів з `ejoosTempArrivalClosure.test.ts`
- `ejoos/plan/absent.test.ts` — кейси з hop fixtures (лише absent_upsert)

**Критерій:** diff ops від `buildEjoosSyncPlan` **ідентичний** для цих kinds (snapshot test на fixture).

---

## PR 6 — Plan: movement helpers (incremental split)

**Мета:** винести перші блоки з `considerMovement` / pre-movement фази без повного розрізу моноліту.

**Зроблено:**
- `ejoos/plan/rankAndContract.ts` — `planRankAndContractOps` (rank_change + contract_update + `latestRankEventByPerson`)
- `ejoos/plan/movements/transferScopeUnclear.ts` — `planTransferScopeUnclearOp`
- `ejoos/plan/movements/manualArrivalMovement.ts` — `planManualArrivalMovementOp` (ПРИБУВ/ЗВІЛЬН)
- `ejoos/plan/movements/movementHelpers.test.ts` — unit-тести для двох movement helpers

**Ще в `ejoosSyncPlan.ts`:** тіло `considerMovement` (~1400 рядків) — на PR 7+.

**Критерій:** `npm run test:ejoos` без нових регресій (6 august/hop — pre-existing).

---

## PR 7 — Plan: absent_close + timesheet_day ✅

**Зроблено:**
- `ejoos/plan/absentCloseFromSh.ts` — `planAbsentCloseFromShOps` (sh «в строю» → закрити «Тимч. відсутні»)
- `ejoos/plan/timesheetDayFromSh.ts` — `planTimesheetDayFromShOps` (коди табеля зі статусу sh)
- `ejoos/plan/timesheetDayFromArchive.ts` — `planTimesheetDayFromArchiveOps` (PAINT_ARCHIVE)
- `ejoos/plan/timesheetDay.test.ts` — unit-тести

**Ще в `ejoosSyncPlan.ts`:** cleanup timesheet_day (stale tab, duplicate rows, false hop exclusion) — на PR 8+.

**Критерій:** `npm run test:ejoos` — 284 passed, 6 failed (pre-existing august/hop).

---

## PR 8 — Plan: exclude_transfer + move_to_disposition ✅

**Зроблено:**
- `ejoos/plan/movements/disposition.ts` — `planDispositionMovementOps` (РОЗПОРЯДЖ → move_to_disposition)
- `ejoos/plan/movements/excludeTransfer.ts` — `planExcludeTransferMovementOps` (зовнішнє ПЕРЕВ/ПОСАДА → exclude_transfer)
- `ejoos/plan/movements/dispositionExclude.test.ts` — smoke unit-тест

**Ще в `considerMovement`:** position_change, arrival, absent_upsert inline — на PR 9+.

**Критерій:** `npm run test:ejoos` — 284 passed, 6 failed (pre-existing august/hop).

---

## PR 9 — Plan: position_change + shpo_occupant + timesheet cleanup ✅

**Зроблено:**
- `ejoos/plan/movements/positionChange.ts` — `planPositionChangeMovementOps`
- `ejoos/plan/shpoOccupant.ts` — `planShpoOccupantOps` + `planShpoReconcileOps`
- `ejoos/plan/timesheetDayCleanup.ts` — stale tab / duplicate exclusions / dup tab rows

**Ще в `buildEjoosSyncPlan`:** absent_upsert inline, data_mismatch, transfer cancel review (~3100 рядків).

**Критерій:** `npm run test:ejoos` — 285 passed, 6 failed (pre-existing august/hop).

---

## PR 10 — Write-plan шар (від’єднати від JSZip)

**Мета:** apply спочатку будує `ZipCellWrite[]`, потім один раз пише в ZIP.

**Дії:**
- `ejoos/apply/writePlan.ts` — тип `EjoosWritePlan = { writes: ZipCellWrite[]; restyle: …; meta: … }`
- `ejoos/apply/buildWritePlan.ts` — `buildWritePlanForOps(ops, snapshot): EjoosWritePlan` (pure, без async)
- `applyConfirmedEjoosOps` → `buildWritePlan` → `applyWritePlanToBlob` (існуючий zip код)
- Unit-тести на `buildWritePlan` для кожного kind (без FileReader/JSZip): порівняння writes з golden JSON.

**Критерій:** bulkApply + august — byte-identical або snapshot-identical workbook (як зараз).

---

## PR 10 — Handlers apply по kind

**Мета:** розрізати `ejoosSyncApply.ts` (2520 рядків).

**Дії:**
- `ejoos/apply/handlers/timesheetDay.ts`
- `ejoos/apply/handlers/excludeTransfer.ts` → делегує `ejoosExcludeTransferZip`
- `ejoos/apply/handlers/disposition.ts` → `ejoosDispositionZip`
- `ejoos/apply/handlers/positionChange.ts` → `ejoosPositionChangeZip`
- `ejoos/apply/applyConfirmedOps.ts` — orchestration: partition batches, sort, validate, handlers
- **Нові тести:** `ejoosDispositionZip.test.ts`, `ejoosSyncApply.handlers.test.ts` (по 2 кейси на handler)

**Критерій:** `ejoosSyncApply.ts` — re-export + deprecated comment; файл < 100 рядків.

---

## PR 11 — Тонкий WorkspaceContext

**Мета:** прибрати бізнес-логіку з 1747-рядкового React-файлу.

**Дії:**
- `ejoos/workspace/actions.ts`:
  - `refreshRosterForExport`
  - `buildPlanFromPb`
  - `applyPersonChanges`
  - `exportStyledWorkbook`
- `EjoosWorkspaceContext.tsx` — лише state, провайдер, виклики actions
- Опційно: `ejoos/workspace/useEjoosWorkspace.ts` hook

**Критерій:** Context < 600 рядків; жодної нової логіки plan/apply всередині JSX.

---

## PR 12 — Прибрати legacy barrel-файли

**Мета:** прямі імпорти з `ejoos/plan`, `ejoos/apply`, `ejoos/parse`.

**Дії:**
- Codemod або ручна заміна імпортів у `Ejoos*.tsx`, тестах.
- Видалити `ejoosSyncPlan.ts` / `ejoosSyncApply.ts` або лишити 10-рядкові re-export з `@deprecated`.

**Критерій:** grep не знаходить `from "./ejoosSyncPlan"` поза `ejoos/` (крім re-export index).

---

## PR 13+ — Довгий хвіст (за пріоритетом)

| PR | Модуль | Тести |
|----|--------|-------|
| 13 | `ejoosAnketaFill.ts` → `ejoos/presentation/anketaFill.ts` | golden anketa rows |
| 14 | `ejoosChangeHistorySheet.ts` → `ejoos/presentation/historySheet.ts` | protocol row snapshot |
| 15 | `ejoosWorkbookSanitize.ts` → `ejoos/apply/sanitize.ts` | broken XML fixtures |
| 16 | `ejoosZipCellWrites.ts` — розбити restyle / inline writes | style tests вже частково є |
| 17 | `ejoosPersonDiff.ts` — UI grouping vs pure diff | розділити `personDiffUi.ts` |

---

## Контракт-тести по op kind (шаблон)

Додавати в кожному PR для відповідного kind:

```ts
// ejoos/plan/arrival.test.ts
describe("planArrivalOps", () => {
  it("contract: temp arrival from movement → arrival op ready", () => {
    const ctx = buildPlanContextFromFixtures("arrival/simple");
    const ops = planArrivalOps(ctx);
    expect(ops).toMatchSnapshot();
    expect(ops.every((op) => op.kind === "arrival")).toBe(true);
  });
});
```

**Мінімальний набір fixtures на kind** (ціль — до PR 8):

| Kind | Fixture джерело |
|------|-----------------|
| `arrival` | tempArrivalClosure |
| `absent_upsert` | hop absence open |
| `absent_close` | hop return |
| `exclude_transfer` | bulkApply exclude |
| `move_to_disposition` | priorMonthDisposition |
| `position_change` | august position |
| `rank_change` | august rank |
| `timesheet_day` | journalHorizon |
| `shpo_occupant` | august shpo |
| `contract_update` | contract.test |

---

## Порядок і оцінка

| PR | Ризик | Час (орієнт.) | Блокує | Статус |
|----|-------|---------------|--------|--------|
| 0 | низький | 0.5 д | — | ✅ `npm run test:ejoos` |
| 1 | низький | 0.5 д | 2–12 | ✅ `ejoos/types/` |
| 2–3 | низький | 2 д | 5–8 | 🟡 parse/pb + cellText; workbook parse — далі |
| 4 | низький | 1 д | 5 | ✅ journalDay + planGuards + opId |
| 5–8 | **високий** | 4–6 д | 9–10 | 🟡 `tempArrivalClose`, `absentFromArchive`, `szchAbsentClose` — далі movement/consider |
| 9–10 | **високий** | 3–4 д | 11 |
| 11 | середній | 2 д | 12 |
| 12 | низький | 1 д | — |

**Рекомендований старт:** PR 0 → PR 1 → PR 2 → PR 5 (перший plan split з snapshot diff).

---

## Що не робити

- Не змішувати plan split і apply split в одному PR.
- Не змінювати правила movement/exclude під час extract — лише move code.
- Не видаляти august/hop fixtures; лише додавати smaller contract tests.
- Не рефакторити UI-панелі (`EjoosChangesPanel`, cards) до PR 11.

---

## Чеклист перед merge кожного PR

- [ ] `npx vitest run src/pages/ejournal/ejoos*.test.ts`
- [ ] Немає нових `any` у public API
- [ ] Старі імпорти працюють (re-export)
- [ ] Додано або оновлено тести для **перенесеного** коду
- [ ] `EJOOS_FUNCTIONAL_PLAN.md` — оновлено статус секції, якщо змінилась структура
