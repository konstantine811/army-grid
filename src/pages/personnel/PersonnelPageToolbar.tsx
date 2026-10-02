import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { FileUploadOutlinedIcon } from "@/components/sci/icons";

export function PersonnelPageToolbar({
  canEdit,
  isLoading,
  isMergingAnketaData,
  isMergingVkTpvDovidky,
  photoExtractCount,
  diskSearchCount,
  onMergeAnketa,
  onImportVkTpv,
  onOpenPhotoExtract,
  onOpenDiskSearch,
  onImportRoster,
  onRefresh,
}: {
  canEdit: boolean;
  isLoading: boolean;
  isMergingAnketaData: boolean;
  isMergingVkTpvDovidky: boolean;
  photoExtractCount: number;
  diskSearchCount: number;
  onMergeAnketa: () => void;
  onImportVkTpv: (file: File | undefined) => void;
  onOpenPhotoExtract: () => void;
  onOpenDiskSearch: () => void;
  onImportRoster: (file: File | undefined) => void;
  onRefresh: () => void;
}) {
  const mergeBusy = isLoading || isMergingAnketaData || isMergingVkTpvDovidky;

  return (
    <header className="topbar analytics-topbar personnel-topbar">
      <Box>
        <Typography component="h1" variant="h4">
          Особовий склад
        </Typography>
        <Typography
          className="personnel-topbar-hint"
          variant="body2"
          color="text.secondary"
        >
          Список із ЕЖООС · ручний файл «Штатка» (.xlsx) імпортується в БД.
        </Typography>
      </Box>
      <Stack className="personnel-topbar-actions" direction="row" spacing={1}>
        <Button
          variant="outlined"
          disabled={!canEdit || mergeBusy}
          onClick={onMergeAnketa}
          title="Доповнити порожні поля з таблиці «Анкети» і додати осіб, яких ще немає в особовому складі"
        >
          {isMergingAnketaData ? "З анкет…" : "З анкетних даних"}
        </Button>
        <Button
          component="label"
          variant="outlined"
          disabled={!canEdit || mergeBusy}
          startIcon={<FileUploadOutlinedIcon />}
          title="Імпорт ВК № в ООС та ІПН в анкетні дані за ПІБ"
        >
          {isMergingVkTpvDovidky ? "ВК ТПВ…" : "ВК ТПВ ДОВІДКИ"}
          <input
            hidden
            type="file"
            accept=".xlsx,.xlsm"
            disabled={!canEdit}
            onChange={(event) => {
              onImportVkTpv(event.target.files?.[0]);
              event.target.value = "";
            }}
          />
        </Button>
        <Button
          variant="outlined"
          disabled={!photoExtractCount || !canEdit}
          onClick={onOpenPhotoExtract}
          title="Витягнути фото з PDF-анкет у БД для осіб без фото"
        >
          Фото з анкет · {photoExtractCount}
        </Button>
        <Button
          variant="outlined"
          disabled={!diskSearchCount}
          onClick={onOpenDiskSearch}
        >
          Пошук усіх анкет
        </Button>
        <Button
          component="label"
          disabled={isLoading || !canEdit}
          startIcon={<FileUploadOutlinedIcon />}
          variant="outlined"
          title="Вибрати файл «Штатка» (.xlsx/.xlsm), імпортувати його та записати в БД"
        >
          Імпорт Штатки в БД
          <input
            hidden
            type="file"
            accept=".xlsx,.xlsm"
            disabled={!canEdit || isLoading}
            onChange={(event) => {
              onImportRoster(event.target.files?.[0]);
              event.target.value = "";
            }}
          />
        </Button>
        <Button variant="outlined" disabled={!canEdit} onClick={onRefresh}>
          Оновити з БД
        </Button>
      </Stack>
    </header>
  );
}
