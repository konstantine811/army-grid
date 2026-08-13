import { Box, Button, Divider, Stack, Switch, Typography } from "@/components/sci/SciPrimitives";
import { SettingsOutlinedIcon } from "@/components/sci/icons";
import { SyncAltOutlinedIcon } from "@/components/sci/icons";

const mappingRows = [
  ["ПІБ", "Особа", "Точний збіг", "136 зіставлено"],
  ["Дата від", "Початок статусу", "Найближча дата", "140 зіставлено"],
  ["Дата до", "Кінець статусу", "Найближча дата", "138 зіставлено"],
  ["Статус", "Поточний статус", "За словником", "136 зіставлено"],
  ["Підрозділ", "Підрозділ", "Точний збіг", "134 зіставлено"],
  ["Звання", "Звання", "Точний збіг", "132 зіставлено"],
  ["Посада", "Посада", "Точний збіг", "130 зіставлено"],
  ["Коментар", "Примітка", "Додати/оновити", "46 зіставлено"],
];

export function MappingPanel() {
  return (
    <section className="panel">
      <div className="panel-heading">
        Відповідність полів та правила об’єднання
      </div>
      <div className="panel-body">
        <div className="mapping-row mapping-header">
          <span />
          <span>Колонка у файлі</span>
          <span>Поле системи</span>
          <span>Правило</span>
          <span>Результат</span>
        </div>
        {mappingRows.map(([fileColumn, systemField, rule, result]) => (
          <div className="mapping-row" key={fileColumn}>
            <SyncAltOutlinedIcon fontSize="small" color="disabled" />
            <div className="fake-select">{fileColumn}</div>
            <div className="fake-select">{systemField}</div>
            <div className="fake-select">{rule}</div>
            <Typography variant="body2" color="primary">
              {result}
            </Typography>
          </div>
        ))}

        <Divider sx={{ my: 2 }} />
        <Typography variant="overline" color="text.secondary">
          Правила об’єднання
        </Typography>
        <div className="merge-rules">
          {[
            [
              "Оновити наявні записи",
              "Оновлювати дані в існуючих записах, якщо знайдено зіставлення.",
            ],
            [
              "Створити нові",
              "Створювати нові записи, якщо зіставлення не знайдено.",
            ],
            [
              "Не перезаписувати порожнім",
              "Не змінювати значення полів, якщо у файлі порожньо.",
            ],
          ].map(([title, text]) => (
            <Stack
              direction="row"
              spacing={1.4}
              sx={{ alignItems: "flex-start" }}
              key={title}
            >
              <Switch size="small" defaultChecked color="primary" />
              <Box>
                <Typography variant="body2">{title}</Typography>
                <Typography variant="caption" className="muted">
                  {text}
                </Typography>
              </Box>
            </Stack>
          ))}
        </div>
        <Stack direction="row" sx={{ justifyContent: "flex-end", mt: 2 }}>
          <Button
            size="small"
            startIcon={<SettingsOutlinedIcon />}
            variant="outlined"
          >
            Налаштування розширеного зіставлення
          </Button>
        </Stack>
      </div>
    </section>
  );
}
