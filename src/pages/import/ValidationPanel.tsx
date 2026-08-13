import { Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { CheckCircleOutlineIcon } from "@/components/sci/icons";
import { ErrorOutlineOutlinedIcon } from "@/components/sci/icons";
import { WarningAmberOutlinedIcon } from "@/components/sci/icons";

export function ValidationPanel() {
  return (
    <section className="panel validation-panel">
      <div className="panel-heading">Перевірка даних</div>
      <div className="panel-body">
        <div className="metric-grid">
          <div className="metric">
            <span className="metric-value">142</span>
            <Typography variant="caption">рядки</Typography>
          </div>
          <div className="metric">
            <span className="metric-value" style={{ color: "#91bc59" }}>
              136
            </span>
            <Typography variant="caption" color="success.main">
              готові
            </Typography>
          </div>
          <div className="metric">
            <span className="metric-value" style={{ color: "#f6a23d" }}>
              4
            </span>
            <Typography variant="caption" color="warning.main">
              дублікати
            </Typography>
          </div>
          <div className="metric">
            <span className="metric-value" style={{ color: "#ef4e3c" }}>
              2
            </span>
            <Typography variant="caption" color="error.main">
              помилки
            </Typography>
          </div>
        </div>

        <Typography
          variant="overline"
          color="text.secondary"
          sx={{ display: "block", mt: 2 }}
        >
          Деталі перевірки
        </Typography>
        <div className="detail-row">
          <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
            <CheckCircleOutlineIcon color="success" fontSize="small" />
            <Typography variant="body2">Готові до імпорту</Typography>
          </Stack>
          <Typography variant="body2" color="success.main">
            136
          </Typography>
        </div>
        <div className="detail-row">
          <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
            <WarningAmberOutlinedIcon color="warning" fontSize="small" />
            <Typography variant="body2">Можливі дублікати</Typography>
          </Stack>
          <Typography variant="body2" color="warning.main">
            4
          </Typography>
        </div>
        {[
          "Іваненко Іван Іванович",
          "Петренко Олексій Васильович",
          "Сидоренко Андрій Петрович",
        ].map((name) => (
          <Typography
            variant="caption"
            className="muted"
            sx={{ display: "block", ml: 3.5, mt: 0.8 }}
            key={name}
          >
            {name} · 2 збіги
          </Typography>
        ))}
        <div className="detail-row">
          <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
            <ErrorOutlineOutlinedIcon color="error" fontSize="small" />
            <Typography variant="body2">Помилки</Typography>
          </Stack>
          <Typography variant="body2" color="error.main">
            2
          </Typography>
        </div>
        <Button
          fullWidth
          color="error"
          variant="outlined"
          startIcon={<WarningAmberOutlinedIcon />}
          sx={{ mt: 2 }}
        >
          Виправити 6 записів
        </Button>
      </div>
    </section>
  );
}
