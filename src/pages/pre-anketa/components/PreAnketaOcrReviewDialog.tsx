import {
  Box,
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import type { PreAnketaOcrProposal } from "../preAnketaOcr";

export function PreAnketaOcrReviewDialog({
  open,
  proposals,
  onChange,
  onClose,
  onApply,
}: {
  open: boolean;
  proposals: PreAnketaOcrProposal[];
  onChange: (next: PreAnketaOcrProposal[]) => void;
  onClose: () => void;
  onApply: () => void;
}) {
  const selectedCount = proposals.filter((item) => item.selected).length;

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle>Gemini: що заповнити</DialogTitle>
      <DialogContent dividers>
        {!proposals.length ? (
          <Typography variant="body2">
            Немає нових полів для заповнення.
          </Typography>
        ) : (
          <Stack spacing={1.5}>
            {proposals.map((proposal, index) => (
              <Box
                key={`${proposal.field}-${index}`}
                sx={{
                  border: "1px solid var(--sci-border)",
                  borderRadius: 1,
                  p: 1.25,
                }}
              >
                <Stack direction="row" spacing={1} alignItems="flex-start">
                  <Checkbox
                    checked={proposal.selected}
                    onCheckedChange={(checked) => {
                      const next = [...proposals];
                      next[index] = {
                        ...proposal,
                        selected: checked === true,
                      };
                      onChange(next);
                    }}
                  />
                  <Box sx={{ flex: 1, minWidth: 0 }}>
                    <Typography variant="subtitle2">
                      {proposal.label}
                      {proposal.source ? ` · ${proposal.source}` : ""}
                    </Typography>
                    <Typography
                      variant="body2"
                      sx={{ whiteSpace: "pre-wrap", mt: 0.5 }}
                    >
                      {proposal.value}
                    </Typography>
                  </Box>
                </Stack>
              </Box>
            ))}
          </Stack>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Скасувати</Button>
        <Button
          variant="contained"
          onClick={onApply}
          disabled={!selectedCount}
        >
          Застосувати {selectedCount ? `(${selectedCount})` : ""}
        </Button>
      </DialogActions>
    </Dialog>
  );
}
