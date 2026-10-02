import { useState } from "react";
import { Box, Stack, Typography } from "@/components/sci/SciPrimitives";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs/tabs";
import { ExcelLabUbdPanel } from "./ubd/ExcelLabUbdPanel";
import { ExcelLabVkCollectPanel } from "./ExcelLabVkCollectPanel";

type ExcelLabTab = "vk" | "ubd";

export function ExcelLabPage() {
  const [tab, setTab] = useState<ExcelLabTab>("vk");

  return (
    <Box className="page-shell excel-lab-page" sx={{ p: 2 }}>
      <Stack spacing={2}>
        <Box>
          <Typography variant="h5">Excel Lab</Typography>
          <Typography variant="body2" color="text.secondary">
            Збір ВК і розбір УБД.
          </Typography>
        </Box>

        <Tabs
          value={tab}
          onValueChange={(value) => setTab(value as ExcelLabTab)}
        >
          <TabsList aria-label="Excel Lab">
            <TabsTrigger value="vk">Збір ВК</TabsTrigger>
            <TabsTrigger value="ubd">УБД</TabsTrigger>
          </TabsList>
        </Tabs>

        <Box sx={{ display: tab === "vk" ? "block" : "none" }}>
          <ExcelLabVkCollectPanel />
        </Box>
        <Box sx={{ display: tab === "ubd" ? "block" : "none" }}>
          <ExcelLabUbdPanel />
        </Box>
      </Stack>
    </Box>
  );
}
