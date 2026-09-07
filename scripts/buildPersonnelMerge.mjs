import * as esbuild from "esbuild";
import { copyFileSync, mkdirSync } from "node:fs";
import path from "node:path";

const outFile = "vendor/personnel-merge.cjs";
const backendVendorDir = path.join("..", "army-backend", "vendor");

await esbuild.build({
  entryPoints: ["src/personnel-merge/server-entry.ts"],
  bundle: true,
  platform: "node",
  format: "cjs",
  outfile: outFile,
  define: {
    "import.meta.env.BASE_URL": '"/"',
    "import.meta.env.VITE_API_BASE_URL": '""',
  },
});

mkdirSync(backendVendorDir, { recursive: true });
copyFileSync(outFile, path.join(backendVendorDir, "personnel-merge.cjs"));
console.log(`Built ${outFile} and copied to army-backend/vendor/`);
