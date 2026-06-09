import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const STORAGE_FILE = path.resolve(__dirname, "../data/storage.json");

const toBuffer = (b64) => Buffer.from(String(b64 || ""), "base64");
const toBase64 = (buf) => Buffer.from(buf).toString("base64");

const runPowerShell = (script) =>
  new Promise((resolve, reject) => {
    const ps = spawn("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script], {
      windowsHide: true
    });
    let stderr = "";
    ps.stderr.on("data", (d) => { stderr += String(d || ""); });
    ps.on("close", (code) => {
      if (code === 0) resolve(true);
      else reject(new Error(stderr || `PowerShell failed ${code}`));
    });
  });

const pdfPlugin = () => ({
  name: "syq-pdf-plugin",
  apply: "serve",
  configureServer(server) {
    server.middlewares.use("/api/quote/convert-pdf", async (req, res) => {
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.end(JSON.stringify({ ok: false, message: "Method Not Allowed" }));
        return;
      }
      let body = "";
      req.on("data", (c) => { body += c; });
      req.on("end", async () => {
        let tempXlsx = "";
        let tempPdf = "";
        try {
          const { xlsxBase64 } = JSON.parse(body || "{}");
          if (!xlsxBase64) throw new Error("xlsxBase64 required");
          const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "syq-"));
          tempXlsx = path.join(tmpDir, `${randomUUID()}.xlsx`);
          tempPdf = path.join(tmpDir, `${randomUUID()}.pdf`);
          await fs.writeFile(tempXlsx, toBuffer(xlsxBase64));
          const ps = `
            $ErrorActionPreference='Stop'
            $excel=New-Object -ComObject Excel.Application
            $excel.Visible=$false; $excel.DisplayAlerts=$false
            try {
              $wb=$excel.Workbooks.Open('${tempXlsx.replace(/\\/g, "\\\\")}')
              $ws=$wb.Worksheets.Item(1)
              $ws.PageSetup.Zoom=$false
              $ws.PageSetup.FitToPagesWide=1
              $ws.PageSetup.FitToPagesTall=1
              $wb.ExportAsFixedFormat(0,'${tempPdf.replace(/\\/g, "\\\\")}')
              $wb.Close($false)
            } finally { $excel.Quit() }
          `;
          await runPowerShell(ps);
          const pdfBuf = await fs.readFile(tempPdf);
          res.statusCode = 200;
          res.setHeader("Content-Type", "application/json");
          res.end(JSON.stringify({ ok: true, pdfBase64: toBase64(pdfBuf) }));
        } catch (e) {
          res.statusCode = 500;
          res.setHeader("Content-Type", "application/json");
          res.end(JSON.stringify({ ok: false, message: e?.message || "PDF failed" }));
        } finally {
          if (tempXlsx) {
            try { await fs.rm(path.dirname(tempXlsx), { recursive: true, force: true }); } catch {}
          }
        }
      });
    });
  }
});

const storagePlugin = () => ({
  name: "syq-storage-plugin",
  apply: "serve",
  configureServer(server) {
    server.middlewares.use("/api/storage/sync", (req, res) => {
      if (req.method !== "POST") { res.statusCode = 405; res.end("{}"); return; }
      let body = "";
      req.on("data", (c) => { body += c; });
      req.on("end", async () => {
        try {
          await fs.writeFile(STORAGE_FILE, body, "utf-8");
          res.statusCode = 200;
          res.setHeader("Content-Type", "application/json");
          res.end(JSON.stringify({ ok: true }));
        } catch (e) {
          res.statusCode = 500;
          res.end(JSON.stringify({ ok: false, message: e.message }));
        }
      });
    });
  }
});

export default defineConfig({
  plugins: [react(), storagePlugin(), pdfPlugin()],
  base: "/SYQ/",
  server: {
    port: 5274
  },
  build: {
    outDir: "dist",
    emptyOutDir: true
  }
});
