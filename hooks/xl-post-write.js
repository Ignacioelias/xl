#!/usr/bin/env node
// Fast front door for the xl PostToolUse hook. Node starts in ~50 ms; Python on a typical corporate
// laptop takes ~1.2 s, so Python only runs when the tool call actually named an .xlsx/.xlsm.
// Everything substantive (mtime check, quick lint, additionalContext) stays in xl-post-write.py.
"use strict";
const { spawnSync } = require("child_process");
const path = require("path");

let raw = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (d) => (raw += d));
process.stdin.on("end", () => {
  if (process.env.XL_HOOK_DISABLE === "1") return;
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    return;
  }
  const tool = data.tool_name || "";
  const ti = data.tool_input || {};
  let text = "";
  if (tool === "Write" || tool === "Edit") text = ti.file_path || "";
  else if (tool === "Bash") text = ti.command || "";
  else return;
  if (!/\.xls[xm]\b/i.test(text)) return;
  // an xl read/verify command is not a write; only --save or exec can change a file
  if (tool === "Bash" && /^\s*(xl|py -3\.14 .*xlcli\.py)\b/.test(text) && !text.includes("--save") && !text.includes(" exec")) return;
  const r = spawnSync("py", ["-3.14", path.join(__dirname, "xl-post-write.py")], {
    input: raw,
    encoding: "utf8",
    timeout: 70000,
    windowsHide: true,
  });
  if (r.stdout) process.stdout.write(r.stdout);
});
