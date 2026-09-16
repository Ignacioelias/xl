"""xl - persistent workbook REPL and Excel verifiers for Claude Code on Windows.

Layers:
  client/server  persistent Python kernel per session (state survives across tool calls)
  wbcache        openpyxl workbooks cached by path+mtime (open once, query many times)
  describe       one-call structural map of a workbook
  lint           openpyxl-only checks (fast, hook-safe)
  excelcom       real Excel engine via a dedicated hidden COM instance: calc, render, check-open
"""
__version__ = "0.9.0"
