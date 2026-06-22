@echo off
setlocal
set NODE=C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe
set APP=C:\Users\ASUS\Documents\Codex\2026-06-20\chrome-plugin-chrome-openai-bundled-file-2\outputs\webapp\server.mjs
cd /d C:\Users\ASUS\Documents\Codex\2026-06-20\chrome-plugin-chrome-openai-bundled-file-2\outputs\webapp
"%NODE%" "%APP%"
