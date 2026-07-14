@echo off
echo Starting local server on http://localhost:8743 ...
start http://localhost:8743
py -3 -m http.server 8743
