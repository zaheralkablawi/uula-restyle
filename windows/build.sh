#!/bin/sh
# Builds the Windows installer (UULA-Restyle-Setup.exe) and the zip for the team.
cd "$(dirname "$0")"
cp ../UULA-Restyle.xml .
GOOS=windows GOARCH=amd64 CGO_ENABLED=0 go build -ldflags "-H windowsgui -s -w" -o UULA-Restyle-Setup.exe .
