#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

echo "[INFO] Advanced Merger Launcher"

if ! command -v node >/dev/null 2>&1; then
    echo "[INFO] Node.js is not installed."

    if command -v apt >/dev/null 2>&1; then
        echo "[INFO] Debian/Ubuntu detected. Install Node.js with:"
        echo "       sudo apt update && sudo apt install -y nodejs npm"
    elif command -v dnf >/dev/null 2>&1; then
        echo "[INFO] Fedora/RHEL detected. Install Node.js with:"
        echo "       sudo dnf install -y nodejs npm"
    elif command -v yum >/dev/null 2>&1; then
        echo "[INFO] RHEL/CentOS detected. Install Node.js with:"
        echo "       sudo yum install -y nodejs npm"
    elif command -v pacman >/dev/null 2>&1; then
        echo "[INFO] Arch Linux detected. Install Node.js with:"
        echo "       sudo pacman -S nodejs npm"
    elif command -v zypper >/dev/null 2>&1; then
        echo "[INFO] openSUSE detected. Install Node.js with:"
        echo "       sudo zypper install nodejs npm"
    else
        echo "[ERROR] Please install Node.js LTS for your Linux distribution."
        echo "        Download: https://nodejs.org/"
    fi

    exit 1
fi

echo "[INFO] Starting Advanced Merger..."
node server.js
