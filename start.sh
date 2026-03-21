#!/usr/bin/env bash
# Convenience script: loads nvm, switches to the project's Node version, and starts the add-in.
set -e

export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && . "$NVM_DIR/nvm.sh"

nvm use
npm start
