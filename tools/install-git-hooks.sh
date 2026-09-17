#!/bin/sh
set -eu
cd "$(git rev-parse --show-toplevel)"
git config core.hooksPath .githooks
printf '%s\n' 'Git hooks aktif: .githooks (artefak offline dibuat ulang sebelum commit).'
