#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
TOOLS_DIR="${ROOT_DIR}/tools"
TMP_DIR="$(mktemp -d)"
PANDOC_VERSION="${PANDOC_VERSION:-3.1.12.3}"
ARCHIVE="pandoc-${PANDOC_VERSION}-linux-amd64.tar.gz"
URL="https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/${ARCHIVE}"

cleanup() {
  rm -rf "${TMP_DIR}"
}
trap cleanup EXIT

mkdir -p "${TOOLS_DIR}"

echo "Downloading pandoc ${PANDOC_VERSION} with proxy disabled..."
env -u http_proxy -u https_proxy -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY -u all_proxy \
  curl -L --fail --retry 3 -o "${TMP_DIR}/${ARCHIVE}" "${URL}"

mkdir -p "${TMP_DIR}/pandoc"
tar -xzf "${TMP_DIR}/${ARCHIVE}" -C "${TMP_DIR}/pandoc" --strip-components=1
cp "${TMP_DIR}/pandoc/bin/pandoc" "${TOOLS_DIR}/pandoc"
chmod +x "${TOOLS_DIR}/pandoc"

echo "Installed local pandoc to: ${TOOLS_DIR}/pandoc"
"${TOOLS_DIR}/pandoc" --version | head -n 3
