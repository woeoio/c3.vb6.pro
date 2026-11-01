#!/usr/bin/env bash
# =============================================================================
# build.sh - fast build entry for C3 (bash side)
#
# Why this exists: build.bat drives the build through cmd.exe, and cmd.exe in a
# vcvars-populated environment spawns child processes at ~82 ms each, versus
# ~17 ms from bash.  A no-op build therefore costs 9.3 s through build.bat and
# ~0.1 s through here; a one-file rebuild costs 8.9 s versus 1.4 s (with PCH).
#
# What it does:
#   1. locates Visual Studio (vswhere, same rule as build.bat) -- cached, the
#      lookup only reruns when the toolchain fingerprint changes
#   2. reuses a cached MSVC environment snapshot (.build/msvc_env.sh); it is
#      regenerated only when the toolchain moves (VS upgraded / repaired /
#      moved, or the Windows SDK version changed)
#   3. runs ninja directly (no `cmake --build` re-entry, no second vcvars)
#
# Usage:
#   ./build.sh                build .build (default)
#   ./build.sh --clean        reconfigure + build
#   ./build.sh --dir <dir>    build a different build dir
#   ./build.sh --ninja <args> pass extra args straight to ninja
#
# On a fresh clone git does not carry the executable bit (core.filemode is off
# on Windows), so use `bash build.sh` if ./build.sh is not runnable.
# =============================================================================
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

BDIR=".build"
CLEAN=0
NINJA_ARGS=()

while [ $# -gt 0 ]; do
  case "$1" in
    --clean) CLEAN=1; shift ;;
    --dir)   BDIR="${2:?--dir needs an argument}"; shift 2 ;;
    --ninja) shift; NINJA_ARGS=("$@"); break ;;
    -h|--help) sed -n '2,25p' "${BASH_SOURCE[0]}"; exit 0 ;;
    *) echo "[build] unknown argument: $1" >&2; exit 1 ;;
  esac
done

mkdir -p "$BDIR"
CACHE="$BDIR/.toolchain_cache"
ENV_SH="$BDIR/msvc_env.sh"
ENV_FP="$BDIR/.msvc_env_fingerprint"

# --- toolchain discovery (cached) -----------------------------------------
# The cache holds where vswhere lives, which VS to use, the MSVC version and
# the vcvars64.bat timestamp.  A hit costs one stat; anything else reruns the
# full discovery.
VSINST=""; MSVC_VER=""; VSWHERE=""; VCVARS=""; MTIME=""
if [ -f "$CACHE" ]; then
  # bash builtin read: spawning `cat` here costs ~100 ms in Git Bash.
  IFS='|' read -r C_VSW C_VSINST C_MVCVER C_MTIME < "$CACHE" || true
  if [ -n "${C_VSINST:-}" ] && [ -f "$C_VSINST/VC/Auxiliary/Build/vcvars64.bat" ]; then
    NOW_MTIME="$(stat -c %Y "$C_VSINST/VC/Auxiliary/Build/vcvars64.bat" 2>/dev/null || echo 0)"
    if [ "$NOW_MTIME" = "${C_MTIME:-}" ]; then
      VSWHERE="$C_VSW"; VSINST="$C_VSINST"; MSVC_VER="$C_MVCVER"; MTIME="$NOW_MTIME"
      VCVARS="$VSINST/VC/Auxiliary/Build/vcvars64.bat"
    fi
  fi
fi

if [ -z "$VSINST" ]; then
  # bash cannot expand ${ProgramFiles(x86)} (parentheses are not valid in a
  # parameter name), so ask cmd.exe for the value and convert it.
  PF86="$(cmd //c 'echo %ProgramFiles(x86)%' 2>/dev/null | tr -d '\r' | tail -1)"
  for cand in \
      "${PF86}/Microsoft Visual Studio/Installer/vswhere.exe" \
      "/c/Program Files (x86)/Microsoft Visual Studio/Installer/vswhere.exe" \
      "/d/Program Files (x86)/Microsoft Visual Studio/Installer/vswhere.exe" \
      "/c/Program Files/Microsoft Visual Studio/Installer/vswhere.exe" \
      "/d/Program Files/Microsoft Visual Studio/Installer/vswhere.exe"; do
    [ -f "$cand" ] && { VSWHERE="$cand"; break; }
  done
  [ -n "$VSWHERE" ] || { echo "[build] ERROR: vswhere.exe not found" >&2; exit 1; }

  VSINST="$("$VSWHERE" -latest -products '*' \
             -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 \
             -property installationPath 2>/dev/null | tr -d '\r' | head -1)"
  [ -n "$VSINST" ] || { echo "[build] ERROR: no Visual Studio C++ toolset found" >&2; exit 1; }

  VCVARS="$VSINST/VC/Auxiliary/Build/vcvars64.bat"
  [ -f "$VCVARS" ] || { echo "[build] ERROR: $VCVARS missing" >&2; exit 1; }

  MSVC_VER="$(ls "$VSINST/VC/Tools/MSVC" 2>/dev/null | sort -V | tail -1)"
  MTIME="$(stat -c %Y "$VCVARS" 2>/dev/null || echo 0)"
  printf '%s|%s|%s|%s' "$VSWHERE" "$VSINST" "$MSVC_VER" "$MTIME" > "$CACHE"
fi

# --- Ninja always comes from the VS install, never from PATH --------------
NINJA="$VSINST/Common7/IDE/CommonExtensions/Microsoft/CMake/Ninja/ninja.exe"
[ -f "$NINJA" ] || { echo "[build] ERROR: ninja.exe not found under $VSINST" >&2; exit 1; }

# --- environment snapshot --------------------------------------------------
FINGERPRINT="$VSINST|$MSVC_VER|$MTIME"
if [ ! -f "$ENV_SH" ] || [ ! -f "$ENV_FP" ] || [ "$(<"$ENV_FP")" != "$FINGERPRINT" ]; then
  echo "[build] regenerating MSVC env snapshot (toolchain changed or missing)..."
  DUMP_WIN="$(cygpath -w "$ROOT/.temp/msvc_env_dump.txt" 2>/dev/null || echo "$ROOT/.temp/msvc_env_dump.txt")"
  BAT_WIN="$(cygpath -w "$ROOT/.temp/_dump_env.bat" 2>/dev/null || echo "$ROOT/.temp/_dump_env.bat")"
  mkdir -p "$(dirname "$DUMP_WIN")"
  # vcvars64.bat only runs under cmd.exe; dump the environment it produces.
  {
    echo "@echo off"
    echo "call \"$VCVARS\" >nul"
    echo "set > \"$DUMP_WIN\""
  } > "$ROOT/.temp/_dump_env.bat"
  cmd //c "$BAT_WIN" >/dev/null 2>&1 || { echo "[build] ERROR: vcvars64.bat failed" >&2; exit 1; }
  [ -f "$DUMP_WIN" ] || { echo "[build] ERROR: env dump missing" >&2; exit 1; }
  python3 "$ROOT/scripts/gen_msvc_env_sh.py" "$DUMP_WIN" "$ENV_SH" "$VSINST" "$MSVC_VER" \
    || { echo "[build] ERROR: env snapshot generation failed" >&2; exit 1; }
  printf '%s' "$FINGERPRINT" > "$ENV_FP"
fi

# --- load the snapshot -----------------------------------------------------
# shellcheck disable=SC1090
. "$ENV_SH"

if [ "$CLEAN" = "1" ]; then
  CMAKE_EXE="$(command -v cmake || true)"
  if [ -z "$CMAKE_EXE" ]; then
    CMAKE_EXE="$VSINST/Common7/IDE/CommonExtensions/Microsoft/CMake/CMake/bin/cmake.exe"
  fi
  "$CMAKE_EXE" -S . -B "$BDIR" -G Ninja -DCMAKE_BUILD_TYPE=Debug
fi

exec "$NINJA" -C "$BDIR" "${NINJA_ARGS[@]+"${NINJA_ARGS[@]}"}"
