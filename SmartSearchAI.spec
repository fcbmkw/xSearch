# -*- mode: python ; coding: utf-8 -*-
#
# v9.0: "fat" build. AI libraries (torch/transformers/sentence-transformers/
# einops/huggingface_hub) are bundled into the exe HERE, at build time, on a
# clean machine (GitHub Actions windows-latest runner) -- instead of being
# pip-installed on the user's machine at runtime. This is the whole reason
# this file no longer has an `excludes=[...]` list for those packages: we
# WANT PyInstaller to find and bundle them now.
#
# Two things this build needs beyond a normal PyInstaller run:
#
# 1. Full stdlib bundling (kept from the previous version of this file).
#    Even now that AI packages are bundled at build time, PyInstaller's
#    static analysis can still miss stdlib modules that those packages only
#    import deep at runtime (this is how we previously hit
#    "No module named 'filecmp'"). Cheap insurance: bundle the entire
#    stdlib instead of guessing module-by-module.
#
# 2. collect_all() for the AI packages. transformers/torch/tokenizers etc.
#    use a LOT of dynamic/lazy importing and ship non-.py data files
#    (tokenizer configs, compiled extensions, version metadata) that plain
#    static analysis does not reliably catch. collect_all() pulls in a
#    package's submodules, binaries, and data files all at once -- this is
#    the standard/recommended way to bundle ML libraries with PyInstaller.

import os
import sysconfig
from PyInstaller.utils.hooks import collect_all

block_cipher = None


def _collect_all_stdlib_modules():
    stdlib_dir = sysconfig.get_path('stdlib')
    mods = []
    skip_dirs = {'test', 'tests', 'idlelib', 'tkinter', 'turtledemo',
                 '__pycache__', 'lib2to3', 'ensurepip', 'venv'}
    for root, dirs, files in os.walk(stdlib_dir):
        dirs[:] = [d for d in dirs if d not in skip_dirs and not d.startswith('.')]
        rel_root = os.path.relpath(root, stdlib_dir)
        for fn in files:
            if not fn.endswith('.py'):
                continue
            if rel_root == '.':
                mod_rel = fn[:-3]
            else:
                mod_rel = os.path.join(rel_root, fn[:-3])
            mod = mod_rel.replace(os.sep, '.')
            if mod.endswith('.__init__'):
                mod = mod[: -len('.__init__')]
            mods.append(mod)
    return sorted(set(mods))


_stdlib_hidden_imports = _collect_all_stdlib_modules()

# Packages that need the full collect_all() treatment (submodules + binaries
# + data files). Add to this list if a future build error says
# "No module named 'X'" or a model complains about a missing tokenizer/
# config file that should have been bundled.
_AI_PACKAGES = [
    'torch',
    'transformers',
    'sentence_transformers',
    'tokenizers',
    'safetensors',
    'huggingface_hub',
    'einops',
    'regex',
]

_datas = []
_binaries = []
_hiddenimports = list(_stdlib_hidden_imports)

for _pkg in _AI_PACKAGES:
    try:
        _d, _b, _h = collect_all(_pkg)
        _datas += _d
        _binaries += _b
        _hiddenimports += _h
    except Exception as _e:
        print(f"[spec] WARNING: collect_all('{_pkg}') failed: {_e}")

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=_binaries,
    datas=_datas,
    hiddenimports=_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SmartSearchAI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,   # UPX-compressing a multi-GB AI build saves little and
                 # sometimes trips antivirus heuristics on corporate
                 # machines even more than an uncompressed exe -- leave off.
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # app is stable now -- hide the console window. Requires
                     # the sys.stdout/stderr None-guard at the top of app.py,
                     # or every print() call would crash the app instantly.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# v9.1: onedir instead of onefile. --onefile bundles everything into a
# single .exe, but that means EVERY launch has to re-extract the whole
# thing (380MB+ here) into a fresh %TEMP%\_MEIxxxxxx folder before it can
# even start running -- this was the exact cause of the 20-30s startup
# delay. onedir instead produces a folder (dist/SmartSearchAI/) with the
# .exe plus all its DLLs/data sitting right next to it, so launching just
# loads them directly -- no extraction step, dramatically faster startup.
# Distribute the whole dist/SmartSearchAI/ folder (zipped) instead of a
# single .exe file; SmartSearchAI.exe inside that folder is still the
# thing users double-click.
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='SmartSearchAI',
)
