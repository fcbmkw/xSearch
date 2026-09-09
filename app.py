import sys
import os
import types as _types
import io as _io

# v9.2/v10.17: required BEFORE anything else once the .spec is built with
# console=False (windowed mode, no visible console window). In that mode
# PyInstaller sets sys.stdout/sys.stderr to None (there is no console to
# write to) -- but this app calls print() extensively throughout for
# logging/debugging. Without this guard, the very first print() call would
# crash the app instantly with "AttributeError: 'NoneType' object has no
# attribute 'write'", before the GUI even has a chance to open.
#
# v10.17 FIX: the original os.devnull redirect below looked safe but
# wasn't -- open(os.devnull, "w") still uses the OS default text encoding
# (cp1252 on most Vietnamese/Windows installs), and print() encodes the
# string BEFORE the (discarded) write happens. Any Vietnamese diacritic
# (any char outside cp1252, e.g. '\u1ebf' = 'ế') in an f-string passed to
# print() then raised UnicodeEncodeError and crashed the whole app at
# import time -- same failure whether running the raw .py in a plain
# Windows console (cp1252 by default there too) or the frozen --noconsole
# exe. Forcing UTF-8 with errors="replace" on stdout/stderr fixes BOTH
# cases: real console output now prints Vietnamese correctly, and the
# devnull/log target can no longer choke on it either.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="replace")
else:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="replace")
else:
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        sys.stderr = _io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import math
import json
import configparser
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import webbrowser

# ═══════════════════════════════════════════════════════════════════════════
# BUILD FLAG — AI Search (offline embedding: Jina/BGE) + Vietnamese
# diacritics restoration + OCR image indexing.
#
# Set to False for the "for friends / nonDS" build that ships without
# torch/transformers/sentence_transformers/easyocr baked in (keeps the exe
# small — no need to remove any code). When False:
#   - The "🤖 AI Search" button and its model dropdown stay permanently
#     greyed out (see _sync_ai_adv_lock), no matter what the ramp light says.
#   - The "Search AI models" and "Vietnamese diacritics restoration"
#     sections in the Update DB dialog are greyed out entirely.
#   - BM25 search and AI Chat (online Gemini/Groq) are UNAFFECTED — this
#     flag only touches the offline embedding-model features above.
#
# Set to True (default) for the internal/full build.
# ═══════════════════════════════════════════════════════════════════════════
ENABLE_AI_SEARCH_FEATURE = True

import subprocess
import re
import sqlite3
import threading
import queue
import string
import unicodedata
import time
import tempfile  # v-new (yêu cầu 4): lưu tạm ảnh vừa dán từ clipboard trước khi đính kèm
from datetime import datetime
import pandas as pd 
import docx # pip install python-docx
import pypdf # pip install pypdf
from pptx import Presentation # pip install python-pptx
import logging
import warnings
warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.CRITICAL)
logging.getLogger().setLevel(logging.CRITICAL)
logging.getLogger("pypdf").setLevel(logging.CRITICAL)
logging.getLogger("pypdf._cmap").setLevel(logging.CRITICAL)
try:
    import xlrd
    xlrd.xlsx.ensure_elementtree_imported(False, False)
except:
    pass
logging.getLogger('xlrd').setLevel(logging.CRITICAL)

# --- CONFIGURATION ---
# Always find DB next to the exe (or .py script), regardless of working directory
if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)  # when running as exe
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # when running as .py

DB_FILE = os.path.join(_BASE_DIR, 'search_data.db')
# v7.10 FIX: Search History used to be logged straight into DB_FILE
# (search_data.db) itself. That meant just typing a query -- even with
# --update data NEVER run -- created/grew search_data.db (a few KB, just
# the 'history' table). On next launch the ramp light saw that file
# "exists" and jumped from Red to Yellow, even though nothing was ever
# actually indexed. Search History now lives in its own separate file so
# merely searching can never make the ramp light think indexing happened.
HISTORY_DB_FILE = os.path.join(_BASE_DIR, 'search_history.db')

# v-new (yêu cầu: 1 file configure.ini duy nhất, user tự mở/sửa bằng
# Notepad được -- gộp cả ngôn ngữ VI/EN/JP, AI Search model, Chat AI
# model, và API Key Gemini/Groq, thay cho app_settings.json + phần lưu
# API key riêng trong search_data.db trước đây): dùng configparser
# (built-in, không cần cài thêm gì) -- format .ini gọn, mỗi setting 1
# dòng "key = value", dễ đọc/sửa tay hơn hẳn json lồng nhau. Đặt cạnh
# .exe/.py giống DB_FILE/HISTORY_DB_FILE để không phụ thuộc working
# directory lúc khởi động.
CONFIG_FILE = os.path.join(_BASE_DIR, 'configure.ini')


def _load_config():
    """Đọc configure.ini, trả về 1 configparser.ConfigParser (rỗng nếu
    chưa tồn tại/lỗi/user sửa tay bị sai cú pháp) -- không bao giờ raise,
    để 1 file config bị gõ sai không thể chặn app khởi động."""
    cfg = configparser.ConfigParser()
    try:
        cfg.read(CONFIG_FILE, encoding='utf-8')
    except Exception as e:
        print(f"[Config] configure.ini có lỗi cú pháp, bỏ qua và dùng "
              f"mặc định cho tới khi được ghi lại: {e}")
    return cfg


def _save_config(mutate_fn):
    """Đọc configure.ini hiện có, cho mutate_fn(cfg) sửa nó (thêm/xoá/đổi
    section & key), rồi ghi lại nguyên file -- giữ lại MỌI setting khác
    (kể cả những dòng user tự thêm tay) mà mutate_fn không đụng tới."""
    try:
        cfg = _load_config()
        mutate_fn(cfg)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            cfg.write(f)
    except Exception as e:
        print(f"[Config] failed to save configure.ini: {e}")


def _config_get(section, key, fallback=None):
    cfg = _load_config()
    if cfg.has_option(section, key):
        return cfg.get(section, key)
    return fallback


def _config_set(section, key, value):
    def _mutate(cfg):
        if not cfg.has_section(section):
            cfg.add_section(section)
        cfg.set(section, key, str(value))
    _save_config(_mutate)


def _numbered_keys_from_config(cfg, prefix):
    """Đọc mọi key dạng '<prefix><N> = <value>' trong section [APIKeys]
    (vd gemini_key_1, gemini_key_2, ...), sắp theo N tăng dần -- cho phép
    user tự thêm 'gemini_key_3 = ...' bằng tay trong configure.ini mà
    không cần mở app."""
    if not cfg.has_section("APIKeys"):
        return []
    items = []
    for k, v in cfg.items("APIKeys"):
        if k.startswith(prefix) and v.strip():
            try:
                n = int(k[len(prefix):])
            except ValueError:
                n = 999
            items.append((n, v.strip()))
    items.sort(key=lambda t: t[0])
    return [v for _, v in items]


def _save_numbered_keys_to_config(prefix, keys):
    """Ghi lại TOÀN BỘ danh sách key của 1 provider (Gemini hoặc Groq)
    vào [APIKeys], xoá sạch các dòng '<prefix>N' cũ trước rồi đánh số lại
    từ 1 -- tránh để lại key rỗng/trùng số khi user xoá bớt 1 key ở giữa
    danh sách trên UI."""
    def _mutate(cfg):
        if not cfg.has_section("APIKeys"):
            cfg.add_section("APIKeys")
        for k in list(cfg.options("APIKeys")):
            if k.startswith(prefix):
                cfg.remove_option("APIKeys", k)
        for i, key in enumerate([k for k in keys if k.strip()], start=1):
            cfg.set("APIKeys", f"{prefix}{i}", key.strip())
    _save_config(_mutate)



# v-merge: outlook_search.py and onenote_search.py used to be separate
# files sitting next to this script (imported normally below). Now that
# testing is done, all three are merged into this ONE file for tidiness --
# their full source is embedded (base64, to sidestep any quoting/escaping
# issues with their own docstrings/strings) and exec'd into real module
# objects at startup, so every existing `outlook_search.xxx` /
# `onenote_search.xxx` call site elsewhere in this file keeps working
# completely unchanged. search_outlook.db / search_onenote.db / search_data.db
# remain three separate files, same as before -- only the .py source moved.
import types as _mod_types
import base64 as _b64

_OUTLOOK_SEARCH_SRC_B64 = (
    "IiIiCm91dGxvb2tfc2VhcmNoLnB5IOKAlCBJbmRleCAmIHNlYXJjaCBPdXRsb29rIG1haWwgY29u"
    "dGVudCBsb2NhbGx5IG92ZXIgQ09NLAp3aXRob3V0IGV2ZXIgZXhwb3J0aW5nIG1haWwgdG8gLm1z"
    "ZyBhbmQgd2l0aG91dCBkZXBlbmRpbmcgb24gd2hldGhlcgpXaW5kb3dzIFNlYXJjaCBoYXMgaW5k"
    "ZXhlZCBPdXRsb29rIG9uIHRoaXMgbWFjaGluZSAodGhhdCdzIHBlci1QQyBhbmQKb2Z0ZW4gZGlz"
    "YWJsZWQvYnJva2VuIG9uIHdvcmsgbGFwdG9wcyDigJQgdGhpcyB3b3JrcyB0aGUgc2FtZSBvbiBl"
    "dmVyeQptYWNoaW5lIHRoYXQganVzdCBoYXMgT3V0bG9vayBpbnN0YWxsZWQgYW5kIHNpZ25lZCBp"
    "bikuCgpEZXNpZ24gbm90ZXMgKHNvIHRoaXMgZml0cyBzbWFydFNlYXJjaF9NRlRfQk0yNV9BSSdz"
    "IGV4aXN0aW5nIGFyY2hpdGVjdHVyZSk6CiAgLSBPd24gc3FsaXRlIGRiIGZpbGUgKHNlYXJjaF9v"
    "dXRsb29rLmRiKSwgc2VwYXJhdGUgZnJvbSBzZWFyY2hfZGF0YS5kYiDigJQKICAgIHNhbWUgcmVh"
    "c29uaW5nIGFzIEhJU1RPUllfREJfRklMRTogc2VhcmNoaW5nIHNob3VsZCBuZXZlciBjcmVhdGUv"
    "Z3JvdwogICAgc2VhcmNoX2RhdGEuZGIsIHNpbmNlIHRoZSByYW1wLWxpZ2h0ICJpcyBEQiBidWls"
    "dD8iIGNoZWNrIGluIHRoZSBtYWluCiAgICBhcHAgbG9va3MgYXQgdGhhdCBmaWxlJ3MgZXhpc3Rl"
    "bmNlL3NpemUuCiAgLSBGVFM1IHdpdGggdGhlICdwb3J0ZXIgdW5pY29kZTYxJyB0b2tlbml6ZXIg"
    "KG1haWwgaXMgbm9ybWFsIHByb3NlLCBub3QKICAgIGZpbGVuYW1lcy9wYXRocywgc28gdGhlIHRy"
    "aWdyYW0gdG9rZW5pemVyIHVzZWQgZm9yIGNvbnRlbnRfaW5kZXggaXNuJ3QKICAgIHRoZSByaWdo"
    "dCBmaXQgaGVyZSDigJQgcG9ydGVyIGdpdmVzIHByb3BlciBCTTI1IHdvcmQtc3RlbSByYW5raW5n"
    "KS4KICAtIEluY3JlbWVudGFsOiBlYWNoIG1haWwncyBMYXN0TW9kaWZpY2F0aW9uVGltZSBpcyBz"
    "dG9yZWQ7IGEgcmUtaW5kZXgKICAgIG9ubHkgdG91Y2hlcyBpdGVtcyB0aGF0IGFjdHVhbGx5IGNo"
    "YW5nZWQgc2luY2UgdGhlIGxhc3QgcnVuLCBzbyByZXBlYXQKICAgICJVcGRhdGUgT3V0bG9vayBJ"
    "bmRleCIgcnVucyBhZnRlciB0aGUgZmlyc3QgYXJlIGZhc3QuCiAgLSBBbGwgT3V0bG9vayBDT00g"
    "Y2FsbHMgaGFwcGVuIG9uIHdoYXRldmVyIHRocmVhZCBjYWxscyBpbnRvIHRoaXMgbW9kdWxlLgog"
    "ICAgQ09NIGFwYXJ0bWVudHMgYXJlIHRocmVhZC1sb2NhbCDigJQgaWYgeW91IGNhbGwgaW5kZXhf"
    "b3V0bG9va19tYWlsKCkgb3IKICAgIHNlYXJjaF9vdXRsb29rKCkvb3Blbl9vdXRsb29rX2l0ZW0o"
    "KSBmcm9tIGEgYmFja2dyb3VuZCB0aHJlYWQgKHdoaWNoCiAgICB5b3Ugc2hvdWxkLCBmb3IgaW5k"
    "ZXhpbmcgYXQgbGVhc3Qg4oCUIHNlZSBpbnRlZ3JhdGlvbiBub3RlcyksIHRoYXQKICAgIHRocmVh"
    "ZCBtdXN0IGJlIHRoZSBvbmUgdGhhdCBvd25zIHRoZSBDb0luaXRpYWxpemUoKSBjYWxsLiBUaGlz"
    "IG1vZHVsZQogICAgaGFuZGxlcyB0aGF0IGludGVybmFsbHkgcGVyLWNhbGwsIHNvIHlvdSBkb24n"
    "dCBuZWVkIHRvIHRoaW5rIGFib3V0IGl0CiAgICBmcm9tIHRoZSBjYWxsZXIncyBzaWRlLgoKUmVx"
    "dWlyZXM6IHB5d2luMzIgKHBpcCBpbnN0YWxsIHB5d2luMzIpIGFuZCBhIGxvY2FsIE91dGxvb2sg"
    "ZGVza3RvcAppbnN0YWxsIHNpZ25lZCBpbnRvIHRoZSBhY2NvdW50IHlvdSB3YW50IHRvIHNlYXJj"
    "aCAoQ2xhc3NpYyBvciBOZXcKT3V0bG9vaydzIHVuZGVybHlpbmcgQ09NIHN1cmZhY2Ug4oCUIE5l"
    "dyBPdXRsb29rIGN1cnJlbnRseSBkb2VzIG5vdCBleHBvc2UKdGhpcyBDT00gQVBJOyBDbGFzc2lj"
    "IE91dGxvb2sgZG9lcykuCiIiIgoKaW1wb3J0IG9zCmltcG9ydCBzcWxpdGUzCmltcG9ydCB0aW1l"
    "CmltcG9ydCB0cmFjZWJhY2sKCnByaW50KCJbb3V0bG9va19zZWFyY2hdIG1vZHVsZSBsb2FkZWQg"
    "4oCUIGJ1aWxkIHRhZzogc3RvcmVpZC1maXgtdjUgKHJvb3QgY2F1c2UgZml4ZWQpIikKCl9pbnNl"
    "cnRfZXJyb3JfY291bnQgPSAwICAjIHRocm90dGxlcyBkaWFnbm9zdGljIGVycm9yIHByaW50aW5n"
    "IOKAlCBzZWUgaW5kZXhfb3V0bG9va19tYWlsKCkKCnRyeToKICAgIGltcG9ydCBweXRob25jb20K"
    "ICAgIGltcG9ydCBweXdpbnR5cGVzCiAgICBpbXBvcnQgd2luMzJjb20uY2xpZW50CiAgICBpbXBv"
    "cnQgd2luMzJjb20uc2VydmVyLnV0aWwKICAgIE9VVExPT0tfQVZBSUxBQkxFID0gVHJ1ZQpleGNl"
    "cHQgSW1wb3J0RXJyb3I6CiAgICBPVVRMT09LX0FWQUlMQUJMRSA9IEZhbHNlCgpfQkFTRV9ESVIg"
    "PSBvcy5wYXRoLmRpcm5hbWUob3MucGF0aC5hYnNwYXRoKF9fZmlsZV9fKSkKT1VUTE9PS19EQl9G"
    "SUxFID0gb3MucGF0aC5qb2luKF9CQVNFX0RJUiwgJ3NlYXJjaF9vdXRsb29rLmRiJykKCiMgb2xN"
    "YWlsID0gNDMgaW4gT3V0bG9vaydzIE9sT2JqZWN0Q2xhc3MgZW51bS4gT3RoZXIgaXRlbSBjbGFz"
    "c2VzCiMgKGFwcG9pbnRtZW50cywgY29udGFjdHMsIHRhc2tzLi4uKSBhcmUgc2tpcHBlZCDigJQg"
    "dGhpcyBpcyBhIG1haWwgc2VhcmNoLgpPTF9NQUlMX0NMQVNTID0gNDMKCiMgUlBDL0NPTSByZWpl"
    "Y3QtY2FsbCByZWFzb24gY29kZXMgdXNlZCBieSBSZXRyeVJlamVjdGVkQ2FsbCBiZWxvdyAoZnJv"
    "bQojIHRoZSBXaW5kb3dzIFNESydzIG9sZWlkbC5oIOKAlCBweXdpbjMyIGRvZXNuJ3QgZXhwb3Nl"
    "IHRoZXNlIGFzIGNvbnN0YW50cykuCl9TRVJWRVJDQUxMX0lTSEFORExFRCA9IDAKX1NFUlZFUkNB"
    "TExfUkVKRUNURUQgPSAxCl9TRVJWRVJDQUxMX1JFVFJZTEFURVIgPSAyCl9QRU5ESU5HTVNHX1dB"
    "SVRERUZQUk9DRVNTID0gMgoKIyB2LWZpeDogYXV0b21hdGluZyBPdXRsb29rIGZyb20gYSBiYWNr"
    "Z3JvdW5kIHRocmVhZCAodGhpcyBtb2R1bGUgYWx3YXlzCiMgcnVucyB1bmRlciBpdHMgb3duIHB5"
    "dGhvbmNvbS5Db0luaXRpYWxpemUoKSwgb24gd2hhdGV2ZXIgdGhyZWFkIHRoZQojIGNhbGxlciB1"
    "c2VzIOKAlCBzZWUgbW9kdWxlIGRvY3N0cmluZykgY2FuIGhpdCBhIHdlbGwta25vd24gQ09NIHJl"
    "ZW50cmFuY3kKIyBzdGFsbDogT3V0bG9vaydzIG93biBTVEEgb2NjYXNpb25hbGx5IHJlamVjdHMg"
    "YW4gaW5jb21pbmcgY2FsbCAoYnVzeQojIHNob3dpbmcgYSBkaWFsb2csIG1pZC1zeW5jLCBtb21l"
    "bnRhcmlseSByZWVudGVyZWQsIGV0Yy4pIGFuZCBhc2tzIHRoZQojIENBTExFUiAidHJ5IGFnYWlu"
    "IGxhdGVyIiB2aWEgUmV0cnlSZWplY3RlZENhbGwuIEEgcHJvY2VzcyB3aXRoIG5vCiMgSU1lc3Nh"
    "Z2VGaWx0ZXIgcmVnaXN0ZXJlZCBnZXRzIENPTSdzIGRlZmF1bHQgaGFuZGxpbmcgZm9yIHRoaXMs"
    "IHdoaWNoIGluCiMgcHJhY3RpY2UgY2FuIGVuZCB1cCB3YWl0aW5nIGluZGVmaW5pdGVseSBvbiBz"
    "b21lIE91dGxvb2svV2luZG93cyBidWlsZHMKIyAtLSBubyBleGNlcHRpb24sIG5vIHRpbWVvdXQs"
    "IGxvdyBDUFUgLS0gZXhhY3RseSBtYXRjaGluZyAicHJvY2Vzc2VkID09CiMgdG90YWwgYnV0IG5l"
    "dmVyIGZpbmlzaGVzLCAuZGIgc2l6ZSBmcm96ZW4gZm9yIGhvdXJzIiB3aGlsZSBPdXRsb29rCiMg"
    "aXRzZWxmIHN0aWxsIGxvb2tzIHBlcmZlY3RseSByZXNwb25zaXZlIHRvIHRoZSB1c2VyLiBSZWdp"
    "c3RlcmluZyB0aGlzCiMgZmlsdGVyIG1ha2VzIHJldHJpZXMgaGFwcGVuIG9uIE9VUiB0ZXJtcyAo"
    "c2hvcnQgYm91bmRlZCBiYWNrb2ZmLCB0aGVuCiMgZ2l2ZSB1cCBvbiB0aGF0IG9uZSBjYWxsKSBp"
    "bnN0ZWFkIG9mIGJsb2NraW5nIGZvcmV2ZXIuCiMKIyB2LWZpeDI6IGJ1aWxkaW5nIHRoZSBJSUQg"
    "YW5kIGRlZmluaW5nL3JlZ2lzdGVyaW5nIHRoaXMgY2xhc3MgdHVybmVkIG91dAojIHRvIGJlIHB5"
    "d2luMzItdmVyc2lvbi1zZW5zaXRpdmUgKHB5dGhvbmNvbS5JSUQgLyBweXRob25jb20uSUlEX01l"
    "c3NhZ2UtCiMgRmlsdGVyIGRvbid0IGV4aXN0IG9uIGV2ZXJ5IGluc3RhbGwpIGFuZCBjcmFzaGVk"
    "IHRoZSBXSE9MRSBBUFAgYXQKIyBpbXBvcnQgdGltZSB0d2ljZSBpbiBhIHJvdyAtLSBjb21wbGV0"
    "ZWx5IGRpc3Byb3BvcnRpb25hdGUgZm9yIHdoYXQncwojIG1lYW50IHRvIGJlIGFuIG9wdGlvbmFs"
    "IHNhZmV0eSBuZXQuIEV2ZXJ5dGhpbmcgYmVsb3cgaXMgbm93IHdyYXBwZWQgc28KIyBBTlkgZmFp"
    "bHVyZSBoZXJlIGp1c3QgZGlzYWJsZXMgdGhpcyBvbmUgZmVhdHVyZSAocHJpbnRzIGEgd2Fybmlu"
    "ZywKIyBfTVNHRklMVEVSX0FWQUlMQUJMRSBzdGF5cyBGYWxzZSkgaW5zdGVhZCBvZiBldmVyIGJs"
    "b2NraW5nIHRoZSBhcHAgb3IKIyBPdXRsb29rIGluZGV4aW5nIGZyb20gd29ya2luZyBhdCBhbGws"
    "IHNhbWUgYXMgT1VUTE9PS19BVkFJTEFCTEUgaXRzZWxmLgpfTVNHRklMVEVSX0FWQUlMQUJMRSA9"
    "IEZhbHNlCmlmIE9VVExPT0tfQVZBSUxBQkxFOgogICAgdHJ5OgogICAgICAgICMgRml4ZWQgSU1l"
    "c3NhZ2VGaWx0ZXIgR1VJRCBmcm9tIHRoZSBXaW5kb3dzIFNESyAob2xlaWRsLmggLwogICAgICAg"
    "ICMgb2JqaWRsLmgpIC0tIHNhbWUgdmFsdWUgb24gZXZlcnkgV2luZG93cyBpbnN0YWxsLiBweXdp"
    "bnR5cGVzLklJRCgpCiAgICAgICAgIyBpcyB0aGUgYWN0dWFsIGNvbnN0cnVjdG9yIHRoYXQgZXhp"
    "c3RzIGFjcm9zcyBweXdpbjMyIHZlcnNpb25zIGZvcgogICAgICAgICMgdHVybmluZyB0aGF0IHN0"
    "cmluZyBpbnRvIGEgcmVhbCBJSUQgb2JqZWN0IChweXRob25jb20uSUlEKC4uLikKICAgICAgICAj"
    "IGFuZCBweXRob25jb20uSUlEX0lNZXNzYWdlRmlsdGVyIGRvIE5PVCBleGlzdCBpbiBldmVyeSB2"
    "ZXJzaW9uKS4KICAgICAgICBfSUlEX0lNZXNzYWdlRmlsdGVyID0gcHl3aW50eXBlcy5JSUQoJ3sw"
    "MDAwMDAxNi0wMDAwLTAwMDAtQzAwMC0wMDAwMDAwMDAwNDZ9JykKCiAgICAgICAgY2xhc3MgX091"
    "dGxvb2tNZXNzYWdlRmlsdGVyOgogICAgICAgICAgICBfY29tX2ludGVyZmFjZXNfID0gW19JSURf"
    "SU1lc3NhZ2VGaWx0ZXJdCiAgICAgICAgICAgIF9wdWJsaWNfbWV0aG9kc18gPSBbJ0hhbmRsZUlu"
    "Q29taW5nQ2FsbCcsICdSZXRyeVJlamVjdGVkQ2FsbCcsICdNZXNzYWdlUGVuZGluZyddCgogICAg"
    "ICAgICAgICBkZWYgSGFuZGxlSW5Db21pbmdDYWxsKHNlbGYsIGR3Q2FsbFR5cGUsIGh0YXNrQ2Fs"
    "bGVyLCBkd1RpY2tDb3VudCwgbHBJbnRlcmZhY2VJbmZvKToKICAgICAgICAgICAgICAgIHJldHVy"
    "biBfU0VSVkVSQ0FMTF9JU0hBTkRMRUQKCiAgICAgICAgICAgIGRlZiBSZXRyeVJlamVjdGVkQ2Fs"
    "bChzZWxmLCBodGFza0NhbGxlZSwgZHdUaWNrQ291bnQsIGR3UmVqZWN0VHlwZSk6CiAgICAgICAg"
    "ICAgICAgICAjIGR3VGlja0NvdW50IGlzIGhvdyBsb25nIFdFJ1ZFIGFscmVhZHkgYmVlbiB3YWl0"
    "aW5nIG9uIHRoaXMKICAgICAgICAgICAgICAgICMgb25lIGNhbGwsIGluIG1zLiBLZWVwIHJldHJ5"
    "aW5nIChzaG9ydCBwYXVzZXMpIGZvciB1cCB0bwogICAgICAgICAgICAgICAgIyB+MiBtaW51dGVz"
    "OyBwYXN0IHRoYXQsIGdpdmUgdXAgKC0xKSBzbyB0aGUgY2FsbCBmYWlscyB3aXRoCiAgICAgICAg"
    "ICAgICAgICAjIGFuIGV4Y2VwdGlvbiBpbnN0ZWFkIG9mIGhhbmdpbmcgZm9yZXZlciAtLSBvdXIg"
    "ZXhpc3RpbmcKICAgICAgICAgICAgICAgICMgdHJ5L2V4Y2VwdCBhcm91bmQgZXZlcnkgcHJvcGVy"
    "dHkgYWNjZXNzIChzZWUKICAgICAgICAgICAgICAgICMgaW5kZXhfb3V0bG9va19tYWlsKSB0aGVu"
    "IGp1c3Qgc2tpcHMgdGhhdCBvbmUgaXRlbS9mb2xkZXIKICAgICAgICAgICAgICAgICMgYW5kIG1v"
    "dmVzIG9uLCByYXRoZXIgdGhhbiB0aGUgd2hvbGUgcnVuIHNpdHRpbmcgZnJvemVuLgogICAgICAg"
    "ICAgICAgICAgaWYgZHdSZWplY3RUeXBlID09IF9TRVJWRVJDQUxMX1JFVFJZTEFURVI6CiAgICAg"
    "ICAgICAgICAgICAgICAgaWYgZHdUaWNrQ291bnQgPiAxMjBfMDAwOgogICAgICAgICAgICAgICAg"
    "ICAgICAgICByZXR1cm4gLTEgICMgY2FuY2VsIHRoZSBjYWxsCiAgICAgICAgICAgICAgICAgICAg"
    "cmV0dXJuIDI1MCAgIyByZXRyeSBhZ2FpbiBpbiAyNTBtcwogICAgICAgICAgICAgICAgcmV0dXJu"
    "IC0xICAjIFNFUlZFUkNBTExfUkVKRUNURUQgb3IgYW55dGhpbmcgZWxzZSAtLSBkb24ndCByZXRy"
    "eQoKICAgICAgICAgICAgZGVmIE1lc3NhZ2VQZW5kaW5nKHNlbGYsIGh0YXNrQ2FsbGVlLCBkd1Rp"
    "Y2tDb3VudCwgZHdQZW5kaW5nVHlwZSk6CiAgICAgICAgICAgICAgICByZXR1cm4gX1BFTkRJTkdN"
    "U0dfV0FJVERFRlBST0NFU1MKCiAgICAgICAgX01TR0ZJTFRFUl9BVkFJTEFCTEUgPSBUcnVlCiAg"
    "ICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgcHJpbnQoZiJbT3V0bG9va10gSU1lc3Nh"
    "Z2VGaWx0ZXIgc2V0dXAgbm90IGF2YWlsYWJsZSBvbiB0aGlzIHB5d2luMzIvV2luZG93cyAiCiAg"
    "ICAgICAgICAgICAgZiJpbnN0YWxsIChub24tZmF0YWwsIGluZGV4aW5nIHN0aWxsIHJ1bnMgd2l0"
    "aG91dCBpdCk6IHtlIXJ9IikKCgpkZWYgX3JlZ2lzdGVyX21lc3NhZ2VfZmlsdGVyKCk6CiAgICAi"
    "IiJCZXN0LWVmZm9ydCAtLSBpZiB0aGlzIGZhaWxzIGZvciBhbnkgcmVhc29uLCBpbmRleGluZyBz"
    "dGlsbCBydW5zCiAgICBleGFjdGx5IGFzIGJlZm9yZSAoanVzdCB3aXRob3V0IHRoZSBleHRyYSBw"
    "cm90ZWN0aW9uIGFnYWluc3QgYSBzdHVjawogICAgcmVlbnRyYW50IGNhbGwpLiBNdXN0IGJlIGNh"
    "bGxlZCBBRlRFUiBweXRob25jb20uQ29Jbml0aWFsaXplKCkgb24gdGhlCiAgICBzYW1lIHRocmVh"
    "ZCB0aGF0IHdpbGwgbWFrZSB0aGUgT3V0bG9vayBDT00gY2FsbHMuIiIiCiAgICBpZiBub3QgX01T"
    "R0ZJTFRFUl9BVkFJTEFCTEU6CiAgICAgICAgcmV0dXJuCiAgICB0cnk6CiAgICAgICAgcHl0aG9u"
    "Y29tLkNvUmVnaXN0ZXJNZXNzYWdlRmlsdGVyKAogICAgICAgICAgICB3aW4zMmNvbS5zZXJ2ZXIu"
    "dXRpbC53cmFwKF9PdXRsb29rTWVzc2FnZUZpbHRlcigpLCBfSUlEX0lNZXNzYWdlRmlsdGVyKSkK"
    "ICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICBwcmludChmIltPdXRsb29rXSBDb3Vs"
    "ZCBub3QgcmVnaXN0ZXIgSU1lc3NhZ2VGaWx0ZXIgKG5vbi1mYXRhbCwgY29udGludWluZyB3aXRo"
    "b3V0IGl0KToge2Uhcn0iKQoKCmRlZiBfY29ubmVjdCgpOgogICAgY29ubiA9IHNxbGl0ZTMuY29u"
    "bmVjdChPVVRMT09LX0RCX0ZJTEUpCiAgICBjID0gY29ubi5jdXJzb3IoKQogICAgIyB2LWZpeCAo"
    "VuG6pW4gxJHhu4E6IGPhu6VtIHThu6sgdGnhur9uZyBOaOG6rXQga2jDtG5nIGPDsyBraG/huqNu"
    "ZyB0cuG6r25nL2ThuqV1IGPDonUgbmfEg24KICAgICMgY8OhY2gsIHbDrSBk4bulICLjgrvjgq3j"
    "g6Xjg6rjg4bjgqPjg5Hjg4Pjg4EiIG7hurFtIEdJ4buuQSBjw6J1IG5oxrAgIi4uLuOCteODvOOD"
    "kOOBuOOBrgogICAgIyDjgrvjgq3jg6Xjg6rjg4bjgqPjg5Hjg4Pjg4HpgannlKjjgavjgaTjgYTj"
    "gaYiLCBraMO0bmcgdMOsbSDEkcaw4bujYyB0cm9uZyBu4buZaSBkdW5nIG1haWwgZMO5CiAgICAj"
    "IGVtYWlsIMSRw7MgdGjhu7FjIHPhu7EgY2jhu6lhIMSRw7puZyBj4bulbSBuw6B5KTogdG9rZW5p"
    "emVyICdwb3J0ZXIgdW5pY29kZTYxJyBn4buZcAogICAgIyBj4bqjIDEgY2h14buXaSBrw70gdOG7"
    "sSBDSksgbGnDqm4gdOG7pWMgKGtow7RuZyBi4buLIG5n4bqvdCBi4bufaSBraG/huqNuZyB0cuG6"
    "r25nL2ThuqV1IGPDonUpCiAgICAjIHRow6BuaCBEVVkgTkjhuqRUIDEgdG9rZW4ga2hpIGluZGV4"
    "IOKAlCBt4buZdCBj4bulbSB04burIGtow7NhIG7hurFtIGdp4buvYSBkw7JuZyBuaMawCiAgICAj"
    "IHbhuq15IHPhur0ga2jDtG5nIGJhbyBnaeG7nSBraOG7m3AgTUFUQ0gsIGvhu4MgY+G6oyBraGkg"
    "bsOzIGto4bubcCBob8OgbiB0b8OgbiB0aGVvIG3huq90CiAgICAjIG5nxrDhu51pIMSR4buNYy4g"
    "Y29udGVudF9pbmRleCAoZmlsZSB0aMaw4budbmcsIHhlbSBhcHAgY2jDrW5oKSDEkcOjIGTDuW5n"
    "CiAgICAjICd0cmlncmFtJyDEkeG7gyBnaeG6o2kgcXV54bq/dCDEkcO6bmcgduG6pW4gxJHhu4Eg"
    "bsOgeSB04burIHRyxrDhu5tjIOKAlCDEkeG7lWkgb3V0bG9va19jb250ZW50CiAgICAjIHNhbmcg"
    "Y8O5bmcgdG9rZW5pemVyIGNobyBuaOG6pXQgcXXDoW4gdsOgIMSR4buDIHRp4bq/bmcgTmjhuq10"
    "L0NKSyB0w6xtIMSRxrDhu6NjIMSRw7puZwogICAgIyBuaMawIGZpbGUgdGjGsOG7nW5nICjEkcOh"
    "bmggxJHhu5VpOiBt4bqldCBraOG6oyBuxINuZyBzdGVtIHRp4bq/bmcgQW5oIGPhu6dhIHBvcnRl"
    "ciwKICAgICMgdsOtIGThu6UgInJ1biIga2jDtG5nIGPDsm4gdOG7sSBraOG7m3AgInJ1bm5pbmci"
    "IOKAlCBjaOG6pXAgbmjhuq1uIMSRxrDhu6NjLCB2w6wgbWFpbAogICAgIyB0aMaw4budbmcgxJHG"
    "sOG7o2Mgc2VhcmNoIGLhurFuZyBj4bulbSB04burL3TDqm4gcmnDqm5nIGNow61uaCB4w6FjIGjG"
    "oW4gbMOgIHThu6sgxJHGoW4pLgogICAgIwogICAgIyBmdHM1IEtIw5RORyBjaG8gxJHhu5VpIHRv"
    "a2VuaXplciBj4bunYSAxIHZpcnR1YWwgdGFibGUgxJHDoyB04buTbiB04bqhaSBi4bqxbmcKICAg"
    "ICMgQUxURVIg4oCUIG7hur91IHBow6F0IGhp4buHbiBEQiBjxakgKMSRxrDhu6NjIHThuqFvIGLh"
    "u59pIGLhuqNuIHRyxrDhu5tjLCB0b2tlbml6ZXIga2jDoWMKICAgICMgJ3RyaWdyYW0nKSwgcGjh"
    "uqNpIERST1AgcuG7k2kgdOG6oW8gbOG6oWkuIFhvw6EgbHXDtG4gb3V0bG9va19zdG9yZSBjw7lu"
    "ZyBsw7pjIMSR4buDCiAgICAjIGluZGV4X291dGxvb2tfbWFpbCgpIGNvaSBUT8OATiBC4buYIG1h"
    "aWwgbMOgICJt4bubaSIg4bufIGzhuqduIGNo4bqheSB0aeG6v3AgdGhlbwogICAgIyAoZOG7sWEg"
    "dHLDqm4gbGFzdF9tb2RpZmllZCBy4buXbmcpIHbDoCB04buxIMSR4buZbmcgcmVidWlsZCBs4bqh"
    "aSDEkeG6p3kgxJHhu6cgbuG7mWkgZHVuZwogICAgIyB0aGVvIHRva2VuaXplciBt4bubaSDigJQg"
    "Y2jhu4kgdOG7kW4gY8O0bmcgMSBs4bqnbiBkdXkgbmjhuqV0IG5nYXkgc2F1IGtoaSBj4bqtcAog"
    "ICAgIyBuaOG6rXQgY29kZSwgc2F1IMSRw7MgY8ahIGNo4bq/IGluY3JlbWVudGFsIHbhuqtuIGhv"
    "4bqhdCDEkeG7mW5nIGLDrG5oIHRoxrDhu51uZyBuaMawIGPFqS4KICAgIHRyeToKICAgICAgICBf"
    "cm93ID0gYy5leGVjdXRlKAogICAgICAgICAgICAiU0VMRUNUIHNxbCBGUk9NIHNxbGl0ZV9tYXN0"
    "ZXIgV0hFUkUgdHlwZT0ndGFibGUnIEFORCBuYW1lPSdvdXRsb29rX2NvbnRlbnQnIgogICAgICAg"
    "ICkuZmV0Y2hvbmUoKQogICAgICAgIGlmIF9yb3cgYW5kIF9yb3dbMF0gYW5kICd0cmlncmFtJyBu"
    "b3QgaW4gX3Jvd1swXToKICAgICAgICAgICAgYy5leGVjdXRlKCJEUk9QIFRBQkxFIElGIEVYSVNU"
    "UyBvdXRsb29rX2NvbnRlbnQiKQogICAgICAgICAgICBjLmV4ZWN1dGUoIkRST1AgVEFCTEUgSUYg"
    "RVhJU1RTIG91dGxvb2tfc3RvcmUiKQogICAgICAgICAgICBjb25uLmNvbW1pdCgpCiAgICBleGNl"
    "cHQgRXhjZXB0aW9uOgogICAgICAgIHBhc3MKICAgIGMuZXhlY3V0ZSgiIiJDUkVBVEUgVklSVFVB"
    "TCBUQUJMRSBJRiBOT1QgRVhJU1RTIG91dGxvb2tfY29udGVudAogICAgICAgICAgICAgICAgIFVT"
    "SU5HIGZ0czUoc3ViamVjdCwgYm9keSwgdG9rZW5pemU9J3RyaWdyYW0gY2FzZV9zZW5zaXRpdmUg"
    "MCcpIiIiKQogICAgYy5leGVjdXRlKCIiIkNSRUFURSBUQUJMRSBJRiBOT1QgRVhJU1RTIG91dGxv"
    "b2tfc3RvcmUgKAogICAgICAgICAgICAgICAgICAgIGVudHJ5X2lkICAgICAgVEVYVCBQUklNQVJZ"
    "IEtFWSwKICAgICAgICAgICAgICAgICAgICBzdG9yZV9pZCAgICAgIFRFWFQsCiAgICAgICAgICAg"
    "ICAgICAgICAgc3ViamVjdCAgICAgICBURVhULAogICAgICAgICAgICAgICAgICAgIHNlbmRlciAg"
    "ICAgICAgVEVYVCwKICAgICAgICAgICAgICAgICAgICByZWNlaXZlZCAgICAgIFRFWFQsCiAgICAg"
    "ICAgICAgICAgICAgICAgZm9sZGVyX3BhdGggICBURVhULAogICAgICAgICAgICAgICAgICAgIHNp"
    "emUgICAgICAgICAgSU5URUdFUiwKICAgICAgICAgICAgICAgICAgICBsYXN0X21vZGlmaWVkIFRF"
    "WFQsCiAgICAgICAgICAgICAgICAgICAgcm93aWRfZnRzICAgICBJTlRFR0VSCiAgICAgICAgICAg"
    "ICAgICAgKSIiIikKICAgICMgdi1maXg6IG9sZGVyIHNlYXJjaF9vdXRsb29rLmRiIGZpbGVzIGNy"
    "ZWF0ZWQgYmVmb3JlIHRoZSBzaXplIGNvbHVtbgogICAgIyBleGlzdGVkIHdvbid0IGhhdmUgaXQg"
    "4oCUIGFkZCBpdCBpbiBwbGFjZSBzbyByZS1ydW5uaW5nIGRvZXNuJ3QgcmVxdWlyZQogICAgIyBk"
    "ZWxldGluZyB0aGUgZGIgZmlsZS4KICAgIHRyeToKICAgICAgICBjLmV4ZWN1dGUoIkFMVEVSIFRB"
    "QkxFIG91dGxvb2tfc3RvcmUgQUREIENPTFVNTiBzaXplIElOVEVHRVIiKQogICAgZXhjZXB0IHNx"
    "bGl0ZTMuT3BlcmF0aW9uYWxFcnJvcjoKICAgICAgICBwYXNzICAjIGFscmVhZHkgaGFzIHRoZSBj"
    "b2x1bW4KICAgICMgdi1maXggKHJvb3QgY2F1c2Ugb2YgImV4ZWN1dGUocmFuaytzb3J0KT0zNHMi"
    "IG9uIHNlYXJjaF9vdXRsb29rKCkpOgogICAgIyBzZWFyY2hfb3V0bG9vaygpJ3MgSk9JTiAocy5y"
    "b3dpZF9mdHMgPSBvYy5yb3dpZCkgaGFkIE5PIGluZGV4IG9uCiAgICAjIHJvd2lkX2Z0cyAtLSBv"
    "bmx5IGVudHJ5X2lkICh1bnJlbGF0ZWQgY29sdW1uKSBpcyBhIFBSSU1BUlkgS0VZIC0tCiAgICAj"
    "IHNvIFNRTGl0ZSBoYWQgdG8gZnVsbC10YWJsZS1zY2FuIG91dGxvb2tfc3RvcmUgZm9yIGV2ZXJ5"
    "IG1hdGNoZWQKICAgICMgRlRTNSByb3cuIE9uIGEgMzUwTUIgZGIgd2l0aCB0ZW5zIG9mIHRob3Vz"
    "YW5kcyBvZiBtYWlsIHRoaXMgdHVybnMKICAgICMgYSBzdXBwb3NlZC10by1iZS1zdWItc2Vjb25k"
    "IEZUUzUgcXVlcnkgaW50byAzMHMrLCBhbmQgd29yc2UgYXMgdGhlCiAgICAjIG1haWxib3ggZ3Jv"
    "d3MuIENoZWFwIHRvIGNyZWF0ZSwgaHVnZSB3aW4gZm9yIGV2ZXJ5IGZ1dHVyZSBzZWFyY2guCiAg"
    "ICBjLmV4ZWN1dGUoIkNSRUFURSBJTkRFWCBJRiBOT1QgRVhJU1RTIGlkeF9vdXRsb29rX3N0b3Jl"
    "X3Jvd2lkX2Z0cyAiCiAgICAgICAgICAgICAgIk9OIG91dGxvb2tfc3RvcmUocm93aWRfZnRzKSIp"
    "CiAgICBjb25uLmNvbW1pdCgpCiAgICByZXR1cm4gY29ubiwgYwoKCmRlZiBfaXRlcl9mb2xkZXJz"
    "KHJvb3RfZm9sZGVyKToKICAgICIiIllpZWxkIHJvb3RfZm9sZGVyIGl0c2VsZiwgdGhlbiBldmVy"
    "eSBzdWJmb2xkZXIsIHJlY3Vyc2l2ZWx5CiAgICAoSW5ib3gsIFNlbnQgSXRlbXMsIGV2ZXJ5IGN1"
    "c3RvbSBmb2xkZXIvc3ViZm9sZGVyLCBldGMuKS4KCiAgICB2LWZpeDogcmV3cml0dGVuIHRvIHVz"
    "ZSBpbmRleGVkIGAuSXRlbShpKWAgYWNjZXNzIGluc3RlYWQgb2YKICAgIGBmb3Igc3ViIGluIHJv"
    "b3RfZm9sZGVyLkZvbGRlcnM6YC4gVGhlIGltcGxpY2l0IGZvci1sb29wIHVzZXMgQ09NJ3MKICAg"
    "IGVudW1lcmF0b3IgcHJvdG9jb2wgdW5kZXIgdGhlIGhvb2QgKElFbnVtVkFSSUFOVC5OZXh0KCkp"
    "IOKAlCB0aGlzIGlzCiAgICB0aGUgZXhhY3QgY2FsbCB0aGF0IHdhcyBvYnNlcnZlZCBoYW5naW5n"
    "IGluZGVmaW5pdGVseSAobm8gZXJyb3IsIG5vCiAgICB0aW1lb3V0LCBsb3cgQ1BVKSB3aGVuIE91"
    "dGxvb2sgaXMgYnVzeS9zeW5jaW5nLiBJbmRleGVkIC5JdGVtKGkpCiAgICBhY2Nlc3MgZ29lcyB0"
    "aHJvdWdoIGEgZGlmZmVyZW50LCBtb3JlIHJlbGlhYmxlIENPTSBjb2RlIHBhdGggYW5kCiAgICBk"
    "b2VzIG5vdCBleGhpYml0IHRoaXMgaGFuZy4iIiIKICAgIHlpZWxkIHJvb3RfZm9sZGVyCiAgICB0"
    "cnk6CiAgICAgICAgc3ViZm9sZGVycyA9IHJvb3RfZm9sZGVyLkZvbGRlcnMKICAgICAgICBjb3Vu"
    "dCA9IHN1YmZvbGRlcnMuQ291bnQKICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgcmV0dXJu"
    "CiAgICBmb3IgaSBpbiByYW5nZSgxLCBjb3VudCArIDEpOgogICAgICAgIHRyeToKICAgICAgICAg"
    "ICAgc3ViID0gc3ViZm9sZGVycy5JdGVtKGkpCiAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAg"
    "ICAgICAgICAgY29udGludWUKICAgICAgICB5aWVsZCBmcm9tIF9pdGVyX2ZvbGRlcnMoc3ViKQoK"
    "CmRlZiBfaXRlcl9hbGxfZm9sZGVycyhucyk6CiAgICAiIiJZaWVsZCBldmVyeSBmb2xkZXIgYWNy"
    "b3NzIGV2ZXJ5IG9wZW4gc3RvcmUgKG1haWxib3gpLCBvbmUgYXQgYSB0aW1lCiAgICDigJQgbGF6"
    "aWx5LCB1c2luZyBpbmRleGVkIGFjY2VzcyAoc2VlIF9pdGVyX2ZvbGRlcnMgZG9jc3RyaW5nIOKA"
    "lCBzYW1lCiAgICBOZXh0KCktaGFuZyBjb25jZXJuIGFwcGxpZXMgdG8gbnMuU3RvcmVzLCB3aGlj"
    "aCBpcyBleGFjdGx5IHdoZXJlIHRoZQogICAgcmVwb3J0ZWQgaGFuZyBoYXBwZW5lZCkuIiIiCiAg"
    "ICB0cnk6CiAgICAgICAgc3RvcmVzID0gbnMuU3RvcmVzCiAgICAgICAgc3RvcmVfY291bnQgPSBz"
    "dG9yZXMuQ291bnQKICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgcmV0dXJuCiAgICBmb3Ig"
    "aSBpbiByYW5nZSgxLCBzdG9yZV9jb3VudCArIDEpOgogICAgICAgIHRyeToKICAgICAgICAgICAg"
    "c3RvcmUgPSBzdG9yZXMuSXRlbShpKQogICAgICAgICAgICByb290ID0gc3RvcmUuR2V0Um9vdEZv"
    "bGRlcigpCiAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgY29udGludWUKICAg"
    "ICAgICB5aWVsZCBmcm9tIF9pdGVyX2ZvbGRlcnMocm9vdCkKCgpjbGFzcyBfT3V0bG9va0Rpc2Nv"
    "bm5lY3RlZChFeGNlcHRpb24pOgogICAgIiIiUmFpc2VkIChpbnRlcm5hbGx5LCB3aXRoaW4gdGhp"
    "cyBtb2R1bGUgb25seSkgd2hlbiBhIENPTSBjYWxsIGZhaWxzCiAgICBpbiBhIHdheSB0aGF0IG1l"
    "YW5zIE91dGxvb2sncyBvd24gcHJvY2VzcyBpcyBnb25lIChjbG9zZWQvY3Jhc2hlZCkgLS0KICAg"
    "IGFzIG9wcG9zZWQgdG8gYSBvbmUtb2ZmIHF1aXJrIHdpdGggYSBzaW5nbGUgaXRlbS9mb2xkZXIu"
    "IERpc3Rpbmd1aXNoaW5nCiAgICB0aGUgdHdvIG1hdHRlcnM6IGEgc2luZ2xlIGJhZCBpdGVtIHNo"
    "b3VsZCBqdXN0IGJlIHNraXBwZWQsIGJ1dCBpZgogICAgT3V0bG9vayBpdHNlbGYgaXMgZ29uZSwg"
    "RVZFUlkgcmVtYWluaW5nIGNhbGwgaW4gdGhlIGN1cnJlbnQgc2NhbiBpcwogICAgYWJvdXQgdG8g"
    "ZmFpbCB0aGUgc2FtZSB3YXkgLS0gYmV0dGVyIHRvIHBhdXNlIG9uY2UgYW5kIHdhaXQgZm9yCiAg"
    "ICBPdXRsb29rIHRvIGNvbWUgYmFjayB0aGFuIHRvIGJ1cm4gdGhyb3VnaCB0aG91c2FuZHMgb2Yg"
    "aXRlbXMgZWFjaAogICAgaW5kaXZpZHVhbGx5IGZhaWxpbmcgKGFuZCwgd29yc2UsIGhhdmluZyBl"
    "YWNoIG9uZSBjb3VudGVkIGFzIGEgbm9ybWFsCiAgICAic2tpcHBlZC91bmNoYW5nZWQiIGl0ZW0s"
    "IHdoaWNoIHVzZWQgdG8gbWFrZSBhbiBpbnRlcnJ1cHRlZCBydW4gbG9vawogICAgZmFsc2VseSBj"
    "b21wbGV0ZSkuIiIiCiAgICBwYXNzCgoKIyBLbm93biBIUkVTVUxUcyBmb3IgInRoZSBvdGhlciBw"
    "cm9jZXNzL3NlcnZlciBpcyBnb25lIiAtLSB0aGVzZSBhcmUgd2hhdAojIGFjdHVhbGx5IGdldCBy"
    "YWlzZWQgd2hlbiBPdXRsb29rLmV4ZSBpcyBjbG9zZWQgd2hpbGUgYSBDT00gY2FsbCB0byBpdCBp"
    "cwojIGluIGZsaWdodCBvciBhYm91dCB0byBiZSBtYWRlLiAoMHg4MDA3MDZCQSBSUENfU19TRVJW"
    "RVJfVU5BVkFJTEFCTEUsCiMgMHg4MDA3MDZCRSBSUENfU19DQUxMX0ZBSUxFRCwgMHg4MDAxMDEw"
    "OCBSUENfRV9ESVNDT05ORUNURUQsCiMgMHg4MDA3MDZCRiBSUENfU19DQUxMX0ZBSUxFRF9ETkUp"
    "IOKAlCBzaWduZWQgMzItYml0IGZvcm1zLCBhcyBweXdpbnR5cGVzCiMgcmVwb3J0cyB0aGVtLgpf"
    "UlBDX0RJU0NPTk5FQ1RfSFJFU1VMVFMgPSB7LTIxNDcwMjMxNzAsIC0yMTQ3MDIzMTY5LCAtMjE0"
    "NzAyMzE2NiwgLTIxNDc0MTc4NDh9CgoKZGVmIF9jaGVja19kaXNjb25uZWN0KGUpOgogICAgIiIi"
    "Q2FsbCBmcm9tIGluc2lkZSBhbiBgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOmAgYmxvY2sgdGhhdCB3"
    "b3VsZAogICAgb3RoZXJ3aXNlIGp1c3Qgc2tpcC1hbmQtY29udGludWUuIFJhaXNlcyBfT3V0bG9v"
    "a0Rpc2Nvbm5lY3RlZCAobGV0dGluZwogICAgaXQgcHJvcGFnYXRlIHBhc3QgdGhlIG5vcm1hbCBz"
    "a2lwIGhhbmRsaW5nKSBpZiBgZWAgbG9va3MgbGlrZSBPdXRsb29rCiAgICBpdHNlbGYgd2VudCBh"
    "d2F5OyBkb2VzIG5vdGhpbmcgKGZhbGxzIHRocm91Z2ggdG8gdGhlIG5vcm1hbAogICAgc2tpcC10"
    "aGlzLW9uZS1pdGVtIGJlaGF2aW9yKSBmb3IgYW55IG90aGVyIGtpbmQgb2YgZXJyb3IuIiIiCiAg"
    "ICBocmVzdWx0ID0gZ2V0YXR0cihlLCAnaHJlc3VsdCcsIE5vbmUpCiAgICBpZiBocmVzdWx0IGlz"
    "IE5vbmUgYW5kIGdldGF0dHIoZSwgJ2FyZ3MnLCBOb25lKToKICAgICAgICBocmVzdWx0ID0gZS5h"
    "cmdzWzBdIGlmIGlzaW5zdGFuY2UoZS5hcmdzWzBdLCBpbnQpIGVsc2UgTm9uZQogICAgaWYgaHJl"
    "c3VsdCBpbiBfUlBDX0RJU0NPTk5FQ1RfSFJFU1VMVFM6CiAgICAgICAgcmFpc2UgX091dGxvb2tE"
    "aXNjb25uZWN0ZWQoc3RyKGUpKQoKCmRlZiBfd2FpdF9mb3Jfb3V0bG9va19yZXN0YXJ0KHN0b3Bf"
    "ZmxhZz1Ob25lLCBwcm9ncmVzc19jYj1Ob25lLCBwb2xsX3NlYz0xNSk6CiAgICAiIiJQb2xsIChl"
    "dmVyeSBwb2xsX3NlYykgdW50aWwgT3V0bG9vayByZXNwb25kcyB0byBhIGZyZXNoLCBjaGVhcCBD"
    "T00KICAgIGNhbGwgYWdhaW4uIFJldHVybnMgVHJ1ZSBvbmNlIGl0J3MgYmFjaywgRmFsc2UgaWYg"
    "c3RvcF9mbGFnIHJlcXVlc3RlZAogICAgYW4gYWJvcnQgd2hpbGUgd2FpdGluZy4gRGVsaWJlcmF0"
    "ZWx5IGhhcyBOTyBvdmVyYWxsIHRpbWVvdXQgLS0gdGhlCiAgICB3aG9sZSBwb2ludCBpcyB0byBz"
    "dXBwb3J0ICJjbG9zZSBPdXRsb29rLCBjb21lIGJhY2sgdG8gaXQgaG91cnMgbGF0ZXIsCiAgICBy"
    "ZW9wZW4gaXQiIGFuZCBoYXZlIGluZGV4aW5nIHBpY2sgYmFjayB1cCBvbiBpdHMgb3duLiIiIgog"
    "ICAgd2FpdGVkID0gMAogICAgd2hpbGUgVHJ1ZToKICAgICAgICBpZiBzdG9wX2ZsYWcgYW5kIHN0"
    "b3BfZmxhZygpOgogICAgICAgICAgICByZXR1cm4gRmFsc2UKICAgICAgICB0cnk6CiAgICAgICAg"
    "ICAgIHRlc3RfYXBwID0gd2luMzJjb20uY2xpZW50LkRpc3BhdGNoKCJPdXRsb29rLkFwcGxpY2F0"
    "aW9uIikKICAgICAgICAgICAgdGVzdF9hcHAuR2V0TmFtZXNwYWNlKCJNQVBJIikKICAgICAgICAg"
    "ICAgcmV0dXJuIFRydWUKICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICBpZiBw"
    "cm9ncmVzc19jYjoKICAgICAgICAgICAgICAgIHByb2dyZXNzX2NiKDAsIDEsIGYiT3V0bG9vayDE"
    "kcOjIMSRw7NuZyDigJQgxJFhbmcgY2jhu50gbeG7nyBs4bqhaS4uLiAoxJHDoyBjaOG7nSB7d2Fp"
    "dGVkIC8vIDYwfXB7d2FpdGVkICUgNjA6MDJkfXMpIikKICAgICAgICAgICAgdGltZS5zbGVlcChw"
    "b2xsX3NlYykKICAgICAgICAgICAgd2FpdGVkICs9IHBvbGxfc2VjCgoKZGVmIGluZGV4X291dGxv"
    "b2tfbWFpbChwcm9ncmVzc19jYj1Ob25lLCBzdG9wX2ZsYWc9Tm9uZSk6CiAgICAiIiJJbmNyZW1l"
    "bnRhbGx5IGluZGV4IGFsbCBPdXRsb29rIG1haWwgKGV2ZXJ5IHN0b3JlL21haWxib3ggY3VycmVu"
    "dGx5CiAgICBvcGVuIGluIE91dGxvb2sg4oCUIHNoYXJlZCBtYWlsYm94ZXMgaW5jbHVkZWQpIGlu"
    "dG8gT1VUTE9PS19EQl9GSUxFLgoKICAgIHByb2dyZXNzX2NiKGRvbmUsIHRvdGFsLCBjdXJyZW50"
    "X3N1YmplY3QpOiBvcHRpb25hbCwgY2FsbGVkIHBlcmlvZGljYWxseQogICAgc28gdGhlIFVJIGNh"
    "biBkcml2ZSBhIHByb2dyZXNzIGJhciB0aGUgc2FtZSB3YXkgaW5kZXhpbmdfd29ya2VyKCkgZG9l"
    "cwogICAgZm9yIHRoZSBtYWluIGZpbGUgREIuCgogICAgc3RvcF9mbGFnOiBvcHRpb25hbCB6ZXJv"
    "LWFyZyBjYWxsYWJsZTsgaWYgaXQgcmV0dXJucyBUcnVlLCBpbmRleGluZwogICAgc3RvcHMgYXQg"
    "dGhlIG5leHQgc2FmZSBwb2ludCAobWlycm9ycyB0aGUgY2FuY2VsbGFibGUgVXBkYXRlIERCIGZs"
    "b3cpLgoKICAgIFJldHVybnMgKGluZGV4ZWRfY291bnQsIHNraXBwZWRfY291bnQsIGVycm9yKSDi"
    "gJQgZXJyb3IgaXMgTm9uZSBvbiBzdWNjZXNzLAogICAgb3IgYSBzaG9ydCBodW1hbi1yZWFkYWJs"
    "ZSBzdHJpbmcgb24gZmFpbHVyZSAoZS5nLiBPdXRsb29rIG5vdCBydW5uaW5nLwogICAgaW5zdGFs"
    "bGVkKSBzbyB0aGUgVUkgY2FuIHNob3cgaXQgZGlyZWN0bHkuCiAgICAiIiIKICAgIGlmIG5vdCBP"
    "VVRMT09LX0FWQUlMQUJMRToKICAgICAgICByZXR1cm4gMCwgMCwgInB5d2luMzIgY2jGsGEgxJHG"
    "sOG7o2MgY8OgaSAocGlwIGluc3RhbGwgcHl3aW4zMikiCgogICAgcHl0aG9uY29tLkNvSW5pdGlh"
    "bGl6ZSgpCiAgICB0cnk6CiAgICAgICAgX3JlZ2lzdGVyX21lc3NhZ2VfZmlsdGVyKCkKICAgICAg"
    "ICB0cnk6CiAgICAgICAgICAgIG91dGxvb2sgPSB3aW4zMmNvbS5jbGllbnQuRGlzcGF0Y2goIk91"
    "dGxvb2suQXBwbGljYXRpb24iKQogICAgICAgICAgICBucyA9IG91dGxvb2suR2V0TmFtZXNwYWNl"
    "KCJNQVBJIikKICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgICAgIHJldHVy"
    "biAwLCAwLCBmIktow7RuZyBr4bq/dCBu4buRaSDEkcaw4bujYyBPdXRsb29rICjEkcOjIGPDoGkv"
    "xJHEg25nIG5o4bqtcCBjaMawYT8pOiB7ZX0iCgogICAgICAgIGNvbm4sIGMgPSBfY29ubmVjdCgp"
    "CgogICAgICAgICMgZW50cnlfaWQgLT4gbGFzdF9tb2RpZmllZCBhbHJlYWR5IGluZGV4ZWQsIHNv"
    "IHVuY2hhbmdlZCBtYWlsIGNhbgogICAgICAgICMgYmUgc2tpcHBlZCBlbnRpcmVseSAobm8gQm9k"
    "eSByZWFkID0gbm8gc2xvdyBDT00gcm91bmQtdHJpcCkuIFRoaXMKICAgICAgICAjIGRpY3QgaXMg"
    "YWxzbyB1cGRhdGVkIGxpdmUgYXMgbmV3IGl0ZW1zIGFyZSBpbmRleGVkIGJlbG93IChub3QganVz"
    "dAogICAgICAgICMgcmVsb2FkZWQgZnJvbSB0aGUgREIpLCBzbyB0aGF0IGlmIE91dGxvb2sgZGlz"
    "Y29ubmVjdHMgYW5kIHRoZSBzY2FuCiAgICAgICAgIyByZXN0YXJ0cyAoc2VlIHRoZSByZXN1bWUg"
    "bG9vcCBiZWxvdyksIGV2ZXJ5dGhpbmcgY29tbWl0dGVkIHNvIGZhcgogICAgICAgICMgaW4gVEhJ"
    "UyBydW4gaXMgc2tpcHBlZCBqdXN0IGFzIGZhc3QgYXMgb2xkZXIsIGFscmVhZHktY29tbWl0dGVk"
    "IG1haWwuCiAgICAgICAgYy5leGVjdXRlKCJTRUxFQ1QgZW50cnlfaWQsIGxhc3RfbW9kaWZpZWQg"
    "RlJPTSBvdXRsb29rX3N0b3JlIikKICAgICAgICBrbm93biA9IGRpY3QoYy5mZXRjaGFsbCgpKQoK"
    "ICAgICAgICAjIGluZGV4ZWQvc2tpcHBlZCBBUkUgY3VtdWxhdGl2ZSBhY3Jvc3MgdGhlIHdob2xl"
    "IHJ1biAoaW5jbHVkaW5nIGFueQogICAgICAgICMgcmVzdGFydHMgYmVsb3cpIC0tIHRoYXQncyB0"
    "aGUgcmVhbCwgY29ycmVjdCB0b3RhbCBmb3IgdGhlIGZpbmFsCiAgICAgICAgIyBzdW1tYXJ5LiBw"
    "cm9jZXNzZWQvdG90YWwgYXJlIE5PVDogdGhleSdyZSByZXNldCBhdCB0aGUgdG9wIG9mCiAgICAg"
    "ICAgIyBlYWNoIHJlc3RhcnQgYXR0ZW1wdCAoc2VlIHYtZml4IGJlbG93KSBzaW5jZSB0aGV5IHJl"
    "cHJlc2VudCAiaG93CiAgICAgICAgIyBmYXIgYWxvbmcgVEhJUyB3YWxrIG9mIHRoZSBmb2xkZXIg"
    "dHJlZSBpcyIsIGFuZCBhIHJlc3RhcnQgaXMgYQogICAgICAgICMgYnJhbmQgbmV3IHdhbGsgZnJv"
    "bSB0aGUgdG9wLgogICAgICAgIGluZGV4ZWQgPSAwCiAgICAgICAgc2tpcHBlZCA9IDAKICAgICAg"
    "ICBkaXNjb25uZWN0X2NvdW50ID0gMAogICAgICAgICMgR2VuZXJvdXMgY2FwIHNvIGEgZ2VudWlu"
    "ZWx5IHBlcnNpc3RlbnQsIHVucmVsYXRlZCBwcm9ibGVtIGNhbid0CiAgICAgICAgIyBzcGluIHRo"
    "aXMgZm9yZXZlciDigJQgYnV0IGhpZ2ggZW5vdWdoIHRoYXQgYSB1c2VyIGNsb3NpbmcgT3V0bG9v"
    "awogICAgICAgICMgc2V2ZXJhbCB0aW1lcyBvdmVyIGEgbG9uZyB1bmF0dGVuZGVkIHJ1biBpcyBu"
    "ZXZlciBhbiBpc3N1ZS4KICAgICAgICBNQVhfT1VUTE9PS19ESVNDT05ORUNUUyA9IDUwMAoKICAg"
    "ICAgICAjIHYtcmVzdW1lOiBvdXRlciByZXRyeSBsb29wLiBJZiBPdXRsb29rJ3MgcHJvY2VzcyBp"
    "dHNlbGYgZ29lcyBhd2F5CiAgICAgICAgIyBtaWQtc2NhbiAoY2xvc2VkIGJ5IHRoZSB1c2VyLCBj"
    "cmFzaGVkLCBldGMuKSwgRVZFUlkgQ09NIG9iamVjdAogICAgICAgICMgaGVsZCBmcm9tIGJlZm9y"
    "ZSB0aGF0IHBvaW50IChmb2xkZXIsIGZvbGRlcl9pdGVtcywgaXRlbSwgZXZlbiB0aGUKICAgICAg"
    "ICAjIGZvbGRlcl9nZW4gZ2VuZXJhdG9yJ3MgaW50ZXJuYWwgc3RhdGUpIGlzIHBlcm1hbmVudGx5"
    "IGludmFsaWQg4oCUCiAgICAgICAgIyBDT00vUlBDIGhhbmRsZXMgZG9uJ3Qgc3Vydml2ZSB0aGUg"
    "c2VydmVyIHByb2Nlc3MgcmVzdGFydGluZy4KICAgICAgICAjIFRoZXJlJ3Mgbm8gd2F5IHRvIHJl"
    "c3VtZSB0aGUgT0xEIHNjYW4gYXQgdGhlIGV4YWN0IGl0ZW0gaXQgd2FzIG9uOwogICAgICAgICMg"
    "dGhlIG9ubHkgb3B0aW9uIGlzIHRvIHdhaXQgZm9yIE91dGxvb2sgdG8gY29tZSBiYWNrLCBnZXQg"
    "YSBGUkVTSAogICAgICAgICMgQXBwbGljYXRpb24vTmFtZXNwYWNlLCBhbmQgcmUtd2FsayBmb2xk"
    "ZXJzIGZyb20gdGhlIHRvcC4gVGhhbmtzIHRvCiAgICAgICAgIyB0aGUgYGtub3duYCBkaWN0IGFi"
    "b3ZlLCB0aGF0IHJlLXdhbGsgaXMgZmFzdDogZXZlcnkgYWxyZWFkeS0KICAgICAgICAjIGluZGV4"
    "ZWQgaXRlbSAoZnJvbSB0aGlzIHJ1biBvciBhbiBlYXJsaWVyIG9uZSkgaXMgc2tpcHBlZCB3aXRo"
    "IGEKICAgICAgICAjIHNpbmdsZSBkaWN0IGxvb2t1cCwgbm8gQm9keSByZWFkIOKAlCBzbyBpbiBw"
    "cmFjdGljZSB0aGlzIGJlaGF2ZXMKICAgICAgICAjIGxpa2UgInBhdXNlIGFuZCByZXN1bWUiLCBq"
    "dXN0IGltcGxlbWVudGVkIGFzICJyZXN0YXJ0LCBidXQgdGhlCiAgICAgICAgIyByZXN0YXJ0IGlz"
    "IGNoZWFwIiwgbm90IGEgdHJ1ZSBsb3ctbGV2ZWwgcmVzdW1lLgogICAgICAgIHdoaWxlIFRydWU6"
    "CiAgICAgICAgICAgICMgdi1maXg6IHByb2Nlc3NlZC90b3RhbCB1c2VkIHRvIGJlIGRlY2xhcmVk"
    "IE9OQ0UsIG91dHNpZGUgdGhpcwogICAgICAgICAgICAjIGxvb3AsIGFuZCBrZXB0IGNsaW1iaW5n"
    "IGFjcm9zcyBldmVyeSByZXN0YXJ0IC0tIHNvIDItMwogICAgICAgICAgICAjIGRpc2Nvbm5lY3Rz"
    "IG9uIGV2ZW4gYSBtb2Rlc3QgbWFpbGJveCBjb3VsZCBzaG93ICJwcm9jZXNzZWQiCiAgICAgICAg"
    "ICAgICMgcmVhY2hpbmcgM3ggdGhlIHJlYWwgZm9sZGVyLXRyZWUgc2l6ZSAoZWFjaCByZXN0YXJ0"
    "IHJlLXdhbGtzCiAgICAgICAgICAgICMgdGhlIHNhbWUgZm9sZGVycyBhbmQgcmUtY291bnRzIHRo"
    "ZW0sIGV2ZW4gdGhvdWdoIHRoZSBhY3R1YWwKICAgICAgICAgICAgIyBwZXItaXRlbSB3b3JrIGlz"
    "IHNraXBwZWQgZmFzdCB2aWEgYGtub3duYCkuIFJlc2V0dGluZyB0aGVtCiAgICAgICAgICAgICMg"
    "aGVyZSBtZWFucyB0aGUgZGlzcGxheWVkIGNvdW50IGFsd2F5cyByZWZsZWN0cyBUSElTIHdhbGsg"
    "b25seS4KICAgICAgICAgICAgcHJvY2Vzc2VkID0gMAogICAgICAgICAgICB0b3RhbCA9IDAgICAg"
    "ICAgICAgIyBncm93cyBhcyBmb2xkZXJzIGFyZSBkaXNjb3ZlcmVkIOKAlCBzZWUgbm90ZSBhYm92"
    "ZQogICAgICAgICAgICBmb2xkZXJfbnVtID0gMAogICAgICAgICAgICBmb2xkZXJfZ2VuID0gX2l0"
    "ZXJfYWxsX2ZvbGRlcnMobnMpCiAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgIHdoaWxl"
    "IFRydWU6CiAgICAgICAgICAgICAgICAgICAgaWYgc3RvcF9mbGFnIGFuZCBzdG9wX2ZsYWcoKToK"
    "ICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgICAgICB0cnk6CiAg"
    "ICAgICAgICAgICAgICAgICAgICAgIGZvbGRlciA9IG5leHQoZm9sZGVyX2dlbikKICAgICAgICAg"
    "ICAgICAgICAgICBleGNlcHQgU3RvcEl0ZXJhdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAg"
    "YnJlYWsKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAg"
    "ICAgICAgICAgICAgICAgIF9jaGVja19kaXNjb25uZWN0KGUpICAjIHJlLXJhaXNlcyBfT3V0bG9v"
    "a0Rpc2Nvbm5lY3RlZCBpZiBhcHBsaWNhYmxlCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIHBy"
    "b2dyZXNzX2NiOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcHJvZ3Jlc3NfY2IocHJvY2Vz"
    "c2VkLCB0b3RhbCBvciAxLCBmIihPdXRsb29rIENPTSBs4buXaSwgZOG7q25nIGluZGV4OiB7ZX0p"
    "IikKICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKCiAgICAgICAgICAgICAgICAgICAgZm9s"
    "ZGVyX251bSArPSAxCiAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAg"
    "ICAgICBmb2xkZXJfcGF0aCA9IGZvbGRlci5Gb2xkZXJQYXRoCiAgICAgICAgICAgICAgICAgICAg"
    "ZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgZm9sZGVyX3BhdGggPSBm"
    "Iihmb2xkZXIgI3tmb2xkZXJfbnVtfSkiCgogICAgICAgICAgICAgICAgICAgICMgdi1maXg6IGl0"
    "ZW0uU3RvcmVJRCBkb2Vzbid0IGV4aXN0IG9uIHRoaXMgT3V0bG9vay93aW4zMmNvbQogICAgICAg"
    "ICAgICAgICAgICAgICMgc2V0dXAgKHNlZSBpbmRleF9vdXRsb29rX21haWwncyBpbnNlcnQtYmxv"
    "Y2sgY29tbWVudCkg4oCUIHVzZQogICAgICAgICAgICAgICAgICAgICMgdGhlIGZvbGRlcidzIFN0"
    "b3JlSUQgaW5zdGVhZCwgZmV0Y2hlZCBvbmNlIHBlciBmb2xkZXIuCiAgICAgICAgICAgICAgICAg"
    "ICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICBmb2xkZXJfc3RvcmVfaWQgPSBmb2xkZXIu"
    "U3RvcmVJRAogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAg"
    "ICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGZvbGRlcl9zdG9y"
    "ZV9pZCA9IGZvbGRlci5TdG9yZS5TdG9yZUlECiAgICAgICAgICAgICAgICAgICAgICAgIGV4Y2Vw"
    "dCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBmb2xkZXJfc3RvcmVfaWQg"
    "PSAiIgoKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIGZv"
    "bGRlcl9pdGVtcyA9IGZvbGRlci5JdGVtcwogICAgICAgICAgICAgICAgICAgICAgICBmb2xkZXJf"
    "Y291bnQgPSBmb2xkZXJfaXRlbXMuQ291bnQKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhj"
    "ZXB0aW9uIGFzIGU6CiAgICAgICAgICAgICAgICAgICAgICAgIF9jaGVja19kaXNjb25uZWN0KGUp"
    "CiAgICAgICAgICAgICAgICAgICAgICAgICMgdi1maXg6IHRoaXMgaXMgZXhhY3RseSB0aGUgY2Fs"
    "bCB0aGF0IGNhbiBoYW5nIGZvciBhIGxvbmcKICAgICAgICAgICAgICAgICAgICAgICAgIyB0aW1l"
    "IG9uIHNvbWUgZm9sZGVycyAobGFyZ2Ugc2hhcmVkIG1haWxib3gsIElNQVAgZm9sZGVyCiAgICAg"
    "ICAgICAgICAgICAgICAgICAgICMgbm90IHlldCBzeW5jZWQsIHB1YmxpYyBmb2xkZXIsIGV0Yyku"
    "IFJlcG9ydCBCRUZPUkUgdGhlCiAgICAgICAgICAgICAgICAgICAgICAgICMgY2FsbCB0b28sIHNv"
    "IGlmIGl0IGRvZXMgaGFuZywgdGhlIGxhc3QgdGV4dCBvbiBzY3JlZW4KICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIyBuYW1lcyB0aGUgZm9sZGVyIHRoYXQncyBzdHVjayBpbnN0ZWFkIG9mIGdvaW5n"
    "IHNpbGVudC4KICAgICAgICAgICAgICAgICAgICAgICAgaWYgcHJvZ3Jlc3NfY2I6CiAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICBwcm9ncmVzc19jYihwcm9jZXNzZWQsIHRvdGFsIG9yIDEsIGYi"
    "KGLhu48gcXVhIOKAlCBs4buXaSBt4bufKSB7Zm9sZGVyX3BhdGh9IikKICAgICAgICAgICAgICAg"
    "ICAgICAgICAgY29udGludWUKCiAgICAgICAgICAgICAgICAgICAgdG90YWwgKz0gZm9sZGVyX2Nv"
    "dW50CiAgICAgICAgICAgICAgICAgICAgaWYgcHJvZ3Jlc3NfY2I6CiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICMgRmlyZXMgaW1tZWRpYXRlbHkgb24gZW50ZXJpbmcgZXZlcnkgZm9sZGVyLCBldmVu"
    "IG9uZXMKICAgICAgICAgICAgICAgICAgICAgICAgIyB3aXRoIDAgbWFpbCDigJQgdGhpcyBpcyB3"
    "aGF0IG1ha2VzIHRoZSBjb3VudGVyIHZpc2libHkKICAgICAgICAgICAgICAgICAgICAgICAgIyBt"
    "b3ZlIHdpdGhpbiB0aGUgZmlyc3Qgc2Vjb25kLCBsb25nIGJlZm9yZSBhbnkgbWFpbCBoYXMKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgIyBhY3R1YWxseSBiZWVuIGluZGV4ZWQuCiAgICAgICAgICAg"
    "ICAgICAgICAgICAgIHByb2dyZXNzX2NiKHByb2Nlc3NlZCwgdG90YWwsIGYixJBhbmcgcXXDqXQ6"
    "IHtmb2xkZXJfcGF0aH0gKHtmb2xkZXJfY291bnR9IG3hu6VjKSIpCgogICAgICAgICAgICAgICAg"
    "ICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgZm9yIGlkeCBpbiByYW5nZSgxLCBmb2xk"
    "ZXJfY291bnQgKyAxKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHByb2Nlc3NlZCArPSAx"
    "CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBzdG9wX2ZsYWcgYW5kIHN0b3BfZmxhZygp"
    "OgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICBpZiBwcm9ncmVzc19jYiBhbmQgaWR4ICUgMyA9PSAwOgogICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICMgdi1maXg6IGZpcmVzIEJFRk9SRSB0b3VjaGluZyB0aGlzIGl0"
    "ZW0ncwogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgcHJvcGVydGllcyAoQ2xhc3Mv"
    "Qm9keS9ldGMg4oCUIGFueSBvZiB3aGljaCBjYW4gYmUKICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAjIHRoZSBjYWxsIHRoYXQgaGFuZ3MsIGUuZy4gYSBuZXR3b3JrIGZldGNoIGluCiAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyBPdXRsb29rIE9ubGluZSBtb2RlKS4gSWYg"
    "aW5kZXhpbmcgZnJlZXplcywgdGhlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyBs"
    "YXN0IHRleHQgb24gc2NyZWVuIG5vdyBuYW1lcyB0aGUgZXhhY3QgZm9sZGVyICsKICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAjIGl0ZW0gcG9zaXRpb24gaXQgZnJvemUgb24sIGluc3Rl"
    "YWQgb2YgYSBzdGFsZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgbWVzc2FnZSBm"
    "cm9tIGh1bmRyZWRzIG9mIGl0ZW1zIGVhcmxpZXIuCiAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgcHJvZ3Jlc3NfY2IocHJvY2Vzc2VkLCB0b3RhbCwgZiLEkGFuZyBt4bufOiB7Zm9sZGVy"
    "X3BhdGh9IFt7aWR4fS97Zm9sZGVyX2NvdW50fV0iKQogICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGl0ZW0gPSBmb2xkZXJfaXRl"
    "bXMuSXRlbShpZHgpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9u"
    "IGFzIGU6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgX2NoZWNrX2Rpc2Nvbm5lY3Qo"
    "ZSkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBza2lwcGVkICs9IDEKICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICBpZiBwcm9jZXNzZWQgJSA1MDAgPT0gMDoKICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgY29ubi5jb21taXQoKQogICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICB0cnk6"
    "CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgaXRlbS5DbGFzcyAhPSBPTF9NQUlM"
    "X0NMQVNTOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBwcm9ncmVzc19j"
    "YiBhbmQgcHJvY2Vzc2VkICUgMjUgPT0gMDoKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIHByb2dyZXNzX2NiKHByb2Nlc3NlZCwgdG90YWwsIGZvbGRlcl9wYXRoKQogICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIHYtZml4OiBjb21taXQgdHJpZ2dlciBi"
    "ZWxvdyAoYWZ0ZXIgdGhpcyBpZi9lbGlmCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICMgY2hhaW4pIHVzZWQgdG8gb25seSBmaXJlIG9uICpuZXdseSBpbmRleGVkKgogICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIG1haWwuIEEgZm9sZGVyIGZ1bGwgb2Ygbm9u"
    "LW1haWwgaXRlbXMgKG9yIG1haWwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "IyBhbHJlYWR5IGluZGV4ZWQvdW5jaGFuZ2VkKSBjb3VsZCBydW4gZm9yIGhvdXJzCiAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgd2l0aG91dCBhIHNpbmdsZSBjb21taXQg4oCU"
    "IERCIGxvb2tlZCBmcm96ZW4gZXZlbgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAjIHRob3VnaCAicHJvY2Vzc2VkIiBrZXB0IGNsaW1iaW5nLiBGYWxsIHRocm91Z2gKICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyB0byB0aGUgc2hhcmVkIHBlcmlvZGljLWNv"
    "bW1pdCBjaGVjayBpbnN0ZWFkIG9mCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICMgY29udGludWluZyBlYXJseS4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "aWYgcHJvY2Vzc2VkICUgNTAwID09IDA6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICBjb25uLmNvbW1pdCgpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "IGNvbnRpbnVlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFz"
    "IGU6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgX2NoZWNrX2Rpc2Nvbm5lY3QoZSkK"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBjb250aW51ZQoKICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBlbnRyeV9p"
    "ZCA9IGl0ZW0uRW50cnlJRAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGxtX3N0ciA9"
    "IHN0cihpdGVtLkxhc3RNb2RpZmljYXRpb25UaW1lKQoKICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICBpZiBrbm93bi5nZXQoZW50cnlfaWQpID09IGxtX3N0cjoKICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgc2tpcHBlZCArPSAxCiAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgIGlmIHByb2dyZXNzX2NiIGFuZCBwcm9jZXNzZWQgJSAyNSA9PSAwOgogICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcHJvZ3Jlc3NfY2IocHJvY2Vzc2Vk"
    "LCB0b3RhbCwgIihi4buPIHF1YSDigJQga2jDtG5nIMSR4buVaSkiKQogICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAjIFNhbWUgdi1maXggYXMgYWJvdmU6IGEgcnVuIG9mIGFscmVh"
    "ZHktaW5kZXhlZAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjICJ1bmNoYW5n"
    "ZWQiIG1haWwgbXVzdCBub3Qgc3RhcnZlIG91dCBjb21taXRzLgogICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICBpZiBwcm9jZXNzZWQgJSA1MDAgPT0gMDoKICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgIGNvbm4uY29tbWl0KCkKICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgY29udGludWUKCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgc3ViamVjdCA9IGl0ZW0uU3ViamVjdCBvciAiIgogICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgIGJvZHkgPSBpdGVtLkJvZHkgb3IgIiIKICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICBzZW5kZXIgPSBpdGVtLlNlbmRlck5hbWUgb3IgIiIKICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHJlY2Vp"
    "dmVkID0gaXRlbS5SZWNlaXZlZFRpbWUuc3RyZnRpbWUoIiVZLSVtLSVkICVIOiVNOiVTIikKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICByZWNlaXZlZCA9ICIiCiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgIyB2LWZpeDogaXRlbS5TdG9yZUlEIHJhaXNlZCBBdHRyaWJ1dGVFcnJv"
    "ciBvbgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgRVZFUlkgbWFpbCBpdGVtIGlu"
    "IHRoaXMgZW52aXJvbm1lbnQgKGNvbmZpcm1lZAogICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICMgdmlhIHRoZSBkaWFnbm9zdGljIGJ1aWxkJ3MgZXJyb3IgbG9nKSDigJQgdGhhdCBvbmUK"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIGxpbmUgd2FzIHNpbGVudGx5IGtpbGxp"
    "bmcgMTAwJSBvZiBpbnNlcnRzIHRoaXMKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAj"
    "IHdob2xlIHRpbWUsIHdoaWNoIGlzIHdoeSB0aGUgREIgc3RheWVkIGF0IDAgcm93cwogICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICMgbm8gbWF0dGVyIGhvdyBtYW55IGl0ZW1zIHdlcmUg"
    "cHJvY2Vzc2VkLiBGb2xkZXJzCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyBleHBv"
    "c2UgU3RvcmVJRCByZWxpYWJseSAoaXQncyBhIGNvcmUgTUFQSUZvbGRlcgogICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICMgcHJvcGVydHkpLCBzbyBnZXQgaXQgZnJvbSB0aGUgZm9sZGVy"
    "IGluc3RlYWQg4oCUCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyBzYW1lIHZhbHVl"
    "LCBqdXN0IGEgbW9yZSByZWxpYWJsZSBzb3VyY2UuIENhY2hlIGl0CiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgIyBwZXIgZm9sZGVyIHNvIHRoaXMgaXNuJ3QgYSByZXBlYXRlZCBDT00g"
    "Y2FsbCBmb3IKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIGV2ZXJ5IHNpbmdsZSBt"
    "YWlsIGluIGEgbGFyZ2UgZm9sZGVyLgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHN0"
    "b3JlX2lkID0gZm9sZGVyX3N0b3JlX2lkCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "dHJ5OgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzaXplID0gaW50KGl0ZW0u"
    "U2l6ZSkgICMgYnl0ZXMg4oCUIHNhbWUgdW5pdCBhcyByZWFsIGZpbGUgc2l6ZXMgZWxzZXdoZXJl"
    "IGluIHRoZSBhcHAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0"
    "aW9uOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzaXplID0gMAoKICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAjIEZUUzUgaGFzIG5vIFVQU0VSVCDigJQgZHJvcCB0"
    "aGUgb2xkIGNvbnRlbnQgcm93IChpZgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMg"
    "YW55KSBiZWZvcmUgaW5zZXJ0aW5nIHRoZSBmcmVzaCBvbmUgb24gcmUtaW5kZXguCiAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgYy5leGVjdXRlKCJTRUxFQ1Qgcm93aWRfZnRzIEZST00g"
    "b3V0bG9va19zdG9yZSBXSEVSRSBlbnRyeV9pZD0/IiwgKGVudHJ5X2lkLCkpCiAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgcm93ID0gYy5mZXRjaG9uZSgpCiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgaWYgcm93IGFuZCByb3dbMF0gaXMgbm90IE5vbmU6CiAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgIGMuZXhlY3V0ZSgiREVMRVRFIEZST00gb3V0bG9va19j"
    "b250ZW50IFdIRVJFIHJvd2lkPT8iLCAocm93WzBdLCkpCgogICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIGMuZXhlY3V0ZSgiSU5TRVJUIElOVE8gb3V0bG9va19jb250ZW50IChzdWJqZWN0"
    "LCBib2R5KSBWQUxVRVMgKD8sPykiLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAoc3ViamVjdCwgYm9keSkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ZnRzX3Jvd2lkID0gYy5sYXN0cm93aWQKCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "Yy5leGVjdXRlKCIiIklOU0VSVCBJTlRPIG91dGxvb2tfc3RvcmUKICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgKGVudHJ5X2lkLCBzdG9yZV9pZCwgc3ViamVjdCwg"
    "c2VuZGVyLCByZWNlaXZlZCwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgIGZvbGRlcl9wYXRoLCBzaXplLCBsYXN0X21vZGlmaWVkLCByb3dpZF9mdHMpCiAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIFZBTFVFUyAoPyw/LD8sPyw/"
    "LD8sPyw/LD8pCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIE9O"
    "IENPTkZMSUNUKGVudHJ5X2lkKSBETyBVUERBVEUgU0VUCiAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgc3RvcmVfaWQ9ZXhjbHVkZWQuc3RvcmVfaWQsIHN1Ympl"
    "Y3Q9ZXhjbHVkZWQuc3ViamVjdCwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICBzZW5kZXI9ZXhjbHVkZWQuc2VuZGVyLCByZWNlaXZlZD1leGNsdWRlZC5yZWNl"
    "aXZlZCwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBmb2xk"
    "ZXJfcGF0aD1leGNsdWRlZC5mb2xkZXJfcGF0aCwgc2l6ZT1leGNsdWRlZC5zaXplLAogICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGxhc3RfbW9kaWZpZWQ9ZXhj"
    "bHVkZWQubGFzdF9tb2RpZmllZCwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICByb3dpZF9mdHM9ZXhjbHVkZWQucm93aWRfZnRzIiIiLAogICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAoZW50cnlfaWQsIHN0b3JlX2lkLCBzdWJqZWN0"
    "LCBzZW5kZXIsIHJlY2VpdmVkLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgZm9sZGVyX3BhdGgsIHNpemUsIGxtX3N0ciwgZnRzX3Jvd2lkKSkKCiAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAga25vd25bZW50cnlfaWRdID0gbG1fc3RyICAjIHYtcmVzdW1l"
    "OiBrZWVwIGluLW1lbW9yeSBza2lwLWxpc3QgY3VycmVudAogICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIGluZGV4ZWQgKz0gMQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlm"
    "IHByb2dyZXNzX2NiIGFuZCBwcm9jZXNzZWQgJSAxMCA9PSAwOgogICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICBwcm9ncmVzc19jYihwcm9jZXNzZWQsIHRvdGFsLCBzdWJqZWN0Wzo2"
    "MF0pCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyBDb21taXQgZXZlcnkgNTAgd3Jp"
    "dGVzIChub3QgMjAwKSBzbyB0aGUgLmRiIGZpbGUKICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAjIHZpc2libHkgZ3Jvd3MgZWFybHkgYW5kIG9mdGVuLCBub3QgaW4gb25lIGxhdGUgbHVt"
    "cC4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBpbmRleGVkICUgNTAgPT0gMDoK"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgY29ubi5jb21taXQoKQogICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgIF9jaGVja19kaXNjb25uZWN0KGUpCiAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgc2tpcHBlZCArPSAxCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgIyB2LWZpeDogdGhpcyBleGNlcHQgYmxvY2sgd2FzIHNpbGVudGx5IHN3YWxsb3dpbmcKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIGV2ZXJ5IGluc2VydC9jb21taXQgZXJyb3Ig"
    "d2l0aCBubyB0cmFjZSBhdCBhbGwg4oCUCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "IyB0aGF0J3MgZXhhY3RseSB3aHkgIjE4Mjk1IHByb2Nlc3NlZCwgMCByb3dzIGluCiAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgIyBEQiIgcHJvZHVjZWQgemVybyBjbHVlcy4gUHJpbnQg"
    "dGhlIHJlYWwgZXJyb3IgZm9yCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyB0aGUg"
    "Zmlyc3Qgc2V2ZXJhbCBvY2N1cnJlbmNlcyBzbyB0aGUgYWN0dWFsIGNhdXNlCiAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgIyAobG9ja2VkIERCLCBkaXNrIGZ1bGwsIHJlYWRvbmx5IGZp"
    "bGUsIGJhZCB2YWx1ZSwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIGV0Yy4pIGJl"
    "Y29tZXMgdmlzaWJsZSBpbnN0ZWFkIG9mIGludmlzaWJsZS4KICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICBnbG9iYWwgX2luc2VydF9lcnJvcl9jb3VudAogICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgIGlmIF9pbnNlcnRfZXJyb3JfY291bnQgPCA4OgogICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICBfaW5zZXJ0X2Vycm9yX2NvdW50ICs9IDEKICAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgcHJpbnQoZiJbT3V0bG9va11bSU5TRVJUIEVSUk9SICN7"
    "X2luc2VydF9lcnJvcl9jb3VudH1dIC0+IHtlIXJ9IikKICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgdHJhY2ViYWNrLnByaW50X2V4YygpCiAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgaWYgcHJvY2Vzc2VkICUgNTAwID09IDA6CiAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgICAgIGNvbm4uY29tbWl0KCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICBjb250aW51ZQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBfT3V0bG9va0Rpc2Nvbm5lY3Rl"
    "ZDoKICAgICAgICAgICAgICAgICAgICAgICAgcmFpc2UKICAgICAgICAgICAgICAgICAgICBleGNl"
    "cHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICBwYXNzCgogICAgICAgICAgICAg"
    "ICAgICAgICMgQ29tbWl0IGF0IHRoZSBlbmQgb2YgZXZlcnkgZm9sZGVyIHRvbywgc28gc3dpdGNo"
    "aW5nIGZvbGRlcnMKICAgICAgICAgICAgICAgICAgICAjIChJbmJveCAtPiBTZW50IEl0ZW1zIC0+"
    "IC4uLikgaXMgYSB2aXNpYmxlIHN0ZXAsIG5vdCBzaWxlbmNlLgogICAgICAgICAgICAgICAgICAg"
    "IGNvbm4uY29tbWl0KCkKCiAgICAgICAgICAgICAgICAjIElubmVyIHNjYW4gZmluaXNoZWQgbmF0"
    "dXJhbGx5IChTdG9wSXRlcmF0aW9uIG9yIHN0b3BfZmxhZykg4oCUCiAgICAgICAgICAgICAgICAj"
    "IG5vIGRpc2Nvbm5lY3QgaGFwcGVuZWQgdGhpcyBhdHRlbXB0LCBzbyB3ZSdyZSBkb25lIGZvciBy"
    "ZWFsLgogICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgZXhjZXB0IF9PdXRsb29rRGlz"
    "Y29ubmVjdGVkIGFzIGU6CiAgICAgICAgICAgICAgICBjb25uLmNvbW1pdCgpICAjIGtlZXAgZXZl"
    "cnl0aGluZyBpbmRleGVkIHNvIGZhciwgbm8gbWF0dGVyIHdoYXQgaGFwcGVucyBuZXh0CiAgICAg"
    "ICAgICAgICAgICBkaXNjb25uZWN0X2NvdW50ICs9IDEKICAgICAgICAgICAgICAgIHByaW50KGYi"
    "W091dGxvb2tdIERpc2Nvbm5lY3RlZCAoI3tkaXNjb25uZWN0X2NvdW50fSkgLS0ge2Uhcn0gLS0g"
    "d2FpdGluZyBmb3IgT3V0bG9vayB0byByZW9wZW4uLi4iKQogICAgICAgICAgICAgICAgaWYgcHJv"
    "Z3Jlc3NfY2I6CiAgICAgICAgICAgICAgICAgICAgcHJvZ3Jlc3NfY2IocHJvY2Vzc2VkLCBwcm9j"
    "ZXNzZWQsCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIk91dGxvb2sgxJHDoyDEkcOz"
    "bmcgZ2nhu69hIGNo4burbmcg4oCUIMSRYW5nIGNo4budIG3hu58gbOG6oWkuLi4iKQogICAgICAg"
    "ICAgICAgICAgaWYgZGlzY29ubmVjdF9jb3VudCA+IE1BWF9PVVRMT09LX0RJU0NPTk5FQ1RTIG9y"
    "IChzdG9wX2ZsYWcgYW5kIHN0b3BfZmxhZygpKToKICAgICAgICAgICAgICAgICAgICBwcmludCgi"
    "W091dGxvb2tdIEdpdmluZyB1cCB3YWl0aW5nIChzdG9wIHJlcXVlc3RlZCBvciB0b28gbWFueSBk"
    "aXNjb25uZWN0cykuIikKICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAg"
    "aWYgbm90IF93YWl0X2Zvcl9vdXRsb29rX3Jlc3RhcnQoc3RvcF9mbGFnLCBwcm9ncmVzc19jYik6"
    "CiAgICAgICAgICAgICAgICAgICAgcHJpbnQoIltPdXRsb29rXSBTdG9wIHJlcXVlc3RlZCB3aGls"
    "ZSB3YWl0aW5nIGZvciBPdXRsb29rLiIpCiAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAg"
    "ICAgICAgICAgIHByaW50KCJbT3V0bG9va10gT3V0bG9vayByZXNwb25kZWQgYWdhaW4g4oCUIHJl"
    "c3VtaW5nIHNjYW4uIikKICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICBv"
    "dXRsb29rID0gd2luMzJjb20uY2xpZW50LkRpc3BhdGNoKCJPdXRsb29rLkFwcGxpY2F0aW9uIikK"
    "ICAgICAgICAgICAgICAgICAgICBucyA9IG91dGxvb2suR2V0TmFtZXNwYWNlKCJNQVBJIikKICAg"
    "ICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgIyBPdXRs"
    "b29rLmV4ZSBhbnN3ZXJlZCB0aGUgbGlnaHR3ZWlnaHQgdGVzdCBjYWxsIGluc2lkZQogICAgICAg"
    "ICAgICAgICAgICAgICMgX3dhaXRfZm9yX291dGxvb2tfcmVzdGFydCBhIG1vbWVudCBhZ28gYnV0"
    "IGlzIHJlZnVzaW5nCiAgICAgICAgICAgICAgICAgICAgIyB0aGlzIG9uZSAoZS5nLiBzdGlsbCBm"
    "aW5pc2hpbmcgaXRzIG93biBzdGFydHVwKSDigJQgbG9vcAogICAgICAgICAgICAgICAgICAgICMg"
    "YmFjayB0byB0aGUgdG9wLCB3aGljaCB3YWl0cyBhZ2FpbiBiZWZvcmUgcmV0cnlpbmcuCiAgICAg"
    "ICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICMgTG9vcCBiYWNrIHRvIHRo"
    "ZSB0b3A6IHJlLXdhbGsgZm9sZGVycyBmcm9tIHNjcmF0Y2ggd2l0aCB0aGUKICAgICAgICAgICAg"
    "ICAgICMgZnJlc2ggYG5zYCwgdXNpbmcgdGhlIG5vdy1sYXJnZXIgYGtub3duYCBkaWN0IHRvIHNr"
    "aXAgZmFzdC4KICAgICAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgIGNvbm4uY29tbWl0KCkK"
    "ICAgICAgICBjb25uLmNsb3NlKCkKICAgICAgICBpZiBwcm9ncmVzc19jYjoKICAgICAgICAgICAg"
    "cHJvZ3Jlc3NfY2IocHJvY2Vzc2VkLCB0b3RhbCwgIkRvbmUiKQogICAgICAgIHJldHVybiBpbmRl"
    "eGVkLCBza2lwcGVkLCBOb25lCiAgICBmaW5hbGx5OgogICAgICAgIHB5dGhvbmNvbS5Db1VuaW5p"
    "dGlhbGl6ZSgpCgoKZGVmIHNlYXJjaF9vdXRsb29rKHF1ZXJ5LCBsaW1pdD0yMDApOgogICAgIiIi"
    "Qk0yNSBzZWFyY2ggb3ZlciBpbmRleGVkIG1haWwgKHN1YmplY3QgKyBib2R5KS4gUmV0dXJucyBh"
    "IGxpc3Qgb2YKICAgIGRpY3RzLCBiZXN0IG1hdGNoIGZpcnN0LiBTYWZlIHRvIGNhbGwgZnJvbSB0"
    "aGUgbWFpbiB0aHJlYWQg4oCUIHRoaXMgaXMKICAgIHB1cmUgc3FsaXRlLCBubyBDT00gaW52b2x2"
    "ZWQgKENPTSBpcyBvbmx5IG5lZWRlZCBmb3IgaW5kZXhpbmcgYW5kIGZvcgogICAgb3BlbmluZyBh"
    "biBpdGVtIGFmdGVyd2FyZHMpLiIiIgogICAgaWYgbm90IHF1ZXJ5IG9yIG5vdCBxdWVyeS5zdHJp"
    "cCgpOgogICAgICAgIHJldHVybiBbXQogICAgaWYgbm90IG9zLnBhdGguZXhpc3RzKE9VVExPT0tf"
    "REJfRklMRSk6CiAgICAgICAgcmV0dXJuIFtdCgogICAgY29ubiA9IHNxbGl0ZTMuY29ubmVjdChP"
    "VVRMT09LX0RCX0ZJTEUpCiAgICBjID0gY29ubi5jdXJzb3IoKQogICAgIyB2LWZpeDogY3JlYXRl"
    "IHRoZSByb3dpZF9mdHMgaW5kZXggaGVyZSB0b28gKGlkZW1wb3RlbnQsIG5lYXItaW5zdGFudAog"
    "ICAgIyBpZiBpdCBhbHJlYWR5IGV4aXN0cykgc28gYW4gRVhJU1RJTkcgc2VhcmNoX291dGxvb2su"
    "ZGIgYnVpbHQgYmVmb3JlCiAgICAjIHRoaXMgZml4IGdldHMgZmFzdCBpbW1lZGlhdGVseSBvbiB0"
    "aGUgbmV4dCBzZWFyY2ggLS0gbm90IG9ubHkgYWZ0ZXIKICAgICMgdGhlIG5leHQgZnVsbCAiVXBk"
    "YXRlIE91dGxvb2sgSW5kZXgiIHJ1biAod2hpY2ggaXMgd2hlbiBfY29ubmVjdCgpCiAgICAjIHdv"
    "dWxkIG90aGVyd2lzZSBiZSB0aGUgb25lIHRvIGNyZWF0ZSBpdCkuCiAgICB0cnk6CiAgICAgICAg"
    "Yy5leGVjdXRlKCJDUkVBVEUgSU5ERVggSUYgTk9UIEVYSVNUUyBpZHhfb3V0bG9va19zdG9yZV9y"
    "b3dpZF9mdHMgIgogICAgICAgICAgICAgICAgICAiT04gb3V0bG9va19zdG9yZShyb3dpZF9mdHMp"
    "IikKICAgIGV4Y2VwdCBzcWxpdGUzLk9wZXJhdGlvbmFsRXJyb3I6CiAgICAgICAgcGFzcyAgIyBl"
    "LmcuIGRiIGJyaWVmbHkgbG9ja2VkIGJ5IGFuIGluLXByb2dyZXNzIGluZGV4IHJ1biAtLSBuZXh0"
    "IHNlYXJjaCB3aWxsIHJldHJ5CiAgICB0cnk6CiAgICAgICAgYy5leGVjdXRlKCIiIgogICAgICAg"
    "ICAgICBTRUxFQ1Qgcy5lbnRyeV9pZCwgcy5zdG9yZV9pZCwgcy5zdWJqZWN0LCBzLnNlbmRlciwg"
    "cy5yZWNlaXZlZCwKICAgICAgICAgICAgICAgICAgIHMuZm9sZGVyX3BhdGgsIHMuc2l6ZSwKICAg"
    "ICAgICAgICAgICAgICAgIHNuaXBwZXQob2Mub3V0bG9va19jb250ZW50LCAxLCAnWycsICddJywg"
    "Jy4uLicsIDEyKSBBUyBzbmlwLAogICAgICAgICAgICAgICAgICAgYm0yNShvYy5vdXRsb29rX2Nv"
    "bnRlbnQpIEFTIHNjb3JlCiAgICAgICAgICAgIEZST00gb3V0bG9va19jb250ZW50IG9jCiAgICAg"
    "ICAgICAgIEpPSU4gb3V0bG9va19zdG9yZSBzIE9OIHMucm93aWRfZnRzID0gb2Mucm93aWQKICAg"
    "ICAgICAgICAgV0hFUkUgb2Mub3V0bG9va19jb250ZW50IE1BVENIID8KICAgICAgICAgICAgT1JE"
    "RVIgQlkgc2NvcmUKICAgICAgICAgICAgTElNSVQgPwogICAgICAgICIiIiwgKHF1ZXJ5LCBsaW1p"
    "dCkpCiAgICAgICAgcm93cyA9IGMuZmV0Y2hhbGwoKQogICAgZXhjZXB0IHNxbGl0ZTMuT3BlcmF0"
    "aW9uYWxFcnJvcjoKICAgICAgICAjIE1hbGZvcm1lZCBGVFM1IHF1ZXJ5IChzdHJheSBxdW90ZXMv"
    "b3BlcmF0b3JzIHR5cGVkIG1pZC1zZWFyY2gpIOKAlAogICAgICAgICMgZmFpbCBzb2Z0IHdpdGgg"
    "bm8gcmVzdWx0cyByYXRoZXIgdGhhbiByYWlzaW5nIGludG8gdGhlIFVJIHRocmVhZC4KICAgICAg"
    "ICByb3dzID0gW10KICAgIGNvbm4uY2xvc2UoKQoKICAgIHJldHVybiBbCiAgICAgICAgewogICAg"
    "ICAgICAgICAiZW50cnlfaWQiOiByWzBdLCAic3RvcmVfaWQiOiByWzFdLCAic3ViamVjdCI6IHJb"
    "Ml0sCiAgICAgICAgICAgICJzZW5kZXIiOiByWzNdLCAicmVjZWl2ZWQiOiByWzRdLCAiZm9sZGVy"
    "X3BhdGgiOiByWzVdLAogICAgICAgICAgICAic2l6ZSI6IHJbNl0sICJzbmlwcGV0Ijogcls3XSwg"
    "InNjb3JlIjogcls4XSwKICAgICAgICB9CiAgICAgICAgZm9yIHIgaW4gcm93cwogICAgXQoKCmRl"
    "ZiBvcGVuX291dGxvb2tfaXRlbShlbnRyeV9pZCwgc3RvcmVfaWQpOgogICAgIiIiT3BlbiB0aGUg"
    "Z2l2ZW4gbWFpbCBpdGVtIGluIE91dGxvb2sncyBvd24gcmVhZGluZyB3aW5kb3cgKGJyaW5ncwog"
    "ICAgT3V0bG9vayB0byBmcm9udCwgZm9jdXNlZCBvbiB0aGF0IGV4YWN0IGVtYWlsKSDigJQgc2Ft"
    "ZSBpZGVhIGFzCiAgICBkb3VibGUtY2xpY2tpbmcgYSBmaWxlIHJlc3VsdCBvcGVucyBFeHBsb3Jl"
    "ci90aGUgZmlsZSBpdHNlbGYuIiIiCiAgICBpZiBub3QgT1VUTE9PS19BVkFJTEFCTEU6CiAgICAg"
    "ICAgcmV0dXJuIEZhbHNlLCAicHl3aW4zMiBjaMawYSDEkcaw4bujYyBjw6BpIgogICAgcHl0aG9u"
    "Y29tLkNvSW5pdGlhbGl6ZSgpCiAgICB0cnk6CiAgICAgICAgb3V0bG9vayA9IHdpbjMyY29tLmNs"
    "aWVudC5EaXNwYXRjaCgiT3V0bG9vay5BcHBsaWNhdGlvbiIpCiAgICAgICAgbnMgPSBvdXRsb29r"
    "LkdldE5hbWVzcGFjZSgiTUFQSSIpCiAgICAgICAgaXRlbSA9IG5zLkdldEl0ZW1Gcm9tSUQoZW50"
    "cnlfaWQsIHN0b3JlX2lkKQogICAgICAgIGl0ZW0uRGlzcGxheSgpCiAgICAgICAgcmV0dXJuIFRy"
    "dWUsIE5vbmUKICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICByZXR1cm4gRmFsc2Us"
    "IHN0cihlKQogICAgZmluYWxseToKICAgICAgICBweXRob25jb20uQ29VbmluaXRpYWxpemUoKQoK"
    "CmRlZiBnZXRfaW5kZXhlZF9tYWlsX2NvdW50KCk6CiAgICAiIiJRdWljayBjb3VudCBmb3IgdGhl"
    "IFVJIChlLmcuIHNob3cgJzEyLDM0MCBtYWlsIGluZGV4ZWQnIG5leHQgdG8gdGhlCiAgICBVcGRh"
    "dGUgT3V0bG9vayBJbmRleCBidXR0b24pLiIiIgogICAgaWYgbm90IG9zLnBhdGguZXhpc3RzKE9V"
    "VExPT0tfREJfRklMRSk6CiAgICAgICAgcmV0dXJuIDAKICAgIHRyeToKICAgICAgICBjb25uID0g"
    "c3FsaXRlMy5jb25uZWN0KE9VVExPT0tfREJfRklMRSkKICAgICAgICBjID0gY29ubi5jdXJzb3Io"
    "KQogICAgICAgIGMuZXhlY3V0ZSgiU0VMRUNUIENPVU5UKCopIEZST00gb3V0bG9va19zdG9yZSIp"
    "CiAgICAgICAgbiA9IGMuZmV0Y2hvbmUoKVswXQogICAgICAgIGNvbm4uY2xvc2UoKQogICAgICAg"
    "IHJldHVybiBuCiAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgIHJldHVybiAw"
)

_ONENOTE_SEARCH_SRC_B64 = (
    "IiIiCm9uZW5vdGVfc2VhcmNoLnB5IOKAlCBJbmRleCAmIHNlYXJjaCBPbmVOb3RlIHBhZ2UgY29u"
    "dGVudCBsb2NhbGx5IG92ZXIgQ09NLgpNaXJyb3JzIG91dGxvb2tfc2VhcmNoLnB5J3MgZGVzaWdu"
    "IChvd24gc3FsaXRlIGRiIGZpbGUsIG93biB0cnkvZXhjZXB0IHNvCmEgbWlzc2luZyBkZXBlbmRl"
    "bmN5IG5ldmVyIGJyZWFrcyB0aGUgcmVzdCBvZiB0aGUgYXBwKSwgYWRhcHRlZCBmb3IgaG93CmRp"
    "ZmZlcmVudGx5IE9uZU5vdGUncyBvYmplY3QgbW9kZWwgd29ya3MgY29tcGFyZWQgdG8gT3V0bG9v"
    "aydzLgoKRGVzaWduIG5vdGVzIChzbyB0aGlzIGZpdHMgc21hcnRTZWFyY2hfTUZUX0JNMjVfQUkn"
    "cyBleGlzdGluZyBhcmNoaXRlY3R1cmUpOgogIC0gT3duIHNxbGl0ZSBkYiBmaWxlIChzZWFyY2hf"
    "b25lbm90ZS5kYiksIHNlcGFyYXRlIGZyb20gc2VhcmNoX2RhdGEuZGIKICAgIEFORCBzZWFyY2hf"
    "b3V0bG9vay5kYiDigJQgc2FtZSByZWFzb25pbmcgYXMgSElTVE9SWV9EQl9GSUxFL09VVExPT0tf"
    "REJfRklMRToKICAgIHNlYXJjaGluZy9pbmRleGluZyBPbmVOb3RlIHNob3VsZCBuZXZlciB0b3Vj"
    "aCBlaXRoZXIgb2YgdGhvc2UgZmlsZXMuCiAgLSBGVFM1IHdpdGggJ3BvcnRlciB1bmljb2RlNjEn"
    "IChPbmVOb3RlIHBhZ2UgY29udGVudCBpcyBwcm9zZSwgc2FtZQogICAgdG9rZW5pemVyIGNob2lj"
    "ZSBhcyBvdXRsb29rX3NlYXJjaC5weSdzIG1haWwgYm9keS9zdWJqZWN0KS4KICAtIEluY3JlbWVu"
    "dGFsOiBlYWNoIHBhZ2UncyBsYXN0TW9kaWZpZWRUaW1lIChmcm9tIE9uZU5vdGUncyBvd24KICAg"
    "IGhpZXJhcmNoeSBYTUwpIGlzIHN0b3JlZDsgdW5jaGFuZ2VkIHBhZ2VzIGFyZSBza2lwcGVkIHdp"
    "dGhvdXQKICAgIHJlLXJlYWRpbmcgdGhlaXIgY29udGVudC4KICAtIFBhc3N3b3JkLXByb3RlY3Rl"
    "ZCBzZWN0aW9ucyBhcmUgc2tpcHBlZCAodGhlaXIgY29udGVudCBsaXRlcmFsbHkKICAgIGNhbid0"
    "IGJlIHJlYWQgdmlhIENPTSB3aXRob3V0IHRoZSB1c2VyIHVubG9ja2luZyB0aGVtIGluIHRoZSBP"
    "bmVOb3RlCiAgICBVSSBmaXJzdCkg4oCUIGNvdW50ZWQgYW5kIHJlcG9ydGVkIGJhY2ssIG5vdCBz"
    "aWxlbnRseSBkcm9wcGVkLgoKICAtIElNUE9SVEFOVCDigJQgZWFybHkgYmluZGluZyByZXF1aXJl"
    "ZDogdGhpcyB1c2VzCiAgICB3aW4zMmNvbS5jbGllbnQuZ2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2go"
    "Ik9uZU5vdGUuQXBwbGljYXRpb24iKSwgTk9UCiAgICBwbGFpbiB3aW4zMmNvbS5jbGllbnQuRGlz"
    "cGF0Y2goKS4gR2V0SGllcmFyY2h5KCkgYW5kIEdldFBhZ2VDb250ZW50KCkKICAgIGJvdGggcmV0"
    "dXJuIHRoZWlyIHJlYWwgcmVzdWx0IHRocm91Z2ggYW4gW291dF0gQlNUUiBwYXJhbWV0ZXIgaW4K"
    "ICAgIE9uZU5vdGUncyBDT00gaW50ZXJmYWNlLiBMYXRlLWJvdW5kIERpc3BhdGNoKCkgaGFzIGEg"
    "bG9uZy1kb2N1bWVudGVkCiAgICBoaXN0b3J5IG9mIG5vdCBtYXJzaGFsaW5nIHRoYXQgcGFydGlj"
    "dWxhciBraW5kIG9mIFtvdXRdIHN0cmluZyBwYXJhbQogICAgYmFjayBjb3JyZWN0bHkgKHNpbGVu"
    "dGx5IHJldHVybnMgTm9uZS9lbXB0eSBmb3IgdGhlc2UgdHdvIGNhbGxzCiAgICBzcGVjaWZpY2Fs"
    "bHksIGV2ZW4gdGhvdWdoIG90aGVyIE9uZU5vdGUgQ09NIGNhbGxzIHdvcmsgZmluZSB3aXRoIGl0"
    "KS4KICAgIGdlbmNhY2hlLkVuc3VyZURpc3BhdGNoIGJ1aWxkcyBhIHJlYWwgd3JhcHBlciBmcm9t"
    "IE9uZU5vdGUncyB0eXBlCiAgICBsaWJyYXJ5IGFoZWFkIG9mIHRpbWUsIHdoaWNoIGhhbmRsZXMg"
    "dGhpcyBjb3JyZWN0bHkuIFRoaXMgaXMgdGhlCiAgICBzaW5nbGUgbW9zdCBjb21tb24gZ290Y2hh"
    "IGluIE9uZU5vdGUgQ09NIGF1dG9tYXRpb24sIHNvIGl0J3MgY2FsbGVkCiAgICBvdXQgZXhwbGlj"
    "aXRseSBoZXJlIHJhdGhlciB0aGFuIGxlZnQgdG8gYmUgcmVkaXNjb3ZlcmVkIHRoZSBoYXJkIHdh"
    "eS4KCiAgLSBVbmxpa2UgT3V0bG9vayAod2hpY2ggaGFzIHRvIGJlIHdhbGtlZCBmb2xkZXIgYnkg"
    "Zm9sZGVyLCBhbmQgd2hlcmUKICAgIHRoYXQgd2FsayBpdHNlbGYgdHVybmVkIG91dCB0byBiZSBh"
    "IHNvdXJjZSBvZiBoYW5ncyDigJQgc2VlCiAgICBvdXRsb29rX3NlYXJjaC5weSdzIGhpc3Rvcnkp"
    "LCBPbmVOb3RlIHJldHVybnMgaXRzIEVOVElSRQogICAgbm90ZWJvb2svc2VjdGlvbi1ncm91cC9z"
    "ZWN0aW9uL3BhZ2UgdHJlZSBpbiBhIFNJTkdMRSBHZXRIaWVyYXJjaHkoKQogICAgY2FsbC4gVGhl"
    "cmUgaXMgbm8gZm9sZGVyLWJ5LWZvbGRlciBkaXNjb3Zlcnkgc3RlcCBoZXJlLCBhbmQg4oCUIGJl"
    "Y2F1c2UKICAgIHRoZSBmdWxsIHBhZ2UgY291bnQgaXMga25vd24gaW1tZWRpYXRlbHkgZnJvbSB0"
    "aGF0IG9uZSBjYWxsIOKAlCBubwogICAgInByb2dyZXNzIGJhciB0b3RhbCBrZWVwcyBncm93aW5n"
    "IiBhbWJpZ3VpdHkgZWl0aGVyLgoKUmVxdWlyZXM6IHB5d2luMzIgKHBpcCBpbnN0YWxsIHB5d2lu"
    "MzIpIGFuZCBPbmVOb3RlIGluc3RhbGxlZCBsb2NhbGx5IOKAlApzcGVjaWZpY2FsbHkgdGhlIGNs"
    "YXNzaWMgZGVza3RvcCB2ZXJzaW9uIHRoYXQgc2hpcHMgd2l0aCBPZmZpY2UgKHNhbWUKZmFtaWx5"
    "IGFzIGNsYXNzaWMgT3V0bG9vaykuIE9uZU5vdGUgZm9yIFdpbmRvd3MgMTAgKHRoZSBNaWNyb3Nv"
    "ZnQgU3RvcmUKYXBwKSBoYXMgTk8gQ09NIGF1dG9tYXRpb24gc3VyZmFjZSBhdCBhbGwgYW5kIGNh"
    "bm5vdCBiZSB1c2VkIHRoaXMgd2F5LgoiIiIKCmltcG9ydCBodG1sCmltcG9ydCBvcwppbXBvcnQg"
    "cmUKaW1wb3J0IHNodXRpbAppbXBvcnQgc3FsaXRlMwppbXBvcnQgc3lzCmltcG9ydCB0aHJlYWRp"
    "bmcKaW1wb3J0IHRpbWUKaW1wb3J0IHRyYWNlYmFjawppbXBvcnQgeG1sLmV0cmVlLkVsZW1lbnRU"
    "cmVlIGFzIEVUCgp0cnk6CiAgICBpbXBvcnQgcHl0aG9uY29tCiAgICBpbXBvcnQgd2luMzJjb20u"
    "Y2xpZW50CiAgICBPTkVOT1RFX0FWQUlMQUJMRSA9IFRydWUKZXhjZXB0IEltcG9ydEVycm9yOgog"
    "ICAgT05FTk9URV9BVkFJTEFCTEUgPSBGYWxzZQoKIyBTYW1lICJ0aGUgb3RoZXIgcHJvY2VzcyBp"
    "cyBnb25lIiBIUkVTVUxUcyBhcyBvdXRsb29rX3NlYXJjaC5weSDigJQgc2VlIHRoYXQKIyBtb2R1"
    "bGUncyBjb21tZW50IGZvciB0aGUgZXhhY3QgbWVhbmluZ3MuIER1cGxpY2F0ZWQgcmF0aGVyIHRo"
    "YW4gaW1wb3J0ZWQKIyBzbyB0aGlzIG1vZHVsZSBrZWVwcyB3b3JraW5nIHN0YW5kYWxvbmUgZXZl"
    "biBpZiBvdXRsb29rX3NlYXJjaC5weSBpcwojIG1pc3NpbmcvYnJva2VuIChtYXRjaGVzIHRoaXMg"
    "ZmlsZSdzIG93biAibmV2ZXIgYnJlYWtzIHRoZSByZXN0IG9mIHRoZQojIGFwcCIgZGVzaWduIG5v"
    "dGUgYWJvdmUpLgpfUlBDX0RJU0NPTk5FQ1RfSFJFU1VMVFMgPSB7LTIxNDcwMjMxNzAsIC0yMTQ3"
    "MDIzMTY5LCAtMjE0NzAyMzE2NiwgLTIxNDc0MTc4NDh9CgoKZGVmIF9pc19kaXNjb25uZWN0KGUp"
    "OgogICAgaHJlc3VsdCA9IGdldGF0dHIoZSwgJ2hyZXN1bHQnLCBOb25lKQogICAgaWYgaHJlc3Vs"
    "dCBpcyBOb25lIGFuZCBnZXRhdHRyKGUsICdhcmdzJywgTm9uZSk6CiAgICAgICAgaHJlc3VsdCA9"
    "IGUuYXJnc1swXSBpZiBpc2luc3RhbmNlKGUuYXJnc1swXSwgaW50KSBlbHNlIE5vbmUKICAgIHJl"
    "dHVybiBocmVzdWx0IGluIF9SUENfRElTQ09OTkVDVF9IUkVTVUxUUwoKCl9PTkVOT1RFX0FQUF9M"
    "T0NLID0gdGhyZWFkaW5nLkxvY2soKQoKCmRlZiBfZW5zdXJlX29uZW5vdGVfYXBwKCk6CiAgICAi"
    "IiJ3aW4zMmNvbS5jbGllbnQuZ2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2goIk9uZU5vdGUuQXBwbGlj"
    "YXRpb24iKSwKICAgIHNlbGYtaGVhbGluZyBhZ2FpbnN0IGEgc3RhbGUvY29ycnVwdGVkIGdlbl9w"
    "eSBjYWNoZSwgYW5kIHNlcmlhbGl6ZWQKICAgIGFnYWluc3QgY29uY3VycmVudCBjYWxsZXJzLgoK"
    "ICAgIHYtZml4IChs4buXaSAiS2jDtG5nIG3hu58gxJHGsOG7o2MgdHJhbmcgbsOgeSB0cm9uZyBP"
    "bmVOb3RlOiAuLi4gTW9kdWxlCiAgICB3aW4zMmNvbS5nZW5fcHkuLi4uaGFzIG5vIGF0dHJpYnV0"
    "ZSAnQ0xTSURUb0NsYXNzTWFwJyIpOiB0aGlzIGlzIGEKICAgIHR3by1wYXJ0IHByb2JsZW0sIGJv"
    "dGggYWRkcmVzc2VkIGhlcmU6CgogICAgMS4gZ2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2ggY2FjaGVz"
    "IGEgUHl0aG9uIHdyYXBwZXIgZ2VuZXJhdGVkIGZyb20KICAgICAgIE9uZU5vdGUncyBDT00gdHlw"
    "ZSBsaWJyYXJ5IG9uIGRpc2ssIHVuZGVyIGEgcGVyLXVzZXIgdGVtcCBmb2xkZXIKICAgICAgICh3"
    "aW4zMmNvbS5fX2dlbl9wYXRoX18sIG5vcm1hbGx5CiAgICAgICAlTE9DQUxBUFBEQVRBJVxcVGVt"
    "cFxcZ2VuX3B5XFw8cHl0aG9uLXZlcnNpb24+XFw8dHlwZWxpYi1ndWlkPikuCiAgICAgICBUaGF0"
    "IGNhY2hlIGNhbiBnbyBzdGFsZS9oYWxmLXdyaXR0ZW4gKE9mZmljZS9PbmVOb3RlIHVwZGF0ZQog"
    "ICAgICAgY2hhbmdlZCB0aGUgdHlwZWxpYidzIHZlcnNpb24vR1VJRCwgYSBwcmV2aW91cyBydW4g"
    "Z290IGtpbGxlZAogICAgICAgbWlkLXdyaXRlLCBldGMpLiBPbmNlIHRoYXQgaGFwcGVucywgRVZF"
    "Ulkgc3Vic2VxdWVudAogICAgICAgRW5zdXJlRGlzcGF0Y2ggY2FsbCBmYWlscyB0aGUgc2FtZSB3"
    "YXkgdW50aWwgdGhlIGZvbGRlciBpcwogICAgICAgY2xlYXJlZCAtLSBoYW5kbGVkIGJlbG93IGJ5"
    "IHdpcGluZyBpdCBhbmQgcmV0cnlpbmcgb25jZS4KCiAgICAyLiBnZW5jYWNoZSdzIGNvZGUtZ2Vu"
    "ZXJhdGlvbiBzdGVwIGlzIE5PVCB0aHJlYWQtc2FmZS4gVGhpcyBhcHAKICAgICAgIGNhbGxzIF9l"
    "bnN1cmVfb25lbm90ZV9hcHAoKSBmcm9tIGJhY2tncm91bmQgdGhyZWFkcyAoZXZlcnkKICAgICAg"
    "IGRvdWJsZS1jbGljay10by1vcGVuIHNwYXducyBpdHMgb3duIHRocmVhZCAtLSBzZWUKICAgICAg"
    "IF9vcGVuX3BhdGhfZm9yX2NoYXQvX29wZW5fb25lbm90ZV9yZXN1bHQpLiBJZiB0d28gdGhyZWFk"
    "cyByZWFjaAogICAgICAgRW5zdXJlRGlzcGF0Y2ggYXQgKG9yIG5lYXIpIHRoZSBzYW1lIG1vbWVu"
    "dCAtLSBlLmcuIGNsaWNraW5nIHR3bwogICAgICAgZGlmZmVyZW50IE9uZU5vdGUgcmVzdWx0cyBi"
    "YWNrIHRvIGJhY2ssIG9yIGNsaWNraW5nIG9uZSB3aGlsZSBhbgogICAgICAgVXBkYXRlIERCIGlu"
    "ZGV4aW5nIHJ1biBpcyBhbHNvIHVzaW5nIE9uZU5vdGUgLS0gZ2VuY2FjaGUgY2FuIHJhY2UKICAg"
    "ICAgIGFuZCBsZWF2ZSBiZWhpbmQgZXhhY3RseSBhIGhhbGYtYnVpbHQgbW9kdWxlIHRoYXQncyBt"
    "aXNzaW5nCiAgICAgICAnQ0xTSURUb0NsYXNzTWFwJy4gX09ORU5PVEVfQVBQX0xPQ0sgYmVsb3cg"
    "c2VyaWFsaXplcyBBTEwgY2FsbHMKICAgICAgIGludG8gdGhpcyBmdW5jdGlvbiBzbyBvbmx5IG9u"
    "ZSB0aHJlYWQgaXMgZXZlciBpbnNpZGUKICAgICAgIEVuc3VyZURpc3BhdGNoL3RoZSBjYWNoZS1y"
    "ZWJ1aWxkIGF0IGEgdGltZSwgd2hpY2ggcmVtb3ZlcyB0aGF0CiAgICAgICByYWNlIGVudGlyZWx5"
    "ICh0aGUgY2FjaGUtY29ycnVwdGlvbiBzeW1wdG9tLCBub3QganVzdCBpdHMKICAgICAgIGNsZWFu"
    "dXApLgoKICAgIElmIEVuc3VyZURpc3BhdGNoIHN0aWxsIGZhaWxzIGFmdGVyIG9uZSBjbGVhbiBy"
    "ZXRyeSwgdGhlIHByb2JsZW0gaXMKICAgIHNvbWV0aGluZyBlbHNlIChPbmVOb3RlIG5vdCBpbnN0"
    "YWxsZWQsIG5vdCB0aGUgY2xhc3NpYyBkZXNrdG9wCiAgICB2ZXJzaW9uLCBldGMuKSBhbmQgdGhl"
    "IGV4Y2VwdGlvbiBpcyBhbGxvd2VkIHRvIHByb3BhZ2F0ZSBhcwogICAgYmVmb3JlLiIiIgogICAg"
    "d2l0aCBfT05FTk9URV9BUFBfTE9DSzoKICAgICAgICB0cnk6CiAgICAgICAgICAgIHJldHVybiB3"
    "aW4zMmNvbS5jbGllbnQuZ2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2goIk9uZU5vdGUuQXBwbGljYXRp"
    "b24iKQogICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgIHRyeToKICAgICAgICAg"
    "ICAgICAgIGdlbl9kaXIgPSBnZXRhdHRyKHdpbjMyY29tLCAiX19nZW5fcGF0aF9fIiwgTm9uZSkK"
    "ICAgICAgICAgICAgICAgIGlmIGdlbl9kaXIgYW5kIG9zLnBhdGguaXNkaXIoZ2VuX2Rpcik6CiAg"
    "ICAgICAgICAgICAgICAgICAgc2h1dGlsLnJtdHJlZShnZW5fZGlyLCBpZ25vcmVfZXJyb3JzPVRy"
    "dWUpCiAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICBwYXNzCiAg"
    "ICAgICAgICAgIGZvciBfbW9kX25hbWUgaW4gW24gZm9yIG4gaW4gbGlzdChzeXMubW9kdWxlcykg"
    "aWYgbi5zdGFydHN3aXRoKCJ3aW4zMmNvbS5nZW5fcHkiKV06CiAgICAgICAgICAgICAgICBzeXMu"
    "bW9kdWxlcy5wb3AoX21vZF9uYW1lLCBOb25lKQogICAgICAgICAgICAjIExldCB0aGlzIHNlY29u"
    "ZCBhdHRlbXB0J3MgZXhjZXB0aW9uIChpZiBhbnkpIHByb3BhZ2F0ZSBhcy1pcwogICAgICAgICAg"
    "ICAjIC0tIGNhbGxlcnMgYWxyZWFkeSBoYW5kbGUgIk9uZU5vdGUgbm90IHJlYWNoYWJsZSIgdmlh"
    "IHRoZWlyCiAgICAgICAgICAgICMgb3duIHRyeS9leGNlcHQgYW5kIHNob3cgYSByZWFsIGVycm9y"
    "IG1lc3NhZ2UgdG8gdGhlIHVzZXIuCiAgICAgICAgICAgIHJldHVybiB3aW4zMmNvbS5jbGllbnQu"
    "Z2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2goIk9uZU5vdGUuQXBwbGljYXRpb24iKQoKCgoKZGVmIF93"
    "YWl0X2Zvcl9vbmVub3RlX3Jlc3RhcnQoc3RvcF9mbGFnPU5vbmUsIHByb2dyZXNzX2NiPU5vbmUs"
    "IHBvbGxfc2VjPTE1KToKICAgICIiIlBvbGwgdW50aWwgT25lTm90ZSByZXNwb25kcyB0byBhIGZy"
    "ZXNoLCBjaGVhcCBDT00gY2FsbCBhZ2FpbiAob3IKICAgIHN0b3BfZmxhZyBzYXlzIGFib3J0KS4g"
    "Tm8gb3ZlcmFsbCB0aW1lb3V0IOKAlCBzdXBwb3J0cyAiY2xvc2UgT25lTm90ZSwKICAgIHJlb3Bl"
    "biBpdCB3aGVuZXZlciIgYW5kIGhhdmluZyBpbmRleGluZyBwaWNrIGJhY2sgdXAgb24gaXRzIG93"
    "bi4iIiIKICAgIHdhaXRlZCA9IDAKICAgIHdoaWxlIFRydWU6CiAgICAgICAgaWYgc3RvcF9mbGFn"
    "IGFuZCBzdG9wX2ZsYWcoKToKICAgICAgICAgICAgcmV0dXJuIEZhbHNlCiAgICAgICAgdHJ5Ogog"
    "ICAgICAgICAgICB3aW4zMmNvbS5jbGllbnQuZ2VuY2FjaGUuRW5zdXJlRGlzcGF0Y2goIk9uZU5v"
    "dGUuQXBwbGljYXRpb24iKQogICAgICAgICAgICByZXR1cm4gVHJ1ZQogICAgICAgIGV4Y2VwdCBF"
    "eGNlcHRpb246CiAgICAgICAgICAgIGlmIHByb2dyZXNzX2NiOgogICAgICAgICAgICAgICAgcHJv"
    "Z3Jlc3NfY2IoMCwgMSwgZiJPbmVOb3RlIMSRw6MgxJHDs25nIOKAlCDEkWFuZyBjaOG7nSBt4buf"
    "IGzhuqFpLi4uICjEkcOjIGNo4budIHt3YWl0ZWQgLy8gNjB9cHt3YWl0ZWQgJSA2MDowMmR9cyki"
    "KQogICAgICAgICAgICB0aW1lLnNsZWVwKHBvbGxfc2VjKQogICAgICAgICAgICB3YWl0ZWQgKz0g"
    "cG9sbF9zZWMKCgpwcmludCgiW29uZW5vdGVfc2VhcmNoXSBtb2R1bGUgbG9hZGVkIOKAlCBidWls"
    "ZCB0YWc6IHYxIChzaW5nbGUtaGllcmFyY2h5LWNhbGwgZGVzaWduKSIpCgpfQkFTRV9ESVIgPSBv"
    "cy5wYXRoLmRpcm5hbWUob3MucGF0aC5hYnNwYXRoKF9fZmlsZV9fKSkKT05FTk9URV9EQl9GSUxF"
    "ID0gb3MucGF0aC5qb2luKF9CQVNFX0RJUiwgJ3NlYXJjaF9vbmVub3RlLmRiJykKCiMgT25lTm90"
    "ZSdzIFhNTCBuYW1lc3BhY2UuIE1pY3Jvc29mdCdzIDIwMTMgc2NoZW1hIGlzIHdoYXQncyBkb2N1"
    "bWVudGVkCiMgYW5kIHVzZWQgYWxtb3N0IHVuaXZlcnNhbGx5IHRvZGF5IOKAlCBpdCByZWFkcyBj"
    "b250ZW50IHdyaXR0ZW4gYnkgYW55CiMgb2xkZXIgT25lTm90ZSB2ZXJzaW9uIGZpbmUgdG9vICh0"
    "aGUgc2NoZW1hIGlzIGFkZGl0aXZlL2NvbXBhdGlibGUpLCBzbwojIHRoZXJlJ3Mgbm8gbmVlZCB0"
    "byBhbHNvIHRyeSAyMDEwJ3MgbmFtZXNwYWNlIFVSSS4KX05TID0gIntodHRwOi8vc2NoZW1hcy5t"
    "aWNyb3NvZnQuY29tL29mZmljZS9vbmVub3RlLzIwMTMvb25lbm90ZX0iCgpfaW5zZXJ0X2Vycm9y"
    "X2NvdW50ID0gMCAgIyB0aHJvdHRsZXMgZGlhZ25vc3RpYyBlcnJvciBwcmludGluZyAobW9kdWxl"
    "LWxldmVsLCBsaWtlIG91dGxvb2tfc2VhcmNoLnB5J3MpCgoKZGVmIF9jb25uZWN0KCk6CiAgICBj"
    "b25uID0gc3FsaXRlMy5jb25uZWN0KE9ORU5PVEVfREJfRklMRSkKICAgIGMgPSBjb25uLmN1cnNv"
    "cigpCiAgICAjIHYtZml4OiBzYW1lIENKSy10b2tlbml6ZXIgZml4IGFzIG91dGxvb2tfc2VhcmNo"
    "LnB5J3MgX2Nvbm5lY3QoKSAtLSBzZWUKICAgICMgdGhhdCBkb2NzdHJpbmcgZm9yIHRoZSBmdWxs"
    "IGV4cGxhbmF0aW9uLiAncG9ydGVyIHVuaWNvZGU2MScgZ+G7mXAgY+G6owogICAgIyBj4bulbSBD"
    "SksgbGnDqm4gdOG7pWMga2jDtG5nIGPDsyBraG/huqNuZyB0cuG6r25nIHRow6BuaCAxIHRva2Vu"
    "LCBraGnhur9uIDEgY+G7pW0gdOG7qwogICAgIyBraMOzYSB0aeG6v25nIE5o4bqtdCBu4bqxbSBn"
    "aeG7r2EgZMOybmcga2jDtG5nIHTDrG0gxJHGsOG7o2MgZMO5IHRyYW5nIE9uZU5vdGUgY8OzIGNo"
    "4bupYQogICAgIyDEkcO6bmcgY+G7pW0gxJHDsy4gxJDhu5VpIHNhbmcgJ3RyaWdyYW0nIGNobyBu"
    "aOG6pXQgcXXDoW4gduG7m2kgY29udGVudF9pbmRleC8KICAgICMgb3V0bG9va19jb250ZW50LCBr"
    "w6htIG1pZ3JhdGlvbiBndWFyZCB0xrDGoW5nIHThu7EgKERST1AgKyByZWJ1aWxkIG7hur91CiAg"
    "ICAjIHBow6F0IGhp4buHbiBEQiBjxakgZMO5bmcgdG9rZW5pemVyIGtow6FjKS4KICAgIHRyeToK"
    "ICAgICAgICBfcm93ID0gYy5leGVjdXRlKAogICAgICAgICAgICAiU0VMRUNUIHNxbCBGUk9NIHNx"
    "bGl0ZV9tYXN0ZXIgV0hFUkUgdHlwZT0ndGFibGUnIEFORCBuYW1lPSdvbmVub3RlX2NvbnRlbnQn"
    "IgogICAgICAgICkuZmV0Y2hvbmUoKQogICAgICAgIGlmIF9yb3cgYW5kIF9yb3dbMF0gYW5kICd0"
    "cmlncmFtJyBub3QgaW4gX3Jvd1swXToKICAgICAgICAgICAgYy5leGVjdXRlKCJEUk9QIFRBQkxF"
    "IElGIEVYSVNUUyBvbmVub3RlX2NvbnRlbnQiKQogICAgICAgICAgICBjLmV4ZWN1dGUoIkRST1Ag"
    "VEFCTEUgSUYgRVhJU1RTIG9uZW5vdGVfc3RvcmUiKQogICAgICAgICAgICBjb25uLmNvbW1pdCgp"
    "CiAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgIHBhc3MKICAgIGMuZXhlY3V0ZSgiIiJDUkVB"
    "VEUgVklSVFVBTCBUQUJMRSBJRiBOT1QgRVhJU1RTIG9uZW5vdGVfY29udGVudAogICAgICAgICAg"
    "ICAgICAgIFVTSU5HIGZ0czUodGl0bGUsIGJvZHksIHRva2VuaXplPSd0cmlncmFtIGNhc2Vfc2Vu"
    "c2l0aXZlIDAnKSIiIikKICAgIGMuZXhlY3V0ZSgiIiJDUkVBVEUgVEFCTEUgSUYgTk9UIEVYSVNU"
    "UyBvbmVub3RlX3N0b3JlICgKICAgICAgICAgICAgICAgICAgICBwYWdlX2lkICAgICAgIFRFWFQg"
    "UFJJTUFSWSBLRVksCiAgICAgICAgICAgICAgICAgICAgdGl0bGUgICAgICAgICBURVhULAogICAg"
    "ICAgICAgICAgICAgICAgIG5vdGVib29rICAgICAgVEVYVCwKICAgICAgICAgICAgICAgICAgICBz"
    "ZWN0aW9uX3BhdGggIFRFWFQsCiAgICAgICAgICAgICAgICAgICAgY3JlYXRlZCAgICAgICBURVhU"
    "LAogICAgICAgICAgICAgICAgICAgIGxhc3RfbW9kaWZpZWQgVEVYVCwKICAgICAgICAgICAgICAg"
    "ICAgICByb3dpZF9mdHMgICAgIElOVEVHRVIKICAgICAgICAgICAgICAgICApIiIiKQogICAgIyB2"
    "LWZpeDogc2FtZSBtaXNzaW5nLWluZGV4IGlzc3VlIGFzIG91dGxvb2tfc2VhcmNoLnB5J3MgX2Nv"
    "bm5lY3QoKSAtLQogICAgIyBzZWFyY2hfb25lbm90ZSgpJ3MgSk9JTiAocy5yb3dpZF9mdHMgPSBv"
    "Yy5yb3dpZCkgaGFzIG5vIGluZGV4IG9uCiAgICAjIHJvd2lkX2Z0cywgc28gaXQgZnVsbC10YWJs"
    "ZS1zY2FucyBvbmVub3RlX3N0b3JlIHBlciBtYXRjaGVkIEZUUzUKICAgICMgcm93LiBOb3QgcGFp"
    "bmZ1bCB5ZXQgYXQgdG9kYXkncyBwYWdlIGNvdW50cywgYnV0IGdyb3dzIHRoZSBzYW1lCiAgICAj"
    "IHdheSBvdXRsb29rX3N0b3JlJ3MgZGlkIC0tIGZpeGVkIHByb2FjdGl2ZWx5IGhlcmUgdG9vLgog"
    "ICAgYy5leGVjdXRlKCJDUkVBVEUgSU5ERVggSUYgTk9UIEVYSVNUUyBpZHhfb25lbm90ZV9zdG9y"
    "ZV9yb3dpZF9mdHMgIgogICAgICAgICAgICAgICJPTiBvbmVub3RlX3N0b3JlKHJvd2lkX2Z0cyki"
    "KQogICAgY29ubi5jb21taXQoKQogICAgcmV0dXJuIGNvbm4sIGMKCgpkZWYgX3dhbGtfaGllcmFy"
    "Y2h5KGVsZW0sIG5vdGVib29rX25hbWUsIHBhdGhfcGFydHMpOgogICAgIiIiUmVjdXJzaXZlbHkg"
    "eWllbGQgKHBhZ2VfZWxlbWVudCwgbm90ZWJvb2tfbmFtZSwgc2VjdGlvbl9wYXRoX2xpc3QpCiAg"
    "ICBmb3IgZXZlcnkgPG9uZTpQYWdlPiB1bmRlciB0aGlzIGVsZW1lbnQgb2YgdGhlIGhpZXJhcmNo"
    "eSBYTUwgdHJlZS4KCiAgICAtIE5vdGVib29rOiBqdXN0IHJlbWVtYmVycyBpdHMgbmFtZSBhbmQg"
    "cmVjdXJzZXMuCiAgICAtIFNlY3Rpb25Hcm91cDogZXh0ZW5kcyB0aGUgZGlzcGxheSBwYXRoIHdp"
    "dGggaXRzIG5hbWUgYW5kIHJlY3Vyc2VzLgogICAgICBPbmVOb3RlIGFsc28gcmVwcmVzZW50cyBp"
    "dHMgb3duIGludGVybmFsIFJlY3ljbGUgQmluIGFzIGEKICAgICAgU2VjdGlvbkdyb3VwIChpc1Jl"
    "Y3ljbGVCaW49InRydWUiKSDigJQgc2tpcHBlZCBzbyBkZWxldGVkIHBhZ2VzIGRvbid0CiAgICAg"
    "IHNob3cgdXAgaW4gc2VhcmNoIHJlc3VsdHMuCiAgICAtIFNlY3Rpb246IGV4dGVuZHMgdGhlIGRp"
    "c3BsYXkgcGF0aCB3aXRoIGl0cyBuYW1lIGFuZCB5aWVsZHMgZWFjaAogICAgICBjaGlsZCBQYWdl"
    "IOKAlCBVTkxFU1MgdGhlIHNlY3Rpb24gaXMgcGFzc3dvcmQtbG9ja2VkIChsb2NrZWQ9InRydWUi"
    "KSwKICAgICAgaW4gd2hpY2ggY2FzZSBpdCdzIHNraXBwZWQgb3V0cmlnaHQgKEdldFBhZ2VDb250"
    "ZW50IHdvdWxkIGp1c3QKICAgICAgdGhyb3cgb24gZXZlcnkgcGFnZSBpbiBpdCBhbnl3YXkpLiIi"
    "IgogICAgdGFnID0gZWxlbS50YWcKICAgIGlmIHRhZyA9PSBmIntfTlN9Tm90ZWJvb2tzIjoKICAg"
    "ICAgICBmb3IgY2hpbGQgaW4gZWxlbToKICAgICAgICAgICAgeWllbGQgZnJvbSBfd2Fsa19oaWVy"
    "YXJjaHkoY2hpbGQsIG5vdGVib29rX25hbWUsIHBhdGhfcGFydHMpCiAgICBlbGlmIHRhZyA9PSBm"
    "IntfTlN9Tm90ZWJvb2siOgogICAgICAgIG5iX25hbWUgPSBlbGVtLmdldCgibmFtZSIsICIiKSBv"
    "ciBub3RlYm9va19uYW1lCiAgICAgICAgZm9yIGNoaWxkIGluIGVsZW06CiAgICAgICAgICAgIHlp"
    "ZWxkIGZyb20gX3dhbGtfaGllcmFyY2h5KGNoaWxkLCBuYl9uYW1lLCBbbmJfbmFtZV0gaWYgbmJf"
    "bmFtZSBlbHNlIFtdKQogICAgZWxpZiB0YWcgPT0gZiJ7X05TfVNlY3Rpb25Hcm91cCI6CiAgICAg"
    "ICAgaWYgZWxlbS5nZXQoImlzUmVjeWNsZUJpbiIpID09ICJ0cnVlIjoKICAgICAgICAgICAgcmV0"
    "dXJuCiAgICAgICAgZ3JwX25hbWUgPSBlbGVtLmdldCgibmFtZSIsICIiKQogICAgICAgIG5ld19w"
    "YXRoID0gcGF0aF9wYXJ0cyArIChbZ3JwX25hbWVdIGlmIGdycF9uYW1lIGVsc2UgW10pCiAgICAg"
    "ICAgZm9yIGNoaWxkIGluIGVsZW06CiAgICAgICAgICAgIHlpZWxkIGZyb20gX3dhbGtfaGllcmFy"
    "Y2h5KGNoaWxkLCBub3RlYm9va19uYW1lLCBuZXdfcGF0aCkKICAgIGVsaWYgdGFnID09IGYie19O"
    "U31TZWN0aW9uIjoKICAgICAgICBpZiBlbGVtLmdldCgibG9ja2VkIikgPT0gInRydWUiOgogICAg"
    "ICAgICAgICByZXR1cm4KICAgICAgICBzZWNfbmFtZSA9IGVsZW0uZ2V0KCJuYW1lIiwgIiIpCiAg"
    "ICAgICAgbmV3X3BhdGggPSBwYXRoX3BhcnRzICsgKFtzZWNfbmFtZV0gaWYgc2VjX25hbWUgZWxz"
    "ZSBbXSkKICAgICAgICBmb3IgY2hpbGQgaW4gZWxlbToKICAgICAgICAgICAgaWYgY2hpbGQudGFn"
    "ID09IGYie19OU31QYWdlIjoKICAgICAgICAgICAgICAgIHlpZWxkIChjaGlsZCwgbm90ZWJvb2tf"
    "bmFtZSwgbmV3X3BhdGgpCiAgICAjIGFueSBvdGhlci91bmV4cGVjdGVkIHRhZyBpcyBzaWxlbnRs"
    "eSBpZ25vcmVkIOKAlCBmb3J3YXJkLWNvbXBhdGlibGUKICAgICMgd2l0aCBzY2hlbWEgYWRkaXRp"
    "b25zIHdlIGRvbid0IGtub3cgYWJvdXQuCgoKX1BTRVVET19UQUdfUkUgPSByZS5jb21waWxlKHIn"
    "PFtePl0rPicpCgoKZGVmIF9jbGVhbl9ydW5fdGV4dChyYXcpOgogICAgIiIiQSA8b25lOlQ+J3Mg"
    "Q0RBVEEgY29udGVudCBpcyBzb21ldGltZXMgTk9UIHBsYWluIHRleHQgLS0gT25lTm90ZQogICAg"
    "d3JhcHMgbWl4ZWQtbGFuZ3VhZ2UvbWl4ZWQtc3R5bGUgcnVucyBpbiBwc2V1ZG8tSFRNTCwgZS5n"
    "LjoKICAgICAgICA8c3BhbiBsYW5nPWphPuOCguOBl+OAgTwvc3Bhbj48c3BhbiBsYW5nPWVuLVVT"
    "PkFiYXF1czwvc3Bhbj48c3BhbiBsYW5nPWphPuOBoOOBkeOBp+OAgTwvc3Bhbj4KICAgIFRoaXMg"
    "bGl2ZXMgSU5TSURFIGEgQ0RBVEEgc2VjdGlvbiwgc28gdGhlIFhNTCBwYXJzZXIgZG9lcyBOT1Qg"
    "dHJlYXQKICAgIGl0IGFzIHJlYWwgY2hpbGQgZWxlbWVudHMgLS0gaXQgaGFuZHMgdGhlIHdob2xl"
    "IHRoaW5nIGJhY2sgYXMgb25lCiAgICBsaXRlcmFsIHRleHQgc3RyaW5nLCA8c3Bhbj4gdGFncyBh"
    "bmQgYWxsLiBMZWZ0IGFzLWlzLCB0aGVzZSBmYWtlCiAgICB0YWdzIGxhbmQgaW4gdGhlIG1pZGRs"
    "ZSBvZiBhIHNlbnRlbmNlIChvZnRlbiBtaWQtd29yZCBmb3IgSmFwYW5lc2UsCiAgICB3aGljaCBo"
    "YXMgbm8gc3BhY2VzKSwgY29ycnVwdGluZyB0aGUgY29udGlndW91cyB0ZXh0IGFuZCBzaWxlbnRs"
    "eQogICAgYnJlYWtpbmcgYW55IHNlYXJjaCBmb3IgYSBwaHJhc2UgdGhhdCBzcGFucyBhIHRhZyBi"
    "b3VuZGFyeSAtLSBlLmcuCiAgICAiQWJhcXVzIENBReOBruODqeOCpOOCu+ODs+OCueOCteODvOOD"
    "kOODvOOBr+OAgUZsZXhMTSIgd2FzIGFjdHVhbGx5IGluZGV4ZWQgYXMKICAgIHNvbWV0aGluZyBs"
    "aWtlICJBYmFxdXMgPHNwYW4gbGFuZz1qYT5DQUXjga4uLi4iLCBzbyBzZWFyY2hpbmcgdGhlIHJl"
    "YWwKICAgIHNlbnRlbmNlIG5ldmVyIG1hdGNoZWQuCgogICAgRml4OiBzdHJpcCB0aGUgcHNldWRv"
    "LXRhZ3MgYW5kIEhUTUwtdW5lc2NhcGUgYW55IGVudGl0aWVzICgmcXVvdDsgZXRjLAogICAgYWxz"
    "byBzdG9yZWQgbGl0ZXJhbGx5IHNpbmNlIENEQVRBIGRpc2FibGVzIGVudGl0eSBleHBhbnNpb24p"
    "IHNvIHRoZQogICAgcmVzdWx0IHJlYWRzIGV4YWN0bHkgbGlrZSB0aGUgc2VudGVuY2UgYSBodW1h"
    "biBzZWVzIGluIE9uZU5vdGUuCiAgICBOT1RFOiB0aGlzIGlzIGEgYmx1bnQgcmVnZXgsIG5vdCBh"
    "IHJlYWwgSFRNTCBwYXJzZXIgLS0gYSByYXJlIG5vdGUKICAgIHRoYXQgZ2VudWluZWx5IHR5cGVz"
    "IGEgbGl0ZXJhbCAiPC4uLj4iIGNvdWxkIGdldCBzdHJpcHBlZCB0b28sIGJ1dAogICAgdGhhdCdz"
    "IGZhciByYXJlciB0aGFuIHRoZSBzeXN0ZW1hdGljIGJyZWFrYWdlIHRoaXMgZml4ZXMuIiIiCiAg"
    "ICBpZiBub3QgcmF3OgogICAgICAgIHJldHVybiAiIgogICAgcmV0dXJuIGh0bWwudW5lc2NhcGUo"
    "X1BTRVVET19UQUdfUkUuc3ViKCcnLCByYXcpKQoKCmRlZiBfZXh0cmFjdF90ZXh0KGNvbnRlbnRf"
    "eG1sKToKICAgICIiIlB1bGwgZXZlcnkgYml0IG9mIHZpc2libGUgdGV4dCBvdXQgb2YgYSBwYWdl"
    "J3MgY29udGVudCBYTUwsIG5vCiAgICBtYXR0ZXIgaG93IGl0J3MgbmVzdGVkIChwbGFpbiBwYXJh"
    "Z3JhcGhzLCB0YWJsZXMsIHRhZ2dlZCBpdGVtcywKICAgIE9DUidkIHRleHQgZnJvbSBpbWFnZXMg"
    "aWYgJ01ha2UgdGV4dCBpbiBpbWFnZXMgc2VhcmNoYWJsZScgaXMgb24sCiAgICBldGMpLgoKICAg"
    "IHYtZml4OiBlYXJsaWVyIHZlcnNpb25zIGpvaW5lZCBFVkVSWSB0ZXh0IGZyYWdtZW50IGZyb20K"
    "ICAgIHJvb3QuaXRlcnRleHQoKSB3aXRoIGFuIGluc2VydGVkICIgIiBiZXR3ZWVuIGVhY2ggb25l"
    "LiBPbmVOb3RlIHZlcnkKICAgIG9mdGVuIHNwbGl0cyBPTkUgY29udGludW91cyBzZW50ZW5jZSBp"
    "bnRvIHNldmVyYWwgPG9uZTpUPiBydW5zCiAgICB3aXRoaW4gdGhlIFNBTUUgbGluZSBwdXJlbHkg"
    "Zm9yIGZvcm1hdHRpbmcgcmVhc29ucyAoc3BlbGwtY2hlY2sKICAgIGZsYWdzLCBoeXBlcmxpbmtz"
    "LCBhdXRvY29ycmVjdCwgYm9sZC9pdGFsaWMgc3BhbnMpIC0tIHdpdGggTk8gc3BhY2UKICAgIGJl"
    "dHdlZW4gdGhlbSBpbiB0aGUgcmVhbCB0ZXh0LiBUaGlzIGlzIGVzcGVjaWFsbHkgZGFtYWdpbmcg"
    "Zm9yCiAgICBKYXBhbmVzZS9DaGluZXNlLCB3aGljaCBoYXZlIG5vIHNwYWNlcyBiZXR3ZWVuIHdv"
    "cmRzIGF0IGFsbDogYQogICAga2V5d29yZCB0aGF0IGhhcHBlbmVkIHRvIHN0cmFkZGxlIG9uZSBv"
    "ZiB0aG9zZSBpbnRlcm5hbCBydW4gc3BsaXRzCiAgICBnb3QgYSBmYWJyaWNhdGVkIHNwYWNlIHNo"
    "b3ZlZCBpbnRvIHRoZSBtaWRkbGUgb2YgaXQsIHNpbGVudGx5CiAgICBicmVha2luZyBzZWFyY2gg"
    "Zm9yIHRoYXQga2V5d29yZCAtLSBldmVuIHRob3VnaCB0aGUgZXhhY3Qgc2FtZQogICAgdGV4dCwg"
    "ZXh0cmFjdGVkIGludGFjdCBlbHNld2hlcmUgKGUuZy4gYW4gT3V0bG9vayBtYWlsIGJvZHksIHdo"
    "aWNoCiAgICBoYXMgbm8gc3VjaCBydW4tc3BsaXR0aW5nKSwgc2VhcmNoZWQgZmluZS4gVGhpcyBp"
    "cyB3aHkgdGhlIHNhbWUKICAgIGtleXdvcmQgY291bGQgYmUgZm91bmQgdmlhIHRoZSBNc2cvT3V0"
    "bG9vayB0YWIgYnV0IG5vdCBPbmVOb3RlLCBldmVuCiAgICB3aGVuIGJvdGggZ2VudWluZWx5IGNv"
    "bnRhaW5lZCBpdC4KCiAgICBGaXggKDEpOiBjb25jYXRlbmF0ZSB0ZXh0IHJ1bnMgV0lUSElOIHRo"
    "ZSBzYW1lIGxpbmUvYnVsbGV0ICh0aGUgZGlyZWN0CiAgICA8b25lOlQ+IGNoaWxkcmVuIG9mIGEg"
    "Z2l2ZW4gPG9uZTpPRT4pIHdpdGggTk8gc2VwYXJhdG9yIC0tIHRoaXMKICAgIHByZXNlcnZlcyB0"
    "aGUgZXhhY3Qgb3JpZ2luYWwgY2hhcmFjdGVyIHNlcXVlbmNlIC0tIGFuZCBvbmx5IGluc2VydAog"
    "ICAgYSBzZXBhcmF0b3IgQkVUV0VFTiBkaWZmZXJlbnQgbGluZXMvYnVsbGV0cyAoZGlmZmVyZW50"
    "IDxvbmU6T0U+CiAgICBlbGVtZW50cyksIHdoaWNoIGlzIGEgcmVhbCwgY29ycmVjdCBsaW5lIGJy"
    "ZWFrLgoKICAgIEZpeCAoMik6IGVhY2ggcnVuJ3MgcmF3IHRleHQgaXMgcGFzc2VkIHRocm91Z2gg"
    "X2NsZWFuX3J1bl90ZXh0KCkKICAgIGZpcnN0LCBzaW5jZSBPbmVOb3RlIG9mdGVuIHN0b3JlcyBt"
    "aXhlZC1sYW5ndWFnZS9taXhlZC1zdHlsZSBydW5zIGFzCiAgICBwc2V1ZG8tSFRNTCA8c3Bhbj4u"
    "Li48L3NwYW4+IHRleHQgKHNlZSBfY2xlYW5fcnVuX3RleHQncyBkb2NzdHJpbmcpCiAgICB0aGF0"
    "IHdvdWxkIG90aGVyd2lzZSBsYW5kIGFzIGxpdGVyYWwgZ2FyYmFnZSBpbiB0aGUgbWlkZGxlIG9m"
    "IHRoZQogICAgc2VudGVuY2UsIGNvcnJ1cHRpbmcgaXQganVzdCBhcyBiYWRseSBhcyBmaXggKDEp"
    "J3MgZmFicmljYXRlZCBzcGFjZXMuCgogICAgRml4ICgzKTogdGV4dCB0aGF0IGxpdmVzIE9VVFNJ"
    "REUgdGhlIDxvbmU6T0U+LzxvbmU6VD4gc3RydWN0dXJlIC0tCiAgICBtb3N0IG5vdGFibHkgT0NS"
    "J2QgdGV4dCBmcm9tIGEgcGFzdGVkIHNjcmVlbnNob3QgKCJNYWtlIHRleHQgaW4KICAgIGltYWdl"
    "cyBzZWFyY2hhYmxlIiksIGJ1dCBhbHNvIHRoaW5ncyBsaWtlIGF0dGFjaG1lbnQgY2FwdGlvbnMg"
    "LS0KICAgIGlzIE5PVCBwYXJ0IG9mIGFueSA8b25lOk9FPidzIGRpcmVjdCA8b25lOlQ+IGNoaWxk"
    "cmVuLCBzbyBmaXggKDEpJ3MKICAgIHN3aXRjaCB0byBvbmx5IHJlYWRpbmcgPG9uZTpPRT4vPG9u"
    "ZTpUPiBzaWxlbnRseSBkcm9wcGVkIGl0IGVudGlyZWx5CiAgICAodGhpcyBpcyB3aHkgcmUtaW5k"
    "ZXhlZCBzZWFyY2hfb25lbm90ZS5kYiBzaHJhbmsgbm90aWNlYWJseSBpbiBzaXplCiAgICBhZnRl"
    "ciB0aGF0IGZpeCkuIFRoaXMgZmFsbGJhY2sgYnJhbmNoIHdhbGtzIGV2ZXJ5IE9USEVSIGVsZW1l"
    "bnQgdG9vCiAgICBhbmQgcGlja3MgdXAgYW55IGxlZnRvdmVyIC50ZXh0IGl0IGZpbmRzLCBzbyBu"
    "b3RoaW5nIHRoYXQgdXNlZCB0byBiZQogICAgc2VhcmNoYWJsZSAocHJlLWZpeC0xKSBnb2VzIG1p"
    "c3NpbmcuIiIiCiAgICB0cnk6CiAgICAgICAgcm9vdCA9IEVULmZyb21zdHJpbmcoY29udGVudF94"
    "bWwpCiAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgIHJldHVybiAiIgoKICAgIGxpbmVzID0g"
    "W10KCiAgICBkZWYgX2hhc19kZXNjZW5kYW50X29lKGVsZW0pOgogICAgICAgIGZvciBjaGlsZCBp"
    "biBlbGVtOgogICAgICAgICAgICBpZiBjaGlsZC50YWcgPT0gZiJ7X05TfU9FIiBvciBfaGFzX2Rl"
    "c2NlbmRhbnRfb2UoY2hpbGQpOgogICAgICAgICAgICAgICAgcmV0dXJuIFRydWUKICAgICAgICBy"
    "ZXR1cm4gRmFsc2UKCiAgICBkZWYgX3dhbGsoZWxlbSk6CiAgICAgICAgaWYgZWxlbS50YWcgPT0g"
    "ZiJ7X05TfU9FIjoKICAgICAgICAgICAgIyBUaGlzIGVsZW1lbnQncyBPV04gdGV4dCBydW5zIG9u"
    "bHkgKGRpcmVjdCA8b25lOlQ+IGNoaWxkcmVuKS4KICAgICAgICAgICAgb3duX3BhcnRzID0gW19j"
    "bGVhbl9ydW5fdGV4dCgiIi5qb2luKHQuaXRlcnRleHQoKSkpIGZvciB0IGluIGVsZW0uZmluZGFs"
    "bChmIntfTlN9VCIpXQogICAgICAgICAgICBvd25fbGluZSA9ICIiLmpvaW4ob3duX3BhcnRzKS5z"
    "dHJpcCgpCiAgICAgICAgICAgIGlmIG93bl9saW5lOgogICAgICAgICAgICAgICAgbGluZXMuYXBw"
    "ZW5kKG93bl9saW5lKQogICAgICAgICAgICAjIFN0aWxsIHJlY3Vyc2UgYmVsb3cgLS0gYW4gT0Ug"
    "Y2FuIGRpcmVjdGx5IGNvbnRhaW4gb3RoZXIKICAgICAgICAgICAgIyB0aGluZ3MgKG5lc3RlZCA8"
    "b25lOk9FQ2hpbGRyZW4+LCBhbiBlbWJlZGRlZCA8b25lOkltYWdlPiwKICAgICAgICAgICAgIyBl"
    "dGMuKSB0aGF0IG5lZWQgdGhlaXIgb3duIGhhbmRsaW5nLCBub3QganVzdCBUIGNoaWxkcmVuLgog"
    "ICAgICAgICAgICBmb3IgY2hpbGQgaW4gZWxlbToKICAgICAgICAgICAgICAgIF93YWxrKGNoaWxk"
    "KQogICAgICAgICAgICByZXR1cm4KCiAgICAgICAgaWYgZWxlbS50YWcgPT0gZiJ7X05TfVQiOgog"
    "ICAgICAgICAgICByZXR1cm4gICMgYWx3YXlzIGNvbnN1bWVkIHZpYSBpdHMgcGFyZW50IDxvbmU6"
    "T0U+IGFib3ZlCgogICAgICAgIGlmIG5vdCBfaGFzX2Rlc2NlbmRhbnRfb2UoZWxlbSk6CiAgICAg"
    "ICAgICAgICMgQSBsZWFmLWlzaCByZWdpb24gd2l0aCBubyA8b25lOk9FPiBsaW5lcyBhbnl3aGVy"
    "ZSBiZWxvdyBpdAogICAgICAgICAgICAjIC0tIGUuZy4gT0NSJ2QgdGV4dCBmcm9tIGEgcGFzdGVk"
    "IHNjcmVlbnNob3QvcHJpbnRvdXQsCiAgICAgICAgICAgICMgYXR0YWNobWVudCBjYXB0aW9ucywg"
    "dGFibGUgbWV0YWRhdGEuIEdyYWIgRVZFUllUSElORyB1bmRlcgogICAgICAgICAgICAjIGl0IGlu"
    "IG9uZSBnbyB2aWEgaXRlcnRleHQoKSAod2hpY2gsIHVubGlrZSBhIHNoYWxsb3cKICAgICAgICAg"
    "ICAgIyAudGV4dCByZWFkLCBhbHNvIHBpY2tzIHVwIGFueSBuZXN0ZWQgc3ViLWVsZW1lbnRzJyB0"
    "ZXh0CiAgICAgICAgICAgICMgQU5EIHRoZWlyIC50YWlsIHRleHQgLS0gaW1wb3J0YW50IGJlY2F1"
    "c2UgbGFyZ2UgIkluc2VydAogICAgICAgICAgICAjIFByaW50b3V0IiBPQ1IgY29udGVudCBpcyBv"
    "ZnRlbiBzdHJ1Y3R1cmVkIHdpdGggaW50ZXJuYWwKICAgICAgICAgICAgIyBzdWItdGFncywgYW5k"
    "IGEgc2hhbGxvdyByZWFkIHdhcyBzaWxlbnRseSB0cnVuY2F0aW5nIGl0CiAgICAgICAgICAgICMg"
    "cmlnaHQgYXQgdGhlIGZpcnN0IHN1Y2ggc3ViLXRhZywgd2hpY2ggaXMgd2h5IGJpZyBPQ1IvCiAg"
    "ICAgICAgICAgICMgcHJpbnRvdXQgcGFnZXMgbG9zdCB0ZW5zIG9mIHRob3VzYW5kcyBvZiBjaGFy"
    "YWN0ZXJzKS4KICAgICAgICAgICAgdHh0ID0gIiIuam9pbihlbGVtLml0ZXJ0ZXh0KCkpCiAgICAg"
    "ICAgICAgIGNsZWFuZWQgPSBfY2xlYW5fcnVuX3RleHQodHh0KS5zdHJpcCgpCiAgICAgICAgICAg"
    "IGlmIGNsZWFuZWQ6CiAgICAgICAgICAgICAgICBsaW5lcy5hcHBlbmQoY2xlYW5lZCkKICAgICAg"
    "ICAgICAgcmV0dXJuICAjIGFscmVhZHkgY29uc3VtZWQgZXZlcnl0aGluZyBiZWxvdzsgZG9uJ3Qg"
    "cmVjdXJzZQoKICAgICAgICAjIE90aGVyd2lzZSB0aGlzIGlzIGEgc3RydWN0dXJhbCBjb250YWlu"
    "ZXIgKE91dGxpbmUsIE9FQ2hpbGRyZW4sCiAgICAgICAgIyBUYWJsZSwgUm93LCBDZWxsLCAuLi4p"
    "IHdpdGggPG9uZTpPRT4gbGluZXMgc29tZXdoZXJlIGJlbmVhdGggaXQKICAgICAgICAjIC0tIGp1"
    "c3QgcmVjdXJzZSBhbmQgbGV0IHRob3NlIE9FcyBoYW5kbGUgdGhlbXNlbHZlcyBpbmRpdmlkdWFs"
    "bHkuCiAgICAgICAgZm9yIGNoaWxkIGluIGVsZW06CiAgICAgICAgICAgIF93YWxrKGNoaWxkKQoK"
    "ICAgIF93YWxrKHJvb3QpCiAgICByZXR1cm4gIiAiLmpvaW4obGluZXMpCgoKZGVmIGluZGV4X29u"
    "ZW5vdGUocHJvZ3Jlc3NfY2I9Tm9uZSwgc3RvcF9mbGFnPU5vbmUpOgogICAgIiIiSW5jcmVtZW50"
    "YWxseSBpbmRleCBldmVyeSBPbmVOb3RlIHBhZ2UgY3VycmVudGx5IHZpc2libGUgdG8KICAgIE9u"
    "ZU5vdGUgKGV2ZXJ5IG9wZW4gbm90ZWJvb2spIGludG8gT05FTk9URV9EQl9GSUxFLgoKICAgIHBy"
    "b2dyZXNzX2NiKGRvbmUsIHRvdGFsLCBjdXJyZW50X2xhYmVsKTogc2FtZSBzaGFwZSBhcwogICAg"
    "b3V0bG9va19zZWFyY2guaW5kZXhfb3V0bG9va19tYWlsKCkncyBjYWxsYmFjay4KICAgIHN0b3Bf"
    "ZmxhZzogb3B0aW9uYWwgemVyby1hcmcgY2FsbGFibGU7IFRydWUgPSBzdG9wIGF0IG5leHQgc2Fm"
    "ZSBwb2ludC4KCiAgICBSZXR1cm5zIChpbmRleGVkX2NvdW50LCBza2lwcGVkX2NvdW50LCBsb2Nr"
    "ZWRfY291bnQsIGVycm9yKSDigJQgZXJyb3IgaXMKICAgIE5vbmUgb24gc3VjY2Vzcywgb3IgYSBz"
    "aG9ydCBodW1hbi1yZWFkYWJsZSBzdHJpbmcgb24gZmFpbHVyZS4KICAgICIiIgogICAgZ2xvYmFs"
    "IF9pbnNlcnRfZXJyb3JfY291bnQKCiAgICBpZiBub3QgT05FTk9URV9BVkFJTEFCTEU6CiAgICAg"
    "ICAgcmV0dXJuIDAsIDAsIDAsICJweXdpbjMyIGNoxrBhIMSRxrDhu6NjIGPDoGkgKHBpcCBpbnN0"
    "YWxsIHB5d2luMzIpIgoKICAgIHB5dGhvbmNvbS5Db0luaXRpYWxpemUoKQogICAgdHJ5OgogICAg"
    "ICAgIHRyeToKICAgICAgICAgICAgIyB2LW5vdGU6IG11c3QgYmUgZ2VuY2FjaGUuRW5zdXJlRGlz"
    "cGF0Y2gg4oCUIHNlZSBtb2R1bGUgZG9jc3RyaW5nLgogICAgICAgICAgICAjIHYtZml4OiBzZWxm"
    "LWhlYWxpbmcgd3JhcHBlciDigJQgc2VlIF9lbnN1cmVfb25lbm90ZV9hcHAoKS4KICAgICAgICAg"
    "ICAgb25lbm90ZSA9IF9lbnN1cmVfb25lbm90ZV9hcHAoKQogICAgICAgIGV4Y2VwdCBFeGNlcHRp"
    "b24gYXMgZToKICAgICAgICAgICAgcmV0dXJuIDAsIDAsIDAsIGYiS2jDtG5nIGvhur90IG7hu5Fp"
    "IMSRxrDhu6NjIE9uZU5vdGUgKMSRw6MgY8OgaSBjaMawYT8pOiB7ZX0iCgogICAgICAgIHRyeToK"
    "ICAgICAgICAgICAgaGllcmFyY2h5X3htbCA9IG9uZW5vdGUuR2V0SGllcmFyY2h5KCIiLCA0KSAg"
    "IyBoc1BhZ2VzPTQgLT4gZnVsbCB0cmVlLCBldmVyeSBub3RlYm9vaywgZG93biB0byBwYWdlcwog"
    "ICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICAgICAgcmV0dXJuIDAsIDAsIDAs"
    "IGYiS2jDtG5nIGzhuqV5IMSRxrDhu6NjIGRhbmggc8OhY2ggT25lTm90ZToge2V9IgoKICAgICAg"
    "ICB0cnk6CiAgICAgICAgICAgIHJvb3QgPSBFVC5mcm9tc3RyaW5nKGhpZXJhcmNoeV94bWwpCiAg"
    "ICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICByZXR1cm4gMCwgMCwgMCwg"
    "ZiJLaMO0bmcgxJHhu41jIMSRxrDhu6NjIGPhuqV1IHRyw7pjIE9uZU5vdGUgKFhNTCBs4buXaSk6"
    "IHtlfSIKCiAgICAgICAgcGFnZXMgPSBsaXN0KF93YWxrX2hpZXJhcmNoeShyb290LCAiIiwgW10p"
    "KQogICAgICAgIHRvdGFsID0gbGVuKHBhZ2VzKQoKICAgICAgICBjb25uLCBjID0gX2Nvbm5lY3Qo"
    "KQogICAgICAgIGMuZXhlY3V0ZSgiU0VMRUNUIHBhZ2VfaWQsIGxhc3RfbW9kaWZpZWQgRlJPTSBv"
    "bmVub3RlX3N0b3JlIikKICAgICAgICBrbm93biA9IGRpY3QoYy5mZXRjaGFsbCgpKQoKICAgICAg"
    "ICBwcm9jZXNzZWQgPSAwCiAgICAgICAgaW5kZXhlZCA9IDAKICAgICAgICBza2lwcGVkID0gMAog"
    "ICAgICAgIGxvY2tlZCA9IDAKCiAgICAgICAgZm9yIHBhZ2VfZWxlbSwgbm90ZWJvb2tfbmFtZSwg"
    "c2VjdGlvbl9wYXRoIGluIHBhZ2VzOgogICAgICAgICAgICBpZiBzdG9wX2ZsYWcgYW5kIHN0b3Bf"
    "ZmxhZygpOgogICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgcHJvY2Vzc2VkICs9IDEK"
    "CiAgICAgICAgICAgIHBhZ2VfaWQgPSBwYWdlX2VsZW0uZ2V0KCJJRCIsICIiKQogICAgICAgICAg"
    "ICB0aXRsZSA9IHBhZ2VfZWxlbS5nZXQoIm5hbWUiLCAiIikgb3IgIih1bnRpdGxlZCkiCiAgICAg"
    "ICAgICAgIGNyZWF0ZWQgPSBwYWdlX2VsZW0uZ2V0KCJkYXRlVGltZSIsICIiKQogICAgICAgICAg"
    "ICBsbSA9IHBhZ2VfZWxlbS5nZXQoImxhc3RNb2RpZmllZFRpbWUiLCAiIikgb3IgY3JlYXRlZAog"
    "ICAgICAgICAgICBzZWN0aW9uX3N0ciA9ICIgPiAiLmpvaW4oW25vdGVib29rX25hbWVdICsgc2Vj"
    "dGlvbl9wYXRoKSBpZiBub3RlYm9va19uYW1lIGVsc2UgIiA+ICIuam9pbihzZWN0aW9uX3BhdGgp"
    "CgogICAgICAgICAgICBpZiBwcm9ncmVzc19jYiBhbmQgcHJvY2Vzc2VkICUgNSA9PSAwOgogICAg"
    "ICAgICAgICAgICAgcHJvZ3Jlc3NfY2IocHJvY2Vzc2VkLCB0b3RhbCwgZiJ7c2VjdGlvbl9zdHJ9"
    "ID4ge3RpdGxlfSIpCgogICAgICAgICAgICBpZiBub3QgcGFnZV9pZDoKICAgICAgICAgICAgICAg"
    "IHNraXBwZWQgKz0gMQogICAgICAgICAgICAgICAgY29udGludWUKCiAgICAgICAgICAgIGlmIGxt"
    "IGFuZCBrbm93bi5nZXQocGFnZV9pZCkgPT0gbG06CiAgICAgICAgICAgICAgICBza2lwcGVkICs9"
    "IDEKICAgICAgICAgICAgICAgIGlmIHByb2Nlc3NlZCAlIDUwMCA9PSAwOgogICAgICAgICAgICAg"
    "ICAgICAgIGNvbm4uY29tbWl0KCkKICAgICAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgICAg"
    "ICB0cnk6CiAgICAgICAgICAgICAgICBjb250ZW50X3htbCA9IG9uZW5vdGUuR2V0UGFnZUNvbnRl"
    "bnQocGFnZV9pZCkKICAgICAgICAgICAgICAgIGJvZHkgPSBfZXh0cmFjdF90ZXh0KGNvbnRlbnRf"
    "eG1sKQogICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgICAgICAgICBp"
    "ZiBfaXNfZGlzY29ubmVjdChlKToKICAgICAgICAgICAgICAgICAgICAjIE9uZU5vdGUncyBvd24g"
    "cHJvY2VzcyB3ZW50IGF3YXkgKGNsb3NlZCBieSB0aGUgdXNlciwKICAgICAgICAgICAgICAgICAg"
    "ICAjIGNyYXNoZWQsIGV0YykuIFVubGlrZSBhIGxvY2tlZC9taWQtc3luYyBwYWdlLCB0aGlzCiAg"
    "ICAgICAgICAgICAgICAgICAgIyBpc24ndCBzcGVjaWZpYyB0byBUSElTIHBhZ2Ug4oCUIGV2ZXJ5"
    "IHN1YnNlcXVlbnQKICAgICAgICAgICAgICAgICAgICAjIEdldFBhZ2VDb250ZW50IGNhbGwgd291"
    "bGQgZmFpbCB0aGUgc2FtZSB3YXkuIFBhdXNlLAogICAgICAgICAgICAgICAgICAgICMgd2FpdCBm"
    "b3IgT25lTm90ZSB0byBjb21lIGJhY2ssIGdldCBhIGZyZXNoIENPTSBoYW5kbGUsCiAgICAgICAg"
    "ICAgICAgICAgICAgIyBhbmQgcmV0cnkgdGhpcyBTQU1FIHBhZ2UgKHBhZ2VfaWQgaXMganVzdCBh"
    "IHN0cmluZywKICAgICAgICAgICAgICAgICAgICAjIG5vdCBhIGxpdmUgQ09NIG9iamVjdCwgc28g"
    "bm90aGluZyBhYm91dCBgcGFnZXNgIG9yIG91cgogICAgICAgICAgICAgICAgICAgICMgcG9zaXRp"
    "b24gaW4gaXQgbmVlZHMgdG8gYmUgcmVkb25lKS4KICAgICAgICAgICAgICAgICAgICBjb25uLmNv"
    "bW1pdCgpCiAgICAgICAgICAgICAgICAgICAgaWYgcHJvZ3Jlc3NfY2I6CiAgICAgICAgICAgICAg"
    "ICAgICAgICAgIHByb2dyZXNzX2NiKHByb2Nlc3NlZCwgdG90YWwsICJPbmVOb3RlIMSRw6MgxJHD"
    "s25nIGdp4buvYSBjaOG7q25nIOKAlCDEkWFuZyBjaOG7nSBt4bufIGzhuqFpLi4uIikKICAgICAg"
    "ICAgICAgICAgICAgICBpZiBub3QgX3dhaXRfZm9yX29uZW5vdGVfcmVzdGFydChzdG9wX2ZsYWcs"
    "IHByb2dyZXNzX2NiKToKICAgICAgICAgICAgICAgICAgICAgICAgY29ubi5jb21taXQoKQogICAg"
    "ICAgICAgICAgICAgICAgICAgICBjb25uLmNsb3NlKCkKICAgICAgICAgICAgICAgICAgICAgICAg"
    "aWYgcHJvZ3Jlc3NfY2I6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBwcm9ncmVzc19jYihw"
    "cm9jZXNzZWQsIHRvdGFsLCAiRG9uZSAoZOG7q25nIGdp4buvYSBjaOG7q25nKSIpCiAgICAgICAg"
    "ICAgICAgICAgICAgICAgIHJldHVybiBpbmRleGVkLCBza2lwcGVkLCBsb2NrZWQsIE5vbmUKICAg"
    "ICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIG9uZW5vdGUgPSBf"
    "ZW5zdXJlX29uZW5vdGVfYXBwKCkKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9u"
    "OgogICAgICAgICAgICAgICAgICAgICAgICBwYXNzICAjIHN0aWxsIG5vdCByZWFkeSAtLSBmYWxs"
    "IHRocm91Z2gsIHJldHJ5IGJlbG93IGFueXdheQogICAgICAgICAgICAgICAgICAgIHRyeToKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgY29udGVudF94bWwgPSBvbmVub3RlLkdldFBhZ2VDb250ZW50"
    "KHBhZ2VfaWQpCiAgICAgICAgICAgICAgICAgICAgICAgIGJvZHkgPSBfZXh0cmFjdF90ZXh0KGNv"
    "bnRlbnRfeG1sKQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZTI6CiAg"
    "ICAgICAgICAgICAgICAgICAgICAgICMgUmV0cmllZCBvbmNlIHJpZ2h0IGFmdGVyIHJlY29ubmVj"
    "dGluZyBhbmQgaXQgc3RpbGwKICAgICAgICAgICAgICAgICAgICAgICAgIyBmYWlsZWQgLS0gdHJl"
    "YXQgYXMgYSBnZW51aW5lIHBlci1wYWdlIGlzc3VlIHRoaXMKICAgICAgICAgICAgICAgICAgICAg"
    "ICAgIyB0aW1lIHJhdGhlciB0aGFuIGxvb3BpbmcgZm9yZXZlciBvbiBvbmUgcGFnZS4KICAgICAg"
    "ICAgICAgICAgICAgICAgICAgbG9ja2VkICs9IDEKICAgICAgICAgICAgICAgICAgICAgICAgaWYg"
    "X2luc2VydF9lcnJvcl9jb3VudCA8IDg6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBfaW5z"
    "ZXJ0X2Vycm9yX2NvdW50ICs9IDEKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHByaW50KGYi"
    "W09uZU5vdGVdW0dldFBhZ2VDb250ZW50IEVSUk9SICN7X2luc2VydF9lcnJvcl9jb3VudH1dIHt0"
    "aXRsZSFyfSAtPiB7ZTIhcn0iKQogICAgICAgICAgICAgICAgICAgICAgICBpZiBwcm9jZXNzZWQg"
    "JSA1MDAgPT0gMDoKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbm4uY29tbWl0KCkKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgIGVsc2U6CiAgICAg"
    "ICAgICAgICAgICAgICAgIyBNb3N0IGNvbW1vbmx5IGEgbG9ja2VkL3Bhc3N3b3JkLXByb3RlY3Rl"
    "ZCBzZWN0aW9uIHRoYXQKICAgICAgICAgICAgICAgICAgICAjIHNsaXBwZWQgdGhyb3VnaCAoZS5n"
    "LiBsb2NrZWQgbWlkLXJ1biksIG9yIGEgcGFnZSBtaWQtc3luYy4KICAgICAgICAgICAgICAgICAg"
    "ICBsb2NrZWQgKz0gMQogICAgICAgICAgICAgICAgICAgIGlmIF9pbnNlcnRfZXJyb3JfY291bnQg"
    "PCA4OgogICAgICAgICAgICAgICAgICAgICAgICBfaW5zZXJ0X2Vycm9yX2NvdW50ICs9IDEKICAg"
    "ICAgICAgICAgICAgICAgICAgICAgcHJpbnQoZiJbT25lTm90ZV1bR2V0UGFnZUNvbnRlbnQgRVJS"
    "T1IgI3tfaW5zZXJ0X2Vycm9yX2NvdW50fV0ge3RpdGxlIXJ9IC0+IHtlIXJ9IikKICAgICAgICAg"
    "ICAgICAgICAgICBpZiBwcm9jZXNzZWQgJSA1MDAgPT0gMDoKICAgICAgICAgICAgICAgICAgICAg"
    "ICAgY29ubi5jb21taXQoKQogICAgICAgICAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgICAg"
    "ICB0cnk6CiAgICAgICAgICAgICAgICBjLmV4ZWN1dGUoIlNFTEVDVCByb3dpZF9mdHMgRlJPTSBv"
    "bmVub3RlX3N0b3JlIFdIRVJFIHBhZ2VfaWQ9PyIsIChwYWdlX2lkLCkpCiAgICAgICAgICAgICAg"
    "ICByb3cgPSBjLmZldGNob25lKCkKICAgICAgICAgICAgICAgIGlmIHJvdyBhbmQgcm93WzBdIGlz"
    "IG5vdCBOb25lOgogICAgICAgICAgICAgICAgICAgIGMuZXhlY3V0ZSgiREVMRVRFIEZST00gb25l"
    "bm90ZV9jb250ZW50IFdIRVJFIHJvd2lkPT8iLCAocm93WzBdLCkpCgogICAgICAgICAgICAgICAg"
    "Yy5leGVjdXRlKCJJTlNFUlQgSU5UTyBvbmVub3RlX2NvbnRlbnQgKHRpdGxlLCBib2R5KSBWQUxV"
    "RVMgKD8sPykiLCAodGl0bGUsIGJvZHkpKQogICAgICAgICAgICAgICAgZnRzX3Jvd2lkID0gYy5s"
    "YXN0cm93aWQKCiAgICAgICAgICAgICAgICBjLmV4ZWN1dGUoIiIiSU5TRVJUIElOVE8gb25lbm90"
    "ZV9zdG9yZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgIChwYWdlX2lkLCB0aXRsZSwgbm90"
    "ZWJvb2ssIHNlY3Rpb25fcGF0aCwgY3JlYXRlZCwgbGFzdF9tb2RpZmllZCwgcm93aWRfZnRzKQog"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgIFZBTFVFUyAoPyw/LD8sPyw/LD8sPykKICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICBPTiBDT05GTElDVChwYWdlX2lkKSBETyBVUERBVEUgU0VU"
    "CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICB0aXRsZT1leGNsdWRlZC50aXRsZSwgbm90"
    "ZWJvb2s9ZXhjbHVkZWQubm90ZWJvb2ssCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBz"
    "ZWN0aW9uX3BhdGg9ZXhjbHVkZWQuc2VjdGlvbl9wYXRoLCBjcmVhdGVkPWV4Y2x1ZGVkLmNyZWF0"
    "ZWQsCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBsYXN0X21vZGlmaWVkPWV4Y2x1ZGVk"
    "Lmxhc3RfbW9kaWZpZWQsIHJvd2lkX2Z0cz1leGNsdWRlZC5yb3dpZF9mdHMiIiIsCiAgICAgICAg"
    "ICAgICAgICAgICAgICAgICAgKHBhZ2VfaWQsIHRpdGxlLCBub3RlYm9va19uYW1lLCBzZWN0aW9u"
    "X3N0ciwgY3JlYXRlZCwgbG0sIGZ0c19yb3dpZCkpCgogICAgICAgICAgICAgICAgaW5kZXhlZCAr"
    "PSAxCiAgICAgICAgICAgICAgICBpZiBpbmRleGVkICUgNTAgPT0gMDoKICAgICAgICAgICAgICAg"
    "ICAgICBjb25uLmNvbW1pdCgpCiAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAg"
    "ICAgICAgICAgICAgIHNraXBwZWQgKz0gMQogICAgICAgICAgICAgICAgaWYgX2luc2VydF9lcnJv"
    "cl9jb3VudCA8IDg6CiAgICAgICAgICAgICAgICAgICAgX2luc2VydF9lcnJvcl9jb3VudCArPSAx"
    "CiAgICAgICAgICAgICAgICAgICAgcHJpbnQoZiJbT25lTm90ZV1bSU5TRVJUIEVSUk9SICN7X2lu"
    "c2VydF9lcnJvcl9jb3VudH1dIC0+IHtlIXJ9IikKICAgICAgICAgICAgICAgICAgICB0cmFjZWJh"
    "Y2sucHJpbnRfZXhjKCkKCiAgICAgICAgICAgICMgdi1ub3RlOiBzYW1lIGxlc3NvbiBsZWFybmVk"
    "IGZyb20gb3V0bG9va19zZWFyY2gucHkg4oCUIGNvbW1pdAogICAgICAgICAgICAjIG9uIGEgcHJv"
    "Y2Vzc2VkLWNvdW50IGNhZGVuY2UgdG9vIChub3QganVzdCBvbiBzdWNjZXNzZnVsCiAgICAgICAg"
    "ICAgICMgaW5zZXJ0cyksIHNvIGEgbG9uZyBydW4gb2Ygc2tpcHBlZC9sb2NrZWQgcGFnZXMgc3Rp"
    "bGwKICAgICAgICAgICAgIyBmbHVzaGVzIHRvIGRpc2sgcGVyaW9kaWNhbGx5IGluc3RlYWQgb2Yg"
    "Z29pbmcgc2lsZW50LgogICAgICAgICAgICBpZiBwcm9jZXNzZWQgJSA1MDAgPT0gMDoKICAgICAg"
    "ICAgICAgICAgIGNvbm4uY29tbWl0KCkKCiAgICAgICAgY29ubi5jb21taXQoKQogICAgICAgIGNv"
    "bm4uY2xvc2UoKQogICAgICAgIGlmIHByb2dyZXNzX2NiOgogICAgICAgICAgICBwcm9ncmVzc19j"
    "Yihwcm9jZXNzZWQsIHRvdGFsLCAiRG9uZSIpCiAgICAgICAgcmV0dXJuIGluZGV4ZWQsIHNraXBw"
    "ZWQsIGxvY2tlZCwgTm9uZQogICAgZmluYWxseToKICAgICAgICBweXRob25jb20uQ29VbmluaXRp"
    "YWxpemUoKQoKCmRlZiBzZWFyY2hfb25lbm90ZShxdWVyeSwgbGltaXQ9MjAwKToKICAgICIiIkJN"
    "MjUgc2VhcmNoIG92ZXIgaW5kZXhlZCBPbmVOb3RlIHBhZ2VzICh0aXRsZSArIGJvZHkpLiBSZXR1"
    "cm5zIGEKICAgIGxpc3Qgb2YgZGljdHMsIGJlc3QgbWF0Y2ggZmlyc3QuIFNhZmUgdG8gY2FsbCBm"
    "cm9tIHRoZSBtYWluIHRocmVhZCDigJQKICAgIHB1cmUgc3FsaXRlLCBubyBDT00gaW52b2x2ZWQu"
    "IiIiCiAgICBpZiBub3QgcXVlcnkgb3Igbm90IHF1ZXJ5LnN0cmlwKCk6CiAgICAgICAgcmV0dXJu"
    "IFtdCiAgICBpZiBub3Qgb3MucGF0aC5leGlzdHMoT05FTk9URV9EQl9GSUxFKToKICAgICAgICBy"
    "ZXR1cm4gW10KCiAgICBjb25uID0gc3FsaXRlMy5jb25uZWN0KE9ORU5PVEVfREJfRklMRSkKICAg"
    "IGMgPSBjb25uLmN1cnNvcigpCiAgICAjIHYtZml4OiBzYW1lIHJlYXNvbmluZyBhcyBvdXRsb29r"
    "X3NlYXJjaC5zZWFyY2hfb3V0bG9vaygpIC0tIGNyZWF0ZQogICAgIyB0aGUgcm93aWRfZnRzIGlu"
    "ZGV4IGhlcmUgdG9vIHNvIGFuIGV4aXN0aW5nIHNlYXJjaF9vbmVub3RlLmRiIGdldHMKICAgICMg"
    "ZmFzdCBpbW1lZGlhdGVseSwgbm90IG9ubHkgYWZ0ZXIgdGhlIG5leHQgZnVsbCByZS1pbmRleC4K"
    "ICAgIHRyeToKICAgICAgICBjLmV4ZWN1dGUoIkNSRUFURSBJTkRFWCBJRiBOT1QgRVhJU1RTIGlk"
    "eF9vbmVub3RlX3N0b3JlX3Jvd2lkX2Z0cyAiCiAgICAgICAgICAgICAgICAgICJPTiBvbmVub3Rl"
    "X3N0b3JlKHJvd2lkX2Z0cykiKQogICAgZXhjZXB0IHNxbGl0ZTMuT3BlcmF0aW9uYWxFcnJvcjoK"
    "ICAgICAgICBwYXNzCiAgICB0cnk6CiAgICAgICAgYy5leGVjdXRlKCIiIgogICAgICAgICAgICBT"
    "RUxFQ1Qgcy5wYWdlX2lkLCBzLnRpdGxlLCBzLm5vdGVib29rLCBzLnNlY3Rpb25fcGF0aCwKICAg"
    "ICAgICAgICAgICAgICAgIHMuY3JlYXRlZCwgcy5sYXN0X21vZGlmaWVkLAogICAgICAgICAgICAg"
    "ICAgICAgc25pcHBldChvYy5vbmVub3RlX2NvbnRlbnQsIDEsICdbJywgJ10nLCAnLi4uJywgMTIp"
    "IEFTIHNuaXAsCiAgICAgICAgICAgICAgICAgICBibTI1KG9jLm9uZW5vdGVfY29udGVudCkgQVMg"
    "c2NvcmUKICAgICAgICAgICAgRlJPTSBvbmVub3RlX2NvbnRlbnQgb2MKICAgICAgICAgICAgSk9J"
    "TiBvbmVub3RlX3N0b3JlIHMgT04gcy5yb3dpZF9mdHMgPSBvYy5yb3dpZAogICAgICAgICAgICBX"
    "SEVSRSBvYy5vbmVub3RlX2NvbnRlbnQgTUFUQ0ggPwogICAgICAgICAgICBPUkRFUiBCWSBzY29y"
    "ZQogICAgICAgICAgICBMSU1JVCA/CiAgICAgICAgIiIiLCAocXVlcnksIGxpbWl0KSkKICAgICAg"
    "ICByb3dzID0gYy5mZXRjaGFsbCgpCiAgICBleGNlcHQgc3FsaXRlMy5PcGVyYXRpb25hbEVycm9y"
    "OgogICAgICAgICMgTWFsZm9ybWVkIEZUUzUgcXVlcnkgKHN0cmF5IHF1b3Rlcy9vcGVyYXRvcnMg"
    "dHlwZWQgbWlkLXNlYXJjaCkg4oCUCiAgICAgICAgIyBmYWlsIHNvZnQgd2l0aCBubyByZXN1bHRz"
    "LCBzYW1lIGFzIG91dGxvb2tfc2VhcmNoLnNlYXJjaF9vdXRsb29rLgogICAgICAgIHJvd3MgPSBb"
    "XQogICAgY29ubi5jbG9zZSgpCgogICAgcmV0dXJuIFsKICAgICAgICB7CiAgICAgICAgICAgICJw"
    "YWdlX2lkIjogclswXSwgInRpdGxlIjogclsxXSwgIm5vdGVib29rIjogclsyXSwgInNlY3Rpb25f"
    "cGF0aCI6IHJbM10sCiAgICAgICAgICAgICJjcmVhdGVkIjogcls0XSwgImxhc3RfbW9kaWZpZWQi"
    "OiByWzVdLCAic25pcHBldCI6IHJbNl0sICJzY29yZSI6IHJbN10sCiAgICAgICAgfQogICAgICAg"
    "IGZvciByIGluIHJvd3MKICAgIF0KCgpkZWYgb3Blbl9vbmVub3RlX3BhZ2UocGFnZV9pZCk6CiAg"
    "ICAiIiJCcmluZyBPbmVOb3RlIHRvIGZyb250LCBuYXZpZ2F0ZWQgdG8gdGhpcyBleGFjdCBwYWdl"
    "IOKAlCBzYW1lIGlkZWEKICAgIGFzIG91dGxvb2tfc2VhcmNoLm9wZW5fb3V0bG9va19pdGVtKCku"
    "CgogICAgdi1maXggKGNodeG7l2kgbOG7l2kgMyDEkeG7nWkga2hpIG3hu58gdHJhbmcgT25lTm90"
    "ZSk6CiAgICAxKSAiLi4uaGFzIG5vIGF0dHJpYnV0ZSAnQ0xTSURUb0NsYXNzTWFwJyIgLS0gZ2Vu"
    "X3B5IGNhY2hlIGjhu49uZyBraGkKICAgICAgIGTDuW5nIGdlbmNhY2hlLkVuc3VyZURpc3BhdGNo"
    "LgogICAgMikgIkNvdWxkIG5vdCBvcGVuIHBhZ2U6IE9uZU5vdGUuQXBwbGljYXRpb24uTmF2aWdh"
    "dGVUbyIKICAgICAgIChBdHRyaWJ1dGVFcnJvcikgLS0gdMaw4bufbmcgxJHDoyBuw6kgxJHGsOG7"
    "o2MgKDEpIGLhurFuZwogICAgICAgd2luMzJjb20uY2xpZW50LkRpc3BhdGNoKCkgKFwibGF0ZS1i"
    "b3VuZFwiKSwgbmjGsG5nCiAgICAgICB3aW4zMmNvbS5jbGllbnQuRGlzcGF0Y2goKSBLSMOUTkcg"
    "aOG7gSBwdXJlIGxhdGUtYm91bmQgbmjGsCB0w6puIGfhu41pOgogICAgICAgbsOzIHThu7EgZMOy"
    "IHhlbSBvYmplY3QgY8OzIGV4cG9zZSB0eXBlIGluZm8ga2jDtG5nIHLhu5NpIMOCTSBUSOG6pk0g"
    "dGjhu60KICAgICAgIGdlbmVyYXRlL2TDuW5nIDEgbW9kdWxlIGdlbmNhY2hlIGNobyBuw7MgKHF1"
    "YQogICAgICAgZ2VuY2FjaGUuR2V0TW9kdWxlRm9yQ0xTSUQpIFRSxq/hu5pDIEtISSByxqFpIHbh"
    "u4EgZHluYW1pYyBkaXNwYXRjaCAtLQogICAgICAgbsOqbiB24bqrbiBkw61uaCDEkcO6bmcgY8O5"
    "bmcgduG6pW4gxJHhu4EgZ2VuX3B5IGNhY2hlIG7hu61hIHbhu51pL2jhu49uZyBuaMawICgxKSwK"
    "ICAgICAgIGNo4buJIGtow6FjIGJp4buDdSBoaeG7h24gdGjDoG5oIEF0dHJpYnV0ZUVycm9yICht"
    "ZW1iZXIgIk5hdmlnYXRlVG8iCiAgICAgICBraMO0bmcgdHJhIMSRxrDhu6NjIHRyw6puIDEgbW9k"
    "dWxlIGdlbmNhY2hlIGLhu4sgc2luaCBs4buXaSkuCiAgICAzKSAiVGhpcyBDT00gb2JqZWN0IGNh"
    "biBub3QgYXV0b21hdGUgdGhlIG1ha2VweSBwcm9jZXNzIC0gcGxlYXNlIHJ1bgogICAgICAgbWFr"
    "ZXB5IG1hbnVhbGx5IGZvciB0aGlzIG9iamVjdCIgLS0gbOG6p24gZmFsbGJhY2sgdHLGsOG7m2Mg"
    "xJHDsyBn4buNaQogICAgICAgdGjhurNuZyBnZW5jYWNoZS5FbnN1cmVEaXNwYXRjaCgpIChxdWEg"
    "X2Vuc3VyZV9vbmVub3RlX2FwcCgpKSDEkeG7gyBuw6kKICAgICAgICgyKSwgbmjGsG5nIGLhuqNu"
    "IGPDoGkgxJHhurd0IE9uZU5vdGUgbsOgeSAocuG6pXQgY8OzIHRo4buDIGzDoCBPbmVOb3RlIGZv"
    "cgogICAgICAgV2luZG93cyBi4bqjbiBVV1AvU3RvcmUgdGhheSB2w6wgT25lTm90ZSAyMDE2IGRl"
    "c2t0b3AgY+G7lSDEkWnhu4NuKSBjw7MKICAgICAgIHR5cGUgbGlicmFyeSBLSMOUTkcgaOG7lyB0"
    "cuG7oyBjb2RlLWdlbmVyYXRpb24gcXVhIG1ha2VweSBDSE8gT0JKRUNUCiAgICAgICBOw4BZIC0t"
    "IGTDuSBjYWNoZSDEkcOjIMSRxrDhu6NjIHhvw6Egc+G6oWNoL3NpbmggbOG6oWksIGdlbmNhY2hl"
    "IHbhuqtuIGtow7RuZyB0aOG7gwogICAgICAgdOG6oW8gbW9kdWxlIGNobyBuw7MsIG7Dqm4gRW5z"
    "dXJlRGlzcGF0Y2ggbHXDtG4gdGjhuqV0IGLhuqFpIGLhuqV0IGvhu4MgdHLhuqFuZwogICAgICAg"
    "dGjDoWkgY2FjaGUuCgogICAgxJBp4buDbSBjaHVuZyBj4bunYSBj4bqjIDMgbOG7l2k6IFThuqRU"
    "IEPhuqIgxJHhu4F1IGLhuq90IG5ndeG7k24gdOG7qyB2aeG7h2MgZ2VuY2FjaGUvbWFrZXB5CiAg"
    "ICBi4buLIMSR4bulbmcgdOG7m2kgdGhlbyBjw6FjaCBuw6B5IGhheSBjw6FjaCBraMOhYy4gRml4"
    "IHRyaeG7h3QgxJHhu4M6IGTDuW5nIHRo4bqzbmcKICAgIGB3aW4zMmNvbS5jbGllbnQuZHluYW1p"
    "Yy5EaXNwYXRjaCgpYCAtLSDEkcOieSBt4bubaSBsw6AgZHluYW1pYyBkaXNwYXRjaAogICAgVEjh"
    "uqxUIFPhu7AgKG1vZHVsZSBjb24gYGR5bmFtaWNgIGPhu6dhIHdpbjMyY29tLmNsaWVudCksIEtI"
    "w5RORyBCQU8gR0nhu5wKICAgIGNo4bqhbSB2w6BvIGdlbmNhY2hlL2dlbl9weS9tYWtlcHkgZMaw"
    "4bubaSBi4bqldCBr4buzIGjDrG5oIHRo4bupYyBuw6BvLCBjaOG7iSBn4buNaQogICAgbWV0aG9k"
    "IHF1YSBJRGlzcGF0Y2g6Okludm9rZSB04bqhaSBydW50aW1lLiBOYXZpZ2F0ZVRvIGtow7RuZyBj"
    "w7MgdGhhbQogICAgc+G7kSBraeG7g3UgW291dF0gQlNUUiBj4bqnbiBtYXJzaGFsIMSR4bq3YyBi"
    "aeG7h3QgKHhlbSBsw70gZG8gZ2VuY2FjaGUgxJHGsOG7o2MKICAgIGTDuW5nIGNobyBHZXRIaWVy"
    "YXJjaHkvR2V0UGFnZUNvbnRlbnQg4bufIMSR4bqndSBtb2R1bGUpIG7Dqm4gaG/DoG4gdG/DoG4g"
    "cGjDuQogICAgaOG7o3AgZMO5bmcgZHluYW1pYyBkaXNwYXRjaCB0aHXhuqduIGNobyByacOqbmcg"
    "aMOgbSBt4bufIHRyYW5nIG7DoHkuIiIiCiAgICBpZiBub3QgT05FTk9URV9BVkFJTEFCTEU6CiAg"
    "ICAgICAgcmV0dXJuIEZhbHNlLCAicHl3aW4zMiBjaMawYSDEkcaw4bujYyBjw6BpIgogICAgcHl0"
    "aG9uY29tLkNvSW5pdGlhbGl6ZSgpCiAgICB0cnk6CiAgICAgICAgIyB2LWZpeCAoVuG6pW4gxJHh"
    "u4E6IGPhuqMgTmF2aWdhdGVUbyBs4bqrbiBHZXRIeXBlcmxpbmtUb09iamVjdCDEkeG7gXUgZmFp"
    "bAogICAgICAgICMgduG7m2kgQXR0cmlidXRlRXJyb3IgcXVhIGR5bmFtaWMuRGlzcGF0Y2gsIFRS"
    "T05HIEtISSBpbmRleC9zZWFyY2gKICAgICAgICAjIE9uZU5vdGUgLS0gduG7kW4gZ+G7jWkgR2V0"
    "SGllcmFyY2h5KCkvR2V0UGFnZUNvbnRlbnQoKSBxdWEKICAgICAgICAjIF9lbnN1cmVfb25lbm90"
    "ZV9hcHAoKSAoZ2VuY2FjaGUpIC0tIHbhuqtuIGNo4bqheSB04buRdCB0csOqbiBDw5lORyBtw6F5"
    "CiAgICAgICAgIyBuw6B5KTogZHluYW1pYy5EaXNwYXRjaCB0cmEgbWV0aG9kIHRoZW8gVMOKTiB0"
    "4bqhaSBydW50aW1lIHF1YQogICAgICAgICMgSURpc3BhdGNoOjpHZXRJRHNPZk5hbWVzIHRyw6pu"
    "IGNow61uaCBvYmplY3QgQ09NIHRy4bqjIHbhu4E7IGdlbmNhY2hlCiAgICAgICAgIyB0aMOsIGTD"
    "uW5nIERJU1BJRCDEkcOjICLEkcOzbmcgY+G7qW5nIiBz4bq1biBsw7pjIGNvZGUtZ2VuIHThu6sg"
    "dHlwZSBsaWJyYXJ5LgogICAgICAgICMgTuG6v3UgYuG6o24gT25lTm90ZSDEkWFuZyBjaOG6oXkg"
    "Y8OzIHR5cGUgbGlicmFyeSDEkcSDbmcga8O9IGtow7RuZyBraOG7m3AKICAgICAgICAjIGhvw6Bu"
    "IHRvw6BuIHbhu5tpIHZp4buHYyB0cmEgY+G7qXUgbGF0ZS1ib3VuZCB0aGVvIHTDqm4gKGtow6Eg"
    "cGjhu5UgYmnhur9uIOG7nwogICAgICAgICMgY8OhYyBi4bqjbiBPbmVOb3RlIG3hu5tpL2Phuq1w"
    "IG5o4bqtdCBxdWEgV2luZG93cyBVcGRhdGUpIG5oxrBuZyB0eXBlCiAgICAgICAgIyBsaWJyYXJ5"
    "IGfhu5FjIChkw7luZyDEkeG7gyBidWlsZCBnZW5jYWNoZSkgduG6q24gxJHDum5nLCB0aMOsIGdl"
    "bmNhY2hlIGPDswogICAgICAgICMgdGjhu4MgZ+G7jWkgxJHGsOG7o2MgTkjhu65ORyBNRVRIT0Qg"
    "TcOAIGR5bmFtaWMgS0jDlE5HIHRyYSByYSBu4buVaSAtLSDEkcO6bmcKICAgICAgICAjIHTDrG5o"
    "IGh14buRbmcg4bufIMSRw6J5OiBHZXRIaWVyYXJjaHkvR2V0UGFnZUNvbnRlbnQgKHF1YSBnZW5j"
    "YWNoZSkKICAgICAgICAjIGNo4bqheSB04buRdCwgY8OybiBOYXZpZ2F0ZVRvL0dldEh5cGVybGlu"
    "a1RvT2JqZWN0IChxdWEgZHluYW1pYykgdGjDrAogICAgICAgICMga2jDtG5nLiBUaOG7rSBnZW5j"
    "YWNoZSBUUsav4buaQyAodMOhaSBkw7luZyDEkcO6bmcgY8ahIGNo4bq/IMSRw6MgY2jhu6luZyBt"
    "aW5oCiAgICAgICAgIyBob+G6oXQgxJHhu5luZyBjaG8gdmnhu4djIGluZGV4KSwgQ0jhu4ggcsah"
    "aSB24buBIGR5bmFtaWMuRGlzcGF0Y2ggbuG6v3UKICAgICAgICAjIGdlbmNhY2hlIHThu7EgbsOz"
    "IGPFqW5nIGZhaWwgaOG6s24gKE9uZU5vdGUgdGjhuq10IHPhu7Ega2jDtG5nIGPDswogICAgICAg"
    "ICMgYXV0b21hdGlvbiBzdXJmYWNlIG7DoG8ga2jhuqMgZOG7pW5nKS4KICAgICAgICB0cnk6CiAg"
    "ICAgICAgICAgIG9uZW5vdGUgPSBfZW5zdXJlX29uZW5vdGVfYXBwKCkKICAgICAgICBleGNlcHQg"
    "RXhjZXB0aW9uIGFzIF9lX2dlbmNhY2hlOgogICAgICAgICAgICBwcmludChmIltPbmVOb3RlXSBn"
    "ZW5jYWNoZSBkaXNwYXRjaCBmYWlsZWQgKHtfZV9nZW5jYWNoZSFyfSksICIKICAgICAgICAgICAg"
    "ICAgICAgZiJmYWxsaW5nIGJhY2sgdG8gZHluYW1pYyBkaXNwYXRjaCBmb3Igb3Blbi4uLiIpCiAg"
    "ICAgICAgICAgIGZyb20gd2luMzJjb20uY2xpZW50IGltcG9ydCBkeW5hbWljIGFzIF93aW4zMl9k"
    "eW5hbWljCiAgICAgICAgICAgIG9uZW5vdGUgPSBfd2luMzJfZHluYW1pYy5EaXNwYXRjaCgiT25l"
    "Tm90ZS5BcHBsaWNhdGlvbiIpCiAgICAgICAgIyB2LWZpeCAoduG6q24gY8OybiAiT25lTm90ZS5B"
    "cHBsaWNhdGlvbi5OYXZpZ2F0ZVRvIiBkw7kgxJHDoyDEkeG7lWkgc2FuZwogICAgICAgICMgZHlu"
    "YW1pYy5EaXNwYXRjaCB0aHXhuqduKTogZHluYW1pYy5EaXNwYXRjaCBW4bqqTiBsw6AgbGF0ZS1i"
    "b3VuZCAtLQogICAgICAgICMgbsOzIHbhuqtuIHRyYSBtZXRob2QgdGhlbyBUw4pOIHF1YSBHZXRJ"
    "RHNPZk5hbWVzIGzDumMgZ+G7jWksIHkgaOG7h3QKICAgICAgICAjIERpc3BhdGNoKCkgdHLGsOG7"
    "m2MgxJHDsyAtLSBuw6puIG7hur91IGLhuqNuIHRow6JuIG9iamVjdCBDT00gdHLhuqMgduG7gSBs"
    "w7pjCiAgICAgICAgIyDEkcOzIENIxq9BIHPhurVuIHPDoG5nLCBs4buXaSBnaeG7kW5nIGjhu4d0"
    "IHbhuqtuIHjhuqN5IHJhLCBraMO0bmcgbGnDqm4gcXVhbiBnw6wKICAgICAgICAjIMSR4bq/biB2"
    "aeG7h2MgY2jhu41uIGVhcmx5L2xhdGUtYm91bmQgbuG7r2EuIE5ndXnDqm4gbmjDom4gdGjhuq10"
    "IHPhu7Egbmhp4buBdQogICAgICAgICMga2jhuqMgbsSDbmcgbMOgIFJBQ0UgbMO6YyBraOG7n2kg"
    "xJHhu5luZzogbuG6v3UgT25lTm90ZSBDSMavQSDEkWFuZyBjaOG6oXkgc+G6tW4sCiAgICAgICAg"
    "IyBEaXNwYXRjaCgiT25lTm90ZS5BcHBsaWNhdGlvbiIpIHPhur0gdOG7sSBraOG7n2kgY2jhuqF5"
    "IDEgdGnhur9uIHRyw6xuaAogICAgICAgICMgT25lTm90ZSBN4buaSSBwaMOtYSBzYXUgLS0gbmjG"
    "sG5nIG7DsyB0cuG6oyB24buBIHBvaW50ZXIgQ09NIGfhuqduIG5oxrAKICAgICAgICAjIG5nYXkg"
    "bOG6rXAgdOG7qWMsIFRSxq/hu5pDIEtISSBPbmVOb3RlLmV4ZSBraOG7n2kgxJHhu5luZyB4b25n"
    "IHbDoCDEkcSDbmcga8O9CiAgICAgICAgIyDEkeG6p3kgxJHhu6cgYuG7gSBt4bq3dCBJRGlzcGF0"
    "Y2ggY+G7p2EgbsOzLCBuw6puIEdldElEc09mTmFtZXMoIk5hdmlnYXRlVG8iKQogICAgICAgICMg"
    "bmdheSBsw7pjIMSRw7MgdHLhuqMgduG7gSBy4buXbmcgKGTDuSB2aeG7h2MgaW5kZXgvR2V0SGll"
    "cmFyY2h5IHbhuqtuIGNo4bqheQogICAgICAgICMgdOG7kXQgdsOsIG7DsyBsdcO0biB44bqjeSBy"
    "YSBTQVUga2hpIE9uZU5vdGUgxJHDoyBr4buLcCBraOG7n2kgxJHhu5luZyB4b25nIHThu6sKICAg"
    "ICAgICAjIDEgbOG6p24gZ+G7jWkgdHLGsOG7m2MgxJHDsykuIFJldHJ5IHbDoGkgbOG6p24gduG7"
    "m2kga2hv4bqjbmcgbmdo4buJIG5n4bqvbiDEkeG7gyBjaG8KICAgICAgICAjIE9uZU5vdGUgxJHh"
    "u6cgdGjhu51pIGdpYW4ga2jhu59pIMSR4buZbmcgeG9uZyAtLSBjaOG7iSByZXRyeSDEkcO6bmcK"
    "ICAgICAgICAjIEF0dHJpYnV0ZUVycm9yIChiaeG7g3UgaGnhu4duIGPhu6dhIGzhu5dpIG7DoHkp"
    "LCBraMO0bmcgcmV0cnkgbOG7l2kga2jDoWMuCiAgICAgICAgX2xhc3RfZXJyID0gTm9uZQogICAg"
    "ICAgIGZvciBfYXR0ZW1wdCBpbiByYW5nZSg2KToKICAgICAgICAgICAgdHJ5OgogICAgICAgICAg"
    "ICAgICAgb25lbm90ZS5OYXZpZ2F0ZVRvKHBhZ2VfaWQpCiAgICAgICAgICAgICAgICByZXR1cm4g"
    "VHJ1ZSwgTm9uZQogICAgICAgICAgICBleGNlcHQgQXR0cmlidXRlRXJyb3IgYXMgZToKICAgICAg"
    "ICAgICAgICAgIF9sYXN0X2VyciA9IGUKICAgICAgICAgICAgICAgIGlmIF9hdHRlbXB0ID09IDA6"
    "CiAgICAgICAgICAgICAgICAgICAgIyB2LW5ldzogY2jhu4kgaW4gMSBs4bqnbiAoa2jDtG5nIGzh"
    "urdwIGzhuqFpIG3hu5dpIHJldHJ5KSAtLQogICAgICAgICAgICAgICAgICAgICMgbGnhu4d0IGvD"
    "qiBjw6FjIGF0dHJpYnV0ZSBjw7MgY2jhu6lhICJOYXZpZ2F0ZSIvIkh5cGVybGluayIKICAgICAg"
    "ICAgICAgICAgICAgICAjIG3DoCBvYmplY3QgQ09NIG7DoHkgVEjhu7BDIFPhu7AgbOG7mSByYSBx"
    "dWEgZGlyKCksIMSR4buDIGJp4bq/dAogICAgICAgICAgICAgICAgICAgICMgY2jhuq9jIMSRw6J5"
    "IGzDoCBkbyBPbmVOb3RlIHRo4bqtdCBz4buxIGtow7RuZyBleHBvc2UgMgogICAgICAgICAgICAg"
    "ICAgICAgICMgbWV0aG9kIMSRw7MgKGRhbmggc8OhY2ggcuG7l25nL2tow7RuZyBjw7MpIGhheSBj"
    "aOG7iSBsw6AgcmFjZQogICAgICAgICAgICAgICAgICAgICMgbMO6YyBraOG7n2kgxJHhu5luZyAo"
    "c+G6vSB04buxIGjhur90IHNhdSB2w6BpIGzhuqduIHJldHJ5KS4KICAgICAgICAgICAgICAgICAg"
    "ICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIF9yZWxldmFudCA9IFthIGZvciBhIGluIGRp"
    "cihvbmVub3RlKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgJ05hdmln"
    "YXRlJyBpbiBhIG9yICdIeXBlcmxpbmsnIGluIGFdCiAgICAgICAgICAgICAgICAgICAgICAgIHBy"
    "aW50KGYiW09uZU5vdGVdW2RlYnVnXSBkaXNwYXRjaCB0eXBlPXt0eXBlKG9uZW5vdGUpIXJ9LCAi"
    "CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGYiTmF2aWdhdGUvSHlwZXJsaW5rIGF0dHJz"
    "IGZvdW5kOiB7X3JlbGV2YW50IG9yICcobm9uZSknfSIpCiAgICAgICAgICAgICAgICAgICAgZXhj"
    "ZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgcGFzcwogICAgICAgICAgICAg"
    "ICAgdGltZS5zbGVlcCgwLjcpCiAgICAgICAgICAgICAgICBjb250aW51ZQogICAgICAgICMgdi1m"
    "aXg6IG7hur91IE5hdmlnYXRlVG8gduG6q24ga2jDtG5nIHJlc29sdmUgxJHGsOG7o2Mgc2F1IGPD"
    "oWMgbOG6p24gcmV0cnkKICAgICAgICAjICh04bupYyBraMO0bmcgcGjhuqNpIGRvIHRpbWluZyBr"
    "aOG7n2kgxJHhu5luZywgbcOgIGRvIG1ldGhvZCBuw6B5IHRo4bqtdCBz4buxCiAgICAgICAgIyBr"
    "aMO0bmcgbOG7mSByYSBxdWEgR2V0SURzT2ZOYW1lcyB0csOqbiBi4bqjbiBPbmVOb3RlIGPDoGkg"
    "dHLDqm4gbcOheSBuw6B5CiAgICAgICAgIyAtLSB2w60gZOG7pSBPbmVOb3RlIGZvciBXaW5kb3dz"
    "IDEwLzExIGtp4buDdSBTdG9yZSB0aGF5IHbDrCBPbmVOb3RlCiAgICAgICAgIyAyMDE2IGRlc2t0"
    "b3AgY+G7lSDEkWnhu4NuKSwgdGjhu60gxJHGsOG7nW5nIEhPw4BOIFRPw4BOIEtIw4FDIGtow7Ru"
    "ZyBxdWEKICAgICAgICAjIE5hdmlnYXRlVG86IEdldEh5cGVybGlua1RvT2JqZWN0KCkgdHLhuqMg"
    "duG7gSAxIGxpbmsgZOG6oW5nCiAgICAgICAgIyAib25lbm90ZTouLi4iLCBy4buTaSBt4bufIGxp"
    "bmsgxJHDsyBxdWEgdHLDrG5oIHjhu60gbMO9IFVSSSAib25lbm90ZToiCiAgICAgICAgIyDEkcOj"
    "IMSRxINuZyBrw70gc+G6tW4gdHJvbmcgV2luZG93cyAob3Muc3RhcnRmaWxlKSAtLSB2aeG7h2Mg"
    "xJBJ4buAVQogICAgICAgICMgSMav4buaTkcgVEjhuqxUIGzDumMgbsOgeSBkbyBjaMOtbmggT25l"
    "Tm90ZS9XaW5kb3dzIHjhu60gbMO9IHF1YSBVUkkKICAgICAgICAjIHByb3RvY29sIGhhbmRsZXIs"
    "IGtow7RuZyBjw7JuIHBo4bulIHRodeG7mWMgZ8OsIHbDoG8gQ09NIGF1dG9tYXRpb24KICAgICAg"
    "ICAjIGPhu6dhIFB5dGhvbiBu4buvYS4KICAgICAgICB0cnk6CiAgICAgICAgICAgIGxpbmsgPSBv"
    "bmVub3RlLkdldEh5cGVybGlua1RvT2JqZWN0KHBhZ2VfaWQsICIiKQogICAgICAgICAgICBpZiBs"
    "aW5rOgogICAgICAgICAgICAgICAgb3Muc3RhcnRmaWxlKGxpbmspCiAgICAgICAgICAgICAgICBy"
    "ZXR1cm4gVHJ1ZSwgTm9uZQogICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZTI6CiAgICAgICAg"
    "ICAgIF9sYXN0X2VyciA9IGUyCiAgICAgICAgcmV0dXJuIEZhbHNlLCBzdHIoX2xhc3RfZXJyKQog"
    "ICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgIHJldHVybiBGYWxzZSwgc3RyKGUpCiAg"
    "ICBmaW5hbGx5OgogICAgICAgIHB5dGhvbmNvbS5Db1VuaW5pdGlhbGl6ZSgpCgoKZGVmIGdldF9p"
    "bmRleGVkX3BhZ2VfY291bnQoKToKICAgICIiIlF1aWNrIGNvdW50IGZvciB0aGUgVUkgKGUuZy4g"
    "c2hvdyAnTiBwYWdlIGluZGV4ZWQnIG5leHQgdG8gdGhlCiAgICBVcGRhdGUgT25lTm90ZSBidXR0"
    "b24pLiIiIgogICAgaWYgbm90IG9zLnBhdGguZXhpc3RzKE9ORU5PVEVfREJfRklMRSk6CiAgICAg"
    "ICAgcmV0dXJuIDAKICAgIHRyeToKICAgICAgICBjb25uID0gc3FsaXRlMy5jb25uZWN0KE9ORU5P"
    "VEVfREJfRklMRSkKICAgICAgICBjID0gY29ubi5jdXJzb3IoKQogICAgICAgIGMuZXhlY3V0ZSgi"
    "U0VMRUNUIENPVU5UKCopIEZST00gb25lbm90ZV9zdG9yZSIpCiAgICAgICAgbiA9IGMuZmV0Y2hv"
    "bmUoKVswXQogICAgICAgIGNvbm4uY2xvc2UoKQogICAgICAgIHJldHVybiBuCiAgICBleGNlcHQg"
    "RXhjZXB0aW9uOgogICAgICAgIHJldHVybiAwCg=="
)


def _load_embedded_module(name, src_b64):
    """Recreate a standalone module from its embedded base64 source, in a
    way that's indistinguishable to the rest of this file from a normal
    `import name`. __file__ is set to THIS script's own path (not a fake
    one) so the embedded module's own `_BASE_DIR = os.path.dirname(...
    __file__)` logic still resolves to the same folder as before (e.g.
    E:/mySearch_ai/), keeping search_outlook.db / search_onenote.db in
    exactly the same place they were when these were separate files."""
    mod = _mod_types.ModuleType(name)
    mod.__file__ = os.path.abspath(__file__)
    src = _b64.b64decode(src_b64).decode("utf-8")
    exec(compile(src, f"<embedded:{name}.py>", "exec"), mod.__dict__)
    return mod


# v-outlook: optional Outlook mail search — lives in its own db file
# (search_outlook.db), same reasoning as HISTORY_DB_FILE above. If pywin32/
# Outlook isn't available, the app just runs without the feature (no "msg"
# results ever appear, ext filter button still shows but finds nothing)
# instead of crashing at startup.
try:
    outlook_search = _load_embedded_module("outlook_search", _OUTLOOK_SEARCH_SRC_B64)
    OUTLOOK_SEARCH_AVAILABLE = True
except Exception as e:
    OUTLOOK_SEARCH_AVAILABLE = False
    print(f"[Outlook] embedded outlook_search module failed to load — Outlook mail search disabled: {e!r}")

# v-onenote: same reasoning as v-outlook above — own db file
# (search_onenote.db), graceful no-op if OneNote isn't available.
try:
    onenote_search = _load_embedded_module("onenote_search", _ONENOTE_SEARCH_SRC_B64)
    ONENOTE_SEARCH_AVAILABLE = True
except Exception as e:
    ONENOTE_SEARCH_AVAILABLE = False
    print(f"[OneNote] embedded onenote_search module failed to load — OneNote search disabled: {e!r}")

def _make_outlook_pseudo_path(entry_id, store_id, subject, folder_path=""):
    """Outlook mail has no real filesystem path. We synthesize one so a
    mail 'row' can flow through the exact same File Content pipeline as a
    real file (extension-based Ext filter, icon lookup, priority sort,
    etc.) with zero special-casing needed in most of that code. Format is
    deliberately NOT a real-looking path (no '://', no drive letter) so it
    can never be mistaken for — or accidentally matched against — an actual
    file on disk anywhere in the existing path-handling code.
    Always ends in '.msg' so os.path.splitext() (used throughout this file
    for ext filtering / icon lookup / priority sort) naturally treats it as
    a message file with zero extra code.

    v-fix (Vấn đề: "Nguồn:" trong AI Chat luôn hiện pseudo-path xấu, không
    bao giờ kịp hiện "Outlook > folder > subject.msg" đẹp): folder_path
    now embedded directly in the pseudo path itself, the same trick
    _make_onenote_pseudo_path already used for section_path/title -- so
    building the nice display string is pure string-splitting, instant,
    and doesn't depend on self._outlook_meta having been populated yet by
    the (sometimes very slow) _merge_mail_notes_async background scan."""
    safe_subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', (subject or "").strip())[:150]
    safe_subject = safe_subject or "(no subject)"
    safe_folder = re.sub(r'[<>:"|?*\x00-\x1f]', '_', (folder_path or "").strip())[:150]
    return f"OUTLOOK::{entry_id}::{store_id}::{safe_folder}::{safe_subject}.msg"

def _is_outlook_pseudo_path(path):
    return isinstance(path, str) and path.startswith("OUTLOOK::")

def _parse_outlook_pseudo_path(path):
    """Returns (entry_id, store_id) from a pseudo path built above."""
    try:
        _, entry_id, store_id, _rest = path.split("::", 3)
        return entry_id, store_id
    except Exception:
        return None, None


def _make_onenote_pseudo_path(page_id, section_str, title):
    """Same trick as _make_outlook_pseudo_path — OneNote pages have no
    real filesystem path either, so a synthetic one lets a page 'row' flow
    through the same File Content pipeline (ext filter/icon/priority
    sort) with no special-casing. Always ends in '.one' so
    os.path.splitext() naturally treats it as a OneNote file."""
    safe_title = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', (title or "").strip())[:150]
    safe_title = safe_title or "(untitled)"
    # v-fix (Vấn đề: breadcrumb hiện "SIMPACK _ SIMPACK _ 弾性体" thay vì
    # "SIMPACK > SIMPACK > 弾性体"): section_str's natural separator between
    # notebook/section levels IS '>' (see onenote_search's own section_path
    # building), but the old sanitizer here stripped '<'/'>' along with the
    # genuinely Windows-illegal characters, silently mangling every multi-
    # level section path into underscores. This is never written to an
    # actual filesystem path (it only lives inside this synthetic '::'-
    # delimited token), so '<'/'>' don't need to be sanitized here at all
    # -- only the characters that would actually break something else
    # downstream (e.g. embedding another '::' would break the split(...)
    # parsing above).
    safe_section = re.sub(r'[:"|?*\x00-\x1f]', '_', (section_str or "").strip())[:150]
    return f"ONENOTE::{page_id}::{safe_section}::{safe_title}.one"


def _is_onenote_pseudo_path(path):
    return isinstance(path, str) and path.startswith("ONENOTE::")


def _parse_onenote_pseudo_path(path):
    """Returns page_id from a pseudo path built above."""
    try:
        _, page_id, _section, _rest = path.split("::", 3)
        return page_id
    except Exception:
        return None


# v-new (theo yêu cầu: cùng TÊN file -- kể cả khác đuôi PDF/PPT/DOC, hay
# khác số version _v1.1/_v1.2/_v1.3... -- chỉ nên chiếm 1 slot nguồn AI
# Chat, vì nội dung thường giống hệt hoặc gần giống nhau): trước đây
# _add_ctx() (trong _online_chat_worker) chỉ coi 2 file là trùng nếu
# basename+extension GIỐNG HỆT nhau -- "...v1.3.pdf" và "...v1.2.pdf" vẫn
# bị tính là 2 nguồn khác nhau, và "...pdf"/"...pptx" cùng tên gốc cũng
# vậy. Giờ chuẩn hoá về "tên gốc, bỏ đuôi file + bỏ hậu tố version" làm
# khoá gộp nhóm, rồi trong mỗi nhóm chỉ giữ lại bản có version CAO NHẤT
# (ví dụ v1.3 thay vì v1.1) -- xem chỗ dùng trong _online_chat_worker.
_CTX_VERSION_SUFFIX_RE = re.compile(r'[\s_\-]*v(\d+(?:\.\d+)*)\s*$', re.IGNORECASE)


def _ctx_dedup_key(path):
    """Normalized grouping key for a real (non Outlook/OneNote) file path:
    lowercased filename, extension stripped, trailing '_v1.3'/' V2'/etc
    version suffix stripped. Two files that only differ by extension
    (.pdf vs .pptx) or by version number map to the SAME key."""
    name = os.path.splitext(os.path.basename(path))[0].lower()
    name = _CTX_VERSION_SUFFIX_RE.sub('', name).strip()
    return name


def _ctx_version_tuple(path):
    """Parses a trailing '_v1.3' style suffix (if any) into a comparable
    tuple, e.g. 'v1.3' -> (1, 3). Files with no version suffix sort as
    the lowest possible version (0,) so any versioned file outranks them."""
    name = os.path.splitext(os.path.basename(path))[0]
    m = _CTX_VERSION_SUFFIX_RE.search(name)
    if not m:
        return (0,)
    try:
        return tuple(int(x) for x in m.group(1).split('.'))
    except Exception:
        return (0,)


# v9.0: AI libraries (torch/transformers/sentence-transformers/einops) are
# now bundled directly into the exe at build time -- no more runtime
# 'ai_libs' folder, no more sys.path surgery, no more heal-attempt
# counters. See the top-of-file note near _load_semantic_model for why.

# v5.8: Stopword filter -- fixes irrelevant OR-fallback matches.
# Without this, a query like "simpack realtime relevant to HILS" splits into
# keywords ["simpack","realtime","relevant","to","hils"]. The 2-letter word
# "to" then substring-matches almost anything ending in ".toc"
# (Analysis-00.toc, EXE-00.toc, PYZ-00.toc, ...) via `"to" in filename`,
# flooding results with junk that has nothing to do with the real query.
# Filler words (English + Vietnamese) carry no filename-matching signal and
# are stripped before AND/OR keyword matching. If stripping would remove
# every keyword (rare edge case: the whole query was filler words), the
# original keyword list is used instead so the search never returns empty.
STOPWORDS = {
    # English fillers
    "a", "an", "the", "to", "of", "in", "on", "for", "and", "or", "is", "are",
    "was", "were", "be", "been", "with", "at", "by", "from", "as", "that",
    "this", "it", "its", "into", "about", "regarding", "re", "vs", "via",
    "relevant", "related", "concerning",
    # Vietnamese fillers
    "và", "của", "cho", "là", "các", "những", "với", "về", "đến", "tới",
    "cái", "này", "đó", "cùng", "hoặc", "trong", "trên", "dưới", "theo",
    "liên", "quan",
}


def _strip_stopwords(keywords):
    """Remove filler words from a keyword list; fall back to the original
    list if that would leave nothing (so a query never becomes empty)."""
    filtered = [k for k in keywords if k not in STOPWORDS]
    return filtered if filtered else keywords


# v9.13.10 fix: plain `.split()` only breaks on whitespace, which is fine
# for space-separated languages (English/Vietnamese/...) but Japanese/
# Chinese don't put spaces between words at all -- a whole sentence like
# "IRMの周波数はどのように決まっている？" became ONE single giant "keyword"
# under `.split()`, which BM25/filename search then required to match
# VERBATIM somewhere -- essentially never happens for a natural-language
# question, so it silently returned zero results (not an error, just
# nothing found) for any CJK query longer than a short single term/phrase.
# This doesn't attempt real Japanese word segmentation (that needs a
# dictionary-based tokenizer like MeCab/Sudachi/fugashi, a much bigger
# addition) -- it just also splits on common CJK/Latin punctuation
# (、。！？「」『』（） etc.), so a long sentence breaks into smaller phrase
# fragments instead of one monolithic blob that can never match. Short
# queries (a single word/term, CJK or not) are unaffected either way.
_QUERY_SPLIT_RE = re.compile(r'[\s、。！？「」『』（）\(\)\.\,\!\?；：:;]+')

def _split_query_tokens(text):
    """CJK-aware replacement for `text.split()` when extracting keyword
    tokens from a search query."""
    return [t for t in _QUERY_SPLIT_RE.split(text) if t]


def _cjk_glued_or_query(token, limit_splits=12):
    """v-new (Vấn đề: "ライセンスサービス" [dính liền, không dấu cách] ra 0
    kết quả cả trong File Content lẫn Outlook/OneNote, nhưng
    "ライセンス　サービス" [có dấu cách giữa 2 từ] lại ra kết quả bình
    thường — cùng 1 ý, chỉ khác cách gõ): tokenizer trigram MATCH cho 1
    token DÍNH LIỀN đòi hỏi đúng chuỗi liên tục đó tồn tại y hệt trong nội
    dung; nếu văn bản gốc chỉ có "ライセンス" và "サービス" xuất hiện GẦN
    nhau chứ không dính sát (cách khoảng trắng, xuống dòng, chữ nối như
    "の", thẻ HTML khi index email...), match thất bại dù đúng là cùng 1
    cụm — người dùng gõ liền tay theo thói quen tiếng Nhật (vốn không có
    khoảng trắng giữa các từ khi viết). App không có bộ tách từ tiếng Nhật
    thật (MeCab/Janome...), nên đây là fallback nhẹ dùng chung cho cả File
    Content (content_index) lẫn Outlook/OneNote (outlook_content/
    onenote_content) — mọi nơi đều dùng FTS5 nên cùng cú pháp MATCH: thử
    TẤT CẢ các điểm cắt hợp lệ của token đó thành 2 nửa (như thể user đã gõ
    khoảng trắng ở đó), rồi OR toàn bộ các cặp AND lại thành 1 câu MATCH
    duy nhất — ví dụ "ライセンスサービス" (8 ký tự) sinh ra:
    ("ライ" AND "センスサービス") OR ("ライセ" AND "ンスサービス") OR ...
    OR ("ライセンス" AND "サービス")  ← đúng ranh giới thật, sẽ khớp ở đây
    OR ...
    Trả về None nếu token quá ngắn (không đủ chỗ cho 2 từ >=3 ký tự) hoặc
    chứa ký tự không an toàn cho FTS5 (dấu ngoặc kép, toán tử...). Số điểm
    cắt bị giới hạn (limit_splits) để câu MATCH không phình quá to với 1
    token rất dài. Chỉ nên gọi khi query CHÍNH XÁC (token dính liền) đã ra
    v-fix (Vấn đề: "鉄道不整" [4 ký tự, ghép 2 từ 2-ký-tự "鉄道"+"不整"] vẫn
    ra 0 kết quả dù "鉄道" và "不整" riêng lẻ đều có kết quả): ngưỡng cũ
    (n<6, mỗi nửa tối thiểu 3 ký tự) chỉ nhắm tới cụm dài như ví dụ
    "ライセンスサービス" (8-9 ký tự) -- không bao giờ áp dụng được cho rất
    nhiều thuật ngữ ghép 4 ký tự kiểu 2+2 rất phổ biến trong tiếng Nhật
    (đặc biệt thuật ngữ kỹ thuật hán-nhật). Hạ ngưỡng xuống n>=4 và cho
    phép mỗi nửa tối thiểu 2 ký tự.
    Trả về None nếu token quá ngắn (không đủ chỗ cho 2 từ >=2 ký tự) hoặc
    chứa ký tự không an toàn cho FTS5 (dấu ngoặc kép, toán tử...). Số điểm
    cắt bị giới hạn (limit_splits) để câu MATCH không phình quá to với 1
    token rất dài. Chỉ nên gọi khi query CHÍNH XÁC (token dính liền) đã ra
    0 kết quả — đây là lượt thử nới rộng thêm, không thay thế lượt đầu."""
    n = len(token)
    if n < 4:
        return None
    if any(ch in '"\\' for ch in token):
        return None   # can't safely embed in a quoted FTS5 term
    cuts = list(range(2, n - 1))
    if len(cuts) > limit_splits:
        step = len(cuts) / limit_splits
        cuts = sorted({cuts[int(i * step)] for i in range(limit_splits)} | {n // 2})
    if not cuts:
        return None
    return " OR ".join(f'("{token[:c]}" AND "{token[c:]}")' for c in cuts)


def _is_word_char(ch):
    """ASCII letters only count as "inside a word" for _kw_matches's
    whole-word boundary check. Plain .isalpha() is ALSO True for CJK
    characters (kanji/hiragana/katakana), which have no real word-boundary
    concept the way English does -- Japanese/Chinese sentences run with no
    spaces at all. Treating a neighboring CJK character as "still inside
    the word" made "Whole word" effectively never match any Japanese/
    Chinese search term unless it sat at the literal start/end of the
    whole string -- e.g. searching "ライセンス開放など仕様について" found
    0 results even inside a folder literally named
    "...のライセンス開放など仕様について", because the "の" right before
    the match counted as "still part of the word" and blocked it."""
    return ch.isascii() and ch.isalpha()


def _kw_matches(kw, text_lower, whole_word=False):
    """Does `kw` match somewhere inside `text_lower`?
    whole_word=False (default): plain substring match -- original behavior,
        e.g. "adas" matches inside "readasync.xml" (buried mid-word).
    whole_word=True: `kw` only counts as a match when both the character
        immediately before and immediately after it are NOT ascii letters
        (i.e. digit/underscore/hyphen/dot/space/parenthesis/start-of-string/
        end-of-string/any CJK character all count as a boundary). This
        keeps matches like "ADAS_systems.pdf" or "VDIM_0ADAS_ESP" (ADAS
        sits next to a digit/underscore/dot) while rejecting
        "readasync.xml" or "ReadAStringExample.mlx" (adas sits buried
        between two ASCII letters) -- and, since a CJK neighbor now also
        counts as a boundary, it works correctly for Japanese/Chinese
        search terms too, not just ASCII ones."""
    if not whole_word:
        return kw in text_lower
    start = 0
    n = len(text_lower)
    klen = len(kw)
    while True:
        idx = text_lower.find(kw, start)
        if idx == -1:
            return False
        before_ok = (idx == 0) or (not _is_word_char(text_lower[idx - 1]))
        after_idx = idx + klen
        after_ok = (after_idx >= n) or (not _is_word_char(text_lower[after_idx]))
        if before_ok and after_ok:
            return True
        start = idx + 1  # keep scanning -- an earlier occurrence might fail while a later one succeeds


def _kw_matches_with_glued_fallback(kw, text_lower, whole_word=False):
    """v-fix (Vấn đề: "鉄道不整" [ghép liền "鉄道"+"不整", không dấu cách
    -- thói quen gõ tiếng Nhật bình thường] ra 0 kết quả ở tab File Name/
    Folder Name dù "鉄道" và "不整" riêng lẻ đều có kết quả): _kw_matches
    thường đòi hỏi ĐÚNG chuỗi liên tục "kw" xuất hiện trong tên file/
    folder -- nếu tên thật sự chỉ chứa "鉄道" và "不整" ở 2 chỗ KHÔNG
    dính liền nhau (vd "...鉄道の不整状態..." hoặc 2 từ nằm cách xa nhau
    trong tên dài), match thất bại dù cả 2 từ rõ ràng đều có mặt. Đây là
    bản áp dụng lại Ý TƯỞNG của _cjk_glued_or_query (vốn chỉ dùng cho
    FTS5/File Content) sang so khớp chuỗi thuần Python cho File Name/
    Folder Name: nếu match nguyên khối thất bại, thử mọi điểm cắt hợp lệ
    tách kw thành 2 nửa, chấp nhận khớp nếu CẢ 2 nửa đều độc lập thoả
    _kw_matches trong cùng 1 tên. Chỉ áp dụng cho token dài >=4 ký tự và
    toàn bộ là ký tự non-ASCII (CJK) -- từ tiếng Anh/số đã có dấu cách
    thật nên không cần fallback này, và áp dụng nó cho từ ASCII ngắn sẽ
    chỉ tạo ra các match rác vô nghĩa."""
    if _kw_matches(kw, text_lower, whole_word):
        return True
    n = len(kw)
    if n < 4 or any(ch.isascii() for ch in kw):
        return False
    for c in range(2, n - 1):
        if _kw_matches(kw[:c], text_lower, whole_word) and _kw_matches(kw[c:], text_lower, whole_word):
            return True
    return False

SMALL_SIZE = "85x35" 
LARGE_SIZE = "390x35"
RESULT_SIZE = "1250x930"  # v1.3: +1/3 taller so the results panes have more room to grow down
HELP_SIZE = "630x630"
BG_COLOR = "#f4f5f7"
ENTRY_BG = "#ffffff"
TEXT_COLOR = "#1c1e21"
PLACE_COLOR = "#8a8d93"

# v-new: Help tab content in 3 languages (VI/EN/JA buttons next to the
# "Help" label -- see _build_help_content). Kept as static hand-written
# translations (not an AI call) so the Help tab stays instant and works
# offline even if the online AI models are unavailable/rate-limited.
HELP_CONTENT = {
    "EN": """
 [ Shortcuts / Commands ]
  - --update data               : Manually update the database (search_data.db) — all Tier 1-4
  - --update data tier 1        : Update Tier 1 only (office/pdf)
  - --update data tier 1,2      : Update Tier 1 and Tier 2 only (comma/space separated, multiple OK)
  - Update DB button             : Type "tier 1,2" (or just "1,2") into the Searchbox first,
                                   then click — updates only that Tier (empty box = all Tier 1-4)
  - ESC, Alt+F4, --exit, --quit  : Exit

 [ File / Folder Name Search ]
  - keyword               : Search file content / file name / folder name
  - space                 : Multiple keywords (AND search)

 [ Search Results ]
  - Double-click           : Open the file's location in Explorer
  - Second click            : Select/copy text
  - Right-click             : Open Folder, Copy File Name, Copy File Path

 [ History ]
  - Help tab (right side)   : Search History is always visible here
  - Right-click              : "Search again" to re-run a past search
  - Export Excel           : Save history to Excel
  - Double-click             : Select/copy text
  - AI Chat History (below)  : Defaults to the MOST RECENT search keyword automatically —
                              no need to click a row first. Keeps following the newest
                              keyword until you click a different row yourself

 [ Ramp Light ]
  - ● Red                    : search_data.db not created yet (run --update data)
  - ● Yellow                 : search_data.db exists but content not fully indexed yet
  - ● Yellow (blinking)      : Update DB is currently scanning/extracting content
  - ● Green                  : BM25 content ready — Search/Advanced usable (AI may still be building)
  - ● Blue                   : BM25 content AND AI embeddings both fully up to date

 [ File Content Search ]
  - keyword               : Text files, Outlook, OneNote, Office & PDF files

 [ AI Search Model ]
  - Model dropdown : Inside the AI search result window, next to the 🤖 AI Search button
  - Options: Jina-v3 / BGE-Gemma2
  - Each model uses its own data table — switching does not delete the other model's data
  - Run --update data once per model to build its index (can switch freely after that)

 [ AI Chat ]
  - Where               : Below the results, on the "File Content" tab only — appears
                          automatically once you click 🤖 AI Search
  - Auto-summary         : As soon as it opens, it automatically asks the AI to summarize
                          the best results found for your current query — no typing needed
  - Model                : Gemini Flash by default; auto-switches to GPT-OSS 120B if
                          Gemini hits a rate limit, so chat never just stops working
  - VI / EN / JP buttons  : Force the AI to always reply in that language, no matter what
                          language you type your question in. Switching language on an
                          existing conversation re-translates everything already shown
  - ◀ / ▶                : Page through older saved AI Chat sessions for the current keyword
  - New Chat icon         : Clears the conversation and starts a fresh session
  - Grounded answers      : The AI only answers from REAL excerpts of your own local files
                          (found via BM25/FTS5) — it says so if it can't find enough
                          information, and cites which file each claim came from ([1],[2]...)
                          with the full source list auto-appended below the answer
  - AI Chat History       : Every session is auto-saved and viewable later in the Help tab
                          (see History above) — read-only there, Copy only, no re-send
""",
    "VI": """
 [ Phím tắt / Lệnh ]
  - --update data               : Cập nhật thủ công database (search_data.db) — toàn bộ Tier 1-4
  - --update data tier 1        : Chỉ cập nhật Tier 1 (office/pdf)
  - --update data tier 1,2      : Chỉ cập nhật Tier 1 và Tier 2 (cách nhau bằng dấu phẩy/khoảng trắng, nhiều tier OK)
  - Nút Update DB                : Gõ "tier 1,2" (hoặc chỉ "1,2") vào ô Search trước,
                                   rồi bấm nút — chỉ cập nhật Tier đó (bỏ trống = tất cả Tier 1-4)
  - ESC, Alt+F4, --exit, --quit  : Thoát

 [ Tìm File / Folder theo tên ]
  - từ khoá               : Tìm theo nội dung file / tên file / tên folder
  - dấu cách              : Nhiều từ khoá (tìm kiểu AND)

 [ Kết quả tìm kiếm ]
  - Double-click            : Mở vị trí file trong Explorer
  - Click lần 2             : Chọn/copy text
  - Chuột phải              : Open Folder, Copy File Name, Copy File Path

 [ Lịch sử ]
  - Tab Help (bên phải)     : Search History luôn hiển thị ở đây
  - Chuột phải              : "Search again" để tìm lại một truy vấn cũ
  - Export Excel            : Lưu lịch sử ra Excel
  - Double-click            : Chọn/copy text
  - AI Chat History (dưới)   : Mặc định hiển thị theo từ khoá tìm kiếm MỚI NHẤT —
                              không cần bấm chọn dòng nào cả. Tự động bám theo từ khoá
                              mới nhất cho đến khi bạn tự bấm chọn dòng khác

 [ Đèn báo trạng thái ]
  - ● Đỏ                    : Chưa tạo search_data.db (chạy --update data)
  - ● Vàng                  : search_data.db đã có nhưng chưa index xong nội dung
  - ● Vàng (nhấp nháy)      : Update DB đang quét/trích xuất nội dung
  - ● Xanh lá               : Nội dung BM25 sẵn sàng — dùng được Search/Advanced (AI có thể vẫn đang build)
  - ● Xanh dương            : Nội dung BM25 VÀ AI embeddings đều đã cập nhật đầy đủ

 [ Tìm theo nội dung file ]
  - từ khoá               : File text, Outlook, OneNote, Office & PDF

 [ Model AI Search ]
  - Dropdown chọn Model : Trong cửa sổ kết quả AI Search, cạnh nút 🤖 AI Search
  - Lựa chọn: Jina-v3 / BGE-Gemma2
  - Mỗi model dùng bảng dữ liệu riêng — đổi model không xoá dữ liệu của model kia
  - Chạy --update data một lần cho mỗi model để build index (sau đó đổi qua lại thoải mái)

 [ AI Chat ]
  - Ở đâu                  : Bên dưới kết quả, chỉ ở tab "File Content" — tự động hiện ra
                             ngay khi bấm 🤖 AI Search
  - Tự tóm tắt              : Vừa mở ra là tự động hỏi AI tóm tắt các kết quả tốt nhất
                             cho từ khoá đang tìm — không cần gõ gì cả
  - Model                  : Mặc định Gemini Flash; tự chuyển sang GPT-OSS 120B nếu
                             Gemini bị giới hạn tần suất (rate limit) — chat không bao giờ
                             bị đứng vì lý do đó
  - Nút VI / EN / JP        : Bắt AI luôn trả lời bằng ngôn ngữ đã chọn, bất kể bạn gõ câu
                             hỏi bằng ngôn ngữ nào. Đổi ngôn ngữ khi đang chat sẽ dịch lại
                             toàn bộ nội dung đã hiển thị
  - ◀ / ▶                  : Xem lại các phiên AI Chat cũ đã lưu cho từ khoá hiện tại
  - Icon New Chat           : Xoá cuộc trò chuyện hiện tại, bắt đầu phiên mới
  - Trả lời có căn cứ       : AI chỉ trả lời dựa trên nội dung THẬT trích từ file trên máy
                             bạn (tìm bằng BM25/FTS5) — nếu không đủ thông tin sẽ nói rõ,
                             và luôn ghi rõ dựa vào file nào ([1],[2]...) kèm danh sách
                             nguồn tự động thêm bên dưới câu trả lời
  - AI Chat History         : Mỗi phiên chat được tự động lưu lại, xem sau trong tab Help
                             (xem mục Lịch sử ở trên) — chỉ xem/copy, không gửi lại được
""",
    "JA": """
 [ ショートカット／コマンド ]
  - --update data               : データベース(search_data.db)を手動更新 — Tier 1〜4 すべて
  - --update data tier 1        : Tier 1 のみ更新（office/pdf）
  - --update data tier 1,2      : Tier 1・2 のみ更新（カンマ/スペース区切り、複数指定可）
  - Update DB ボタン             : 検索欄に先に "tier 1,2"（または "1,2"）と入力してから
                                   クリック — その Tier のみ更新（空欄なら Tier 1〜4 すべて）
  - ESC, Alt+F4, --exit, --quit  : 終了

 [ ファイル／フォルダ名検索 ]
  - キーワード             : ファイル内容／ファイル名／フォルダ名で検索
  - スペース               : 複数キーワード（AND 検索）

 [ 検索結果 ]
  - ダブルクリック           : Explorer でファイルの場所を開く
  - 2回目のクリック          : テキストを選択／コピー
  - 右クリック               : Open Folder, Copy File Name, Copy File Path

 [ 履歴 ]
  - Help タブ（右側）        : Search History は常にここに表示
  - 右クリック               : "Search again" で過去の検索を再実行
  - Export Excel           : 履歴を Excel に保存
  - ダブルクリック            : テキストを選択／コピー
  - AI Chat History（下）    : 最新の検索キーワードを自動で表示 —
                              行をクリックする必要なし。ユーザーが別の行を
                              クリックするまで、常に最新キーワードを自動追従

 [ ランプ表示 ]
  - ● 赤                    : search_data.db がまだ未作成（--update data を実行）
  - ● 黄                    : search_data.db はあるが内容の索引化が未完了
  - ● 黄（点滅）             : Update DB が内容をスキャン／抽出中
  - ● 緑                    : BM25 の内容準備完了 — Search/Advanced 利用可（AI はまだ構築中の場合あり）
  - ● 青                    : BM25 の内容と AI embeddings の両方が最新

 [ ファイル内容検索 ]
  - キーワード             : テキストファイル、Outlook、OneNote、Office & PDF

 [ AI Search モデル ]
  - モデル選択 : AI Search 結果ウィンドウ内、🤖 AI Search ボタンの隣
  - 選択肢：Jina-v3 / BGE-Gemma2
  - モデルごとに専用のデータテーブルを使用 — 切り替えても他方のデータは削除されない
  - モデルごとに一度 --update data を実行してインデックスを構築（以後は自由に切替可）

 [ AI Chat ]
  - 表示場所                : 結果の下、"File Content" タブのみ — 🤖 AI Search を
                              クリックすると自動で表示される
  - 自動要約                : 表示された瞬間、現在のキーワードで見つかった
                              最良の結果を自動で要約するよう AI に質問（入力不要）
  - モデル                  : デフォルトは Gemini Flash。Gemini がレート制限に達すると
                              自動で GPT-OSS 120B に切替 — チャットが止まることはない
  - VI / EN / JP ボタン      : 入力した質問の言語に関係なく、AI の返答言語を強制指定。
                              会話中に言語を切り替えると、表示済みの内容もすべて
                              その言語に翻訳し直される
  - ◀ / ▶                  : 現在のキーワードで保存済みの過去の AI Chat セッションを閲覧
  - New Chat アイコン        : 会話をクリアして新しいセッションを開始
  - 根拠のある回答           : AI はユーザーのローカルファイルから実際に抽出された内容
                              （BM25/FTS5 検索）のみに基づいて回答 — 情報が不十分な場合は
                              その旨を明示し、根拠となったファイルを [1][2]... で明記。
                              回答の下に出典一覧が自動で追加される
  - AI Chat History         : 各セッションは自動保存され、後で Help タブから確認可能
                              （上記の履歴を参照）— 閲覧・コピーのみ、再送信は不可
""",
}
HELP_DEFAULT_LANG = "EN"

READY_PH = "Type a keyword to search..."
READY_PH1 = "File/Folder/File content..."
READY_PH2 = "ESC, --exit, --quit"
WAIT_PH = "Updating Database, please wait..."
CHAT_ASK_PH = "Ask anything..."
CHAT_THINKING_PH = "Thinking..."
# v-new (giới hạn auto AI Chat/ngày): sau khi Gemini free-tier bị Google
# cắt quota rất mạnh (~20 request/ngày cho Flash, từ cuối 2025) -- việc
# app tự bắn 1 lượt AI Chat SAU MỖI LẦN SEARCH đốt hết quota đó chỉ trong
# vài phút test bình thường. Giới hạn số lần TỰ ĐỘNG gọi mỗi ngày (không
# giới hạn khi user TỰ gõ hỏi -- đó vẫn luôn được phép), để dành phần
# quota còn lại cho lúc user thật sự cần hỏi thủ công. Chỉnh số này nếu
# muốn nhiều/ít lượt auto hơn.
AUTO_AI_CHAT_DAILY_LIMIT = 10
CHAT_AUTO_LIMIT_REACHED_PH = "Đã hết lượt tự động hôm nay — gõ để hỏi 🙂"

# v-new: language toggle for the AI Chat panel (VI/EN/JP buttons next to the
# "🌐 AI Chat" header -- see _build_online_chat_panel). Picking a language
# switches BOTH the auto-generated summary question sent to the AI AND the
# instruction that forces the AI to answer in that language, regardless of
# what language the user happens to type their own follow-up questions in.
CHAT_DEFAULT_LANG = "VI"
CHAT_LANG_NAMES = {"VI": "Vietnamese", "EN": "English", "JP": "Japanese"}
# v-new: module-level (not a local inside _build_online_chat_panel) so
# _draw_send_button/_set_send_btn_enabled can reference the same colors.
_CHAT_SEND_BG_DEFAULT = "#d9640a"
_CHAT_SEND_BG_HOVER = "#b5540a"
_CHAT_SEND_BG_DISABLED = "#e0b691"
CHAT_TRANSLATING_PH = {
    "VI": "Đang dịch sang Tiếng Việt...",
    "EN": "Translating to English...",
    "JP": "日本語に翻訳中...",
}
# v-new (Vấn đề 1): warning shown when NO API key is configured at all --
# the online AI itself (needed to translate anything) is unavailable in
# exactly this situation, so this can't go through the normal AI-translate
# path (see _translate_chat_worker) like every other chat message does.
# Pre-written in all 3 languages instead and picked directly by
# self._chat_lang wherever it's shown/re-shown.
CHAT_NO_API_KEY_MSG = {
    "VI": "⚠️ Chưa có GEMINI_API_KEY / GROQ_API_KEY. Bấm nút 🔑 Update API để nhập key.",
    "EN": "⚠️ No GEMINI_API_KEY / GROQ_API_KEY set yet. Click the 🔑 Update API button to enter a key.",
    "JP": "⚠️ GEMINI_API_KEY / GROQ_API_KEY が未設定です。🔑 Update API ボタンからキーを入力してください。",
}
CHAT_LANG_OPTIONS = {
    "VI": {
        "label": "VI",
        "auto_msg": lambda query: (
            f"Hãy phân tích và tóm tắt những kết quả tốt nhất tìm được "
            f"cho truy vấn: \"{query}\"."
        ),
        "system_prompt": (
            "Bạn là trợ lý tìm kiếm nội bộ (internal file search assistant), đang trò "
            "chuyện NHIỀU LƯỢT với người dùng. Bạn CHỈ được trả lời dựa trên các đoạn "
            "trích dẫn nội dung file bên dưới -- các đoạn này lấy THẬT từ file trên máy "
            "người dùng qua tìm kiếm BM25/FTS5 cục bộ, không phải do bạn tạo ra. KHÔNG "
            "được tự bịa thêm thông tin không có trong các đoạn trích. Nếu không đủ "
            "thông tin để trả lời, hãy nói rõ là không tìm thấy trong các file đã quét "
            "-- đừng đoán. LUÔN LUÔN trả lời bằng TIẾNG VIỆT (kể cả khi người dùng gõ "
            "câu hỏi bằng ngôn ngữ khác), ngắn gọn, rõ ràng, và chỉ rõ bạn dựa vào file "
            "nào bằng cách chèn số thứ tự [1], [2]... NGAY TRONG câu văn khi trích dẫn -- "
            "LUÔN dùng dấu ngoặc vuông nửa rộng [ ], TUYỆT ĐỐI KHÔNG đổi sang kiểu 【 】 "
            "dù đang trả lời bằng tiếng Nhật hay trích dẫn thuật ngữ tiếng Nhật khác trong "
            "cùng câu trả lời. "
            "QUAN TRỌNG: KHÔNG tự liệt kê danh sách nguồn/tài liệu tham khảo ở cuối câu "
            "trả lời (không viết mục 'Nguồn:', 'Tài liệu trích dẫn:', v.v.) -- hệ thống "
            "sẽ TỰ ĐỘNG thêm danh sách nguồn kèm đường dẫn file đầy đủ ngay bên dưới câu "
            "trả lời của bạn, bạn chỉ cần dùng đúng số thứ tự [N] trong văn bản."
        ),
        "no_context": "(Chưa có kết quả BM25 nào -- hãy tìm bằng từ khoá trước, rồi chat để AI đọc nội dung file.)",
        "no_extract": "(không trích xuất được nội dung)",
        "source_label": "Nguồn",
        "no_answer": "(AI không trả lời được, thử câu hỏi khác nhé.)",
    },
    "EN": {
        "label": "EN",
        "auto_msg": lambda query: (
            f"Please analyze and summarize the best results found "
            f"for the query: \"{query}\"."
        ),
        "system_prompt": (
            "You are an internal file-search assistant, having a MULTI-TURN "
            "conversation with the user. You may ONLY answer based on the file "
            "content excerpts below -- these are REAL excerpts pulled from the "
            "user's own local files via BM25/FTS5 search, not made up by you. Do "
            "NOT invent information that isn't in the excerpts. If there isn't "
            "enough information to answer, clearly say it wasn't found in the "
            "scanned files -- don't guess. ALWAYS answer in ENGLISH (even if the "
            "user types their question in another language), keep it concise and "
            "clear, and indicate which file you're basing each claim on by "
            "inserting the reference number [1], [2]... DIRECTLY in the sentence "
            "when citing -- ALWAYS use standard half-width square brackets [ ], "
            "NEVER switch to full-width 【 】 brackets even when quoting Japanese "
            "terms elsewhere in the same answer. "
            "IMPORTANT: do NOT list a source/reference section "
            "yourself at the end of the answer (no 'Sources:', 'References:', "
            "etc.) -- the system will AUTOMATICALLY append a source list with "
            "full file paths right below your answer; you only need to use the "
            "correct [N] number in the text."
        ),
        "no_context": "(No BM25 results yet -- search with a keyword first, then chat so the AI can read the file content.)",
        "no_extract": "(content could not be extracted)",
        "source_label": "Source",
        "no_answer": "(The AI couldn't answer, please try another question.)",
    },
    "JP": {
        "label": "JP",
        "auto_msg": lambda query: (
            f"クエリ「{query}」で見つかった最も良い検索結果を分析し、要約してください。"
        ),
        "system_prompt": (
            "あなたは社内ファイル検索アシスタントであり、ユーザーと複数回にわたって会話"
            "しています。以下のファイル内容の抜粋のみに基づいて回答してください -- これ"
            "らはユーザーのローカルファイルからBM25/FTS5検索によって実際に取得された抜粋"
            "であり、あなたが作り出したものではありません。抜粋にない情報を勝手に作り出"
            "してはいけません。回答するのに十分な情報がない場合は、推測せずに、スキャン"
            "済みのファイルの中には見つからなかったと明確に伝えてください。ユーザーが別"
            "の言語で質問しても、必ず日本語で回答してください。回答は簡潔かつ明確にし、"
            "引用する際は文中に直接 [1]、[2]... のように番号を挿入して、どのファイルに"
            "基づいているかを示してください。番号は必ず半角の角括弧 [ ] を使い、文中で"
            "他の日本語の用語を【 】(全角) で強調している場合でも、引用番号だけは絶対に"
            "全角の【 】に変えないでください。重要: 回答の最後に自分で出典/参考文献のリ"
            "スト(「出典:」「参考文献:」など)を書かないでください -- システムが回答の"
            "すぐ下に完全なファイルパス付きの出典リストを自動的に追加します。本文中で正"
            "しい [N] 番号を使うだけで構いません。"
        ),
        "no_context": "(まだBM25の検索結果がありません -- 先にキーワードで検索してから、チャットでAIにファイル内容を読ませてください。)",
        "no_extract": "(内容を抽出できませんでした)",
        "source_label": "出典",
        "no_answer": "(AIが回答できませんでした。別の質問を試してください。)",
    },
}

# v-new (nút Online): appended to system_prompt only when the user has the
# 🌐 Online toggle ON (see _online_chat_worker). The base system_prompt in
# CHAT_LANG_OPTIONS above hard-restricts the AI to ONLY the local file
# excerpts -- that restriction is exactly right by default, but the user
# explicitly asked for an opt-in way to loosen it for deeper research.
# Kept as a separate suffix (rather than editing the base prompt) so the
# strict local-only behavior is still the OFF/default path, unchanged.
CHAT_ONLINE_SUFFIX = {
    "VI": (
        "\n\nCHẾ ĐỘ ONLINE ĐANG BẬT: ngoài các đoạn trích file cục bộ ở trên, "
        "giờ bạn ĐƯỢC PHÉP bổ sung thêm kiến thức chung/internet để phân tích "
        "sâu và đầy đủ hơn cho người dùng. Vẫn dùng [1],[2]... khi trích dẫn "
        "từ file cục bộ như trước. Nếu có phần thông tin KHÔNG đến từ file cục "
        "bộ (kiến thức chung hoặc kết quả tìm kiếm web), hãy đặt nó vào một "
        "mục RIÊNG Ở CUỐI CÙNG câu trả lời (sau tất cả các mục phân tích file "
        "cục bộ khác) với tiêu đề CHÍNH XÁC là \"Thông tin bổ sung (internet) "
        "🌐\" (không viết \"ngoài nội bộ\" hay cách gọi khác), không gắn số "
        "[N] cho phần đó. Trong mục này, viết vài dòng phân tích/tóm tắt nội "
        "dung thật sự (không chỉ liệt kê link suông), rồi BẮT BUỘC nêu ÍT "
        "NHẤT 3 nguồn internet KHÁC NHAU (không được ít hơn 3) kèm đường "
        "link thật cho mỗi nguồn, để người dùng có thể mở ra đọc thêm."
    ),
    "EN": (
        "\n\nONLINE MODE IS ON: besides the local file excerpts above, you "
        "are now ALLOWED to add general/internet knowledge for a deeper, "
        "more thorough analysis. Keep using [1],[2]... for anything cited "
        "from the local excerpts as before. If there is information that did "
        "NOT come from the local files (general knowledge or web search "
        "results), put it in its OWN section at the very END of the answer "
        "(after every other local-file analysis section), headed EXACTLY "
        "\"Additional information (internet) 🌐\" (do not say \"outside the "
        "internal files\" or any other wording), without an [N] number. In "
        "that section, write a few sentences of real analysis/summary (not "
        "just a bare list of links), then you MUST cite AT LEAST 3 "
        "DIFFERENT internet sources (never fewer than 3) with a real link "
        "for each, so the user can open them to read more."
    ),
    "JP": (
        "\n\nオンラインモードがオンです: 上記のローカルファイルの抜粋に加え、"
        "より深く十分な分析のために一般知識やインターネットの情報も使用してよ"
        "いです。ローカルファイルからの引用は今まで通り [1]、[2]... を使って"
        "ください。ローカルファイルに由来しない情報(一般知識やWeb検索結果)が"
        "ある場合は、回答の一番最後(他のローカルファイル分析セクションすべて"
        "の後)に独立したセクションを設け、見出しは必ず「追加情報(インターネ"
        "ット) 🌐」という正確な表記にしてください(「内部以外の情報」など他の"
        "言い方は使わないこと)。このセクションには [N] 番号を付けないでくだ"
        "さい。このセクションには、リンクを並べるだけでなく実際の分析・要約"
        "を数行書いたうえで、必ず3件以上の異なるインターネットソースを、そ"
        "れぞれ実際のリンク付きで挙げてください(3件未満は不可)。"
    ),
}

# v-new (yêu cầu 4 -- nút 🌐 không cần phân tích lại nguồn cục bộ): trước
# đây MỌI lượt gọi _online_chat_worker (kể cả lượt bấm riêng nút 🌐) đều
# build lại system_prompt kèm TOÀN BỘ context cục bộ (đọc lại snippet từng
# file trong context_paths) rồi bắt model tóm tắt/phân tích lại từ đầu --
# dư thừa, vì lượt phân tích cục bộ đã có sẵn ngay phía trên trong cùng
# khung chat (tự động chạy ngay khi search), và đây vốn dĩ là nút CHỈ ĐỂ
# tra cứu internet. CHAT_ONLINE_SUFFIX (phía trên) cũng ép model viết thêm
# 1 mục riêng "Thông tin bổ sung (internet)" bọc ngoài -- không cần thiết
# nữa vì bây giờ TOÀN BỘ câu trả lời của lượt 🌐 đã là "thông tin từ
# internet" rồi, không có gì cục bộ để tách riêng ra nữa. Prompt riêng này
# thay hẳn cho system_prompt (không nối thêm CHAT_ONLINE_SUFFIX) khi
# force_web_search=True (xem _online_chat_worker): không đọc/gửi bất kỳ
# đoạn trích file cục bộ nào, chỉ yêu cầu trả lời thẳng bằng thông tin/kiến
# thức internet cho đúng từ khoá, không có mục "bổ sung" nào bọc ngoài.
CHAT_WEB_ONLY_SYSTEM_PROMPT = {
    "VI": (
        "Bạn là trợ lý tìm kiếm internet, đang trò chuyện NHIỀU LƯỢT với "
        "người dùng. Lượt này người dùng CHỦ ĐỘNG bấm nút tra cứu internet "
        "cho 1 từ khoá cụ thể -- KHÔNG cần đọc/phân tích lại file cục bộ "
        "nào (việc đó đã làm xong ở các lượt chat trước đó rồi). Hãy trả "
        "lời THẲNG bằng thông tin/kiến thức thật tìm được trên internet cho "
        "đúng từ khoá đó -- viết vài đoạn phân tích/tóm tắt nội dung thật "
        "sự (không chỉ liệt kê link suông), rồi nêu ÍT NHẤT 3 nguồn "
        "internet KHÁC NHAU kèm đường link thật cho mỗi nguồn. KHÔNG tự bịa "
        "thông tin nếu không tìm được gì -- nói rõ là không tìm thấy. "
        "KHÔNG đặt câu trả lời vào bất kỳ mục/tiêu đề 'bổ sung' nào -- viết "
        "thẳng như 1 câu trả lời bình thường. KHÔNG dùng số trích dẫn kiểu "
        "[1],[2]... (đó là quy ước dành riêng cho file cục bộ, lượt này "
        "không có file nào cả). LUÔN LUÔN trả lời bằng TIẾNG VIỆT (kể cả "
        "khi người dùng gõ câu hỏi bằng ngôn ngữ khác), ngắn gọn và rõ ràng."
    ),
    "EN": (
        "You are an internet-search assistant, having a MULTI-TURN "
        "conversation with the user. This turn, the user explicitly clicked "
        "the internet-search button for one specific keyword -- you do NOT "
        "need to read/re-analyze any local files (that was already done in "
        "earlier turns of this chat). Answer DIRECTLY using real "
        "information/knowledge found on the internet for that keyword -- "
        "write a few paragraphs of real analysis/summary (not just a bare "
        "list of links), then cite AT LEAST 3 DIFFERENT internet sources "
        "with a real link for each. Do NOT invent information if you find "
        "nothing -- clearly say so instead. Do NOT wrap the answer in any "
        "'additional/supplementary' section or heading -- just write it as "
        "a normal, direct answer. Do NOT use [1],[2]... style citation "
        "numbers (that convention is only for local files, and this turn "
        "has none). ALWAYS answer in ENGLISH (even if the user types their "
        "question in another language), and keep it concise and clear."
    ),
    "JP": (
        "あなたはインターネット検索アシスタントであり、ユーザーと複数回にわ"
        "たって会話しています。この回では、ユーザーが特定のキーワードに対し"
        "て自らインターネット検索ボタンを押しました -- ローカルファイルを読"
        "み直したり再分析したりする必要はありません(それはこのチャットの前"
        "の回ですでに済んでいます)。そのキーワードについて、インターネット"
        "上で見つかった実際の情報・知識を使って直接回答してください -- リン"
        "クを並べるだけでなく実際の分析・要約を数段落書いたうえで、必ず3件"
        "以上の異なるインターネットソースを、それぞれ実際のリンク付きで挙げ"
        "てください。何も見つからない場合は情報を作り出さず、見つからなか"
        "ったと明確に伝えてください。回答を「追加情報」などの見出しで囲ま"
        "ず、通常の回答としてそのまま書いてください。[1]、[2]...のような引"
        "用番号は使わないでください(それはローカルファイル専用の表記で、"
        "この回には該当ファイルがありません)。ユーザーが別の言語で質問して"
        "も、必ず日本語で簡潔かつ明確に回答してください。"
    ),
}

# v-fix (Vấn đề: comment/thông báo cũ nói sai là "GPT-OSS không có khả năng
# tìm internet"): đã kiểm tra lại docs Groq hiện tại
# (console.groq.com/docs/tool-use/built-in-tools/browser-search) -- Groq
# THỰC SỰ có built-in tool browser_search cho đúng openai/gpt-oss-120b (và
# -20b), dùng Exa để duyệt web tương tác chứ không chỉ lấy snippet như Web
# Search thường. _call_online_ai_chat giờ tự bật tools=[{"type":
# "browser_search"}] khi web_search=True và model là "gptoss". Model DUY
# NHẤT thực sự không có khả năng tìm internet là Qwen 3.6 27B ("qwen") --
# xem WEB_SEARCH_CAPABLE_MODELS. Note này giờ chỉ còn xuất hiện khi: (a)
# Qwen đang là model active và không thể tự chuyển sang Gemini/GPT-OSS
# (cả 2 đều hết quota/thiếu key), hoặc (b) trường hợp hiếm SDK groq cũ chưa
# hỗ trợ tools= (xem fail-soft trong _call_online_ai_chat).
CHAT_NO_WEB_SEARCH_NOTE = {
    "VI": ("\n\n🌐 _Lưu ý: model đang dùng hiện không có khả năng tìm kiếm "
           "internet thật -- câu trả lời trên chỉ dựa vào các file cục bộ "
           "đã quét (không có tìm kiếm web thật sự). Muốn bật tìm kiếm "
           "internet thật, hãy chuyển sang Gemini Flash hoặc GPT-OSS 120B ở "
           "ô chọn model, và đảm bảo GEMINI_API_KEY/GROQ_API_KEY còn hoạt "
           "động (xem nút 🔑 Update API)._"),
    "EN": ("\n\n🌐 _Note: the active model has no real internet-search "
           "ability -- the answer above is based only on the local files "
           "scanned (no live web search was actually performed). To enable "
           "real internet search, switch to Gemini Flash or GPT-OSS 120B in "
           "the model picker, and make sure GEMINI_API_KEY/GROQ_API_KEY is "
           "valid (see the 🔑 Update API button)._"),
    "JP": ("\n\n🌐 _注: 現在使用中のモデルには実際のインターネット検索機能が"
           "ありません -- 上記の回答はスキャン済みのローカルファイルのみに"
           "基づいています(実際のWeb検索は行われていません)。実際のインタ"
           "ーネット検索を使うには、モデル選択で Gemini Flash か GPT-OSS "
           "120B に切り替え、GEMINI_API_KEY/GROQ_API_KEY が有効か確認して"
           "ください(🔑 Update API ボタン参照)。_"),
}

# v-new (yêu cầu: text link "Click here..." ở cuối phần phân tích cục bộ,
# thay thế/bổ sung cho nút 🌐 riêng -- xem _online_chat_worker, nơi dòng
# này được nối vào cuối answer, và _insert_chat_text_with_links, nơi nó
# được nhận diện qua _CHAT_INTERNET_LINK_MARKER và biến thành link click
# được, gọi lại _send_web_search_chat_msg() y hệt bấm nút 🌐).
# _CHAT_INTERNET_LINK_MARKER dùng 2 ký tự ZERO WIDTH SPACE (U+200B) làm
# "dấu vân tay" vô hình bọc quanh dòng -- không hiển thị trên UI, nhưng
# đủ để phân biệt chắc chắn với bất kỳ dòng text thường nào khác (kể cả
# nếu tình cờ trùng y hệt nội dung hiển thị) mà không cần so khớp toàn
# bộ câu chữ theo từng ngôn ngữ.
_CHAT_INTERNET_LINK_MARKER = "\u200bINTERNET_LINK\u200b"
CHAT_INTERNET_LINK_TEXT = {
    "VI": "🌐 Bấm vào đây để xem thêm thông tin từ Internet",
    # v-note: câu tiếng Anh gốc user đề xuất ("Click here to have further
    # more information from Internet") bị thừa từ + sai giới từ -- sửa lại
    # cho đúng ngữ pháp tự nhiên.
    "EN": "🌐 Click here for more information from the Internet",
    "JP": "🌐 ここをクリックすると、インターネットからさらに詳しい情報が見られます",
}

# v-new (yêu cầu: thêm 1 label "🤖 ..." NGAY PHÍA TRÊN link "🌐 ..." ở
# trên, để AI Chat có thể phân tích lại bằng CHÍNH kết quả AI Search
# offline -- Jina/BGE, self._ai_cont_res -- thay vì chỉ BM25 như thường
# lệ, phong phú hơn hẳn nhưng chậm hơn nên KHÔNG tự động, chỉ khi user
# chủ động bấm). Dùng chung cơ chế marker vô hình như
# _CHAT_INTERNET_LINK_MARKER ở trên, nhưng có thêm 1 ký tự TRẠNG THÁI
# ngay sau marker ('1' = đã chạy 🤖 AI Search cho keyword này rồi, label
# hiện màu xanh bấm được; '0' = CHƯA bấm 🤖 AI Search lần nào cho keyword
# này, label hiện màu xám, không bấm được) -- xem
# _insert_chat_text_with_links (nơi đọc ký tự này) và _online_chat_worker
# (nơi quyết định '1' hay '0' tại thời điểm sinh câu trả lời, dựa vào
# self._ai_mode_active/self._ai_active_query/self._ai_cont_res).
_CHAT_AISEARCH_LINK_MARKER = "\u200bAISEARCH_LINK\u200b"
CHAT_AI_SEARCH_LINK_TEXT = {
    "VI": "🤖 Bấm vào đây để xem thêm thông tin từ AI Search offline",
    "EN": "🤖 Click here for more information from AI Search offline",
    "JP": "🤖 ここをクリックすると、AI Search offline からさらに詳しい情報が見られます",
}


MAX_CHARS_TO_INDEX = 60000 

CMD_PATH = r"C:\Windows\System32\cmd.exe /k"

SKIP_FOLDERS = {
    # Windows system
    "windows", "$recycle.bin", "system volume information", "appdata",
    # Program folders
    "program files", "program files (x86)", "programdata",
    # Python installations
    "python39", "python311", "python312", "python314", "python3", "python",
    # Database / middleware
    "oracle19c", "oracle", "mysql", "postgresql",
    # Cloud sync — files may be cloud-only placeholders, opening triggers download
    "onedrive", "onedrive - personal", "sharepoint", "google drive", "dropbox", "box",
    # Dev / build artifacts
    "node_modules", ".git", "__pycache__", ".venv", "venv", "env",
    # 3rd party app data
    "3ds",
    # IT / Security
    "_it", "it", "security", "credentials", "vault", "secrets",
    "password", "passwords", "private", "confidential", "restricted",
}
# Files whose names match these patterns are never opened for content indexing
import re as _re_sec
_SENSITIVE_FILE_RE = _re_sec.compile(
    r'(?i)(password|passwd|credential|secret|private[_\-]?key|'
    r'id_rsa|vault|keystore|\.kdbx|keepass|lastpass|api[_\-]?key)'
)
BINARY_EXT = {".exe", ".dll", ".lib", ".obj", ".pyc", ".bin", ".jpg", ".png", ".gif", ".zip", ".7z", ".rar"}

# v2.5: Recognize a downloaded HuggingFace / SentenceTransformers model repo by
# its file signature (config.json alongside tokenizer/weight files) — this way
# ANY local model cache (jina, bge, e5-large, e5-base, or any future model,
# regardless of folder name) gets skipped during --update data, instead of
# being walked and partially indexed as if it were user documents.
_MODEL_REPO_WEIGHT_OR_TOKENIZER_FILES = {
    "tokenizer.json", "tokenizer_config.json", "pytorch_model.bin",
    "model.safetensors", "modules.json", "sentence_bert_config.json",
    "adapter_config.json",
}
def _looks_like_model_repo(dir_path):
    try:
        names = set(os.listdir(dir_path))
    except Exception:
        return False
    return "config.json" in names and bool(names & _MODEL_REPO_WEIGHT_OR_TOKENIZER_FILES)

# ── Semantic Search config ────────────────────────────────────────────────────
# ── Multi-model AI registry ───────────────────────────────────────────────────
# Each model has: weight folder (auto-detected under D:\mySearch\models or next to .py/.exe),
# its own table in search_data.db (BM25/content still share ONE DB — only
# embedding tables are split because each model produces different-dimension/space vectors),
# and its own encode convention (each model uses different prefix/task/instruction).
SEMANTIC_MODELS = {
    "jina_v3": {
        "label":        "Jina-v3(570M)",
        "dir_names":    ["jina-embeddings-v3", "jina-v3", "jina_v3"],
        "table":        "semantic_index_jina_v3",
        "trust_remote": True,   # jina-v3 requires trust_remote_code=True for LoRA task adapter
        "hf_repo":      "jinaai/jina-embeddings-v3",
        # v-new: rough resident-memory footprint, used by _pick_device_for_model
        # and _dual_resident_allowed() to decide GPU/CPU + whether it's safe
        # to keep both models loaded at once. jina-v3 is a small (~570M
        # param) encoder, light on both VRAM and RAM.
        "est_vram_mb":  1500,
        "est_ram_mb":   2000,
    },
    "bge_gemma2": {
        "label":        "BGE-Gemma2(9B)",
        "dir_names":    ["bge-multilingual-gemma2", "bge-gemma2", "bge_gemma2"],
        "table":        "semantic_index_bge_gemma2",
        "trust_remote": True,
        # v-new: force_cpu is now an OPTIONAL manual override, not the
        # primary mechanism. Device is decided automatically at runtime in
        # _pick_device_for_model() by comparing actual free VRAM against
        # est_vram_mb below -- e.g. on this laptop's ~4GB T1200 GPU it will
        # still land on CPU (same practical result as before), but on a
        # machine with a bigger GPU it would automatically get to run on
        # GPU too, with no code change needed. Set force_cpu: True here
        # only if you want to hard-pin this model to CPU regardless of
        # what the estimate says.
        "est_vram_mb":  20000,  # Gemma2-9B-based, ~18-20GB — won't fit most laptop GPUs
        "est_ram_mb":   20000,
        "hf_repo":      "BAAI/bge-multilingual-gemma2",
    },
}
DEFAULT_SEMANTIC_MODEL = "jina_v3"

# ══════════════════════════════════════════════════════════════════════════
# ONLINE AI BLOCK (v1.2) — Gemini Flash & GPT-OSS 120B (via Groq)
#
# v1.2: no longer tied to the "AI Search" button (that button is now 100%
# offline, Jina-v3/BGE-Gemma2) -- these 2 online models now power the
# separate CHAT panel shown below the results (RealtimeSmartSearchApp.
# _build_online_chat_panel / _online_chat_worker). Gemini Flash is the
# default, auto-switching to GPT-OSS 120B when Gemini hits a rate limit
# (429/RESOURCE_EXHAUSTED).
#
# Anti-hallucination principle -- same as how app_astro-weather forces
# Gemini to call a tool instead of guessing the weather: these 2 online
# models are NOT allowed to make things up. They only act as "read +
# summarize" over REAL file content pulled from content_store (already
# built into search_data.db by the earlier local BM25/FTS5 search) -- see
# _online_chat_worker() and _fetch_content_snippets() below. FINDING files
# (retrieval) stays 100% local/offline as before; only the natural-language
# ANSWER GENERATION needs the online API call. If the API key is missing or
# there's no network, the offline BM25/AI search (Jina-v3/BGE-Gemma2) is
# completely unaffected.
# ══════════════════════════════════════════════════════════════════════════
ONLINE_GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
ONLINE_GROQ_API_KEY   = os.environ.get("GROQ_API_KEY", "")
# v-fix: dummy placeholders "aaa"/"bbb" removed -- bool("aaa") is always
# True, so _online_ai_available() previously reported a key as "present"
# even when nothing real had ever been set. Empty string now correctly
# evaluates to "no key".

# v-fix (yêu cầu 1 -- v9.5 bị Gemini 429 ngay từ request đầu, "chuyển
# thẳng" sang GPT-OSS, dù v4.0 chạy Gemini OK và cả 2 bản "tưởng" là cùng
# 1 key): nguyên nhân THẬT SỰ nằm ở đây, không phải ở model/thứ tự
# candidate. v4.0 KHÔNG có tính năng "🔑 Update API" -- nó không hề đọc
# search_data.db cho API key, mà luôn dùng đúng 1 key hard-code cố định
# ngay trong code (xem ONLINE_GEMINI_API_KEY của v4.0). v9.5 thì khác:
# nó có thêm bảng `settings` trong search_data.db để nút "Update API" lưu
# key (xem _load_api_keys_from_db/_save_api_keys_to_db bên dưới) -- và
# key trong DB đó được ưu tiên NGAY SAU biến môi trường. Nếu TỪNG có 1
# key Gemini nào được nhập/lưu qua dialog đó ở 1 phiên trước (kể cả nhập
# thử/nhầm), nó sẽ NẰM Ở ĐÓ MÃI MÃI và được dùng cho mọi lần mở app sau
# này -- HOÀN TOÀN có thể là 1 key KHÁC với key hard-code trong v4.0, dù
# nhìn hai bên tưởng "giống nhau". Nếu đúng key đó đã hết quota free-tier
# trong ngày, v9.5 sẽ luôn 429 ngay lập tức trong khi v4.0 (dùng key
# hard-code riêng, chưa từng bị đụng tới) vẫn chạy bình thường.
# => Cách kiểm tra/khắc phục NGAY: mở nút "🔑 Update API" trong v9.5, xem
# giá trị đang nằm trong ô GEMINI_API_KEY -- nếu nó không phải key bạn
# muốn dùng, bấm "Delete API Key" (xoá key đã lưu trong DB) rồi Save, HOẶC
# dán đè bằng key đang hoạt động rồi Save. _save_api_keys_to_db() đã tự
# xoá cache client + reset cờ rate-limited nên không cần khởi động lại app.
#
# Fix bổ sung ở đây: nếu CẢ biến môi trường LẪN DB đều trống (cài mới,
# hoặc vừa bấm "Delete API Key"), ONLINE_GEMINI_API_KEY ở lại rỗng --
# _online_ai_available() sẽ báo "chưa có key", và người dùng cần tự nhập
# key của mình qua nút "🔑 Update API" (hoặc biến môi trường GEMINI_API_KEY)
# trước khi dùng AI Chat online. Không còn hard-code sẵn 1 key thật trong
# code nữa (để repo có thể public/lên GitHub an toàn).


def _ensure_settings_table():
    """v-new (Update API): tiny key/value table inside search_data.db
    (DB_FILE) so GEMINI_API_KEY / GROQ_API_KEY entered via the Update API
    dialog survive an app restart -- no more re-entering them every time.
    Kept in the same DB the user already has (search_data.db) rather than
    a separate file, per user's own suggestion. Safe/cheap to call
    repeatedly (CREATE TABLE IF NOT EXISTS)."""
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Settings] could not init settings table: {e}")


def _load_api_keys_from_db():
    """v-new: called once at startup (after _ensure_settings_table). Env
    vars still take priority if set (advanced/CI use case) -- the DB value
    only fills in when the env var is empty."""
    global ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        rows = dict(conn.execute(
            "SELECT key, value FROM settings WHERE key IN ('gemini_api_key','groq_api_key')").fetchall())
        conn.close()
    except Exception as e:
        print(f"[Settings] could not load API keys: {e}")
        rows = {}
    if not ONLINE_GEMINI_API_KEY and rows.get("gemini_api_key"):
        ONLINE_GEMINI_API_KEY = rows["gemini_api_key"]
        print("[Settings] GEMINI_API_KEY: dùng key đã LƯU TRONG search_data.db "
              "(nhập trước đó qua nút 'Update API') -- có thể KHÁC với key mặc "
              "định hard-code trong v4.0. Nếu Gemini bị 429 liên tục, mở '🔑 "
              "Update API' để kiểm tra/xoá/đổi key này.")
    if not ONLINE_GROQ_API_KEY and rows.get("groq_api_key"):
        ONLINE_GROQ_API_KEY = rows["groq_api_key"]
    # v-fix (yêu cầu 1): trước đây có 1 key hard-code làm fallback cuối
    # cùng khi cả env var lẫn DB đều trống -- đã bỏ (không hard-code key
    # thật trong code nữa). Nếu vẫn trống tới đây, ONLINE_GEMINI_API_KEY
    # ở lại "" và _online_ai_available() sẽ báo thiếu key cho người dùng.


def _save_api_keys_to_db(gemini_key=None, groq_key=None):
    """v-new: writes whichever key(s) are non-empty to the settings table,
    updates the in-memory globals immediately, and resets the cached API
    clients so the VERY NEXT chat call already uses the new key (no
    restart needed)."""
    global ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
    global _online_gemini_client, _online_groq_client
    _ensure_settings_table()
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        if gemini_key:
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('gemini_api_key', ?)", (gemini_key,))
            ONLINE_GEMINI_API_KEY = gemini_key
            _online_gemini_client = None
        if groq_key:
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('groq_api_key', ?)", (groq_key,))
            ONLINE_GROQ_API_KEY = groq_key
            _online_groq_client = None
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Settings] could not save API key(s): {e}")

def _auto_ai_chat_count_today():
    """v-new (giới hạn auto AI Chat/ngày): trả về số lần AI Chat đã TỰ
    ĐỘNG gọi hôm nay -- lưu trong bảng settings (search_data.db) với key
    khoá theo NGÀY hiện tại ('auto_ai_chat_count_YYYY-MM-DD'), nên: (1)
    đếm đúng dù app tắt/mở lại nhiều lần trong ngày (không phải biến RAM
    mất khi tắt app), và (2) tự "reset về 0" khi sang ngày mới mà không
    cần dọn dẹp gì -- key của ngày cũ đơn giản không còn được đọc tới
    nữa."""
    _ensure_settings_table()
    key = f"auto_ai_chat_count_{datetime.now().strftime('%Y-%m-%d')}"
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        conn.close()
        return int(row[0]) if row else 0
    except Exception:
        return 0  # lỗi đọc DB -- coi như chưa dùng lượt nào, không chặn oan


def _increment_auto_ai_chat_count():
    """v-new: +1 vào bộ đếm auto AI Chat của NGÀY HIỆN TẠI -- xem
    _auto_ai_chat_count_today()."""
    _ensure_settings_table()
    key = f"auto_ai_chat_count_{datetime.now().strftime('%Y-%m-%d')}"
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        new_val = (int(row[0]) if row else 0) + 1
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(new_val)))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Settings] could not bump auto_ai_chat_count: {e}")


def _delete_api_keys_from_db(clear_gemini=True, clear_groq=True):
    """v-new (nút "Delete API Key" trong dialog Update API): xoá key đã lưu
    trong bảng settings (search_data.db) và reset lại biến global về giá
    trị mặc định từ biến môi trường (nếu có) hoặc rỗng, đồng thời reset
    client đã cache -- đối xứng với _save_api_keys_to_db() ở trên."""
    global ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
    global _online_gemini_client, _online_groq_client
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        if clear_gemini:
            conn.execute("DELETE FROM settings WHERE key = 'gemini_api_key'")
            # v-fix (yêu cầu 1): trước đây rơi về key hard-code mặc định
            # sau khi xoá key trong DB -- đã bỏ hard-code, giờ rơi về
            # chuỗi rỗng (hoặc biến môi trường nếu có).
            ONLINE_GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
            _online_gemini_client = None
        if clear_groq:
            conn.execute("DELETE FROM settings WHERE key = 'groq_api_key'")
            ONLINE_GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
            _online_groq_client = None
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"[Settings] could not delete API key(s): {e}")


# ══════════════════════════════════════════════════════════════════════════
# v-new (yêu cầu: nhiều API key/provider, đọc/ghi qua configure.ini, tự
# động XOAY sang key tiếp theo khi key hiện tại bị rate-limit/hết quota):
# _GEMINI_API_KEYS / _GROQ_API_KEYS giữ TOÀN BỘ danh sách key đã lưu
# trong [APIKeys] của configure.ini (gemini_key_1, gemini_key_2, ... /
# groq_key_1, groq_key_2, ...) -- ONLINE_GEMINI_API_KEY/ONLINE_GROQ_API_KEY
# (2 biến global đã có sẵn từ trước, dùng khắp nơi trong app) LUÔN LÀ key
# đang "active" = list[idx hiện tại]. _rotate_api_key() được
# _call_online_ai()/_call_online_ai_chat() tự gọi ngay khi phát hiện lỗi
# 429/quota (xem _is_online_ai_rate_limit_err) -- không cần user tự vào
# Update API đổi key giữa chừng nữa.
# ══════════════════════════════════════════════════════════════════════════
_GEMINI_API_KEYS = []
_GROQ_API_KEYS = []
_gemini_key_idx = 0
_groq_key_idx = 0


def _load_api_keys_from_config():
    """Đọc danh sách key Gemini/Groq từ configure.ini. Nếu configure.ini
    CHƯA có key nào (lần đầu chạy sau khi cập nhật lên bản có tính năng
    này), tự động migrate 1 LẦN DUY NHẤT từ nguồn cũ (search_data.db /
    biến môi trường / key hard-code mặc định -- xem _load_api_keys_from_db)
    sang configure.ini, để không ai mất key đang dùng khi cập nhật app."""
    global _GEMINI_API_KEYS, _GROQ_API_KEYS, _gemini_key_idx, _groq_key_idx
    global ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
    cfg = _load_config()
    _GEMINI_API_KEYS = _numbered_keys_from_config(cfg, "gemini_key_")
    _GROQ_API_KEYS = _numbered_keys_from_config(cfg, "groq_key_")
    if not _GEMINI_API_KEYS or not _GROQ_API_KEYS:
        _load_api_keys_from_db()  # fills ONLINE_GEMINI_API_KEY/ONLINE_GROQ_API_KEY globals from the old source
        if not _GEMINI_API_KEYS and ONLINE_GEMINI_API_KEY:
            _GEMINI_API_KEYS = [ONLINE_GEMINI_API_KEY]
            _save_numbered_keys_to_config("gemini_key_", _GEMINI_API_KEYS)
        if not _GROQ_API_KEYS and ONLINE_GROQ_API_KEY:
            _GROQ_API_KEYS = [ONLINE_GROQ_API_KEY]
            _save_numbered_keys_to_config("groq_key_", _GROQ_API_KEYS)
    _gemini_key_idx = 0
    _groq_key_idx = 0
    ONLINE_GEMINI_API_KEY = _GEMINI_API_KEYS[0] if _GEMINI_API_KEYS else ""
    ONLINE_GROQ_API_KEY = _GROQ_API_KEYS[0] if _GROQ_API_KEYS else ""


def _rotate_api_key(provider):
    """Xoay sang API key TIẾP THEO trong danh sách của `provider`
    ("gemini" hoặc "groq") khi key hiện tại bị 429/hết quota. Trả về True
    nếu vừa xoay sang 1 key THỰC SỰ KHÁC (còn >=2 key đã lưu); False nếu
    chỉ có 1 key (hoặc 0 key) -- không có gì để xoay, caller nên báo lỗi
    như bình thường thay vì thử lại vô ích."""
    global _gemini_key_idx, _groq_key_idx, ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
    if provider == "gemini" and len(_GEMINI_API_KEYS) > 1:
        _gemini_key_idx = (_gemini_key_idx + 1) % len(_GEMINI_API_KEYS)
        ONLINE_GEMINI_API_KEY = _GEMINI_API_KEYS[_gemini_key_idx]
        _reset_online_ai_clients()
        print(f"[API Key] Gemini key #{_gemini_key_idx} hết quota -- tự động "
              f"chuyển sang key #{_gemini_key_idx + 1}/{len(_GEMINI_API_KEYS)}")
        return True
    if provider == "groq" and len(_GROQ_API_KEYS) > 1:
        _groq_key_idx = (_groq_key_idx + 1) % len(_GROQ_API_KEYS)
        ONLINE_GROQ_API_KEY = _GROQ_API_KEYS[_groq_key_idx]
        _reset_online_ai_clients()
        print(f"[API Key] Groq key #{_groq_key_idx} hết quota -- tự động "
              f"chuyển sang key #{_groq_key_idx + 1}/{len(_GROQ_API_KEYS)}")
        return True
    return False



# falls through to the next candidate (same idea as
# ASTRO_AI_MODEL_CANDIDATES in app_astro-weather).
# v-fix (theo yêu cầu -- Gemini 3.6-flash bị rate-limit ngay từ request đầu
# tiên trong ngày, mọi ngày, trong khi cùng API key vẫn chạy bình thường ở
# app khác): 3.6-flash/3.5-flash-lite là dòng model MỚI NHẤT (ra mắt
# 21/07/2026) -- quota Free-tier riêng cho dòng 3.x thường eo hẹp hơn hẳn
# so với dòng 2.5 cũ hơn, đã ổn định lâu (xác nhận: gemini-2.5-flash /
# gemini-2.5-flash-lite vẫn còn hoạt động, chưa bị tắt -- ngày tắt sớm
# nhất theo Google hiện là 16/10/2026). Đổi thứ tự: thử 2.5 (ổn định, quota
# rộng hơn) TRƯỚC, rồi mới rơi xuống 3.6/3.5-lite/latest nếu 2.5 cũng hết
# quota hoặc bị tắt.
ONLINE_GEMINI_MODEL_CANDIDATES = [
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-3.6-flash",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
]
ONLINE_GROQ_MODEL = "openai/gpt-oss-120b"
# v-new (yêu cầu: thêm model Qwen 3.6 27B, thứ tự Gemini Flash -> GPT-OSS ->
# Qwen 3.6 27B): Groq hosts qwen/qwen3.6-27b on the same account/API key as
# GPT-OSS -- see _get_online_groq_client(). It does NOT support Groq's
# browser_search tool (only openai/gpt-oss-20b and openai/gpt-oss-120b do,
# per console.groq.com/docs/tool-use/built-in-tools/browser-search), so it's
# excluded from WEB_SEARCH_CAPABLE_MODELS below.
ONLINE_GROQ_QWEN_MODEL = "qwen/qwen3.6-27b"
# Maps an ONLINE_AI_MODELS key to the actual Groq model id / a short label
# used in error messages -- lets _call_online_ai / _call_online_ai_chat
# share one code path for every Groq-hosted model instead of duplicating
# the whole branch per model.
GROQ_MODEL_IDS = {"gptoss": ONLINE_GROQ_MODEL, "qwen": ONLINE_GROQ_QWEN_MODEL}
GROQ_MODEL_LABELS = {"gptoss": "GPT-OSS 120B", "qwen": "Qwen 27B"}

ONLINE_AI_MODELS = {
    "gemini": {"label": "Gemini Flash (online)"},
    "gptoss": {"label": "GPT-OSS 120B (online)"},
    "qwen": {"label": "Qwen 27B (online)"},
}
DEFAULT_ONLINE_MODEL = "gemini"
# v-new: the priority order used everywhere a model is auto-picked or
# auto-switched to on rate-limit/missing-key (see _next_alt_model) --
# Gemini Flash first, then GPT-OSS 120B, then Qwen 3.6 27B.
MODEL_FALLBACK_ORDER = ["gemini", "gptoss", "qwen"]
# v-new: which models can actually honor "Online" (real internet search)
# mode -- Gemini via its google_search grounding tool, GPT-OSS via Groq's
# browser_search built-in tool (see _call_online_ai_chat). Qwen has neither,
# so Online mode falls back to one of these two when Qwen is active.
WEB_SEARCH_CAPABLE_MODELS = {"gemini", "gptoss"}
# v-new: module-level so both _build_online_chat_panel (initial dropdown
# text) and _sync_chat_model_picker (keeping it in sync with auto-fallback)
# use the exact same short labels.
MODEL_SHORT_LABELS = {"gemini": "Gemini Flash", "gptoss": "GPT-OSS 120B", "qwen": "Qwen 27B"}

# Number of top BM25 files whose content is read to build context for the
# online AI. Kept small so the prompt stays short, fast, and cheap (no need
# to index the entire DB).
ONLINE_AI_TOP_N_FILES = 8
# v-fix: cap on TOTAL files sent as context once we also mix in files that
# match the current follow-up question's keywords (see _online_chat_worker).
# Slightly higher than ONLINE_AI_TOP_N_FILES alone so a relevant match found
# further down the result list doesn't just bump a top-8 file out entirely.
# v-fix (theo yêu cầu mới: muốn source phong phú như v4.0 trở lại): v4.0
# dùng 14 cho Gemini (không có GPT-OSS budget riêng nên vô tình cũng gồng
# luôn cho cả Groq, dẫn tới lỗi 413 khi dùng GPT-OSS -- xem GPTOSS_MAX_
# CONTEXT_FILES bên dưới). Từng bị giảm xuống 10 theo 1 yêu cầu trước đó để
# prompt gọn hơn -- nay tăng lại về 14 CHỈ cho Gemini: context window Gemini
# Flash rất lớn (hàng trăm nghìn token), quota generateContent thường (khác
# quota grounding) cũng đủ rộng, nên không có lý do kỹ thuật nào phải giữ
# thấp. GPT-OSS/Qwen (GPTOSS_MAX_CONTEXT_FILES) GIỮ NGUYÊN 6 -- 2 model đó
# vẫn bị giới hạn 8000 TPM thật của Groq (xem lỗi 413 đã gặp trước đây).
ONLINE_AI_MAX_CONTEXT_FILES = 14
# v-fix: guaranteed number of Outlook / OneNote items each reserved a slot
# in AI Chat's context, on top of the top ONLINE_AI_TOP_N_FILES normal
# files -- see _online_chat_worker. Without this, mail/notes results almost
# never made it into context at all once there were >= ONLINE_AI_TOP_N_FILES
# normal files ranked ahead of them.
ONLINE_AI_MAIL_RESERVE_FILES = 3
# v-new (yêu cầu: chia rõ số source theo từng model thay vì suy ra bằng
# phép chia nguyên _max_ctx_files // 4 -- công thức cũ tình cờ ra đúng số
# này (6 GPT-OSS/Qwen -> 6//4=1, 10 Gemini -> 10//4=2) nhưng không tường
# minh và dễ lệch nếu GPTOSS_MAX_CONTEXT_FILES/ONLINE_AI_MAX_CONTEXT_FILES
# đổi sau này): GPT-OSS/Qwen = 4 pdf/ppt/doc... + 1 Outlook + 1 OneNote (=6
# tổng theo GPTOSS_MAX_CONTEXT_FILES). Gemini = 10 pdf/ppt/doc... + 2 Outlook
# + 2 OneNote (=14 tổng theo ONLINE_AI_MAX_CONTEXT_FILES, tăng lại từ 10 lên
# 14 theo yêu cầu -- xem comment ở ONLINE_AI_MAX_CONTEXT_FILES). Đây là số TỐI ĐA
# mỗi loại được lấy -- nếu Outlook/OneNote có ÍT HƠN số này (kể cả 0), phần
# dư luôn tự động chảy về cho pdf/ppt/doc (xem _normal_budget trong
# _online_chat_worker) chứ không bỏ phí. Nếu KHÔNG thấy nguồn OneNote nào
# trong câu trả lời dù đã chỉnh ở đây, nhiều khả năng là do OneNote chưa
# được Update DB (search_onenote.db rỗng/thiếu) hoặc không có trang OneNote
# nào thật sự khớp với câu hỏi -- không phải do slot bị chiếm.
ONLINE_AI_MAIL_RESERVE_BY_MODEL = {"gemini": 2, "gptoss": 1, "qwen": 1}
ONLINE_AI_SNIPPET_CHARS = 700

# v-new (yêu cầu: nút "🤖 Bấm vào đây để xem thêm thông tin từ AI Search
# offline" phải phân tích SÂU HƠN hẳn chat thường -- nhiều nguồn hơn, đọc
# nhiều chữ hơn mỗi nguồn, chấp nhận chậm hơn 1 chút): các hằng số
# *_DETAILED này CHỈ áp dụng cho đúng lượt bấm nút đó (use_ai_search_context
# =True trong _online_chat_worker), không đụng tới chat thường/nút 🌐 --
# xem chỗ dùng trong _online_chat_worker. Gemini được tăng mạnh (context
# window rất lớn, quota generateContent thoải mái). GPT-OSS/Qwen chỉ tăng
# nhẹ để tránh lặp lại lỗi 413 "Request too large" (giới hạn 8000 TPM thật
# của Groq -- xem comment ở GPTOSS_MAX_CONTEXT_FILES bên dưới).
ONLINE_AI_MAX_CONTEXT_FILES_DETAILED = 20
ONLINE_AI_SNIPPET_CHARS_DETAILED = 1400
ONLINE_AI_MAIL_RESERVE_BY_MODEL_DETAILED = {"gemini": 3, "gptoss": 1, "qwen": 1}
# v-fix (Vấn đề: lỗi 413 "Request too large" liên tục trên Groq): Groq's
# free/on-demand tier caps at 8000 TOKENS PER MINUTE for openai/gpt-oss-120b
# -- MUCH smaller than Gemini's context window. The context block built for
# Gemini (ONLINE_AI_MAX_CONTEXT_FILES=10 files x ONLINE_AI_SNIPPET_CHARS=700
# chars, plus the full running chat history) regularly blew straight past
# that 8000-token ceiling in a single request once a conversation had a few
# turns, which is exactly the "Requested 8378 / Limit 8000" error the user
# hit. Smaller caps below apply whenever a Groq-hosted model (GPT-OSS OR
# Qwen -- both share the same account/TPM ceiling) is actually the model
# being called (see _online_chat_worker) -- Gemini calls are unaffected.
# v-new (theo yêu cầu: GPT-OSS/Qwen dùng 6 source thay vì 5): tăng nhẹ từ 5
# lên 6, vẫn đủ an toàn so với ngưỡng TPM hiện tại của Groq.
GPTOSS_MAX_CONTEXT_FILES = 6
GPTOSS_SNIPPET_CHARS = 350
GPTOSS_MAX_CONTEXT_FILES_DETAILED = 8    # nút "AI Search offline" -- tăng nhẹ, vẫn an toàn với TPM của Groq
GPTOSS_SNIPPET_CHARS_DETAILED = 450
# v-new (nút "+" Add file): người dùng tự đính kèm 1 file bất kỳ (không
# cần nằm trong index BM25) để AI Chat luôn đọc nội dung file đó, không
# phụ thuộc thứ hạng BM25 -- giải quyết đúng vấn đề "file muốn hỏi nằm ở
# hạng 15~20/300 kết quả nên không lọt vào context". Giới hạn 2MB theo
# yêu cầu -- đủ cho hầu hết pdf/docx/pptx/xlsx văn bản, chặn sớm trước
# khi tốn thời gian extract 1 file quá nặng.
CHAT_ATTACH_MAX_BYTES = 5 * 1024 * 1024
CHAT_ATTACH_ALLOWED_EXT = {
    ".txt", ".md", ".csv", ".log", ".pdf", ".docx", ".doc",
    ".xlsx", ".xls", ".pptx", ".ppt", ".one",
}
# v-new (yêu cầu 4 -- dán ảnh vào ô chat): các đuôi file ảnh được chấp
# nhận, cả khi chọn qua nút "+" LẪN khi dán trực tiếp từ clipboard (xem
# _try_paste_image_from_clipboard). Gộp vào CHAT_ATTACH_ALLOWED_EXT để
# nút "+" (file picker) cũng liệt kê được ảnh, không chỉ paste mới dùng
# được.
CHAT_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
CHAT_ATTACH_ALLOWED_EXT |= CHAT_IMAGE_EXT
CHAT_IMAGE_MIME_BY_EXT = {
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".webp": "image/webp", ".bmp": "image/bmp", ".gif": "image/gif",
}
# v-new (yêu cầu 4): chỉ Gemini đọc được ảnh (multimodal qua inline_data,
# xem _call_online_ai_chat) -- GPT-OSS 120B/Qwen 27B trên Groq chỉ hỗ trợ
# chat completions TEXT-ONLY cho 2 model này. Khi model đang trả lời KHÔNG
# phải Gemini mà vẫn có ảnh đính kèm, chèn ghi chú này vào system prompt
# thay vì lặng lẽ bỏ qua ảnh -- để model không bịa nội dung ảnh, và câu
# trả lời tự giải thích lý do "không xem được ảnh" cho user hiểu.
CHAT_IMAGE_NOT_SUPPORTED_NOTE = {
    "VI": "\n\n[Lưu ý: user đã đính kèm 1 ẢNH, nhưng model hiện tại không đọc được ảnh (chỉ Gemini Flash mới hỗ trợ) -- đừng đoán/bịa nội dung ảnh, hãy nói rõ điều này cho user.]",
    "EN": "\n\n[Note: the user attached an IMAGE, but the current model cannot read images (only Gemini Flash supports vision) -- do not guess/invent the image's content, tell the user this instead.]",
    "JP": "\n\n[注意: ユーザーが画像を添付しましたが、現在のモデルは画像を読み取れません（対応しているのは Gemini Flash のみ）。画像の内容を推測せず、その旨をユーザーに伝えてください。]",
}
GPTOSS_MAX_HISTORY_MSGS = 6   # last ~3 user/assistant turn pairs
# v-fix (Vấn đề: dịch cả cuộc hội thoại dài luôn fail -- "could not parse
# translated array" / "Translate failed" / lỗi 413 Groq): _translate_chat_worker
# used to send EVERY message in the conversation, no matter how long it had
# grown, as ONE giant "reply with a JSON array" request. Once a conversation
# has many turns (each with a full multi-paragraph answer + citation list),
# that blows the SAME way past Groq's 8000-TPM cap the chat-context fix
# above addressed (see the 413 "Requested 12476 / Limit 8000" error) -- but
# ALSO makes the reply itself huge, which is fragile even on Gemini: a
# giant translated JSON array is much more likely to get cut off mid-string
# by an output-token limit (-> unrecoverably broken JSON, "could not parse
# translated array") or to trip Gemini's RECITATION filter on its own
# output (-> empty response, "Translate failed"). Capping how many of the
# MOST RECENT messages get sent/translated per request keeps every
# translate call small and reliable regardless of how long the overall
# conversation has grown; older messages simply stay in their current
# language until translated in a later, smaller batch.
TRANSLATE_MAX_MSGS_GEMINI = 20
TRANSLATE_MAX_MSGS_GPTOSS = GPTOSS_MAX_HISTORY_MSGS

_online_gemini_client = None
_online_groq_client = None


def _online_ai_available(model_choice):
    if model_choice == "gemini":
        return bool(ONLINE_GEMINI_API_KEY)
    if model_choice in GROQ_MODEL_IDS:  # "gptoss", "qwen" -- share the same Groq key
        return bool(ONLINE_GROQ_API_KEY)
    return False


def _get_online_gemini_client():
    global _online_gemini_client
    if _online_gemini_client is None:
        from google import genai as _genai  # pip install google-genai
        _online_gemini_client = _genai.Client(api_key=ONLINE_GEMINI_API_KEY)
    return _online_gemini_client


def _get_online_groq_client():
    global _online_groq_client
    if _online_groq_client is None:
        from openai import OpenAI as _OpenAI  # pip install openai (Groq uses the standard OpenAI-compatible API)
        _online_groq_client = _OpenAI(base_url="https://api.groq.com/openai/v1",
                                       api_key=ONLINE_GROQ_API_KEY)
    return _online_groq_client


def _reset_online_ai_clients():
    """v-fix (Vấn đề: cả Gemini lẫn GPT-OSS đều báo 429, nhưng bấm 'Update
    API' -> Save -- KHÔNG đổi key, chỉ bấm Save lại -- là hết lỗi ngay):
    _save_api_keys_to_db() tự set _online_gemini_client/_online_groq_client
    về None, buộc lần gọi kế tiếp phải tạo lại client SDK (google-genai /
    openai) từ đầu thay vì tái dùng client cũ đã dính request 429 -- các
    SDK này giữ 1 HTTP session/connection tái sử dụng bên trong client
    instance, và có vẻ như session đó "nhớ" trạng thái backoff/lỗi khiến
    request tiếp theo trên CÙNG client vẫn fail dù quota server đã đủ
    trở lại. Tạo client mới "xoá" session cũ đó. Hàm này tái tạo đúng
    hiệu ứng ấy để app tự làm, không cần user vào Settings bấm Save."""
    global _online_gemini_client, _online_groq_client
    _online_gemini_client = None
    _online_groq_client = None


def _best_snippet_window(content, keywords, max_chars):
    """Pick the excerpt actually likely to answer the question, instead of
    always the document's opening paragraph.

    v-fix (AI Chat không tìm được nội dung rõ ràng có trong file): the old
    behavior was content[:max_chars] -- literally just the first ~700
    characters of the ENTIRE extracted text, no matter what was asked. For
    a multi-page PDF/manual, that's usually just the title page/table of
    contents; a topic covered on page 40 (e.g. "Inertia Relief Modes")
    was never in the excerpt sent to the AI at all, so it correctly-but-
    unhelpfully reported "not found in the excerpt I have" even while the
    file itself was correctly identified and cited as a source -- the
    citation only proves the FILE was in context, not that the RELEVANT
    PART of it was. Now: find the first occurrence of any keyword from
    the current question in the full text, and center the window there
    instead. Falls back to the start of the document when no keyword
    matches inside it (e.g. a generic "summarize this" with nothing
    specific to anchor on)."""
    if not content:
        return ""
    if keywords:
        low = content.lower()
        best_pos = -1
        for kw in keywords:
            pos = low.find(kw)
            if pos != -1 and (best_pos == -1 or pos < best_pos):
                best_pos = pos
        if best_pos != -1:
            start = max(0, best_pos - max_chars // 4)
            return content[start:start + max_chars]
    return content[:max_chars]


def _fetch_content_snippets(paths, keywords=None, max_chars=ONLINE_AI_SNIPPET_CHARS):
    """Read the REAL content of the files (already stored in content_store,
    built earlier by Update DB -- not guessed by the AI) to use as context
    for the online AI. Only a short excerpt is sent, never the whole file
    -- see _best_snippet_window for how that excerpt is now chosen."""
    out = []
    if not paths:
        return out
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        c = conn.cursor()
        ph = ",".join("?" * len(paths))
        c.execute(f"SELECT path, content FROM content_store WHERE path IN ({ph})", list(paths))
        rows = {r[0]: r[1] for r in c.fetchall()}
        conn.close()
    except Exception as _e:
        print(f"[Online AI] snippet fetch error: {_e}")
        rows = {}
    kw_list = [k for k in (keywords or []) if k]
    for p in paths:
        content = (rows.get(p) or "").strip()
        out.append({"path": p, "snippet": _best_snippet_window(content, kw_list, max_chars)})
    return out


def _call_online_ai_impl(model_choice, system_prompt, user_prompt, max_output_tokens=None):
    """Call the online model -- this is a SYNCHRONOUS (blocking) function,
    the caller MUST run it inside a background threading.Thread, never
    directly on the main UI thread (same as every other heavy task in this
    app).

    max_output_tokens (v-fix, Vấn đề: dịch sang JP hay bị "could not parse
    translated array" hơn hẳn VI/EN): optional override for callers whose
    reply is expected to be unusually long relative to the input (e.g.
    _translate_chat_worker re-emitting a whole JSON array of translated
    messages) -- Japanese in particular tends to need noticeably MORE
    output tokens than the same content in Vietnamese/English (CJK text
    tokenizes less efficiently per character for these models), so the
    same message batch that fits fine in EN/VI can get cut off mid-JSON
    specifically for JP under the SDK's default output cap -- an
    unrecoverably broken/truncated array, not a parsing bug. None (the
    default) leaves the SDK's normal default untouched for every other
    caller of this function."""
    if model_choice == "gemini":
        if not ONLINE_GEMINI_API_KEY:
            return None, "Thiếu GEMINI_API_KEY (biến môi trường)."
        try:
            from google.genai import types as _genai_types
            client = _get_online_gemini_client()
        except Exception as e:
            return None, f"Chưa cài thư viện google-genai: {e}"
        last_err = None
        for model in ONLINE_GEMINI_MODEL_CANDIDATES:
            try:
                cfg_kwargs = {"system_instruction": system_prompt}
                if max_output_tokens:
                    cfg_kwargs["max_output_tokens"] = max_output_tokens
                resp = client.models.generate_content(
                    model=model, contents=user_prompt,
                    config=_genai_types.GenerateContentConfig(**cfg_kwargs),
                )
                text = (resp.text or "").strip()
                if text:
                    return text, None
                # v-fix (Vấn đề: "🤖 ⚠️ Translate failed" dù key đang OK):
                # resp.text can come back EMPTY with no exception at all --
                # most often Gemini's RECITATION safety filter, which is
                # very easy to trip specifically for a translate request
                # like this one (asked to reproduce the SAME structure --
                # headings, [N] citation markers, numbered lists -- just in
                # another language, which looks a lot like "verbatim
                # reproduction" to that filter). The old code returned
                # ("", None) right here treating this as a normal success
                # with an empty answer -- the caller then saw raw="" AND
                # err=None and fell through to the generic, unhelpful
                # "Translate failed" text. Now: pull the actual
                # finish_reason (RECITATION / SAFETY / MAX_TOKENS / ...)
                # out of the response for a real diagnostic, and try the
                # NEXT candidate model instead of giving up immediately --
                # same as the exception path below.
                reason = "unknown"
                try:
                    cand = (resp.candidates or [None])[0]
                    reason = getattr(cand, "finish_reason", None) or reason
                except Exception:
                    pass
                last_err = f"Gemini trả về nội dung rỗng (finish_reason={reason})"
                continue
            except Exception as e:
                last_err = e
                continue
        return None, f"Lỗi Gemini (đã thử {len(ONLINE_GEMINI_MODEL_CANDIDATES)} model): {last_err}"

    if model_choice in GROQ_MODEL_IDS:  # "gptoss" or "qwen"
        if not ONLINE_GROQ_API_KEY:
            return None, "Thiếu GROQ_API_KEY (biến môi trường)."
        _groq_label = GROQ_MODEL_LABELS[model_choice]
        try:
            client = _get_online_groq_client()
            create_kwargs = {
                "model": GROQ_MODEL_IDS[model_choice],
                "messages": [{"role": "system", "content": system_prompt},
                             {"role": "user", "content": user_prompt}],
            }
            if max_output_tokens:
                create_kwargs["max_tokens"] = max_output_tokens
            resp = client.chat.completions.create(**create_kwargs)
            return (resp.choices[0].message.content or "").strip(), None
        except Exception as e:
            # v-fix (Vấn đề: dịch EN bị lỗi 413 "Request too large ... on
            # tokens per minute (TPM): Limit 8000, Requested 11324" dù key
            # vẫn hợp lệ): tài khoản Groq (tier "on_demand") giới hạn tổng
            # token PROMPT + max_tokens trong MỘT request theo TPM khá thấp
            # (8000 ở tier hiện tại). Khi max_output_tokens truyền vào
            # (ước lượng generous cho câu trả lời dài) cộng với token của
            # prompt vượt mức này, Groq trả 413 ngay lập tức, chưa hề gọi
            # model. Thay vì đoán mò lại theo tỉ lệ ký tự/token (không
            # chính xác, nhất là VI/JP có mật độ token khác hẳn EN), Groq
            # đã trả sẵn đúng 2 con số cần thiết ("Limit" và "Requested")
            # ngay trong error message -- dùng thẳng 2 số đó để tính lại
            # budget chính xác rồi tự động thử lại MỘT lần với max_tokens
            # đã giảm vừa đủ.
            err_text = str(e)
            m = re.search(r"Limit\s+(\d+),\s*Requested\s+(\d+)", err_text)
            if m and max_output_tokens:
                limit, requested = int(m.group(1)), int(m.group(2))
                prompt_tokens = max(0, requested - max_output_tokens)
                # Chừa margin an toàn 200 token; sàn tối thiểu 512 token để
                # bản dịch/trả lời còn có ý nghĩa -- nếu ngân sách còn lại
                # quá nhỏ thì báo lỗi rõ ràng luôn thay vì gửi request chắc
                # chắn lại fail lần nữa.
                safe_budget = limit - prompt_tokens - 200
                if safe_budget >= 512:
                    try:
                        create_kwargs["max_tokens"] = min(safe_budget, max_output_tokens)
                        resp = client.chat.completions.create(**create_kwargs)
                        return (resp.choices[0].message.content or "").strip(), None
                    except Exception as e2:
                        return None, f"Lỗi {_groq_label} (Groq) sau khi giảm max_tokens còn {safe_budget}: {e2}"
                return None, (
                    f"Lỗi {_groq_label} (Groq): nội dung quá dài so với giới "
                    f"hạn TPM {limit} của tài khoản hiện tại (prompt đã tốn "
                    f"~{prompt_tokens} token, không còn đủ ngân sách cho "
                    f"output). Cân nhắc nâng Groq Dev Tier hoặc giảm độ dài "
                    f"câu trả lời."
                )
            return None, f"Lỗi {_groq_label} (Groq): {e}"

    return None, f"Model online không hợp lệ: {model_choice}"


def _call_online_ai(model_choice, system_prompt, user_prompt, max_output_tokens=None):
    """v-new (Multi-key rotation): wrapper mỏng quanh _call_online_ai_impl
    -- khi trả lời lỗi VÀ đó là lỗi rate-limit/quota (429), tự động xoay
    sang API key TIẾP THEO của cùng provider (nếu configure.ini có lưu
    nhiều key -- xem _rotate_api_key) rồi thử lại NGAY với ĐÚNG model đó,
    trước khi để caller coi đây là rate-limit và chuyển sang model khác
    hẳn. Không đổi chữ ký/giá trị trả về so với bản cũ."""
    provider = "gemini" if model_choice == "gemini" else ("groq" if model_choice in GROQ_MODEL_IDS else None)
    _tried_keys = 1
    while True:
        result = _call_online_ai_impl(model_choice, system_prompt, user_prompt, max_output_tokens)
        answer, err = result
        if answer is not None or not provider:
            return result
        if not _is_online_ai_rate_limit_err(str(err)):
            return result
        if not _rotate_api_key(provider):
            return result  # no more untried keys for this provider
        _tried_keys += 1
        print(f"[API Key] Retrying {model_choice} with key #{_tried_keys} after quota error")


def _log_online_ai_fallback(text):
    """v-new (Vấn đề: Gemini chỉ dùng được 1 lần rồi tự chuyển sang GPT-OSS,
    không rõ lý do thật sự): các dòng print(f"[Online Chat][Debug] ...")
    quanh _online_chat_worker RẤT hữu ích để xem lỗi Gemini thật sự trả về
    là gì (429 hết quota thật, hay model name sai/404, hay lỗi khác bị
    _is_online_ai_rate_limit_err() nhận nhầm) -- nhưng khi build bằng
    PyInstaller với console=False (windowed mode), sys.stdout bị trỏ vào
    os.devnull ngay từ đầu file (xem đầu file) nên TOÀN BỘ các dòng print()
    đó biến mất, không cách nào xem được. Ghi thêm ra search_error.log
    (cùng thư mục DB, giống cách Indexing Error đã làm) để dù chạy bản .exe
    đóng gói, người dùng vẫn mở file này lên xem được lý do thật sự."""
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(DB_FILE)), "search_error.log")
        with open(log_path, "a", encoding="utf-8") as lf:
            lf.write(f"\n[{datetime.now()}] ONLINE AI FALLBACK:\n{text}\n")
    except Exception:
        pass


def _is_online_ai_rate_limit_err(err_text):
    """Detect a rate-limit/quota error (Gemini 429/RESOURCE_EXHAUSTED, or an
    equivalent error from Groq) so the online chat panel can automatically
    switch from Gemini Flash to GPT-OSS 120B -- see _call_online_ai_chat /
    RealtimeSmartSearchApp._online_chat_worker."""
    if not err_text:
        return False
    t = str(err_text).lower()
    return ("429" in t or "resource_exhausted" in t or "rate limit" in t
            or "rate_limit" in t or "quota" in t or "too many requests" in t)


def _extract_json_str_array(raw, expected_len=None):
    """v-fix (Vấn đề: 🤖 ⚠️ Translate failed): robustly pull a JSON array of
    strings out of an LLM reply that is SUPPOSED to be "only a JSON array"
    but in practice often comes wrapped in ```json fences, with a leading/
    trailing sentence like "Here is the translated array:" or "Hope this
    helps!", or with a trailing explanation after the array. Instead of a
    single strict json.loads() (which threw on any of the above and always
    surfaced as "Translate failed"), this tries several increasingly
    lenient strategies and only gives up if none of them yield a list of
    strings of the expected length.
    Returns (list_of_str, None) on success, or (None, error_message) on
    failure -- mirrors the (result, err) convention used elsewhere in this
    file (see _call_online_ai)."""
    if not raw or not str(raw).strip():
        return None, "empty response"
    text = str(raw).strip()

    def _try(candidate):
        # strict=False: chat message text legitimately contains literal
        # newlines (multi-paragraph answers, citation lists) -- real
        # models very often emit those newlines RAW inside the JSON string
        # values instead of escaping them as \n. Under strict JSON that is
        # an "Invalid control character" parse error and used to fail
        # EVERY strategy below identically (hence the parse always died
        # with the same "could not parse translated array" message no
        # matter what shape the reply had). strict=False allows literal
        # control characters inside strings, matching what real models
        # actually send.
        try:
            val = json.loads(candidate, strict=False)
        except Exception:
            # v-fix (Vấn đề: dịch luôn thất bại từ khi câu trả lời có kèm
            # đường dẫn Windows trong "Nguồn:"): the instruction tells the
            # model to keep every file path EXACTLY as written, and a
            # Windows path is full of single backslashes (C:\Users\...\
            # file.pdf). The model dutifully copies them byte-for-byte
            # into the JSON string -- but a lone backslash there is an
            # INVALID JSON escape (only \", \\, \/, \b, \f, \n, \r, \t,
            # \uXXXX are legal), so json.loads rejects the whole payload
            # outright, every single time a path shows up in the
            # conversation -- which explains why translate now fails
            # consistently instead of occasionally. Escape any backslash
            # that isn't already part of a valid escape sequence, then
            # retry once.
            try:
                repaired_bs = re.sub(r'\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})', r'\\\\', candidate)
                val = json.loads(repaired_bs, strict=False)
            except Exception:
                return None
        if not isinstance(val, list):
            # some models wrap the array in an object (e.g.
            # {"translations": [...]}) despite the "ONLY a JSON array"
            # instruction -- unwrap the first list-valued field found.
            if isinstance(val, dict):
                for v in val.values():
                    if isinstance(v, list):
                        val = v
                        break
                else:
                    return None
            else:
                return None
        if expected_len is not None and len(val) != expected_len:
            return None
        return val

    # 1) as-is
    out = _try(text)
    if out is not None:
        return out, None

    # 2) strip ```json ... ``` / ``` ... ``` code fences
    fenced = re.sub(r"^```(?:json)?\s*", "", text)
    fenced = re.sub(r"\s*```$", "", fenced.strip())
    out = _try(fenced)
    if out is not None:
        return out, None

    # 3) try EVERY '[' as a possible array start (not just the first) --
    # a preamble sentence before the array can itself contain a '[' (e.g.
    # the model echoing "giữ nguyên [1], [2]..." from the instructions
    # before actually starting the array, which is common specifically
    # when the target language makes the model more talkative, e.g. JP).
    # Naively using the FIRST '[' in the whole text then grabs that stray
    # bracket instead of the real array start, and the LAST ']' in the
    # text is still fine as an end point since nothing legitimate follows
    # the array. For each '[' position, try both the plain slice and a
    # trailing-comma-repaired version before moving to the next '['.
    end = text.rfind("]")
    if end != -1:
        search_from = 0
        while True:
            start = text.find("[", search_from, end + 1)
            if start == -1 or start >= end:
                break
            candidate = text[start:end + 1]
            out = _try(candidate)
            if out is not None:
                return out, None
            repaired = re.sub(r",\s*([\]}])", r"\1", candidate)
            out = _try(repaired)
            if out is not None:
                return out, None
            search_from = start + 1

    # 4) last resort: pull out every top-level double-quoted string in
    # source order via regex (handles a truncated/near-miss array where
    # exact bracket matching failed but the strings themselves are intact).
    # re.DOTALL so a literal newline inside a string doesn't end the match
    # early (same raw-newline issue strict=False works around above).
    if expected_len is not None:
        strs = re.findall(r'"((?:[^"\\]|\\.)*)"', text, re.DOTALL)
        if len(strs) == expected_len:
            try:
                return [json.loads(f'"{s}"', strict=False) for s in strs], None
            except Exception:
                try:
                    return [json.loads(
                        '"' + re.sub(r'\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})', r'\\\\', s) + '"',
                        strict=False) for s in strs], None
                except Exception:
                    pass

    return None, "could not parse translated array"


def _extract_web_sources(resp):
    """v-new (nguồn internet trong mục Source): when Gemini's built-in
    Google Search grounding tool actually kicked in for this turn, the
    response carries candidates[0].grounding_metadata.grounding_chunks --
    each chunk has a .web.uri / .web.title for a real page the model drew
    on. Pulled out here so the caller can list them in the chat's "Nguồn"/
    "Source" section alongside local file citations, instead of the
    external info showing up with no link back to where it came from.
    Best-effort: returns [] (never raises) if the SDK version/response
    shape doesn't have this, or grounding simply wasn't used for this
    particular answer (e.g. the question was fully answered locally)."""
    out = []
    try:
        cand = resp.candidates[0]
        gm = getattr(cand, "grounding_metadata", None) or getattr(cand, "groundingMetadata", None)
        if not gm:
            return out
        chunks = getattr(gm, "grounding_chunks", None) or getattr(gm, "groundingChunks", None) or []
        seen = set()
        for ch in chunks:
            web = getattr(ch, "web", None)
            if not web:
                continue
            uri = getattr(web, "uri", None)
            title = getattr(web, "title", None) or uri
            if uri and uri not in seen:
                seen.add(uri)
                out.append({"title": title, "url": uri})
    except Exception:
        pass
    return out


def _extract_groq_web_sources(resp):
    """v-new (yêu cầu: AI Chat phân tích GPT-OSS không thấy Source nào từ
    internet): khi Groq's browser_search tool THỰC SỰ chạy,
    response.choices[0].message.executed_tools chứa (các) tool call đã
    thực thi -- mỗi tool call browser_search/web_search mang theo
    .search_results (danh sách trang đã duyệt/ghé qua), mỗi kết quả có
    .url/.title (một số phiên bản SDK trả về dict thay vì object, hoặc gói
    trong .output dạng JSON string/dict). Hàm này cố lấy ra danh sách
    [{"title","url"}] giống hệt _extract_web_sources() bên Gemini, để
    _append_web_sources() có thể in ra mục "Nguồn (internet) 🌐" thật sự có
    link, thay vì luôn rỗng như trước đây (_call_online_ai_chat cũ luôn
    return [] cho nhánh Groq bất kể browser_search có chạy hay không).
    Best-effort: không bao giờ raise, trả về [] nếu SDK không có field này
    hoặc browser_search không thực sự được gọi lượt này."""
    out = []
    try:
        msg = resp.choices[0].message
        executed = getattr(msg, "executed_tools", None) or []
        seen = set()
        for t in executed:
            results = getattr(t, "search_results", None)
            if results is None:
                output = getattr(t, "output", None)
                if isinstance(output, str):
                    try:
                        import json as _json
                        parsed = _json.loads(output)
                        results = parsed.get("results") if isinstance(parsed, dict) else parsed
                    except Exception:
                        results = None
                elif isinstance(output, dict):
                    results = output.get("results")
            if not results:
                continue
            for r in results:
                if isinstance(r, dict):
                    url = r.get("url") or r.get("uri")
                    title = r.get("title") or url
                else:
                    url = getattr(r, "url", None) or getattr(r, "uri", None)
                    title = getattr(r, "title", None) or url
                if url and url not in seen:
                    seen.add(url)
                    out.append({"title": title, "url": url})
    except Exception:
        pass
    return out


def _call_online_ai_chat_impl(model_choice, system_prompt, history_msgs, user_message, web_search=False, image=None,
                          on_chunk=None, on_reset=None):
    """Same as _call_online_ai() above but supports a MULTI-TURN
    conversation -- used by the online chat panel below the AI Search
    results, instead of a single question/single answer like before.

    history_msgs: list[{"role": "user"|"assistant", "text": str}] -- the
    PREVIOUS turns in this same chat session (does not include the current
    user_message). Returns (answer_text, err, web_sources) -- web_sources
    is a list of {"title","url"} dicts (possibly empty), see
    _extract_web_sources().

    web_search (v-new, nút Online -- giờ luôn True mặc định, xem
    _online_chat_worker): for model_choice == "gemini", turns on Gemini's
    built-in Google Search grounding tool. For model_choice == "gptoss",
    turns on Groq's built-in browser_search tool (only openai/gpt-oss-20b
    and openai/gpt-oss-120b support it -- see
    console.groq.com/docs/tool-use/built-in-tools/browser-search). Qwen
    (model_choice == "qwen") has neither -- callers should switch
    model_choice to "gemini" or "gptoss" first when web_search is wanted
    for that reason (see _online_chat_worker / WEB_SEARCH_CAPABLE_MODELS)
    -- this function does NOT do that switch itself, it just silently
    ignores web_search for "qwen" (falls back to normal, non-grounded
    behavior).

    image (v-new, yêu cầu 4 -- dán ảnh vào ô chat): dict tuỳ chọn
    {"bytes": <raw bytes>, "mime": "image/png"}. CHỈ Gemini
    (client SDK google-genai hỗ trợ multimodal qua Part.from_bytes) thực
    sự "nhìn" được ảnh này -- Groq/gpt-oss-120b và qwen3.6-27b (chat
    completions text-only trên 2 model này) KHÔNG đọc được ảnh, nên nhánh
    Groq bên dưới cố tình bỏ qua `image`, hàm gọi (_online_chat_worker) tự
    chèn 1 dòng ghi chú vào system prompt cho 2 model đó thay vì gửi ảnh
    đi vô ích.

    on_chunk / on_reset (v-new, streaming -- theo yêu cầu "hiệu ứng gõ
    chữ dần giống ChatGPT/web AI"): khi truyền on_chunk (callable(str)),
    hàm chuyển sang gọi API ở chế độ STREAM và gọi on_chunk(delta_text)
    cho MỖI mẩu text ngay khi model sinh ra, thay vì đợi trả lời xong mới
    trả về 1 cục -- caller (xem _online_chat_worker) dùng nó để đổ chữ
    dần vào khung chat cho cảm giác nhanh. Hàm vẫn trả về
    (answer_text, err, web_sources, used_web_search) y hệt như không
    streaming SAU KHI stream kết thúc, để pipeline hậu xử lý hiện có
    (đổi số trích dẫn, gắn "Nguồn:", gắn nguồn internet...) hoạt động
    trên văn bản ĐẦY ĐỦ như trước giờ -- streaming chỉ ảnh hưởng tới
    CÁCH hiển thị tạm thời, không đổi kết quả cuối cùng. on_reset (nếu có)
    được gọi khi 1 lượt thử đang stream dở bị lỗi giữa chừng và hàm sắp
    thử lại (model khác / bỏ tool grounding) -- để caller xoá phần text
    thô đã lỡ đổ vào UI trước khi thử tiếp, tránh 2 câu trả lời dở dang
    dính liền nhau."""
    if model_choice == "gemini":
        if not ONLINE_GEMINI_API_KEY:
            return None, "Thiếu GEMINI_API_KEY.", [], False
        try:
            from google.genai import types as _genai_types
            client = _get_online_gemini_client()
        except Exception as e:
            return None, f"Chưa cài thư viện google-genai: {e}", [], False
        contents = []
        for m in history_msgs:
            role = "user" if m.get("role") == "user" else "model"
            contents.append({"role": role, "parts": [{"text": m.get("text", "")}]})
        _user_parts = [{"text": user_message}]
        if image and image.get("bytes"):
            # v-new (yêu cầu 4): ảnh vừa dán được đính kèm NGAY vào lượt hỏi
            # hiện tại dưới dạng inline_data -- Gemini đọc trực tiếp pixel,
            # không cần OCR/text-extract như file thường.
            try:
                # v-note: theo đúng schema REST của Gemini (inlineData.data),
                # "data" phải là base64 STRING, không phải raw bytes --
                # image["bytes"] luôn là raw bytes (xem _read_image_bytes_b64
                # ở _online_chat_worker), encode lại ở đây.
                _img_b64 = _b64.b64encode(image["bytes"]).decode("ascii")
                _user_parts.append({
                    "inline_data": {
                        "mime_type": image.get("mime") or "image/png",
                        "data": _img_b64,
                    }
                })
            except Exception as _img_e:
                print(f"[Online Chat] could not attach image to Gemini request: {_img_e}")
        contents.append({"role": "user", "parts": _user_parts})
        gen_config_kwargs = {"system_instruction": system_prompt}
        if web_search:
            try:
                gen_config_kwargs["tools"] = [
                    _genai_types.Tool(google_search=_genai_types.GoogleSearch())
                ]
            except Exception as _tool_e:
                # v-fix: older google-genai versions may not expose
                # GoogleSearch() the same way -- fail SOFT (answer without
                # grounding) rather than losing the whole chat turn over it.
                print(f"[Online Chat] Google Search tool unavailable, "
                      f"continuing without web grounding: {_tool_e}")

        def _try_gemini_models(kwargs, candidates):
            """Try each Gemini model candidate in turn (streamed if
            on_chunk was given, plain otherwise) and return
            (answer, sources_resp, last_err) -- answer is None only if
            EVERY candidate failed. On a streaming candidate that fails
            AFTER some text was already emitted via on_chunk, on_reset is
            called first so the caller's UI doesn't keep a dangling
            partial answer around before the next candidate is tried."""
            _last_err = None
            for _model in candidates:
                try:
                    if on_chunk:
                        stream = client.models.generate_content_stream(
                            model=_model, contents=contents,
                            config=_genai_types.GenerateContentConfig(**kwargs),
                        )
                        _parts, _last_chunk, _emitted = [], None, False
                        for _ch in stream:
                            _last_chunk = _ch
                            _t = getattr(_ch, "text", None)
                            if _t:
                                _parts.append(_t)
                                _emitted = True
                                on_chunk(_t)
                        return "".join(_parts).strip(), _last_chunk, None
                    resp = client.models.generate_content(
                        model=_model, contents=contents,
                        config=_genai_types.GenerateContentConfig(**kwargs),
                    )
                    return (resp.text or "").strip(), resp, None
                except Exception as _e:
                    _last_err = _e
                    if on_chunk and on_reset:
                        on_reset()  # discard any partial text from this failed attempt
                    continue
            return None, None, _last_err

        answer_text, resp_for_sources, last_err = _try_gemini_models(gen_config_kwargs, ONLINE_GEMINI_MODEL_CANDIDATES)
        if answer_text is not None:
            sources = _extract_web_sources(resp_for_sources) if resp_for_sources is not None else []
            return answer_text, None, sources, web_search

        # v-fix (yêu cầu: 429 grounding không nên nhảy thẳng sang GPT-OSS):
        # quota "Google Search grounding" là 1 quota RIÊNG BIỆT, thấp hơn
        # nhiều so với quota generateContent thường -- và áp dụng CHUNG cho
        # cả project/API key, không tính riêng từng model, nên khi 1 model
        # bị 429 grounding thì gần như chắc chắn cả 5 model candidate ở
        # trên đều 429 y hệt nhau (đúng như log debug user báo cáo -- "đã
        # thử 5 model" đều RESOURCE_EXHAUSTED). Trước khi bỏ cuộc và để
        # _online_chat_worker coi đây là rate-limit rồi chuyển hẳn sang
        # GPT-OSS, thử lại NGAY 1 lần nữa với CHÍNH Gemini nhưng bỏ tool
        # grounding đi -- quota generateContent thường rộng rãi hơn nhiều
        # nên khả năng cao vẫn trả lời được (đúng như v4.0, vốn không hề
        # dùng grounding, luôn chạy êm với cùng 1 API key). Chỉ làm việc
        # này khi lượt gọi VỪA RỒI thực sự có bật tools (nếu không có tools
        # từ đầu thì đây đã là "thử không grounding" rồi, không có gì để
        # rút bớt nữa).
        if "tools" in gen_config_kwargs and _is_online_ai_rate_limit_err(str(last_err)):
            print("[Online Chat] Gemini grounding quota exhausted -- "
                  "retrying without google_search tool before giving up")
            _no_tool_kwargs = {k: v for k, v in gen_config_kwargs.items() if k != "tools"}
            answer_text, _, last_err = _try_gemini_models(_no_tool_kwargs, ONLINE_GEMINI_MODEL_CANDIDATES)
            if answer_text is not None:
                # used_web_search=False -- caller may want to append a "no
                # real internet access this turn" note, since this answer
                # is ungrounded even though web_search was asked.
                return answer_text, None, [], False

        return None, f"Lỗi Gemini (đã thử {len(ONLINE_GEMINI_MODEL_CANDIDATES)} model): {last_err}", [], web_search

    if model_choice in GROQ_MODEL_IDS:  # "gptoss" or "qwen"
        if not ONLINE_GROQ_API_KEY:
            return None, "Thiếu GROQ_API_KEY.", [], False
        _groq_label = GROQ_MODEL_LABELS[model_choice]
        try:
            client = _get_online_groq_client()
            msgs = [{"role": "system", "content": system_prompt}]
            for m in history_msgs:
                role = "user" if m.get("role") == "user" else "assistant"
                msgs.append({"role": role, "content": m.get("text", "")})
            msgs.append({"role": "user", "content": user_message})
            create_kwargs = {"model": GROQ_MODEL_IDS[model_choice], "messages": msgs}
            # v-new (Groq browser_search): xác nhận lại từ docs Groq hiện
            # tại -- openai/gpt-oss-120b (và -20b) THỰC SỰ có built-in tool
            # browser_search (duyệt web tương tác qua Exa, không chỉ trả
            # snippet như Web Search thường) -- không phải "không có khả
            # năng tìm internet" như comment cũ giả định nhầm. Chỉ bật cho
            # "gptoss" (Qwen không hỗ trợ tool này).
            if web_search and model_choice == "gptoss":
                create_kwargs["tools"] = [{"type": "browser_search"}]
                # v-new (yêu cầu: phải thực sự tham khảo internet, không chỉ
                # "được phép"): không có tool_choice, model TỰ QUYẾT có gọi
                # browser_search hay không -- nhiều lượt nó bỏ qua luôn nếu
                # nghĩ context cục bộ đã đủ, nên mục "Nguồn (internet)" biến
                # mất. "required" ép Groq luôn chạy browser_search khi Online
                # đang bật (đúng theo ví dụ Quick Start của Groq docs).
                create_kwargs["tool_choice"] = "required"
            if on_chunk:
                # v-new (streaming): Groq/OpenAI-compatible SDK trả về 1
                # iterator các "chunk" (mỗi chunk có .choices[0].delta.content
                # là 1 mẩu text, có thể None cho các sự kiện không phải text
                # -- ví dụ tool-call events của browser_search) khi
                # stream=True.
                create_kwargs["stream"] = True
                try:
                    stream = client.chat.completions.create(**create_kwargs)
                except Exception as _tool_e:
                    if "tools" in create_kwargs:
                        print(f"[Online Chat] browser_search tool unavailable, "
                              f"retrying without it: {_tool_e}")
                        create_kwargs.pop("tools", None)
                        create_kwargs.pop("tool_choice", None)
                        stream = client.chat.completions.create(**create_kwargs)
                    else:
                        raise
                _parts = []
                for _ev in stream:
                    try:
                        _delta = _ev.choices[0].delta
                        _piece = getattr(_delta, "content", None)
                    except Exception:
                        _piece = None
                    if _piece:
                        _parts.append(_piece)
                        on_chunk(_piece)
                answer_text = "".join(_parts).strip()
                # v-note: SDK streaming events don't reliably surface
                # executed_tools (browser_search page list) the same way
                # the non-streamed response.choices[0].message does across
                # SDK versions -- best-effort trade-off: "Nguồn (internet)"
                # list is skipped ([]) on the streaming path even when
                # browser_search did run; the answer text itself is
                # unaffected (still grounded), only the extra 🌐 link list
                # is omitted this turn.
                _used_web_search = bool(web_search and model_choice == "gptoss" and "tools" in create_kwargs)
                return answer_text, None, [], _used_web_search
            try:
                resp = client.chat.completions.create(**create_kwargs)
            except Exception as _tool_e:
                if "tools" in create_kwargs:
                    # fail-soft: SDK groq trên máy quá cũ có thể chưa hỗ
                    # trợ tham số tools=/tool_choice= cho model này -- thử
                    # lại KHÔNG có tool thay vì làm hỏng cả lượt chat.
                    print(f"[Online Chat] browser_search tool unavailable, "
                          f"retrying without it: {_tool_e}")
                    create_kwargs.pop("tools", None)
                    create_kwargs.pop("tool_choice", None)
                    resp = client.chat.completions.create(**create_kwargs)
                else:
                    raise
            # v-fix (Vấn đề: mục "Nguồn (internet) 🌐" luôn rỗng dù
            # browser_search có chạy): trước đây luôn return [] cứng ở đây
            # bất kể tool có thực thi hay không -- giờ lấy URL thật từ
            # executed_tools qua _extract_groq_web_sources() (chỉ có kết
            # quả khi model_choice=="gptoss" và tool THỰC SỰ chạy, best-
            # effort/không raise).
            web_sources = _extract_groq_web_sources(resp) if (web_search and model_choice == "gptoss") else []
            _used_web_search = bool(web_search and model_choice == "gptoss" and "tools" in create_kwargs)
            return (resp.choices[0].message.content or "").strip(), None, web_sources, _used_web_search
        except Exception as e:
            return None, f"Lỗi {_groq_label} (Groq): {e}", [], False

    return None, f"Model online không hợp lệ: {model_choice}", [], False
# ══════════════════════════════════════════════════════════ (end of ONLINE AI config block)


def _call_online_ai_chat(model_choice, system_prompt, history_msgs, user_message, web_search=False, image=None,
                          on_chunk=None, on_reset=None):
    """v-new (Multi-key rotation): wrapper mỏng quanh
    _call_online_ai_chat_impl -- cùng logic như _call_online_ai() ở trên,
    riêng cho chat đa lượt: lỗi 429/quota -> xoay key -> thử lại NGAY
    cùng model, trước khi coi là rate-limit và nhảy sang model khác. Chữ
    ký/giá trị trả về giữ nguyên y hệt bản cũ."""
    provider = "gemini" if model_choice == "gemini" else ("groq" if model_choice in GROQ_MODEL_IDS else None)
    _tried_keys = 1
    while True:
        result = _call_online_ai_chat_impl(model_choice, system_prompt, history_msgs, user_message,
                                            web_search=web_search, image=image,
                                            on_chunk=on_chunk, on_reset=on_reset)
        answer, err, sources, used_web = result
        if answer is not None or not provider:
            return result
        if not _is_online_ai_rate_limit_err(str(err)):
            return result
        if not _rotate_api_key(provider):
            return result  # no more untried keys for this provider
        _tried_keys += 1
        print(f"[API Key] Retrying {model_choice} chat with key #{_tried_keys} after quota error")

# Root folder for models — relative to wherever the app itself lives
# (next to the .py / .exe), so it works the same on any machine/user
# without editing the source. Models must sit in a "models" subfolder
# next to the app; _BASE_DIR falls back to the app's own folder for
# older layouts where the model folders sit alongside the script directly.
_MODEL_ROOT_CANDIDATES = [
    os.path.join(_BASE_DIR, "models"),
    _BASE_DIR,
]

def _find_model_dir_for(model_key):
    """Auto-detect the weight folder for a model from _MODEL_ROOT_CANDIDATES."""
    info = SEMANTIC_MODELS.get(model_key, {})
    names = info.get("dir_names", [])
    candidates = []
    for root in _MODEL_ROOT_CANDIDATES:
        for n in names:
            candidates.append(os.path.join(root, n))
    for p in candidates:
        if os.path.isdir(p):
            return p
    return candidates[0] if candidates else ""  # return first candidate so load error is explicit

# Pre-resolve each model's path at startup (lazy but cached once)
SEMANTIC_MODEL_DIRS = {k: _find_model_dir_for(k) for k in SEMANTIC_MODELS}

SEMANTIC_MIN_WORDS = 3      # query needs at least N words to trigger semantic search
SEMANTIC_TOP_K     = 200    # retrieve top K semantic results — raised from 50:
                            # the full similarity computation + sort already
                            # happens over ALL documents regardless of this
                            # number (this only controls the final slice size),
                            # so raising it costs virtually nothing, but a
                            # hard cap of 50 was pushing genuinely relevant
                            # documents out of the results entirely whenever
                            # enough loosely-related files (e.g. spreadsheets
                            # with vaguely matching column headers) happened
                            # to score just high enough to fill all 50 slots
                            # first.

# v9.3: per-model threshold instead of one shared value. jina_v3 and
# bge_gemma2 are different architectures with different cosine-similarity
# score distributions -- the same cutoff can behave very differently
# between them (this is why bge_gemma2 was returning ZERO results for
# some queries where jina_v3 still returned some: bge_gemma2's best score
# for that query was likely just under 0.25). If the debug print in
# _semantic_search shows a model's best score consistently sitting just
# below its threshold, lower that model's entry here.
SEMANTIC_THRESHOLD_BY_MODEL = {
    "jina_v3":    0.25,
    # v9.3.1: lowered from 0.25 based on real observed data -- a genuinely
    # relevant result for a Vietnamese-no-diacritics query scored 0.231
    # with bge_gemma2 and was being filtered out entirely by the old shared
    # 0.25 threshold (while jina_v3 cleared it fine for the same query).
    # 0.20 leaves a little headroom below that observed score. If you see
    # bge_gemma2 results that are clearly NOT relevant sneaking in now,
    # raise this back up a bit; if good results are still being cut off,
    # check the [Semantic] debug log for the new best-score numbers and
    # lower it further.
    "bge_gemma2": 0.20,
}
SEMANTIC_THRESHOLD_DEFAULT = 0.25  # fallback for any model not listed above

# ── Currently active model for AI Search ────────────────────────────────────
# Changed via UI dropdown (see RealtimeSmartSearchApp._on_ai_model_change)
# v-new (yêu cầu: nhớ AI Search model đã chọn qua configure.ini): xem
# _config_set("Models","ai_search_model",...) trong _on_ai_model_change.
_saved_ai_search_model = _config_get("Models", "ai_search_model")
_sem_model_key = _saved_ai_search_model if _saved_ai_search_model in SEMANTIC_MODELS else DEFAULT_SEMANTIC_MODEL

_sem_model  = None
_sem_ready  = False
_sem_device = "cpu"
_sem_loaded_key = None   # model key currently loaded in _sem_model (to detect when reload is needed)

# ── Secondary model slot (v-new) ────────────────────────────────────────────
# Used ONLY for blending two models in AI Search (_semantic_search_blend).
# The primary slot above (_sem_model/_sem_loaded_key) is left completely
# untouched by this -- indexing, --update data, and the single-model search
# path all keep working exactly as before. This second slot exists so that,
# when the machine has enough spare RAM (see _dual_resident_allowed below),
# BOTH models can stay resident at once -- one on GPU, one on CPU -- instead
# of unloading/reloading a model from disk on every blended search.
_sem_model2      = None
_sem_ready2      = False
_sem_device2     = "cpu"
_sem_loaded_key2 = None

_dual_resident_checked = False   # cache the RAM check so it only runs once per app run
_dual_resident_ok      = False

# ── Vietnamese diacritics restoration (v9.4) ────────────────────────────────
# Fixes AI Search returning nothing/wrong results for Vietnamese queries typed
# without diacritics (e.g. "he thong kiem soat" instead of "hệ thống kiểm
# soát") -- jina_v3/bge_gemma2 were trained almost entirely on text WITH
# diacritics, so diacritics-less Vietnamese tokenizes into something close to
# gibberish for them, regardless of how "smart" the model otherwise is. This
# restores the diacritics on the QUERY text before it gets embedded.
#
# Model: yammdd/vietnamese-error-correction (LoRA adapter over
# vinai/bartpho-syllable, MIT license, ~0.4B params). v9.5: like jina_v3/
# bge_gemma2, this is now downloaded ON DEMAND into models/vi-diacritics/
# next to the exe (via the "Install online..." button in the Update DB
# window, using the same safe snapshot_download-into-a-folder mechanism
# already used for the embedding models), NOT baked into the exe at build
# time. An earlier version baked it into _internal, but bartpho-syllable's
# weights alone are ~1.5-2GB (it inherits mBART's huge multilingual
# vocabulary/embedding table) even after stripping the redundant TF/Flax/
# ONNX copies, which made the "small download" idea (~500MB) unrealistic to
# bake in -- moving it to the same on-demand download flow as the other
# models keeps the base exe small again and treats this as what it really
# is: an optional download, not core infrastructure.
#
# If it's not installed, diacritics restoration is just silently skipped
# (AI Search still works, just without this enhancement) -- it's not
# treated as a fatal error.
# next to the exe (via the "Install online..." button in the Update DB
# window, using the same safe snapshot_download-into-a-folder mechanism
# already used for the embedding models), NOT baked into the exe at build
# time.
#
# v9.13.4 correction: this is a single, self-contained, already-merged
# seq2seq model -- NOT a bare LoRA adapter that needs a separate base model
# (vinai/bartpho-syllable) + PEFT merge, despite what the model card text
# suggested. Confirmed by inspecting the actual downloaded files: the repo
# contains its own config.json + model.safetensors + full tokenizer, no
# adapter_config.json anywhere. This matches the user's own original
# standalone test script, which loaded it with a plain
# AutoModelForSeq2SeqLM.from_pretrained(repo_id) and worked immediately, no
# PEFT code needed. An earlier version of this file wrongly assumed it
# needed vinai/bartpho-syllable as a separate base (adding ~1.5-2GB of
# unnecessary download) and used PeftModel.from_pretrained() to merge them
# -- both are gone now. Simpler, correct, and much smaller download.
#
# If it's not installed, diacritics restoration is just silently skipped
# (AI Search still works, just without this enhancement) -- it's not
# treated as a fatal error.
_vi_dia_pipe = None
_vi_dia_ready = False
_vi_dia_load_attempted = False

VI_DIA_REPO = "yammdd/vietnamese-error-correction"
# Files we never need at runtime -- see _install_vi_diacritics_online,
# which applies the same filtering when actually downloading these.
VI_DIA_IGNORE_PATTERNS = [
    "*.h5", "tf_model*", "*.msgpack", "flax_model*", "*.onnx", "*.ot",
    "*.md", "*.png", "*.jpg",
]

def _vi_dia_local_dir():
    """Local folder this gets downloaded into -- same models/ root as
    jina_v3/bge_gemma2, NOT inside the exe/_internal."""
    return os.path.join(_MODEL_ROOT_CANDIDATES[0], "vi-diacritics")

def _vi_dia_resolve_dir(root, marker_filename):
    """Return the folder that actually CONTAINS marker_filename under root
    (searching recursively -- different huggingface_hub versions lay out
    local_dir downloads differently, flat vs. nested inside a hashed
    snapshot/.cache subfolder), or `root` itself if not found
    (from_pretrained will then raise its own clear error rather than us
    silently guessing wrong)."""
    if os.path.isfile(os.path.join(root, marker_filename)):
        return root
    for r, _, files in os.walk(root):
        if marker_filename in files:
            return r
    return root

def _vi_dia_installed():
    return os.path.isfile(os.path.join(
        _vi_dia_resolve_dir(_vi_dia_local_dir(), "config.json"), "config.json"))

def _load_vi_diacritics_model():
    """Lazily load the Vietnamese diacritics-restoration model from its
    local models/vi-diacritics/ folder. Safe to call repeatedly -- only
    does real work once. Returns True if the model is ready to use, False
    if unavailable (not installed / failed to load), in which case callers
    should just skip the restoration step rather than error out -- this is
    a nice-to-have enhancement, not a hard requirement for AI Search to
    function."""
    global _vi_dia_pipe, _vi_dia_ready, _vi_dia_load_attempted
    if _vi_dia_ready:
        return True
    if _vi_dia_load_attempted:
        return False   # already tried and failed this run -- don't retry every keystroke
    _vi_dia_load_attempted = True
    try:
        if not _vi_dia_installed():
            print(f"[VI-diacritics] Not installed (see {_vi_dia_local_dir()}) — "
                  f"restoration disabled. Install it from the Update DB window "
                  f"if you want this. AI Search still works normally otherwise.")
            return False
        from transformers import AutoTokenizer, AutoModelForSeq2SeqLM, pipeline
        model_dir = _vi_dia_resolve_dir(_vi_dia_local_dir(), "config.json")
        tok = AutoTokenizer.from_pretrained(model_dir)
        mdl = AutoModelForSeq2SeqLM.from_pretrained(model_dir)
        _vi_dia_pipe = pipeline("text2text-generation", model=mdl, tokenizer=tok)
        _vi_dia_ready = True
        print("[VI-diacritics] Model loaded OK.")
        return True
    except Exception as _e:
        print(f"[VI-diacritics] Load failed ({_e}) — restoration disabled, "
              f"AI Search still works normally otherwise.")
        return False

# Matches any character Vietnamese diacritics actually use (tone marks +
# accented vowels + đ/Đ). If a query already contains any of these, it's
# either already fully accented or not Vietnamese at all -- either way,
# running it through the restoration model is more likely to make it worse
# than better, so we only trigger restoration when NONE of these appear.
_VI_DIACRITIC_CHARS = re.compile(
    "[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợ"
    "ùúủũụưừứửữựỳýỷỹỵđ"
    "ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢ"
    "ÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]"
)

# v9.13.11 CRITICAL FIX: a Japanese query like "IRMの周波数はどのように
# 決まっている？" contains zero Vietnamese diacritic characters AND does
# contain at least one Latin letter ("IRM") -- both old checks passed, so
# the WHOLE Japanese sentence got fed into the Vietnamese diacritics model
# and came back as garbage ("IRMN Chu Ba?"), which then got embedded
# instead of the real query, wrecking both Jina (wrong results) and
# BGE (0 results, since the garbled text scored even lower than usual).
# "No Vietnamese diacritics" is nowhere near enough to conclude "this is
# Vietnamese text missing its diacritics" -- CJK/Korean/Thai/Arabic/etc.
# text ALSO has no Vietnamese diacritics, for the obvious reason that it
# isn't Vietnamese at all. Explicitly bail out if the query contains any
# character from a non-Latin script -- Vietnamese (with or without
# diacritics) is always written in Latin script, so this is a safe,
# conservative filter with no false-negatives for real Vietnamese input.
_NON_LATIN_SCRIPT_CHARS = re.compile(
    "["
    "\u3040-\u30ff"    # Hiragana + Katakana
    "\u3400-\u4dbf"    # CJK Extension A
    "\u4e00-\u9fff"    # CJK Unified Ideographs (Kanji/Hanzi)
    "\uf900-\ufaff"    # CJK Compatibility Ideographs
    "\uac00-\ud7af"    # Hangul (Korean)
    "\u0e00-\u0e7f"    # Thai
    "\u0600-\u06ff"    # Arabic
    "\u0750-\u077f"    # Arabic Supplement
    "\u0590-\u05ff"    # Hebrew
    "\u0400-\u04ff"    # Cyrillic
    "]"
)

def _maybe_restore_diacritics(query):
    """If `query` looks like it might be Vietnamese typed without diacritics
    (no diacritic characters at all, but does contain letters, AND isn't
    written in some other non-Latin script entirely), try restoring it via
    the bundled model. Always returns a usable query -- falls back to the
    original text unchanged on any failure/skip, so this can never make
    search worse than not calling it at all, only better or neutral."""
    if not query or not query.strip():
        return query
    if _NON_LATIN_SCRIPT_CHARS.search(query):
        return query   # Japanese/Chinese/Korean/Thai/Arabic/Hebrew/Cyrillic
                        # text -- definitely not Vietnamese, don't touch it
    if _VI_DIACRITIC_CHARS.search(query):
        return query   # already has diacritics -- leave it alone
    if not re.search(r"[a-zA-Z]", query):
        return query   # no letters at all (e.g. pure numbers/symbols) -- nothing to restore
    if not _load_vi_diacritics_model():
        return query   # model unavailable -- silent no-op, not an error
    try:
        restored = _vi_dia_pipe(query, max_new_tokens=128)[0]["generated_text"]
        if restored and restored.strip():
            print(f"[VI-diacritics] '{query}' -> '{restored}'")
            return restored
    except Exception as _e:
        print(f"[VI-diacritics] Restoration failed ({_e}), using original query.")
    return query

def _find_model_dir(base_dir):
    """Walk up to 4 subdirectory levels to find config.json (HuggingFace model)."""
    if base_dir and os.path.isfile(os.path.join(base_dir, "config.json")):
        return base_dir
    if not base_dir or not os.path.isdir(base_dir):
        return base_dir
    for root, dirs, files in os.walk(base_dir):
        depth = root.replace(base_dir, "").count(os.sep)
        if depth > 4: break
        if "config.json" in files:
            return root
    return base_dir

def _semantic_table_for(model_key=None):
    """Return the embedding table name for the current (or specified) model."""
    k = model_key if model_key is not None else _sem_model_key
    return SEMANTIC_MODELS.get(k, SEMANTIC_MODELS[DEFAULT_SEMANTIC_MODEL])["table"]

def _encode_query_with(model_obj, text, model_key):
    """Same encode-convention logic as _encode_query, but takes the model
    object explicitly instead of always reading the global _sem_model —
    lets the secondary (blend) slot reuse the exact same per-model
    conventions without duplicating them."""
    if model_key == "jina_v3":
        return model_obj.encode([text], task="retrieval.query",
                                 convert_to_numpy=True, normalize_embeddings=True)[0]
    if model_key == "bge_gemma2":
        instruction = "Given a web search query, retrieve relevant passages that answer the query."
        return model_obj.encode([text], prompt=instruction,
                                 convert_to_numpy=True, normalize_embeddings=True)[0]
    return model_obj.encode([text], convert_to_numpy=True, normalize_embeddings=True)[0]

def _encode_query(text, model_key):
    """Encode a single query string using the correct convention for each model."""
    return _encode_query_with(_sem_model, text, model_key)


def _encode_passages(texts, model_key, batch_size=32):
    """Encode a list of passages/documents using the correct convention for each model."""
    if model_key == "jina_v3":
        return _sem_model.encode(texts, task="retrieval.passage", batch_size=batch_size,
                                  convert_to_numpy=True, normalize_embeddings=True,
                                  show_progress_bar=False)
    if model_key == "bge_gemma2":
        # bge-gemma2 does not need instruction for passages, only for queries
        return _sem_model.encode(texts, batch_size=batch_size, convert_to_numpy=True,
                                  normalize_embeddings=True, show_progress_bar=False)
    return _sem_model.encode(texts, batch_size=batch_size, convert_to_numpy=True,
                              normalize_embeddings=True, show_progress_bar=False)

# v9.0: Removed the entire runtime pip-install / self-heal / self-relaunch
# system that used to live here (_bootstrap_embeddable_python,
# _is_real_python, _find_pip_python, _pip_install, _heal_into_shadow_dir,
# _spawn_healer_and_exit). AI libraries (torch/transformers/
# sentence-transformers/einops) are now baked into the exe at BUILD time
# (see requirements.txt + the GitHub Actions workflow), not downloaded at
# runtime on the user's machine. This was necessary because self-
# relaunching / spawning hidden child processes / downloading GB-scale
# files at runtime kept getting silently blocked on locked-down corporate
# machines (privilege management / EDR software), with no usable error
# message. A plain `import sentence_transformers` below now just works,
# same as importing any other bundled dependency.


def _check_semantic_table_count(model_key):
    """Return how many embedding rows exist for a given model's table (0 if
    the table doesn't exist yet / --update data was never run for it).
    Module-level so both the (now-hidden) model dropdown code and the
    auto-blend search below share one source of truth instead of two
    copies of the same SQLite query drifting apart."""
    try:
        table = SEMANTIC_MODELS.get(model_key, {}).get("table", "semantic_index")
        conn = sqlite3.connect(DB_FILE, timeout=5)
        c = conn.cursor()
        c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table,))
        has_table = c.fetchone()[0] > 0
        count = 0
        if has_table:
            c.execute(f"SELECT count(*) FROM {table}")
            count = c.fetchone()[0]
        conn.close()
        return count
    except Exception:
        return 0


def _get_free_vram_mb():
    """Free GPU VRAM right now, in MB. 0 if no CUDA GPU available."""
    try:
        import torch
        if not torch.cuda.is_available():
            return 0
        free_b, _total_b = torch.cuda.mem_get_info()
        return free_b / (1024 * 1024)
    except Exception:
        return 0

def _get_free_ram_mb():
    """Free system RAM right now, in MB. Uses ctypes on Windows (no new
    bundled dependency required, unlike psutil which isn't currently baked
    into the exe build) with a psutil fallback if it happens to be present.
    Returns 0 ("unknown") if neither works -- callers treat 0 as "not
    enough", which is the safe default (falls back to swap-per-search
    instead of risking the OS paging under memory pressure)."""
    try:
        import ctypes
        class _MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        stat = _MEMORYSTATUSEX()
        stat.dwLength = ctypes.sizeof(_MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
        return stat.ullAvailPhys / (1024 * 1024)
    except Exception:
        try:
            import psutil
            return psutil.virtual_memory().available / (1024 * 1024)
        except Exception:
            return 0

def _pick_device_for_model(model_key):
    """v-new: replaces the old hardcoded force_cpu flag with a runtime
    decision based on ACTUALLY MEASURED free VRAM vs. this model's
    estimated footprint (SEMANTIC_MODELS[...]["est_vram_mb"]). Practical
    effect on a small laptop GPU (e.g. the T1200's ~4GB) is the same as
    force_cpu used to be for bge_gemma2 -- but on a machine with a bigger
    GPU, bge_gemma2 would automatically get to run on GPU too, with zero
    code changes needed. `force_cpu: True` on a model's config (if set)
    still wins as a manual pin, for when you know better than the estimate."""
    info = SEMANTIC_MODELS.get(model_key, {})
    if info.get("force_cpu"):
        return "cpu"
    try:
        import torch
        if not torch.cuda.is_available():
            return "cpu"
    except Exception:
        return "cpu"
    need_vram = info.get("est_vram_mb", 2000)
    free_vram = _get_free_vram_mb()
    SAFETY_MARGIN_MB = 800
    if free_vram >= need_vram + SAFETY_MARGIN_MB:
        return "cuda"
    print(f"[Semantic] {model_key}: GPU has {free_vram:.0f}MB free, needs "
          f"~{need_vram}MB (+{SAFETY_MARGIN_MB}MB margin) -> using CPU instead")
    return "cpu"

def _dual_resident_allowed():
    """One-time (cached) check of whether this machine has enough spare RAM
    to keep BOTH semantic models loaded simultaneously, so AI-Search blend
    can encode with both models in parallel instead of swapping one out of
    memory and reloading it from disk on every blended search. Deliberately
    conservative: if unsure, returns False and blend falls back to the
    slower (but always-safe) swap-per-search path."""
    global _dual_resident_checked, _dual_resident_ok
    if _dual_resident_checked:
        return _dual_resident_ok
    _dual_resident_checked = True
    free_ram = _get_free_ram_mb()
    need_ram = sum(SEMANTIC_MODELS[k].get("est_ram_mb", 2000) for k in SEMANTIC_MODELS)
    SAFETY_MARGIN_MB = 4000  # headroom for OS + this app + its SQLite DB + everything else
    _dual_resident_ok = free_ram > 0 and free_ram >= (need_ram + SAFETY_MARGIN_MB)
    if free_ram <= 0:
        print("[Semantic] Dual-resident check: could not measure free RAM -> "
              "playing safe, will swap models per search instead of keeping both loaded.")
    else:
        verdict = "OK, keeping both models loaded" if _dual_resident_ok else "not enough, will swap per search"
        print(f"[Semantic] Dual-resident check: free RAM={free_ram:.0f}MB, "
              f"need~{need_ram}MB (+{SAFETY_MARGIN_MB}MB margin) -> {verdict}")
    return _dual_resident_ok


def _load_semantic_model(model_key=None):
    """Load embedding model by key ('jina_v3'/'bge_gemma2'). Returns immediately
    if the correct model is already loaded. If the user switched to a different model,
    unloads the old model first to avoid holding two heavy models in RAM."""
    global _sem_model, _sem_ready, _sem_device, _sem_loaded_key, _sem_model_key
    want_key = model_key if model_key is not None else _sem_model_key
    if want_key not in SEMANTIC_MODELS:
        want_key = DEFAULT_SEMANTIC_MODEL
    if _sem_ready and _sem_loaded_key == want_key:
        return True
    try:
        # v9.0: sentence_transformers/transformers/torch/einops are baked
        # into the exe at build time (see requirements.txt, pinned to
        # transformers<5 there specifically because jina_v3's custom
        # trust_remote_code modeling file, XLMRobertaLoRA, is written
        # against the transformers v4.x tied-weights API and crashes with
        # "'XLMRobertaLoRA' object has no attribute 'all_tied_weights_keys'"
        # on transformers v5+). No install/version-check/self-heal needed
        # here anymore -- if this import ever fails, it means the exe was
        # built wrong (missing dependency at build time), not something
        # this running process can fix on its own; see the error message
        # below for what to do in that case.
        try:
            import torch
            from sentence_transformers import SentenceTransformer
            import importlib
            if SEMANTIC_MODELS.get(want_key, {}).get("trust_remote"):
                importlib.import_module("einops")
        except Exception as _imp_e:
            print(f"[Semantic] Thu vien AI bi thieu hoac loi ({_imp_e}). "
                  f"Day la loi build, khong phai loi cau hinh may nay -- "
                  f"ban can build lai exe voi requirements.txt day du roi "
                  f"phat hanh lai, khong the tu sua trong luc chay.")
            raise

        _cuda_ok = torch.cuda.is_available()
        _sem_device = _pick_device_for_model(want_key)
        if _sem_device == "cuda":
            print(f"[Semantic] GPU: {torch.cuda.get_device_name(0)}")
        else:
            print("[Semantic] No CUDA GPU, using CPU" if not _cuda_ok else f"[Semantic] {want_key}: using CPU")
        # jina-v3 / bge-gemma2 are heavy on CPU — still allowed, just slower
        if want_key == "bge_gemma2" and _sem_device == "cpu":
            print(f"[Semantic] Note: {want_key} running on CPU this session — slower than GPU but stable.")

        # Free old model before loading new one (saves RAM, especially with bge-gemma2)
        if _sem_model is not None:
            try:
                del _sem_model
                _sem_model = None
                import gc; gc.collect()
                try:
                    if torch.cuda.is_available():
                        torch.cuda.empty_cache()
                except Exception:
                    pass
            except Exception:
                pass

        model_info = SEMANTIC_MODELS[want_key]
        model_base_dir = SEMANTIC_MODEL_DIRS.get(want_key, "")
        model_path = _find_model_dir(model_base_dir)
        print(f"[Semantic] Loading model ({want_key}) from: {model_path}  device={_sem_device}")
        if not model_path or not os.path.isfile(os.path.join(model_path, "config.json")):
            raise FileNotFoundError(f"config.json not found for '{want_key}' in: {model_path}")
        load_kwargs = {"device": _sem_device}
        if model_info.get("trust_remote"):
            load_kwargs["trust_remote_code"] = True
        _sem_model = SentenceTransformer(model_path, **load_kwargs)
        _sem_ready = True
        _sem_loaded_key = want_key
        _sem_model_key = want_key
        print(f"[Semantic] Model ({want_key}) loaded OK on {_sem_device.upper()}")
        return True
    except Exception as _e:
        print(f"[Semantic] Load failed: {_e}")
        _sem_ready = False
        # v3.4: a failed load (e.g. missing trust_remote_code file, OOM mid-init)
        # can leave partially-allocated GPU tensors behind even though _sem_model
        # was never assigned. Left uncleaned, this "leaks" VRAM into the NEXT
        # model's load attempt (e.g. jina_v3 fails -> bge_gemma2 then OOMs on a
        # GPU that should have had plenty of free memory). Force-clear here too.
        _sem_model = None
        _sem_loaded_key = None
        try:
            import gc; gc.collect()
            import torch
            if torch.cuda.is_available():
                torch.cuda.empty_cache()
                torch.cuda.ipc_collect()
        except Exception:
            pass
        return False
# ─────────────────────────────────────────────────────────────────────────────


def _load_semantic_model_secondary(model_key):
    """v-new: loads a model into the SECOND slot (_sem_model2), used only by
    _semantic_search_blend when _dual_resident_allowed() says there's
    enough spare RAM to keep both models loaded at once. Deliberately does
    NOT touch the primary slot (_sem_model/_sem_loaded_key) at all -- every
    other existing code path (indexing, --update data, single-model
    search, the dropdown-driven flow) keeps working exactly as before,
    completely unaware this second slot exists."""
    global _sem_model2, _sem_ready2, _sem_device2, _sem_loaded_key2
    if model_key not in SEMANTIC_MODELS:
        return False
    if _sem_ready2 and _sem_loaded_key2 == model_key:
        return True
    try:
        import torch
        from sentence_transformers import SentenceTransformer
        import importlib
        if SEMANTIC_MODELS.get(model_key, {}).get("trust_remote"):
            importlib.import_module("einops")

        _sem_device2 = _pick_device_for_model(model_key)
        print(f"[Semantic] (secondary slot) Loading model ({model_key}) device={_sem_device2}")

        if _sem_model2 is not None:
            try:
                del _sem_model2
                _sem_model2 = None
                import gc; gc.collect()
                if torch.cuda.is_available():
                    torch.cuda.empty_cache()
            except Exception:
                pass

        model_info = SEMANTIC_MODELS[model_key]
        model_path = _find_model_dir(SEMANTIC_MODEL_DIRS.get(model_key, ""))
        if not model_path or not os.path.isfile(os.path.join(model_path, "config.json")):
            raise FileNotFoundError(f"config.json not found for '{model_key}' in: {model_path}")
        load_kwargs = {"device": _sem_device2}
        if model_info.get("trust_remote"):
            load_kwargs["trust_remote_code"] = True
        _sem_model2 = SentenceTransformer(model_path, **load_kwargs)
        _sem_ready2 = True
        _sem_loaded_key2 = model_key
        print(f"[Semantic] (secondary slot) Model ({model_key}) loaded OK on {_sem_device2.upper()}")
        return True
    except Exception as _e:
        print(f"[Semantic] (secondary slot) Load failed: {_e}")
        _sem_ready2 = False
        _sem_model2 = None
        _sem_loaded_key2 = None
        return False
# ─────────────────────────────────────────────────────────────────────────────



def is_text_file(filepath):
    if not filepath: return False
    if os.path.splitext(filepath)[1].lower() in BINARY_EXT: return False
    try:
        with open(filepath, 'rb') as f:
            chunk = f.read(1024)
            if b'\0' in chunk: return False
        return True
    except: return False

# ── OCR (v9.13) ──────────────────────────────────────────────────────────
# Optional, off by default (see the "OCR images" checkbox in Update DB).
# Images have no literal embedded text -- unlike .txt/.py (plain read) or
# .docx/.pdf (format-specific parser), a screenshot/scan needs an actual
# computer-vision model to "read" the pixels. EasyOCR is used here (not
# Tesseract) because it needs no separate native binary installed
# alongside Python -- just a pip package -- and has solid out-of-the-box
# support for Japanese + Vietnamese + English together, matching this
# app's document mix. It downloads its own model weights (~a few hundred
# MB per language) on first use.
_OCR_IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.tiff', '.tif'}
# v-fix (BUG: file .pdf/.docx/.xlsx/... đôi khi bị đọc như text file thô,
# ra toàn ký tự rác của header nhị phân -- xem comment chi tiết trong
# get_file_content()): mọi đuôi file ĐÃ có handler chuyên dụng riêng trong
# get_file_content() liệt kê ở đây, để is_text_file() không bao giờ được
# dùng nhầm cho chúng nữa, bất kể 1024 byte đầu của file đó có tình cờ
# không chứa byte \0 hay không.
_DEDICATED_HANDLER_EXTS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.csv',
    '.pptx', '.ppt', '.one', '.msg',
} | _OCR_IMAGE_EXTS
# v9.16: non-image extensions get_file_content() already knows how to read —
# used by the "Search files" button (formerly "Search image") to widen its
# file picker beyond pictures now that it's no longer OCR-only.
_SEARCH_FILES_DOC_EXTS = {'.txt', '.log', '.ini', '.json', '.xml', '.md',
                          '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.csv',
                          '.pptx', '.ppt', '.one', '.msg'}
OCR_ENABLED = False   # set from the Update DB dialog's "OCR images" checkbox
_ocr_readers = {}          # {'ja_en': Reader, 'vi_en': Reader} once loaded
_ocr_load_attempted = False

def _load_ocr_readers():
    """Lazily load both EasyOCR readers. Safe to call repeatedly -- only
    does real work once. Returns True if at least one reader loaded OK."""
    global _ocr_load_attempted
    if _ocr_readers:
        return True
    if _ocr_load_attempted:
        return False
    _ocr_load_attempted = True
    try:
        import easyocr
        # v9.13.6 fix: use GPU if available -- this machine already runs
        # jina_v3 on CUDA successfully, but OCR was hardcoded to CPU-only
        # (gpu=False), which was the main cause of ~30-60s per image. Try
        # GPU first, fall back to CPU only if that genuinely fails (no
        # GPU, out of VRAM, etc.) rather than assuming CPU-only up front.
        _gpu_ok = False
        try:
            import torch
            _gpu_ok = torch.cuda.is_available()
        except Exception:
            pass
        print(f"[OCR] Loading EasyOCR readers on {'GPU' if _gpu_ok else 'CPU'} "
              f"(first run downloads model weights)...")
        try:
            _ocr_readers['ja_en'] = easyocr.Reader(['ja', 'en'], gpu=_gpu_ok)
            _ocr_readers['vi_en'] = easyocr.Reader(['vi', 'en'], gpu=_gpu_ok)
        except Exception as _e:
            if _gpu_ok:
                print(f"[OCR] GPU init failed ({_e}), retrying on CPU...")
                _ocr_readers['ja_en'] = easyocr.Reader(['ja', 'en'], gpu=False)
                _ocr_readers['vi_en'] = easyocr.Reader(['vi', 'en'], gpu=False)
            else:
                raise
        print("[OCR] EasyOCR readers loaded OK.")
        return True
    except Exception as _e:
        print(f"[OCR] Failed to load EasyOCR ({_e}) — OCR disabled for this run, "
              f"indexing continues normally for everything else.")
        return False

_OCR_MAX_DIM = 1800   # v9.13.6: cap the longer side at this many pixels before
                      # OCR — most screenshots/scans don't need more than this
                      # to read text reliably, and large images (e.g. a
                      # 4000px-wide photo) were a real contributor to slow
                      # OCR times regardless of CPU/GPU.

def _run_ocr(filepath):
    """Try both JA/EN and VI/EN readers on `filepath`, keep whichever gives
    the higher average confidence — lets one code path handle screenshots
    in either language without knowing ahead of time which one a given
    image is in. Returns "" (not an error) on any failure — a bad/corrupt
    image should never stop the rest of indexing."""
    if not _load_ocr_readers():
        return ""
    try:
        from PIL import Image
        import numpy as _np
        img = Image.open(filepath).convert("RGB")
        w, h = img.size
        if max(w, h) > _OCR_MAX_DIM:
            scale = _OCR_MAX_DIM / max(w, h)
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        ocr_input = _np.array(img)   # decode+resize once, reuse for both readers below
    except Exception as _e:
        print(f"[OCR] Could not preprocess {filepath} ({_e}), falling back to raw file path.")
        ocr_input = filepath
    best_text, best_conf = "", -1.0
    for reader in _ocr_readers.values():
        try:
            results = reader.readtext(ocr_input, detail=1)
            if not results:
                continue
            avg_conf = sum(r[2] for r in results) / len(results)
            if avg_conf > best_conf:
                best_conf = avg_conf
                best_text = " ".join(r[1] for r in results)
        except Exception as _e:
            print(f"[OCR] reader failed on {filepath}: {_e}")
    return best_text

def get_file_icon(filepath, is_folder=False):
    if is_folder: return "📁 "
    if not filepath: return "📎 "
    try:
        ext = os.path.splitext(str(filepath))[1].lower()
        if ext == ".pdf": return "📕 "
        elif ext in [".doc", ".docx"]: return "📝 "
        elif ext in [".xls", ".xlsx", ".csv"]: return "📊 "
        elif ext in [".ppt", ".pptx"]: return "📉 "
        elif ext == ".one": return "🔮 "
        elif ext in [".txt", ".log", ".ini", ".json", ".xml"]: return "📄 "
        elif ext == ".msg": return "📧 "
    except: pass
    return "📎 "

# ─────────────────────────────────────────────────────────────────────────
# v4.9: REAL Windows shell icons (same icons Explorer/Everything show) ──
# Extracted once per file-extension via the Windows Shell API and cached
# as Tk PhotoImage objects, so the (slow-ish) icon lookup only happens once
# per unique extension for the whole app session, not once per row.
#
# Requires on Windows:  pip install pywin32 pillow
# If either package is missing, or we're not running on Windows, every
# lookup silently returns None and callers fall back to the old emoji
# icons -- the app keeps working either way, it just looks plainer.
# ─────────────────────────────────────────────────────────────────────────
_ICON_PHOTO_CACHE = {}   # key: ext (".pdf") or "__folder__" / "__file__" -> ImageTk.PhotoImage
_ICON_BACKEND_OK = False
_com_local = threading.local()   # tracks whether CoInitializeEx has run on THIS thread
if sys.platform == "win32":
    try:
        import ctypes
        from ctypes import wintypes
        import win32gui
        import win32ui
        import win32con
        from PIL import Image, ImageTk
        _ICON_BACKEND_OK = True

        class _SHFILEINFOW(ctypes.Structure):
            _fields_ = [
                ("hIcon", wintypes.HANDLE),
                ("iIcon", ctypes.c_int),
                ("dwAttributes", wintypes.DWORD),
                ("szDisplayName", ctypes.c_wchar * 260),
                ("szTypeName", ctypes.c_wchar * 80),
            ]

        _SHGFI_ICON = 0x000000100
        _SHGFI_SMALLICON = 0x000000001
        _SHGFI_USEFILEATTRIBUTES = 0x000000010
        _FILE_ATTRIBUTE_NORMAL = 0x80
        _FILE_ATTRIBUTE_DIRECTORY = 0x10

        class _BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD),
            ]
        _DI_NORMAL = 0x0003

        def _hicon_to_photoimage(hicon, size=16):
            """Convert a Win32 HICON handle to a Tk PhotoImage (real RGBA alpha).
            v5.2 fix: the old version used win32ui's high-level DrawIcon() into a
            device-dependent bitmap (CreateCompatibleBitmap). Two bugs came from
            that: (1) a DDB has no real per-pixel alpha channel, so reading it
            back as RGBA gave alpha=0 almost everywhere -- icons rendered flat
            and washed-out ("no color"). (2) DrawIcon() paints the icon at its
            *native* resolution with no stretching, so on any DPI scale where
            the system small-icon isn't exactly `size` px, the icon came out
            cropped/misaligned ("wrong size" vs Explorer/Everything).
            Fix: draw into a real 32-bit top-down DIB section via DrawIconEx,
            which both stretches to the exact requested size AND preserves the
            icon's real per-pixel alpha (needed for colored folder/app icons)."""
            try:
                hdc_screen = win32gui.GetDC(0)
                hdc_mem = win32gui.CreateCompatibleDC(hdc_screen)
                bmi = _BITMAPINFOHEADER()
                bmi.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
                bmi.biWidth = size
                bmi.biHeight = -size          # negative = top-down (matches PIL row order)
                bmi.biPlanes = 1
                bmi.biBitCount = 32
                bmi.biCompression = 0         # BI_RGB
                ppv_bits = ctypes.c_void_p()
                hbmp = ctypes.windll.gdi32.CreateDIBSection(
                    hdc_mem, ctypes.byref(bmi), 0, ctypes.byref(ppv_bits), None, 0
                )
                if not hbmp or not ppv_bits.value:
                    win32gui.DeleteDC(hdc_mem); win32gui.ReleaseDC(0, hdc_screen)
                    return None
                old_obj = win32gui.SelectObject(hdc_mem, hbmp)
                # DI_NORMAL draws mask+color honoring per-pixel alpha on 32-bit
                # icons, and (unlike DrawIcon) stretches to exactly size x size.
                ctypes.windll.user32.DrawIconEx(
                    hdc_mem, 0, 0, hicon, size, size, 0, None, _DI_NORMAL
                )
                buf = ctypes.string_at(ppv_bits, size * size * 4)
                img = Image.frombuffer("RGBA", (size, size), buf, "raw", "BGRA", 0, 1)
                win32gui.SelectObject(hdc_mem, old_obj)
                win32gui.DeleteObject(hbmp)
                win32gui.DeleteDC(hdc_mem)
                win32gui.ReleaseDC(0, hdc_screen)
                return ImageTk.PhotoImage(img)
            except Exception:
                return None

        def _ensure_com_initialized():
            """SHGetFileInfoW drives the Shell's icon machinery, which needs
            COM initialized (CoInitialize) on whichever thread calls it. The
            Tk main thread usually already has COM up from Tkinter/pywin32
            startup -- but real-time search inserts rows from background
            worker threads that never call CoInitialize, so the FIRST icon
            lookup from each new thread silently failed. The previous fix
            "only extract icons on the main thread" was treating that
            symptom, not the cause -- it made background-thread rows never
            get real icons at all (only Advanced/AI Search, which happen to
            run on the main thread, showed them). The actual fix: initialize
            COM once per thread, lazily, right before that thread's first
            Shell icon call -- then let every thread extract icons freely.
            Safe to call repeatedly: it's a no-op after the first successful
            call on a given thread."""
            if getattr(_com_local, "initialized", False):
                return
            try:
                # COINIT_APARTMENTTHREADED (0x2) -- Shell icon APIs are
                # documented as STA-only. RPC_E_CHANGED_MODE / S_FALSE here
                # just mean this thread already has COM up in some form
                # (e.g. the main thread, via Tkinter/pywin32) -- harmless,
                # the Shell call still works either way.
                ctypes.windll.ole32.CoInitializeEx(None, 0x2)
            except Exception:
                pass
            _com_local.initialized = True

        def _extract_shell_icon(path, is_folder, size=16):
            """Ask the Windows Shell for the icon it would show in Explorer for
            this file/folder, using the real path when it exists on disk so we
            get the exact registered app icon (Word/Excel/PDF reader/etc.)."""
            shfi = _SHFILEINFOW()
            flags = _SHGFI_ICON | _SHGFI_SMALLICON
            use_path = path if (path and os.path.exists(path)) else None
            if use_path:
                target = use_path
                attr = 0
            else:
                # v-fix (icon Outlook/OneNote thỉnh thoảng hiện icon "màn
                # hình"/file lạ thay vì icon .msg thật): SHGFI_USEFILEATTRIBUTES
                # chỉ dựa vào PHẦN ĐUÔI (.msg/.one) của `target` để chọn icon
                # -- toàn bộ phần còn lại của chuỗi không quan trọng. Nhưng
                # pseudo-path Outlook/OneNote (xem _make_outlook_pseudo_path/
                # _make_onenote_pseudo_path) nhúng nguyên entry_id + store_id +
                # folder_path + subject vào path -- với email có StoreID dài
                # (Exchange) và/hoặc subject dài, tổng chuỗi dễ dàng vượt quá
                # MAX_PATH (260 ký tự) của Windows. SHGetFileInfoW nhận 1
                # chuỗi quá dài như vậy có hành vi KHÔNG ổn định -- đôi khi
                # âm thầm trả về icon "unknown file type" chung chung (nhìn
                # giống icon màn hình/trang trắng) thay vì icon .msg thật.
                # Vì chỉ có đuôi file là quan trọng ở nhánh này, luôn dùng 1
                # placeholder NGẮN cùng đuôi thay vì path gốc (có thể dài
                # 300+ ký tự) để tra cứu luôn ổn định, bất kể subject/ID dài
                # cỡ nào.
                flags |= _SHGFI_USEFILEATTRIBUTES
                attr = _FILE_ATTRIBUTE_DIRECTORY if is_folder else _FILE_ATTRIBUTE_NORMAL
                if is_folder:
                    target = "folder"
                else:
                    _ext = os.path.splitext(str(path))[1] if path else ""
                    target = f"placeholder{_ext}" if _ext else "file.txt"
            res = ctypes.windll.shell32.SHGetFileInfoW(
                target, attr, ctypes.byref(shfi), ctypes.sizeof(shfi), flags
            )
            if not res or not shfi.hIcon:
                return None
            photo = _hicon_to_photoimage(shfi.hIcon, size=size)
            win32gui.DestroyIcon(shfi.hIcon)
            return photo
    except Exception as _icon_backend_err:
        _ICON_BACKEND_OK = False
        print(f"[icon] Real Windows icons NOT available -- falling back to emoji icons. Reason: {_icon_backend_err}")
        print("[icon] Fix: run  pip install pywin32 pillow  then restart the app.")
else:
    print("[icon] Real Windows icons only work on Windows -- falling back to emoji icons.")



def get_tree_icon_image(filepath, is_folder=False, size=16):
    """Return a cached ImageTk.PhotoImage with the REAL Windows icon for this
    file/folder (same icon Explorer/Everything display), or None if the
    Shell icon backend isn't available (non-Windows, or pywin32/Pillow not
    installed) -- callers should fall back to get_file_icon() emoji text."""
    if not _ICON_BACKEND_OK:
        return None
    if is_folder:
        key = "__folder__"
    else:
        key = os.path.splitext(str(filepath))[1].lower() or "__file__"
    if key in _ICON_PHOTO_CACHE:
        return _ICON_PHOTO_CACHE[key]
    # v5.6 fix: previously restricted to the main thread only, as a
    # workaround for a failure that was actually caused by missing
    # per-thread COM init (see _ensure_com_initialized() above), not by
    # which thread was calling. That restriction was the reason icons never
    # showed up in the File Name / Folder Name / File Content tabs on first
    # load (real-time search fills those from background threads) while
    # Advanced/AI Search -- which run on the main thread -- looked fine.
    # Now every thread initializes its own COM apartment once, lazily, and
    # is free to extract icons.
    if _ICON_BACKEND_OK:
        _ensure_com_initialized()
    try:
        photo = _extract_shell_icon(filepath, is_folder, size=size)
    except Exception as _e:
        photo = None
        if not getattr(get_tree_icon_image, "_warned", False):
            get_tree_icon_image._warned = True
            print(f"[icon] Icon extraction failed for '{filepath}': {_e}")
    if photo is None:
        if not getattr(get_tree_icon_image, "_warned_none", False):
            get_tree_icon_image._warned_none = True
            print(f"[icon] SHGetFileInfoW returned no icon for '{filepath}' (is_folder={is_folder}) -- check the file/folder actually exists on disk.")
        return None  # do NOT cache the miss -- retry next call
    _ICON_PHOTO_CACHE[key] = photo
    return photo

def format_size(bytes_val):
    """Fix 0KB display bug — convert bytes to human-readable unit accurately."""
    if bytes_val is None: return ""
    try:
        b_val = float(bytes_val)
    except:
        return ""
    if b_val <= 0: return "0 B"
    for unit in ['B', 'KB', 'MB', 'GB']:
        if b_val < 1024.0:
            return f"{b_val:.1f} {unit}".replace(".0 ", " ")
        b_val /= 1024.0
    return f"{b_val:.1f} TB"

def get_live_size(filepath, db_fallback=0):
    """Get the real file size directly from disk.
    Falls back to the DB value if the file is offline/cloud/deleted."""
    try:
        if filepath and os.path.isfile(filepath):
            real = os.path.getsize(filepath)
            return real  # Return actual size, including genuine 0-byte files
    except Exception:
        pass
    return db_fallback

def get_file_type(filepath):
    """Return file extension as type label, e.g. 'PDF', 'DOCX', 'Folder'"""
    if not filepath: return ""
    if os.path.isdir(filepath): return "Folder"
    ext = os.path.splitext(str(filepath))[1]
    return ext.lstrip(".").upper() if ext else ""

def get_live_mtime(filepath):
    """Get file modification time as formatted string. Returns '' on failure."""
    try:
        if filepath and os.path.exists(filepath):
            ts = os.path.getmtime(filepath)
            return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
    except Exception:
        pass
    return ""

def format_meta_datetime(raw):
    """v-fix: Outlook/OneNote metadata (received/last_modified/created)
    comes back from the Graph/COM APIs as raw ISO-8601 UTC strings, e.g.
    "2022-02-15T07:15:36.000Z" -- shown as-is in the tree's Modified time
    column, wildly inconsistent with every real file's "%Y-%m-%d %H:%M"
    (see get_live_mtime above). Reformat to the same local, human style;
    fall back to the raw string untouched if parsing fails for any reason
    (better a slightly-ugly date than a silently blank column)."""
    if not raw:
        return ""
    try:
        s = raw.strip()
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is not None:
            dt = dt.astimezone().replace(tzinfo=None)  # UTC -> local wall-clock time
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return raw

def parse_size_filter(query_str):
    """Fix: strip extra whitespace that would leave the SQL LIKE condition empty."""
    pattern = r'(?:size\s*)?(>=|<=|>|<|=)\s*([0-9.]+)\s*(B|KB|MB|GB)?'
    match = re.search(pattern, query_str, re.IGNORECASE)
    if not match:
        return query_str.strip(), None, None
    
    op, val_str, unit = match.groups()
    try:
        val = float(val_str)
    except:
        return query_str.strip(), None, None
        
    unit = (unit or 'B').upper()
    multiplier = {'B': 1, 'KB': 1024, 'MB': 1024**2, 'GB': 1024**3}.get(unit, 1)
    bytes_val = int(val * multiplier)
    
    # Clean the text after stripping the operator, collapse all extra whitespace
    cleaned_query = re.sub(pattern, '', query_str, flags=re.IGNORECASE).strip()
    cleaned_query = " ".join(cleaned_query.split())
    
    return cleaned_query, op, bytes_val

# ── CJK-aware keyword anchor helpers ─────────────────────────────────────────
# The search logic requires an "anchor" keyword of len(k) >= 3 before it will
# run a filename/content query, to avoid noise from short English words/stop
# words ('a', 'of', ...). But CJK text has no spaces, so a whole query like
# "解析" (2 characters, a complete/specific word meaning "analysis") becomes
# ONE token of length 2 and used to get silently rejected — while "解析条件"
# (4 chars) passed and returned results. Unlike English, 1-2 CJK characters
# are often already a complete, specific, meaningful word, so we treat any
# token containing CJK characters as an anchor regardless of length.
def _contains_cjk(s):
    for ch in s:
        cp = ord(ch)
        if (0x3040 <= cp <= 0x30FF   # Hiragana + Katakana
                or 0x3400 <= cp <= 0x9FFF   # CJK Unified Ideographs (+ Ext A)
                or 0xF900 <= cp <= 0xFAFF   # CJK Compatibility Ideographs
                or 0xAC00 <= cp <= 0xD7A3   # Hangul syllables
                or 0xFF00 <= cp <= 0xFFEF): # Fullwidth forms
            return True
    return False

def _norm_txt(s):
    """NFC-normalize text before comparing/substring-matching it.

    Fixes the Name Filter silently matching nothing for filenames containing
    Vietnamese/Japanese diacritics or accented characters: two strings that
    *look* identical on screen can still be different byte sequences --
    e.g. "a" + combining acute accent (NFD, decomposed) vs the single
    precomposed "á" character (NFC). Depending on the Vietnamese IME/input
    method used when the file was created vs. when the filter text is typed,
    the DB path and the typed filter text can end up in different forms, so
    `needle in haystack` silently fails even though a human reading both
    strings sees an exact match. Normalizing both sides to NFC before
    comparing makes the match reliable regardless of which form each side
    happened to arrive in.
    """
    try:
        return unicodedata.normalize('NFC', s)
    except Exception:
        return s

def _is_anchor_kw(k, min_len=3):
    """True if keyword k is specific enough to safely drive a search:
    either it meets the normal min_len for Latin-script text, or it
    contains any CJK character (in which case length doesn't matter)."""
    return len(k) >= min_len or _contains_cjk(k)
# ──────────────────────────────────────────────────────────────────────────

def setup_context_menu(target_widget, entry_widget, search_cmd):
    def show_popup_menu(event):
        entry_widget.focus_set()
        m = tk.Menu(target_widget, tearoff=0)
        m.add_command(label="Cut", command=lambda: entry_widget.event_generate("<<Cut>>"))
        m.add_command(label="Copy", command=lambda: entry_widget.event_generate("<<Copy>>"))
        m.add_command(label="Paste", command=lambda: entry_widget.event_generate("<<Paste>>"))
        m.add_separator()
        m.add_command(label="Search", command=search_cmd)
        m.tk_popup(event.x_root, event.y_root)
    target_widget.bind("<Button-3>", show_popup_menu)

def add_only_copy_menu(widget):
    m = tk.Menu(widget, tearoff=0)
    m.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
    widget.bind("<Button-3>", lambda e: m.post(e.x_root, e.y_root))

def _extract_citation_ids(text):
    """Return the set of citation numbers found in `text`, recognizing both
    the standard half-width "[N]" markers the system prompt asks for AND
    full-width "【N】" markers -- Gemini sometimes switches to the
    full-width bracket style on its own when the answer is mostly Japanese
    (it's visually consistent with how it quotes Japanese terms elsewhere
    in the same answer, e.g. "【共通ソルバーライセンス】"), which the old
    ASCII-only regex silently failed to match -- so _append_citation_filenames
    found zero citations and appended NO "Nguồn:"/"Source:" list at all,
    even though the answer clearly did cite files. Handles single numbers,
    comma/semicolon-separated groups ("[4, 5]"), and simple ranges
    ("[4-6]") in either bracket style."""
    ids = set()
    for grp in re.findall(r"[\[【]([0-9][0-9,;\-–~\s]*)[\]】]", text or ""):
        for part in re.split(r"[,;~]", grp):
            part = part.strip()
            if not part:
                continue
            m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", part)
            if m:
                lo, hi = int(m.group(1)), int(m.group(2))
                if lo <= hi and hi - lo < 50:
                    ids.update(range(lo, hi + 1))
            elif part.isdigit():
                ids.add(int(part))
    return ids


def add_tooltip(widget, text):
    """Small English tooltip shown on hover, e.g. for toolbar buttons.
    Flips to appear below the widget instead of above it when there isn't
    enough room above (e.g. the searchbox is docked near the top of the
    screen), so the tooltip is never clipped off-screen.
    `text` may be a plain string, or a zero-arg callable returning a string
    (evaluated fresh every time the tooltip is shown) -- used e.g. for the
    ramp-light status dot whose tooltip text depends on the current color."""
    tip = {"win": None}
    def show(_e=None):
        if tip["win"] is not None:
            return
        _text = text() if callable(text) else text
        w = tk.Toplevel(widget)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        lbl = tk.Label(w, text=_text, font=("Segoe UI", 8), bg="#ffffe0", fg="#111111",
                        relief="solid", bd=1, padx=5, pady=2)
        lbl.pack()
        w.update_idletasks()
        tip_h = w.winfo_reqheight(); tip_w = w.winfo_reqwidth()
        wx = widget.winfo_rootx(); wy = widget.winfo_rooty()
        ww = widget.winfo_width(); wh = widget.winfo_height()
        above_y = wy - tip_h - 4
        # Not enough room above (tooltip would go off the top of the screen,
        # or above the widget's containing monitor edge) -> show below instead.
        if above_y < 0:
            y = wy + wh + 4
        else:
            y = above_y
        x = wx + ww // 2 - tip_w // 2
        screen_w = widget.winfo_screenwidth()
        x = max(0, min(x, screen_w - tip_w))
        w.geometry(f"+{x}+{y}")
        tip["win"] = w
    def hide(_e=None):
        if tip["win"] is not None:
            tip["win"].destroy()
            tip["win"] = None
    widget.bind("<Enter>", show, add="+")
    widget.bind("<Leave>", hide, add="+")
    widget.bind("<ButtonPress>", hide, add="+")

class HistoryPanel:
    """Search History — used to be a separate popup opened via the small 'H'
    button below the ramp light. That button is gone now; this same UI is
    built directly into the right pane of the Help tab instead, so it's
    always visible (rebuilt fresh, like the other tabs, each time results
    are shown — so it's always up to date)."""
    def __init__(self, container, parent_app):
        self.parent_app = parent_app; self.last_selected = None; self.last_time = 0; self.edit_entry = None
        # v-new: default AI Chat History (bottom-right pane) to the most
        # recent Search History entry instead of leaving it blank until
        # the user clicks a row -- see _apply_auto_follow(). Stays True
        # (keeps following the newest entry as fresh searches come in)
        # until the user manually clicks a row themselves (see on_click).
        self._auto_follow_latest = True
        self._view_mode = "filter"  # v-fix: "filter" (by yr/mo/da) or "all" (Show All) —
        # the periodic auto-refresh below needs to know which one to keep
        # re-applying, otherwise it always fell back to re-running the day
        # filter every 5s and silently stomped "Show All" moments after
        # the user clicked it.
        f = tk.Frame(container, bg=BG_COLOR); f.pack(fill="x", padx=8, pady=8)
        now = datetime.now()
        self.yr_v = tk.StringVar(value=str(now.year)); self.mo_v = tk.StringVar(value=now.strftime("%m")); self.da_v = tk.StringVar(value=now.strftime("%d"))
        ttk.OptionMenu(f, self.yr_v, self.yr_v.get(), *[str(y) for y in range(now.year, now.year-3, -1)]).pack(side="left")
        ttk.OptionMenu(f, self.mo_v, self.mo_v.get(), *[f"{m:02d}" for m in range(1, 13)]).pack(side="left", padx=2)
        ttk.OptionMenu(f, self.da_v, self.da_v.get(), *[f"{d:02d}" for d in range(1, 32)]).pack(side="left")
        tk.Button(f, text="Filter", command=self.refresh, bg="#444", fg="white").pack(side="left", padx=5)
        tk.Button(f, text="Show All", command=self.show_all, bg="#5c5c5c", fg="white").pack(side="left", padx=2)
        tk.Button(f, text="Export Excel", command=self.export_to_excel, bg="#2e7d32", fg="white").pack(side="left", padx=5)
        tk.Button(f, text="Clear All", command=self.clear_all, bg="#c62828", fg="white").pack(side="right", padx=5)
        t_f = tk.Frame(container, bg=BG_COLOR); t_f.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree = ttk.Treeview(t_f, columns=("d", "q"), show="headings"); self.tree.heading("d", text="Date Time"); self.tree.heading("q", text="Search History")
        self.tree.column("d", width=140, stretch=False); self.tree.column("q", width=300, stretch=True)
        sb = ttk.Scrollbar(t_f, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=sb.set); self.tree.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_click); self.tree.bind("<Button-3>", self.show_menu); self.refresh()
        # v1.10: selecting a row also shows that query's AI Chat (if any) in
        # the AI Chat History pane below -- no need for a separate keyword
        # list over there anymore, this list already has it.
        self.tree.bind("<<TreeviewSelect>>", self._notify_chat_history_panel)
        # v-new: this pane used to only ever load once at creation time —
        # a query typed into the search box gets auto-saved to history
        # after ~10s idle (see _maybe_save_hist_idle), but nothing told
        # this pane to re-query afterwards, so it looked "stuck" until the
        # user manually clicked Filter/Show All. Poll every 5s instead so
        # freshly-saved searches just show up on their own.
        self._auto_refresh_job = None
        self._schedule_auto_refresh()

    def _schedule_auto_refresh(self):
        try:
            if not self.tree.winfo_exists():
                return
            # Don't yank the tree out from under an in-progress rename edit box.
            editing = self.edit_entry is not None and self.edit_entry.winfo_exists()
            if not editing:
                # v-fix: re-apply whichever view the user is currently on,
                # instead of always calling self.refresh() (day filter) —
                # that used to silently revert "Show All" back to
                # today-only within a few seconds of clicking it.
                if self._view_mode == "all":
                    self.show_all()
                else:
                    self.refresh()
        except Exception:
            return
        try:
            self._auto_refresh_job = self.tree.after(5000, self._schedule_auto_refresh)
        except Exception:
            pass

    def on_click(self, e):
        now = time.time(); row = self.tree.identify_row(e.y); col = self.tree.identify_column(e.x)
        if row and col == "#2" and row == self.last_selected and (now - self.last_time) > 0.4: self.show_edit_box(row, col)
        # v-new: a genuine user click on a row means they picked it on
        # purpose -- stop auto-following the latest entry so their choice
        # sticks (see _apply_auto_follow) instead of getting overridden by
        # the next 5s periodic refresh.
        if row:
            self._auto_follow_latest = False
        self.last_selected = row; self.last_time = now

    def _apply_auto_follow(self):
        """v-new: default the AI Chat History pane to the most recent
        Search History entry instead of requiring the user to click a row
        first. Re-applied after every refresh()/show_all() so it keeps
        following the newest entry as fresh searches get saved -- but only
        while the user hasn't manually picked a different row themselves
        (see on_click, which turns this off)."""
        if not self._auto_follow_latest:
            return
        try:
            children = self.tree.get_children()
            if children:
                self.tree.selection_set(children[0])
                # v-fix: selection_set() doesn't reliably re-fire
                # <<TreeviewSelect>> when the item was already selected
                # (e.g. the very first call, right after refresh() just
                # inserted it) -- call the notify directly too, so AI Chat
                # History always ends up in sync instead of depending on
                # whether the virtual event happened to fire this time.
                self._notify_chat_history_panel()
        except Exception:
            pass

    def _notify_chat_history_panel(self, e=None):
        """v1.10: tell ChatHistoryPanel (AI Chat History, bottom half of the
        same Help-tab pane) which query got selected, so it can show that
        query's saved AI Chat using its own full space -- see
        ChatHistoryPanel.show_for_query()."""
        sel = self.tree.selection()
        if not sel:
            return
        try:
            val = self.tree.item(sel[0], "values")[1]
        except Exception:
            return
        ch_panel = getattr(self.parent_app, "chat_history_panel", None)
        if ch_panel is not None:
            try:
                ch_panel.show_for_query(val)
            except Exception:
                pass

    def show_edit_box(self, row, col):
        if self.edit_entry: self.edit_entry.destroy()
        bbox = self.tree.bbox(row, col); x, y, w, h = bbox; val = self.tree.item(row, "values")[1]
        self.edit_entry = tk.Entry(self.tree, font=("Segoe UI", 9), bd=0); self.edit_entry.insert(0, val); self.edit_entry.place(x=x, y=y, width=w, height=h); self.edit_entry.focus_set()
        add_only_copy_menu(self.edit_entry); self.edit_entry.bind("<FocusOut>", lambda e: self.edit_entry.destroy())
    def show_all(self):
        self._view_mode = "all"
        for i in self.tree.get_children(): self.tree.delete(i)
        # v7.10 fix: Search History now lives in its own HISTORY_DB_FILE,
        # completely separate from search_data.db (DB_FILE) -- so opening/
        # reading history can never affect the indexing ramp light. Also
        # guarded against sqlite3.connect() silently creating an empty file.
        if not os.path.exists(HISTORY_DB_FILE): return
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor(); c.execute("SELECT date, query FROM history ORDER BY id DESC")
            [self.tree.insert("", tk.END, values=r) for r in c.fetchall()]; conn.close()
        except: pass
        self._apply_auto_follow()
    def export_to_excel(self):
        data = [self.tree.item(i)['values'] for i in self.tree.get_children()]
        if data:
            f_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SearchHistory.xlsx")
            if f_path: pd.DataFrame(data, columns=["Date Time", "Keyword"]).to_excel(f_path, index=False)
    def show_menu(self, e):
        row = self.tree.identify_row(e.y)
        if row: val = self.tree.item(row, "values")[1]; m = tk.Menu(self.tree, tearoff=0); m.add_command(label="Search again", command=lambda: self.parent_app.use_query_from_hist(val)); m.post(e.x_root, e.y_root)
    def refresh(self):
        self._view_mode = "filter"
        for i in self.tree.get_children(): self.tree.delete(i)
        # v7.10 fix: reads from HISTORY_DB_FILE now, not search_data.db --
        # this is the call that runs automatically at app startup
        # (__init__ calls self.refresh() before the user ever opens the
        # Help tab), so it was the main thing silently creating a phantom
        # search_data.db before this fix.
        if not os.path.exists(HISTORY_DB_FILE): return
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor(); dt = f"{self.yr_v.get()}-{self.mo_v.get()}-{self.da_v.get()}"
            c.execute("SELECT date, query FROM history WHERE date LIKE ? ORDER BY id DESC", (f"{dt}%",)); [self.tree.insert("", tk.END, values=r) for r in c.fetchall()]; conn.close()
        except: pass
        self._apply_auto_follow()
    def clear_all(self):
        if not os.path.exists(HISTORY_DB_FILE): return  # v7.10 fix: nothing to clear, don't create a phantom DB
        if messagebox.askyesno("Confirm", "Clear all history?"):
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor(); c.execute("DELETE FROM history"); conn.commit(); conn.close(); self.refresh()

class ChatHistoryPanel:
    """AI Chat History (v1.10) -- no longer keeps its own keyword/date list
    (that duplicated Search History right next to it); instead this pane
    is driven BY Search History: selecting a query up there (see
    HistoryPanel._notify_chat_history_panel) shows that query's saved AI
    Chat session(s) here, using the pane's FULL space. Reads from the
    'chat_history' table in HISTORY_DB_FILE (grouped by session_id -- see
    RealtimeSmartSearchApp._chat_session_id/_new_online_chat), so saved
    chats persist across searches and app restarts just like Search
    History does.

    Deliberately simpler than the live AI Chat panel: no New Chat icon, no
    input box, and its right-click menu is Copy-only -- explicitly NOT the
    Search History pane's "Search again" (a saved chat isn't a search
    query, so re-running it as one wouldn't make sense)."""
    def __init__(self, container, parent_app):
        self.parent_app = parent_app
        self._current_query = None
        self._last_shown_session_ids = None   # v-fix: tracks which sessions were last rendered, so we know when NEW content actually arrived vs. just a same-content refresh (see show_for_query)
        f = tk.Frame(container, bg=BG_COLOR); f.pack(fill="x", padx=8, pady=(4, 4))
        self.title_lbl = tk.Label(f, text="Chọn 1 dòng trong Search History để xem AI Chat tương ứng",
                                   bg=BG_COLOR, fg=PLACE_COLOR, font=("Segoe UI", 8, "italic"), anchor="w")
        self.title_lbl.pack(side="left", fill="x", expand=True)

        t_f = tk.Frame(container, bg=BG_COLOR)
        t_f.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        sb = ttk.Scrollbar(t_f)
        sb.pack(side="right", fill="y")
        self.preview_txt = tk.Text(t_f, bg=ENTRY_BG, fg=TEXT_COLOR, font=("Segoe UI", 9),
                                    wrap="word", bd=0, padx=6, pady=4, state="disabled",
                                    yscrollcommand=sb.set)
        self.preview_txt.pack(side="left", fill="both", expand=True)
        sb.config(command=self.preview_txt.yview)
        self.preview_txt.tag_config("user_icon", foreground="#0b5fa5", font=("Segoe UI", 9, "bold"))
        self.preview_txt.tag_config("user", foreground="#0b5fa5", font=("Segoe UI", 9))
        self.preview_txt.tag_config("assistant_icon", foreground=TEXT_COLOR, font=("Segoe UI", 9, "bold"))
        self.preview_txt.tag_config("body", foreground=TEXT_COLOR, font=("Segoe UI", 9))
        self.preview_txt.tag_config("sep", foreground=PLACE_COLOR, font=("Segoe UI", 8, "italic"))
        add_only_copy_menu(self.preview_txt)   # copy only -- no "Search again" here, unlike Search History

        # Same "poll every 5s" idea as HistoryPanel -- if the currently-shown
        # query gets a new/updated chat saved elsewhere, refresh it in place.
        self._auto_refresh_job = None
        self._schedule_auto_refresh()

    def _schedule_auto_refresh(self):
        try:
            if not self.preview_txt.winfo_exists():
                return
            if self._current_query:
                self.show_for_query(self._current_query, _force=False)
        except Exception:
            return
        try:
            self._auto_refresh_job = self.preview_txt.after(5000, self._schedule_auto_refresh)
        except Exception:
            pass

    @staticmethod
    def _extract_query(text):
        """Recover the search query a saved first-message was about. The
        auto-summary prompt ("Hãy phân tích và tóm tắt... cho truy vấn:
        "X".") is parsed down to just X; anything else (a manually-typed
        question) is returned as-is."""
        if not text:
            return ""
        m = re.match(r'^Hãy phân tích và tóm tắt.*?cho truy vấn:\s*"(.+)"\.?$', text.strip())
        return m.group(1).strip() if m else text.strip()

    def _find_sessions_for_query(self, query_text):
        if not os.path.exists(HISTORY_DB_FILE):
            return []
        q_norm = (query_text or "").strip().lower()
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='chat_history'")
            if c.fetchone()[0] == 0:
                conn.close()
                return []
            # v1.11 FIX: primary match is the 'query' column saved directly
            # alongside each row (see _save_chat_hist) -- exact, no
            # reparsing needed, so it can't drift out of sync with what's
            # actually in Search History.
            try:
                c.execute("SELECT DISTINCT session_id, MIN(id) as first_id FROM chat_history "
                          "WHERE lower(trim(query))=? GROUP BY session_id ORDER BY first_id ASC", (q_norm,))
                direct_matches = [r[0] for r in c.fetchall()]
            except Exception:
                direct_matches = []  # older DB without the 'query' column yet
            if direct_matches:
                conn.close()
                return direct_matches
            # Fallback for legacy rows saved before this fix (no 'query'
            # column value): reconstruct the query from the first message.
            c.execute("SELECT session_id, MIN(id) as first_id, "
                      "(SELECT text FROM chat_history c2 WHERE c2.session_id=c1.session_id "
                      "AND c2.role='user' ORDER BY c2.id ASC LIMIT 1) as first_msg "
                      "FROM chat_history c1 GROUP BY session_id ORDER BY first_id ASC")
            rows = c.fetchall()
            conn.close()
        except Exception:
            return []
        return [session_id for session_id, _fid, first_msg in rows
                if self._extract_query(first_msg).lower() == q_norm]

    def show_for_query(self, query_text, _force=True):
        """Called by HistoryPanel (Search History) when a row is selected --
        finds the AI Chat session(s) auto-summarized for this exact query
        and renders their full Q&A here.

        v-fix (thanh cuộn nhảy về đầu): this used to unconditionally call
        preview_txt.see("1.0") on every redraw, including the periodic 5s
        auto-refresh tick (_schedule_auto_refresh) -- so a user who
        scrolled down to read older turns got yanked back to the top every
        few seconds, before they could even finish reading.

        v-fix (Vấn đề 2 vẫn hiện nội dung cũ): the first attempt at fixing
        the above tied "should I jump to show the latest content?" to
        `query_text != self._current_query` (i.e. only on switching to a
        different query). That broke the exact case being reported here:
        _auto_ai_chat_after_search calls this pane twice for the SAME new
        query -- once immediately when a search starts (nothing saved yet,
        _current_query flips to the new query right away), and again once
        the reply actually lands a moment later. By that second call
        `query_text == self._current_query` ALREADY (set by the first
        call), so it looked like "just a refresh" and only restored the
        OLD scroll position instead of jumping to the reply that had just
        been saved -- the pane kept showing whatever was on screen before
        the new session existed at all, i.e. old content from a previous
        round of chatting about this same keyword.

        Now the jump decision is based on whether the actual SET OF
        SESSIONS for this query changed since the last render (a real new
        session appeared), not on whether the query string itself changed.
        A true no-op periodic refresh (nothing new saved) still preserves
        scroll position; a switch to a different query OR a new session
        landing for the query already being viewed both jump to show the
        latest session.

        v-fix (thanh cuộn nhảy lên 1 tí khi đang ở dưới cùng): even a true
        no-op tick used to still delete()+insert() the exact same text
        back into the widget every 5s, then restore the scroll position
        via yview_moveto(prev_fraction) -- that's a fractional restore,
        not pixel-perfect, so a user parked at the very bottom kept
        getting nudged up by a hair on every tick. Fix: when nothing
        actually changed, skip touching the Text widget entirely instead
        of redrawing identical content and trying to restore the scroll
        position afterwards."""
        if not query_text:
            return
        if not _force and query_text != self._current_query:
            return  # a periodic refresh tick, but the user has since picked a different query
        is_new_query = (query_text != self._current_query)
        self._current_query = query_text
        session_ids = self._find_sessions_for_query(query_text)
        sessions_changed = is_new_query or (session_ids != self._last_shown_session_ids)
        if not sessions_changed:
            self._last_shown_session_ids = session_ids
            return  # nothing new to show -- leave the widget (and scrollbar) untouched
        self._last_shown_session_ids = session_ids
        try:
            self.title_lbl.config(text=f' AI Chat History — "{query_text}" ', fg="#1a5fb4")
        except Exception:
            pass
        try:
            self.preview_txt.config(state="normal")
            self.preview_txt.delete("1.0", "end")
            last_session_start = "1.0"
            if not session_ids:
                self.preview_txt.insert("end", f'(Chưa có AI Chat nào cho truy vấn "{query_text}")', "sep")
            else:
                for si, session_id in enumerate(session_ids):
                    if si > 0:
                        self.preview_txt.insert("end", "\n\n───────────\n\n", "sep")
                    if si == len(session_ids) - 1:
                        last_session_start = self.preview_txt.index("end-1c")
                    self._insert_session(session_id, query_text)
            self.preview_txt.config(state="disabled")
            # We only ever reach here when sessions_changed is True (see the
            # early-return above), so always jump to the newest session.
            if session_ids:
                self.preview_txt.see(last_session_start)
        except Exception:
            pass

    def _sanitize_legacy_pseudo_paths(self, text):
        """v-fix (Vấn đề: AI Chat History hiện cả Bad lẫn Good cho MỌI
        keyword sau khi bấm "Search again" từ Search History): các session
        được lưu TRƯỚC KHI có fix format đẹp (_display_label_for_source /
        _make_outlook_pseudo_path) đã đóng băng chuỗi thô
        "OUTLOOK::entry::store::folder::subject.msg" /
        "ONENOTE::page::section::title.one" ngay trong cột `text` của
        HISTORY_DB_FILE -- vĩnh viễn xấu, vì code hiện tại chỉ format đẹp
        cho câu trả lời MỚI, không bao giờ sửa lại text cũ đã lưu.
        show_for_query() xếp TẤT CẢ session của cùng 1 query chồng lên
        nhau (cũ trên, mới dưới) -- nên bất kỳ keyword nào từng được test
        trước fix đều có 1 session xấu (Bad) đứng trên 1 session đẹp
        (Good) mới hơn, dù chúng trỏ tới cùng 1 email/note.

        Sửa tại thời điểm HIỂN THỊ (không đụng DB): quét lại text cũ, tìm
        token pseudo-path còn sót và convert qua _display_label_for_source
        (hàm format đẹp đã có sẵn) trước khi insert vào preview_txt -- mọi
        session (cũ hay mới) từ nay luôn hiện nhất quán.

        v-fix (Vấn đề: bật lại "— {pseudo_path}" ở _append_citation_
        filenames để có data click cho Outlook/OneNote thì hàm này lại
        xoá nhầm luôn phần path vừa thêm, vì regex bên dưới match CẢ
        pseudo-path hợp lệ nằm sau "  —  " của dòng đã format đẹp, không
        chỉ token cũ/hỏng nó được viết ra để dọn): dòng MỚI, đúng format
        luôn có dạng "[N] label  —  OUTLOOK::..." / "...ONENOTE::..." --
        token pseudo-path nằm NGAY sau "  —  " (2 space, em-dash, 2
        space). Dòng CŨ/hỏng (trước khi _display_label_for_source tồn
        tại) không có cấu trúc "label — " đó đứng trước -- token pseudo-
        path nằm trơ trọi một mình, ví dụ ngay sau "[N] ". Dùng negative
        lookbehind để CHỈ sanitize token không đứng ngay sau "  —  ",
        tức chỉ dọn dòng thật sự cũ/hỏng, giữ nguyên dòng mới đã đẹp (và
        giữ nguyên path thật cần cho click)."""
        if not text or ("OUTLOOK::" not in text and "ONENOTE::" not in text):
            return text
        def _fix(m):
            try:
                return self.parent_app._display_label_for_source(m.group(0))
            except Exception:
                return m.group(0)
        text = re.sub(r'(?<!  —  )OUTLOOK::\S.*?\.msg', _fix, text)
        text = re.sub(r'(?<!  —  )ONENOTE::\S.*?\.one', _fix, text)
        return text

    def _insert_session(self, session_id, query_text=None):
        """v-fix (session_id collision hardening): also filter by query
        here, not just session_id. session_id is now seeded from
        MAX(session_id)+1 at startup (see
        RealtimeSmartSearchApp._load_next_chat_session_id) so a collision
        shouldn't happen anymore, but this row-level query filter is kept
        as a second line of defense -- without it, ANY future collision
        (e.g. two people sharing one HISTORY_DB_FILE, or manual DB edits)
        would silently blend an unrelated old conversation's rows into
        whatever session_id happens to match, with no query check at all."""
        q_norm = (query_text or "").strip().lower()
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            if q_norm:
                c.execute("SELECT role, text, date FROM chat_history "
                          "WHERE session_id=? AND lower(trim(query))=? ORDER BY id ASC",
                          (session_id, q_norm))
            else:
                c.execute("SELECT role, text, date FROM chat_history WHERE session_id=? ORDER BY id ASC", (session_id,))
            rows = c.fetchall()
            conn.close()
        except Exception:
            rows = []
        for role, text, date in rows:
            who_icon = "👤" if role == "user" else "🤖"
            icon_tag = "user_icon" if role == "user" else "assistant_icon"
            body_tag = "user" if role == "user" else "body"
            text = self._sanitize_legacy_pseudo_paths(text)
            if self.preview_txt.index("end-1c") != "1.0":
                self.preview_txt.insert("end", "\n\n")
            self.preview_txt.insert("end", f"{who_icon} ", icon_tag)
            # v-fix (Vấn đề: link "Nguồn:" trong AI Chat History không mở
            # được file, chỉ có ở live chat): the live AI Chat panel turns
            # each "[N] name — path" line into a clickable link (see
            # RealtimeSmartSearchApp._insert_chat_text_with_links), but
            # this saved-history preview just inserted the raw text --
            # right file, right citation, just no way to click it here.
            # Same tag-based approach, reusing the app's existing
            # _open_path_for_chat() (which already knows how to route
            # Outlook/OneNote pseudo-paths, not just plain files).
            if role == "assistant":
                self._insert_text_with_source_links(text or "", body_tag)
            else:
                self.preview_txt.insert("end", text or "", body_tag)

    def _insert_text_with_source_links(self, text, body_tag):
        """Same "[N] name — path" -> clickable link rendering as the live
        AI Chat panel's _insert_chat_text_with_links, adapted for this
        pane's own preview_txt widget.

        v-fix (Vấn đề: Outlook/OneNote không click mở được ở đây, giống
        panel AI Chat sống): cùng logic ẩn pseudo-path xấu -- xem
        RealtimeSmartSearchApp._insert_chat_text_with_links.

        v-new: link internet (http/https, trong mục "Thông tin bổ sung
        (internet)") cũng click mở được, mở bằng webbrowser.open() --
        cùng logic với bản sống, xem _insert_chat_text_with_links."""
        citation_line_re = re.compile(r"^(\[\d+\]\s+.+?)\s+—\s+(.+)$")
        url_re = re.compile(r"https?://[^\s<>\"')]+")
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if i > 0:
                self.preview_txt.insert("end", "\n")
            m = citation_line_re.match(line)
            if m:
                label, path = m.group(1), m.group(2)
                self._chp_link_seq = getattr(self, "_chp_link_seq", 0) + 1
                tag_name = f"chplink_{self._chp_link_seq}"
                if _is_outlook_pseudo_path(path) or _is_onenote_pseudo_path(path):
                    # Ẩn hẳn pseudo-path xấu -- chỉ hiện label thân thiện,
                    # gạch dưới, click bằng đúng path thật (ẩn) phía sau.
                    self.preview_txt.insert("end", label, (body_tag, tag_name))
                else:
                    self.preview_txt.insert("end", label + "  —  ", body_tag)
                    self.preview_txt.insert("end", path, (body_tag, tag_name))
                self.preview_txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)
                self.preview_txt.tag_bind(
                    tag_name, "<Button-1>", lambda e, p=path: self.parent_app._open_path_for_chat(p))
                self.preview_txt.tag_bind(
                    tag_name, "<Enter>", lambda e: self.preview_txt.config(cursor="hand2"))
                self.preview_txt.tag_bind(
                    tag_name, "<Leave>", lambda e: self.preview_txt.config(cursor=""))
                continue
            url_matches = list(url_re.finditer(line))
            if not url_matches:
                self.preview_txt.insert("end", line, body_tag)
                continue
            pos = 0
            for um in url_matches:
                if um.start() > pos:
                    self.preview_txt.insert("end", line[pos:um.start()], body_tag)
                url = um.group(0)
                self._chp_link_seq = getattr(self, "_chp_link_seq", 0) + 1
                tag_name = f"chplink_{self._chp_link_seq}"
                self.preview_txt.insert("end", url, (body_tag, tag_name))
                self.preview_txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)
                self.preview_txt.tag_bind(tag_name, "<Button-1>", lambda e, u=url: webbrowser.open(u))
                self.preview_txt.tag_bind(tag_name, "<Enter>", lambda e: self.preview_txt.config(cursor="hand2"))
                self.preview_txt.tag_bind(tag_name, "<Leave>", lambda e: self.preview_txt.config(cursor=""))
                pos = um.end()
            if pos < len(line):
                self.preview_txt.insert("end", line[pos:], body_tag)

    def clear_all(self):
        if not os.path.exists(HISTORY_DB_FILE):
            return
        if messagebox.askyesno("Confirm", "Clear all AI Chat History?"):
            try:
                conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
                c.execute("DELETE FROM chat_history"); conn.commit(); conn.close()
            except Exception:
                pass
            self._current_query = None
            try:
                self.preview_txt.config(state="normal")
                self.preview_txt.delete("1.0", "end")
                self.preview_txt.config(state="disabled")
                self.title_lbl.config(text="Chọn 1 dòng trong Search History để xem AI Chat tương ứng", fg=PLACE_COLOR)
            except Exception:
                pass

class RealtimeSmartSearchApp:
    def __init__(self, root):
        self.root = root; self.root.overrideredirect(True); self.root.attributes("-topmost", True)
        # v2.9: don't stay pinned above every other window forever -- release
        # "always on top" shortly after launch (same pattern already used by
        # HistoryWindow below). This still pops the search box to the front the
        # instant it appears, but once the user opens a file/folder/browser
        # from a result, that new window is free to sit above this one instead
        # of the search box permanently covering it.
        self.root.after(500, lambda: self.root.attributes("-topmost", False))
        self.root.configure(bg=BG_COLOR)
        self.args = sys.argv[1:]; self.has_args = len(self.args) > 0
        self.ph_list = [READY_PH, READY_PH1, READY_PH2]; self.ph_index = 0
        self.x_pos = (root.winfo_screenwidth() // 4) - 120
        self.root.geometry(f"{SMALL_SIZE}+{self.x_pos}+5")
        
        self.active_result_win = None
        self.search_timer = None 
        self.current_search_id = 0
        # v-fix (Outlook/OneNote search bị chậm dần 61s -> 297s -> 1660s):
        # trước đây mỗi lần _smart_search_realtime chạy (mỗi lần user dừng gõ
        # >=350ms) sẽ spawn MỘT thread _merge_mail_notes_async mới gọi
        # search_outlook()/search_onenote() (mở connection sqlite MỚI tới file
        # 350MB search_outlook.db). Nếu user gõ nhiều "cụm" cách nhau >350ms
        # trong lúc thread cũ (đang mất hàng chục giây) CHƯA xong, các thread
        # này CHỒNG LÊN NHAU và cùng lúc mở/đọc file DB đó -- trên Windows,
        # antivirus/Defender quét lại toàn bộ file mỗi khi có handle mở mới,
        # nên càng nhiều connection mở đồng thời, mỗi cái càng chậm thêm =>
        # thời gian phình ra theo cấp số nhân đúng như log của bạn. version
        # check trong _merge_mail_notes_async chỉ bỏ KẾT QUẢ cũ lúc render,
        # KHÔNG ngăn nó tốn hàng chục giây chạy trước đó.
        # Fix: dùng 1 worker thread DUY NHẤT + queue depth=1 -- mỗi query mới
        # sẽ THAY THẾ query đang chờ (nếu có), và không bao giờ có 2 lệnh gọi
        # search_outlook()/search_onenote() chạy đồng thời.
        self._mail_notes_queue = queue.Queue(maxsize=1)
        threading.Thread(target=self._mail_notes_worker_loop, daemon=True).start()
        # v5.8: auto-save to Search History after 10s of no typing/Enter --
        # see on_key_release / _maybe_save_hist_idle for details.
        self._hist_idle_timer = None
        self._last_saved_hist_query = ""
        # v-fix (AI Chat phân tích quá sớm khi gõ tiếng Nhật nhiều từ):
        # timestamp of the most recent KeyRelease in the Search box --
        # _auto_ai_chat_after_search waits until this has been quiet for
        # _CHAT_AUTO_QUIET_SEC before actually sending, so a multi-word JP
        # query typed in bursts (e.g. らいせんす [変換/Enter] ...1s...
        # さーばー) doesn't get auto-summarized on just the first word
        # while the user is still mid-typing the rest. See on_key_release.
        self._last_keystroke_ts = 0.0
        self.db_conn = None   # persistent connection — opened once, reused every keystroke
        # v2.6: guards window creation against races. Japanese/CJK IME input fires
        # several KeyRelease events in a burst (once per romaji key during
        # composition, then again on conversion/commit) — each one used to spawn
        # its own _mft_scan_search thread with no coordination, so more than one
        # thread could see "no result window yet" before the first one's
        # scheduled show_results() actually ran, creating two duplicate windows.
        self._win_create_lock = threading.Lock()
        self._opening_result_win = False
        # Smart size filter
        self.size_op_var   = tk.StringVar(value=">")
        self.size_num_var  = tk.StringVar(value="")
        self.size_unit_var = tk.StringVar(value="MB")
        # v7.10: "Whole word" toggle -- ON by default (v7.10b: switched from
        # OFF after testing showed it filters out buried-substring false
        # positives well, e.g. "adas" no longer matches "readasync.xml" or
        # "ReadAStringExample.mlx"). When ON, a keyword like "adas" only
        # counts as a match when it's bounded by non-letter characters
        # (start/end of name, or a digit/underscore/hyphen/dot/space/
        # parenthesis) on both sides -- so "ADAS_systems.pdf" or
        # "VDIM_0ADAS_ESP" still match, but names with the keyword buried
        # mid-word don't. Trade-off: a name like "ADASDemoInstructions.pdf"
        # (no separator between "ADAS" and "Demo") also won't match while
        # ON, since both sides are letters.
        self.whole_word_var = tk.BooleanVar(value=True)
        self._all_files_data  = []   # cache for client-side filter (File Name tab)
        self._all_content_data = []  # cache for client-side filter (File Content tab)
        self.filter_count_label = None
        self.content_filter_count_label = None
        self._last_query = ""
        self._last_bm25_cont_res = []   # BM25-only cont_res cache for AI merge

        # v1.2: the "AI Search" button (🤖) now always runs the offline model
        # (Jina-v3/BGE-Gemma2) -- the old "Online" mode tied to this button
        # is gone. Chatting with the online AI (Gemini Flash / GPT-OSS) now
        # lives in its own chat panel below the results -- see
        # _build_online_chat_panel().
        self._chat_history = []                         # [{"role": "user"/"assistant", "text": ...}]
        self._chat_active_model = DEFAULT_ONLINE_MODEL   # "gemini" | "gptoss" -- auto-switches on rate-limit
        # v-new (Vấn đề: tự chuyển sang model còn lại "cho đến hết" thay vì
        # quay lại thử model vừa hết quota mỗi lượt chat): trước đây,
        # _chat_active_model chỉ đổi model TRONG lượt chat bị lỗi -- nhưng
        # nếu chế độ Online (web_search) đang bật, khối "Online toggle" bên
        # dưới trong _online_chat_worker sẽ ép quay lại Gemini ở lượt kế
        # tiếp CHỈ VÌ key Gemini vẫn còn tồn tại (_online_ai_available chỉ
        # kiểm tra có key, không biết quota đã hết hay chưa) -- khiến app
        # cứ lặp lại: gọi Gemini -> 429 -> fallback GPT-OSS mỗi lượt, tốn
        # thêm 1 API call + độ trễ mỗi tin nhắn, và nếu GPT-OSS cũng hết
        # quota cùng lúc thì lỗi 429 gốc lọt thẳng ra ngoài không fallback
        # được nữa (đúng lỗi user gặp). Dict này nhớ model nào VỪA bị
        # rate-limit trong phiên hiện tại, để mọi nơi chọn model (đầu lượt
        # chat, khối ép-Gemini-cho-Online, dịch chat...) đều tự động BỎ QUA
        # model đó và ở lại với model còn lại, cho đến khi 1 lần gọi tới
        # model đó THÀNH CÔNG trở lại (xem _online_chat_worker / hàm
        # _mark_model_rate_limited).
        self._chat_model_rate_limited = {"gemini": False, "gptoss": False, "qwen": False}
        # v-new: the user's manually-picked model via the header dropdown
        # (see _build_online_chat_panel) -- kept separate from
        # _chat_active_model so a rate-limit auto-fallback mid-conversation
        # doesn't overwrite what they actually asked for, and so "New Chat"
        # restarts on the model they picked instead of silently reverting
        # to Gemini every time (see _new_online_chat).
        # v-new (yêu cầu: nhớ Chat AI model đã chọn qua configure.ini):
        # xem _config_set("Models","chat_ai_model",...) trong _on_model_pick.
        _saved_chat_model = _config_get("Models", "chat_ai_model")
        self._chat_preferred_model = _saved_chat_model if _saved_chat_model in ONLINE_AI_MODELS else DEFAULT_ONLINE_MODEL
        self._chat_auto_sent_for = None                  # query string already auto-summarized (avoid repeats)
        self._chat_citation_paths = []   # v-fix (Vấn đề 2): [N] -> path CỐ ĐỊNH trong suốt 1 phiên chat -- xem _online_chat_worker
        # v-fix (session_id collision): this used to always start at 1 on
        # every app launch. session_id is only unique WITHIN one run, but
        # chat_history rows persist forever in HISTORY_DB_FILE across many
        # runs -- so run #2's session_id=1 silently collided with run #1's
        # already-saved session_id=1 (very likely for a repeatedly-searched
        # keyword like "EPS"). ChatHistoryPanel._insert_session() selects
        # ALL rows for a given session_id with no query filter, so two
        # unrelated conversations sharing that same reused integer got
        # displayed concatenated as if they were one session ("Base + New
        # Chat + very old History" all blended together), while a brand
        # new session that also happened to collide could even get its
        # rows interleaved into that old block instead of standing out as
        # its own entry. Seeding from MAX(session_id) already in the DB
        # instead of a hardcoded 1 guarantees every session_id used in this
        # run is one this DB has never seen before.
        self._chat_session_id = self._load_next_chat_session_id()   # v1.9: groups chat_history rows into "sessions" for AI Chat History (ChatHistoryPanel)
        # v-new (yêu cầu: nhớ ngôn ngữ VI/EN/JP qua các lần mở app): đọc
        # lại lựa chọn đã lưu lần trước từ configure.ini (section
        # [General], key language) -- xem CONFIG_FILE/_load_config ở đầu
        # file. Nếu chưa từng lưu (lần đầu chạy) hoặc giá trị không hợp
        # lệ, vẫn rơi về CHAT_DEFAULT_LANG ("VI") như cũ. Việc LƯU lại
        # mỗi khi user đổi nút VI/EN/JP nằm trong _select_chat_lang
        # (_build_online_chat_panel).
        _saved_lang = _config_get("General", "language")
        self._chat_lang = _saved_lang if _saved_lang in CHAT_LANG_OPTIONS else CHAT_DEFAULT_LANG   # v-new: "VI"/"EN"/"JP" -- selected via the header buttons, see _build_online_chat_panel
        # v-fix (AI Chat trả lời trùng lặp cho cùng 1 câu hỏi): the "AI
        # Search" button both calls _auto_ai_chat_after_search() directly
        # AND (via the split-window it opens once semantic search finishes)
        # schedules ANOTHER call to the same function for the same query --
        # meant to run the chat in parallel with the offline embedding
        # search rather than one blocking the other. Normally _chat_busy
        # blocks the second one, but this single-flight timestamp is a
        # second, timing-independent guard: two attempts for the identical
        # query within a few seconds of each other collapse into one,
        # regardless of exactly how each was triggered.
        self._last_auto_chat_attempt = None   # (query, time.time()) of the most recent auto-chat kickoff attempt
        # v-new: ◀/▶ session preview navigation (see _refresh_chat_preview_nav
        # / _chat_preview_prev / _chat_preview_next / _render_chat_preview).
        # -1 == showing the LIVE conversation (self._chat_history); a
        # non-negative index previews an older saved session read-only.
        self._chat_preview_query = None
        self._chat_preview_sessions = []   # older (non-live) session_ids for _chat_preview_query, oldest first
        self._chat_preview_idx = -1
        self._mail_notes_merge_pending = False   # v1.10: True while _merge_mail_notes_async is running -- see _auto_ai_chat_after_search
        self._online_chat_panel = None
        # v-new (nút Online): default OFF -- AI Chat only uses the local
        # BM25/content_store excerpts (unchanged behavior). When turned ON,
        # AI Chat is allowed to also draw on the model's own general
        # knowledge, and -- for Gemini -- a REAL web-search grounding tool,
        # to go deeper than what the local scan found. See
        # _build_online_chat_panel (toggle button) and _online_chat_worker /
        # _call_online_ai_chat (web_search plumbing).
        self._chat_online_enabled = True  # v-fix: mặc định luôn ON, không còn nút toggle -- xem _online_chat_worker
        self._last_bm25_file_res = []   # BM25-only file_res cache for toggle
        self._ai_search_btn = None      # reference to AI Search button
        self._ai_mode_active = False    # True = currently showing Hybrid results
        # v-fix (AI Search/AI Chat panes flashing back to Default on a new
        # keyword): the sticky-AI retrigger in _smart_search_realtime
        # briefly sets _ai_mode_active back to False so _ai_search_and_update
        # (which TOGGLES on _ai_mode_active) re-arms instead of reverting to
        # BM25 -- but update_or_show_results runs its OWN, separate
        # "_same_ai_query" check in that same window and reads
        # _ai_mode_active as False, so it thought AI mode had genuinely
        # ended and closed the AI split window / AI chat pane and reset the
        # button back to Default, right before the background thread turned
        # AI mode back on a moment later -- hence the pane vanishing and
        # needing a manual "AI Search" click to bring it back. This flag
        # records "a sticky AI re-run for THIS query is in flight" so
        # update_or_show_results can tell that apart from a real exit out
        # of AI mode. See _smart_search_realtime / update_or_show_results /
        # _ai_search_and_update.
        self._ai_pending_query = None
        self._ai_cont_res  = []         # AI merged content results cache
        self._ai_file_res  = []         # AI merged file results cache
        # Track split-pane trees per tab so filters can update all panes
        # Each dict: {"main": tree_widget, "adv": tree_widget, "ai": tree_widget}
        self._c_pane_trees   = {}
        self._f_pane_trees   = {}
        self._fol_pane_trees = {}
        self._adv_search_btn = None      # reference to Advanced button
        self._update_db_btn  = None      # reference to Update DB button (result window only)
        self._update_db_running = False  # True while indexing_worker() is running (button or --update data)
        self._outlook_progress_text = None  # v-outlook: live "Indexing Outlook mail... X/Y" text shown on the Update DB button while running
        self._onenote_progress_text = None  # v-onenote: same idea as _outlook_progress_text, for OneNote
        self._mailnotes_update_running = False  # v-onenote: True while _start_mail_notes_update() (Outlook and/or OneNote "mail/notes only" flow) is running — single master flag covering either/both stages; the Update DB BUTTON greys out for either this or a file-DB run (see _refresh_update_db_btn_lock)
        self._last_index_status_text = None  # v9.11: last real progress text (e.g. "AI 2/2: 28%"),
                                              # survives the button widget being recreated on minimize/reopen
        self._mft_file_res   = []        # v2.3: live MFT scan results (files)
        self._mft_folder_res = []        # v2.3: live MFT scan results (folders)
        self._mft_render_pending_f   = False  # v2.4: coalesced re-render flag (File Name tab)
        self._mft_render_pending_fol = False  # v2.4: coalesced re-render flag (Folder Name tab)
        # v7.7 FIX: race between the DB-backed search (_smart_search_realtime,
        # scans the FULL indexed corpus) and the live disk scan (_mft_scan_search,
        # only walks the user's home folder + non-C: drives -- a much narrower
        # scope). Whichever finished LAST used to blindly overwrite tree_f/
        # tree_fol, so a single generic keyword could flash a full DB result
        # set and then immediately get stomped down to the live scan's much
        # smaller subset the moment its (slower) os.walk finished. This flag
        # records which search id the DB results were last painted for, so
        # the live-scan renderer below can MERGE with them instead of wiping
        # them out.
        self._db_rendered_sid = -1
        self._db_rendered_file_res = []   # snapshot of DB file_res (files+folders) for merge
        self._adv_mode_active = False   # True = currently showing full (unfiltered) results
        self._adv_page = 0              # 0=realtime only, 1=all results shown
        self._adv_all_cont  = []        # full content result set from Advanced search
        self._adv_all_files = []        # full file result set from Advanced search
        self._adv_split_win = None      # Advanced split window (top/bottom)
        self._ai_split_win  = None      # AI Search split window (left/right)
        # Extension filter
        self.ext_filter_var = tk.StringVar(value="")
        # Content tab filter vars (separate from File Name tab)
        self.c_size_op_var   = tk.StringVar(value=">")
        self.c_size_num_var  = tk.StringVar(value="")
        self.c_size_unit_var = tk.StringVar(value="MB")
        self.c_ext_filter_var = tk.StringVar(value="")
        self.c_whole_word_dummy_var = tk.BooleanVar(value=True)  # display-only, greyed out (File Content tab — layout parity with File Name tab)
        self._outlook_meta = {}  # v-outlook: pseudo_path -> dict(entry_id, store_id, subject, sender, received, folder_path) for rows currently on screen
        self._onenote_meta = {}  # v-onenote: pseudo_path -> dict(page_id, title, notebook, section_path, created, last_modified) for rows currently on screen
        # Name filter vars (substring filter on filename) — both tabs
        self.name_filter_var   = tk.StringVar(value="")   # File Name tab
        self.c_name_filter_var = tk.StringVar(value="")   # File Content tab
        # AI model selector: key in SEMANTIC_MODELS ("jina_v3"/"bge_gemma2")
        self.ai_model_var = tk.StringVar(value=_sem_model_key)
        # v-new (yêu cầu: mở configure.ini ra thấy trống -> khó hiểu -- ghi
        # NGAY giá trị hiện tại (dù là mặc định hay đã có sẵn từ lần trước)
        # vào [General]/[Models] lúc khởi động, thay vì chỉ ghi khi user
        # ĐỔI 1 trong 3 lựa chọn này. [APIKeys] không cần ghi ở đây --
        # _load_api_keys_from_config() (đã gọi trước __init__, xem cuối
        # file) tự ghi nếu vừa migrate từ nguồn cũ; nếu user chưa từng
        # nhập key nào thì để trống là đúng, không có gì để "bịa" ra ghi.
        def _mutate_startup_snapshot(cfg):
            if not cfg.has_section("General"):
                cfg.add_section("General")
            cfg.set("General", "language", self._chat_lang)
            if not cfg.has_section("Models"):
                cfg.add_section("Models")
            cfg.set("Models", "ai_search_model", _sem_model_key)
            cfg.set("Models", "chat_ai_model", self._chat_preferred_model)
        _save_config(_mutate_startup_snapshot)

        # v2.8: this top bar (self.bg_f) is now the ONE persistent widget for the
        # search box across every state -- idle, expanded/typing, and full results.
        # Previously, the moment the first result arrived, a brand-new Toplevel
        # window was created (with its own icon/entry/history/close row) while this
        # window got hidden -- visually a "popup swap" that felt like a stutter.
        # Now show_results() just resizes THIS window and adds a results frame
        # below; this bar itself is never destroyed/recreated.
        # v4.9: the outer bar itself has NO border (icon + ramp light + close
        # button live here too) -- the border belongs only to the Entry below,
        # see the Entry creation further down.
        self.bg_f = tk.Frame(self.root, bg=BG_COLOR, height=35)
        self.bg_f.pack(fill="x", side="top")
        # v2.8.1: lock the bar's height so it can NEVER grow, no matter what gets
        # packed into it later (the AI Search / Advanced / Update DB buttons are
        # visually a bit taller than the entry due to their border+padding, and
        # without this they were quietly stretching the whole bar taller the
        # moment results opened — the "pop" the user was seeing).
        self.bg_f.pack_propagate(False)
        self._draw_search_icon(self.bg_f, size=16).pack(side="left", padx=(6, 3), pady=5)

        # v5.9: r_p sits flush against the Entry with no gap (padx below) so
        # the Entry's border reads as reaching all the way to the ramp dot
        # instead of visibly stopping short with dead space before it. r_p
        # itself stays borderless (a border around the ramp dot alone looked
        # like a separate boxed-in element, which wasn't wanted).
        r_p = self._r_p = tk.Frame(self.bg_f, bg=BG_COLOR)
        r_p.pack(side="right", fill="y", padx=(1, 5), pady=5)
        # v3.4: the small "H" (Search History) button that used to live below
        # the ramp light is gone -- Search History now lives permanently in
        # the "Help" tab instead. With only the ramp light left in this
        # column, it no longer needs to hug the top -- center it vertically
        # and make it bigger since it's now the only thing here.
        self.status_label = tk.Label(r_p, text="●", fg="#444", bg=BG_COLOR, font=("Arial", 13))
        self.status_label.pack(expand=True)
        # Ramp-light tooltip: Red/Yellow -> DB not fully indexed yet, Green -> ready.
        # While an Update DB run is actively in progress, show "Updating DB"
        # regardless of the ramp's current color (it can already read Green/
        # Yellow mid-run depending on stage) instead of the misleading
        # "Need Update DB". Callable so the text is re-evaluated fresh every
        # time the mouse hovers.
        def _ramp_tooltip_text():
            if getattr(self, "_update_db_running", False):
                return "Updating DB..."
            _fg = self.status_label.cget("fg")
            if _fg == "#2196f3":
                return "Data + AI Updated"
            if _fg == "#4caf50":
                return "Data Updated (AI pending)"
            return "Need Update DB"
        add_tooltip(self.status_label, _ramp_tooltip_text)

        # "✕" close button -- only shown once results are being displayed (there's
        # nothing to close in idle mode). Re-packed with before=r_p each time it's
        # shown so it always lands as the outermost-right widget on the bar.
        self.close_btn = tk.Button(self.bg_f, text="✕", font=("Arial", 9), bg=BG_COLOR, fg="#888",
                                    bd=0, activebackground=BG_COLOR)

        # "▼/▲" collapse/expand button -- sits between the "✕" close button and
        # the AI-model dropdown cluster. Unlike "✕" (which tears the whole
        # results window down), this only HIDES the results_frame (search
        # results / AI Search panes / AI Chat, notebook etc. -- nothing is
        # destroyed) and shrinks the window down to just the top bar (search
        # box + Update DB / AI Search / AI-model row). Clicking it again
        # re-packs results_frame and restores the window to its previous
        # size/position, with everything still exactly as it was. Only shown
        # once results are being displayed, same as close_btn.
        self.collapse_btn = tk.Button(self.bg_f, text="▼", font=("Arial", 9), bg=BG_COLOR, fg="#888",
                                       bd=0, activebackground=BG_COLOR, command=self._toggle_collapse)
        add_tooltip(self.collapse_btn, lambda: "Expand" if getattr(self, "_collapsed", False) else "Collapse")

        self.entry_var = tk.StringVar()
        # v7.3: the border used to be drawn via the Entry's own
        # highlightthickness ring -- at SMALL_SIZE (85px total window) that
        # ring's right-hand column was getting clipped/not painted (a Tk/
        # Win32 rendering quirk that only shows up once the Entry's own
        # allotted width gets very small; it was fine at LARGE_SIZE where
        # there's slack width). Drawing the border as an actual Frame
        # background color, with the Entry inset 1px inside it, means the
        # border is just a normal widget background -- it can't be clipped
        # the way an overlay ring can, at any window size.
        self.entry_border = tk.Frame(self.bg_f, bg="#8a8a8a")
        self.entry_border.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        self.entry = tk.Entry(self.entry_border, textvariable=self.entry_var, font=("Segoe UI", 10),
                               bg=ENTRY_BG, fg=TEXT_COLOR, insertbackground=TEXT_COLOR,
                               bd=0, highlightthickness=0)
        self.entry.pack(fill="both", expand=True, padx=1, pady=1)
        self.placeholder = tk.Label(self.entry, text=WAIT_PH, fg=PLACE_COLOR, bg=ENTRY_BG, font=("Segoe UI", 9, "italic"))
        if not self.has_args: self.placeholder.place(x=2, y=2)
        
        setup_context_menu(self.entry, self.entry, self.handle_action)
        setup_context_menu(self.placeholder, self.entry, self.handle_action)
        # v2.9 fix: the placeholder Label sits ON TOP of the Entry (via .place())
        # whenever the box is empty, so a left-click there was hitting the Label
        # instead of the Entry underneath -- the box looked clickable but nothing
        # got focus, so typing did nothing until the user closed (✕) and reopened.
        # Forward left-clicks (and click-drag-select) on the placeholder straight
        # to the real entry so it always gets focus + the caret.
        def _focus_entry_from_placeholder(e=None):
            self.entry.focus_set()
        self.placeholder.bind("<Button-1>", _focus_entry_from_placeholder)
        self.entry.bind("<FocusIn>", self.on_expand); self.entry.bind("<FocusOut>", self.on_shrink)
        self.entry.bind("<Return>", self._on_search_entry_return); self.entry.bind("<Escape>", lambda e: self.root.destroy())
        self.entry_var.trace_add("write", self.toggle_placeholder)
        
        self.entry.bind("<KeyRelease>", self.on_key_release)
        # v-fix: paste (right-click context menu's "Paste", which does
        # entry_widget.event_generate("<<Paste>>") -- see setup_context_menu
        # -- as well as some Ctrl+V paths) does NOT fire <KeyRelease>, so the
        # realtime debounced search silently never ran after a paste; the
        # user had to press Enter manually to see any results. <<Paste>>
        # fires as the paste is starting (text not inserted into the Entry
        # yet), so defer one tick before reading the box's new contents.
        self.entry.bind("<<Paste>>", lambda e: self.root.after(1, lambda: self.on_key_release(_types.SimpleNamespace(keysym=""))))
        self.bg_f.bind("<Button-1>", self.start_drag); self.bg_f.bind("<B1-Motion>", self.do_drag)

        # Results-mode state -- the notebook/trees/filter bars etc. built by
        # show_results() live in this frame, packed below self.bg_f only while
        # results are showing. self._results_extra_bar holds the AI Search /
        # Advanced / Update DB buttons that also live on the search bar row.
        self.results_frame = None
        self._results_extra_bar = None
        self._in_results_mode = False
        # v-new: collapse/expand state for the "▼/▲" button -- see
        # _toggle_collapse(). _collapsed tracks whether results_frame is
        # currently hidden; _pre_collapse_geometry remembers the window's
        # exact size/position (including any manual resize via the "◢"
        # grip) from right before collapsing, so expanding restores it
        # exactly instead of snapping back to some default size.
        self._collapsed = False
        self._pre_collapse_geometry = None
        # v4.4: set True right before a "Search again" re-run from the
        # History tab so the results notebook jumps to the File Name tab
        # once results are (re)rendered, instead of silently staying on
        # whatever tab (Help) the user triggered the re-search from.
        self._force_file_tab = False

        self.root.after(10, lambda: self.root.geometry(f"{SMALL_SIZE}+{self.x_pos}+5"))
        # v5.9c: also force a repaint shortly after the very first paint --
        # previously _force_repaint only ran on later shrink events (focus
        # loss, closing results), so the initial launch frame could show the
        # same cut-off-border/missing-ramp glitch with nothing to fix it.
        self.root.after(90, self._force_repaint)
        self.root.after(100, self.start_logic)
        self.root.after(10000, self.rotate_placeholder)
        if self.has_args: self.entry_var.set(" ".join(self.args)); self.root.after(300, self.handle_action)
        else: self.entry.focus_set()

    def rotate_placeholder(self):
        if not self.entry_var.get() and self.status_label.cget("fg") == "#4caf50":
            self.ph_index = (self.ph_index + 1) % len(self.ph_list)
            self.placeholder.config(text=self.ph_list[self.ph_index])
        self.root.after(10000, self.rotate_placeholder)

    def use_query_from_hist(self, val):
        # v3.4 FIX: this used to call self.active_result_win.destroy() first to
        # "force a fresh window" -- but since v2.8, active_result_win IS
        # self.root (results are shown in the same window, not a separate
        # Toplevel anymore). Destroying it destroyed the entire app, which is
        # exactly the crash ("application has been destroyed") seen when
        # right-clicking a history row -> "Search again". handle_action()
        # already fully rebuilds the results tabs on every search (same as
        # typing a new query + Enter while results are open), so no destroy
        # step is needed at all.
        #
        # v7.4 FIX: _force_file_tab used to be the ONLY way the tab switch
        # happened, and it was only consumed inside update_or_show_results --
        # which only runs once the DB-backed _smart_search_realtime thread
        # finishes. If db_conn wasn't ready yet (or that thread was just
        # slow), the switch silently never happened and the view stayed on
        # whatever tab (Help) "Search again" was clicked from. Since this
        # function only ever runs from a right-click inside the already-open
        # results window, self.nb is guaranteed to exist here -- so switch
        # tabs directly instead of waiting on an async callback.
        self._force_file_tab = True
        try:
            # v7.7: default results tab is now File Content (index 2), not
            # File Name (index 0) -- see matching change in show_results.
            self.nb.select(2)
        except Exception:
            pass
        self.entry_var.set(val)
        self.handle_action()
    def on_expand(self, e=None): 
        if not self.has_args and not self._in_results_mode: self.root.geometry(f"{LARGE_SIZE}+{self.x_pos}+5")
    def on_shrink(self, e=None):
        self.root.after(200, self._real_shrink)
    def _real_shrink(self):
        # v2.8: results mode manages its own window size now -- the entry losing
        # focus (e.g. clicking into the results tree) must NOT shrink the window
        # back down, since it's the same window the results are displayed in.
        if self._in_results_mode: return
        try:
            focused = self.root.focus_get()
        except KeyError:
            # Tkinter bug on Windows: Combobox popdown widget causes KeyError in focus_get()
            return
        if not self.has_args and focused != self.entry:
            self.root.geometry(f"{SMALL_SIZE}+{self.x_pos}+5")
            self.root.after(60, self._force_repaint)
    def _force_repaint(self):
        try:
            # v5.9e: THE confirmed actual fix (via the [RAMP DEBUG] dump) --
            # r_p's on-screen x-position was stale from a much wider layout
            # (e.g. left over from LARGE/results mode: rp_geom showed
            # x=1229 in a 115px-wide window, which is why it reported
            # unmapped -- that position is nowhere near this window at all).
            # Re-packing the PARENT (bg_f, as this function used to do) does
            # NOT force Tk to recompute a CHILD's position -- only
            # explicitly pack_forget()+pack()'ing r_p (and the Entry, same
            # risk) itself does that, recalculating against bg_f's CURRENT
            # (now narrow) width instead of whatever it was mid-results.
            try:
                self._r_p.pack_forget()
                self._r_p.pack(side="right", fill="y", padx=(1, 5), pady=5)
            except Exception:
                pass
            try:
                self.entry_border.pack_forget()
                self.entry_border.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
            except Exception:
                pass
            self.root.update_idletasks()
            # v5.9: previous attempts here (v5.2's alpha nudge, plus this
            # session's lift()/geometry nudges) all failed specifically for
            # the "user clicks empty Desktop" case -- the window loses OS
            # foreground focus, and Tk-side nudges apparently aren't enough
            # to make Windows actually recomposite it. Use the native Win32
            # RedrawWindow call instead: this directly tells Windows to
            # invalidate + immediately repaint the window (and all its
            # children), bypassing whatever DWM optimization was skipping
            # the repaint. Falls back to the old alpha nudge if this ever
            # fails (e.g. non-Windows, or winfo_id() unavailable).
            try:
                import ctypes
                hwnd = self.root.winfo_id()
                RDW_INVALIDATE, RDW_ERASE, RDW_UPDATENOW, RDW_ALLCHILDREN = 0x1, 0x4, 0x100, 0x80
                ok = ctypes.windll.user32.RedrawWindow(
                    hwnd, None, None,
                    RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN)
                if not ok:
                    raise RuntimeError("RedrawWindow returned 0")
            except Exception:
                self.root.attributes("-alpha", 0.99)
                self.root.after(30, lambda: self.root.attributes("-alpha", 1.0))
            try:
                self._r_p.lift()
                self.status_label.lift()
                self.status_label.config(bg=self.status_label.cget("bg"))
                self.status_label.update_idletasks()
            except Exception:
                pass
        except Exception:
            pass
    def start_drag(self, e): self._offsetx = e.x; self._offsety = e.y
    def do_drag(self, e):
        # v2.8: active_result_win is now the same window as self.root (results
        # render inline instead of in a separate Toplevel), so there is no longer
        # a second window to keep in sync here -- moving root IS moving results.
        x = self.root.winfo_x() + e.x - self._offsetx; y = self.root.winfo_y() + e.y - self._offsety
        self.root.geometry(f"+{x}+{y}"); self.x_pos = x

    def toggle_placeholder(self, *args):
        if self.entry_var.get(): self.placeholder.place_forget()
        else: self.placeholder.place(x=2, y=2)

    def _open_db_conn(self):
        """Open (or reopen) the persistent read connection with performance PRAGMAs."""
        try:
            if self.db_conn:
                try: self.db_conn.close()
                except: pass
            self.db_conn = sqlite3.connect(DB_FILE, timeout=10, check_same_thread=False)
            c = self.db_conn.cursor()
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("PRAGMA cache_size=-131072")   # 128 MB read cache
            c.execute("PRAGMA temp_store=MEMORY")
            c.execute("PRAGMA mmap_size=4294967296") # 4GB memory-mapped I/O
        except Exception as e:
            print(f"DB open error: {e}")
            self.db_conn = None

    def _ramp_watchdog_tick(self):
        """Safety net: re-apply _sync_ai_adv_lock() once a second for the
        life of the app. Cheap (a few cget/config calls, all guarded by
        winfo_exists) and makes the Advanced/AI Search button state
        self-healing — if any single call site that's supposed to trigger a
        re-sync ever gets missed (e.g. a future code path forgets to call
        it), the buttons still catch up within ~1s instead of staying wrong
        until the next unrelated sync happens to fire."""
        try:
            self._sync_ai_adv_lock()
        except Exception:
            pass
        try:
            self.root.after(1000, self._ramp_watchdog_tick)
        except Exception:
            pass

    def start_logic(self):
        # v3.5: 5-state lamp —
        #   Red    (#ff0000) : search_data.db doesn't exist yet — never indexed
        #   Yellow (#ffcc00) : search_data.db exists but content indexing never
        #                      finished (files table empty/missing)
        #   Yellow (blinking): an Update DB run is actively scanning/extracting
        #                      content right now (see indexing_worker/_ramp_blink_*)
        #   Green  (#4caf50) : BM25 content indexing finished — search/Advanced
        #                      are usable — but AI embeddings for the currently
        #                      selected model aren't fully built yet
        #   Blue   (#2196f3) : BM25 content AND AI embeddings for the currently
        #                      selected model are both fully up to date
        db_exists = os.path.exists(DB_FILE)
        has_data = False
        if db_exists:
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                # Migration: add size column if old DB doesn't have it yet
                c.execute("PRAGMA table_info(files)")
                existing_cols = [row[1] for row in c.fetchall()]
                if existing_cols and 'size' not in existing_cols:
                    c.execute("ALTER TABLE files ADD COLUMN size INTEGER DEFAULT 0")
                    conn.commit()
                c.execute("SELECT count(*) FROM files")
                if c.fetchone()[0] > 0: has_data = True
                conn.close()
            except: pass

        if not db_exists:
            # Red: no database.db at all yet — user should run --update data
            self.status_label.config(fg="#ff0000")
            self.placeholder.config(text=READY_PH)
        elif not has_data:
            # Yellow: database.db exists but hasn't finished being indexed
            # (e.g. app was closed mid-update, or DB was created but never
            # populated) — MFT filename search still works, content search won't
            self.status_label.config(fg="#ffcc00")
            self.placeholder.config(text=READY_PH)
        else:
            self._open_db_conn()
            self.status_label.config(fg="#4caf50"); self.placeholder.config(text=READY_PH)
            threading.Thread(target=self._preload_semantic_model, daemon=True).start()
            # v3.5: if the currently selected AI model's embeddings already
            # cover every content_store row (e.g. a previous --update data run
            # finished AI too), jump straight to Blue instead of staying Green
            # until the next Update DB run notices.
            # v7.10 FIX: the dropdown always starts back at DEFAULT_SEMANTIC_
            # MODEL ("jina_v3") on every app launch -- it isn't persisted. So
            # if Jina's build had failed in a previous run but another model
            # (e.g. BGE-Gemma2) succeeded, every future launch kept checking
            # Jina, found it incomplete, and stayed stuck on Green/AI-Search-
            # greyed-out forever, even though a perfectly usable AI index
            # existed under BGE. Now: if the current selection isn't fully
            # indexed, check every other model and auto-switch to the first
            # one that is.
            def _check_ai_startup():
                global _sem_model_key
                _cur_key = self.ai_model_var.get() if self.ai_model_var.get() in SEMANTIC_MODELS else DEFAULT_SEMANTIC_MODEL
                if not self._ai_fully_indexed(_cur_key):
                    for _mk in SEMANTIC_MODELS.keys():
                        if _mk != _cur_key and self._ai_fully_indexed(_mk):
                            print(f"[Semantic] '{_cur_key}' isn't fully indexed but '{_mk}' is -- "
                                  f"switching the AI model selection to '{_mk}'")
                            _cur_key = _mk
                            _sem_model_key = _mk
                            self.ai_model_var.set(_mk)

                            def _update_combo_display(k=_mk):
                                try:
                                    _combo = getattr(self, "_ai_model_combo", None)
                                    if _combo and _combo.winfo_exists():
                                        _combo.set(SEMANTIC_MODELS[k]["label"])
                                except Exception:
                                    pass
                            self.root.after(0, _update_combo_display)
                            break
                if self._ai_fully_indexed(_cur_key):
                    self.root.after(0, lambda: (self.status_label.config(fg="#2196f3"), self._sync_ai_adv_lock()))
            threading.Thread(target=_check_ai_startup, daemon=True).start()
        self._sync_ai_adv_lock()
        self.root.after(1000, self._ramp_watchdog_tick)
    
    def _search_by_image(self):
        """File Content tab 'Search files' button (was 'Search image') —
        pick ONE file of ANY supported type, extract its text (OCR for
        images, direct read for text files, pypdf/docx/openpyxl/pptx for
        documents — same extractor `get_file_content()` uses for indexing),
        and search directly using that text. No popup/confirmation — the
        point is a quick "does anything in my index look like THIS" lookup,
        so it goes straight from picking the file to seeing results. The
        file also gets added to the index along the way (so it's findable
        again later too), but that's a side effect, not the main point of
        this button.

        v9.16: previously this only ran OCR (PIL image decode), so picking
        a non-image file via the "All files" filter silently failed (PIL
        can't open a .txt as an image) and returned no text. Now routed
        through get_file_content(), which already dispatches by extension
        to the right extractor for text/PDF/Office/OneNote/Outlook/images —
        so the button (and its name) now genuinely covers "any file", not
        just pictures."""
        path = filedialog.askopenfilename(
            title="Search by file content",
            filetypes=[("All supported", " ".join(
                            f"*{e}" for e in sorted(_OCR_IMAGE_EXTS | _SEARCH_FILES_DOC_EXTS))),
                       ("Image files", " ".join(f"*{e}" for e in sorted(_OCR_IMAGE_EXTS))),
                       ("Document files", " ".join(f"*{e}" for e in sorted(_SEARCH_FILES_DOC_EXTS))),
                       ("All files", "*.*")])
        if not path:
            return
        try:
            self.entry.config(state="disabled", fg="#888888")
            self.entry_var.set("Scanning file...")
        except Exception:
            pass
        threading.Thread(target=self._search_by_image_worker, args=(path,), daemon=True).start()

    def _search_by_image_worker(self, path):
        """Runs off the UI thread. Extracts the file's text (see
        get_file_content() for per-extension handling), indexes it (files/
        content_store/content_index + best-effort AI embedding into any
        already-built model, same as Update DB does), then hands the
        extracted text back to the UI thread to apply to the search box
        and actually run the search."""
        content = ""
        try:
            content = self.get_file_content(path, force_ocr=True)
            conn = self.db_conn
            if conn is not None:
                c = conn.cursor()
                size = os.path.getsize(path)
                name = os.path.basename(path)
                mtime = os.path.getmtime(path)
                # v10.19 FIX: sanitize path/name/content AFTER the real
                # filesystem calls above (getsize/getmtime need the raw,
                # unmodified path to actually find the file on disk) but
                # BEFORE anything goes into SQLite -- see _sanitize_utf8.
                name = self._sanitize_utf8(name)
                path = self._sanitize_utf8(path)
                content = self._sanitize_utf8(content)
                c.execute("INSERT OR REPLACE INTO files (type, name, path, size) VALUES (?,?,?,?)",
                          ("File", name, path, size))
                c.execute("INSERT OR REPLACE INTO content_store (path, content, mtime) VALUES (?,?,?)",
                          (path, content, mtime))
                c.execute("DELETE FROM content_index WHERE path=?", (path,))
                if content:
                    c.execute("INSERT INTO content_index (path, content) VALUES (?,?)", (path, content))
                conn.commit()
                embed_text = (name + " " + content[:300]).replace("\n", " ")
                # v9.13.8 fix: only embed into whichever model is ALREADY
                # loaded in memory right now (_sem_loaded_key), instead of
                # looping through every model in SEMANTIC_MODELS. Only one
                # semantic model fits in RAM at a time (_load_semantic_model
                # evicts the previous one to load a new one) -- looping
                # through all of them here was causing a load/evict
                # ping-pong (load jina_v3 -> evict it to load bge_gemma2 ->
                # next real search needs jina_v3 again -> reload from disk)
                # that was the actual dominant cost behind the 30-60s
                # per-image time, not the OCR step itself. This is also a
                # nice-to-have side effect of Search Image, not its main
                # point (BM25 already finds the OCR'd text regardless of
                # AI embedding status), so skipping models that aren't
                # already loaded is the right tradeoff, not a regression.
                if _sem_loaded_key and _sem_ready:
                    try:
                        info = SEMANTIC_MODELS.get(_sem_loaded_key, {})
                        table = info.get("table")
                        if table:
                            c.execute(f"SELECT COUNT(*) FROM {table}")
                            if c.fetchone()[0] > 0:   # only if that model was actually built via Update DB
                                vec = _encode_passages([embed_text], _sem_loaded_key, batch_size=1)[0]
                                c.execute(f"INSERT OR REPLACE INTO {table} VALUES (?,?,?)",
                                          (path, embed_text, vec.astype("float32").tobytes()))
                                conn.commit()
                    except Exception as _e:
                        print(f"[Search files] embed failed for {_sem_loaded_key}: {_e}")
        except Exception as _e:
            print(f"[Search files] failed for {path}: {_e}")

        def _done():
            try:
                self.entry.config(state="normal", fg=TEXT_COLOR)
            except Exception:
                pass
            if not content or not content.strip():
                self.entry_var.set("")
                print(f"[Search files] No text could be read from {path}")
                return
            snippet = " ".join(content.strip().split()[:30])   # first ~30 words
            self.entry_var.set(snippet)
            self.entry.focus_set()
            self.entry.icursor(tk.END)
            class _FakeEvent:
                keysym = ""
            self.on_key_release(_FakeEvent())
        self.root.after(0, _done)

    def _open_office_doc_with_timeout(self, open_fn, kind, timeout=12):
        """v-new (yêu cầu: Update DB gặp file .doc/.ppt có mật khẩu thì phải
        đứng chờ hộp thoại nhập password/Cancel mới chạy tiếp): khi mở file
        .doc/.ppt có mật khẩu qua Word/PowerPoint COM automation mà không
        có đúng mật khẩu, Office hiển thị một hộp thoại MODAL THẬT của
        Windows đứng chờ người dùng bấm OK/Cancel -- đây KHÔNG phải là lỗi/
        exception Python nên try/except bình thường KHÔNG bắt được, nó
        treo cứng cả luồng Update DB cho tới khi có người tương tác với
        hộp thoại đó.

        Chạy toàn bộ open_fn() (Dispatch + Open + đọc text + Close + Quit)
        trên MỘT thread nền (daemon) riêng và chỉ đợi tối đa `timeout`
        giây. File mở bình thường (không mật khẩu) luôn xong trong
        khoảng dưới 1-2 giây nên timeout dài (12s) không ảnh hưởng gì tới
        file thường. Nếu quá `timeout` mà thread vẫn chưa xong -- gần như
        chắc chắn đang bị chặn bởi hộp thoại mật khẩu -- không đợi thêm
        nữa, coi như file này bị khoá và trả về "" ngay để Update DB tiếp
        tục file kế tiếp. Thread bị bỏ lại phía sau là daemon nên không
        giữ app/tiến trình chính; tiến trình WINWORD.EXE/POWERPNT.EXE mồ
        côi (nếu có) tự dọn khi Windows/app chính đóng, không làm treo
        Update DB nữa (đánh đổi chấp nhận được so với treo vô thời hạn)."""
        result = {}
        def _worker():
            try:
                result["val"] = open_fn()
                result["ok"] = True
            except Exception as e:
                result["err"] = e
                result["ok"] = False
        th = threading.Thread(target=_worker, daemon=True)
        th.start()
        th.join(timeout)
        if th.is_alive():
            print(f"[Extract] {kind}: có vẻ đang bị khoá bởi hộp thoại mật khẩu "
                  f"(quá {timeout}s không phản hồi) -- bỏ qua, tiếp tục file kế tiếp.")
            return ""
        if not result.get("ok"):
            return ""
        return result.get("val") or ""

    @staticmethod
    def _sanitize_utf8(s):
        """v10.19 FIX: on Windows, os.walk/os.listdir can return a filename
        containing a "lone surrogate" character (\\udc00-\\udfff) -- this
        happens for the rare file whose name isn't valid UTF-16 on disk
        (leftover from an old tool, a broken sync client, a different
        filesystem mounted in, etc.). Python deliberately allows this in
        path strings for filesystem round-tripping, but such a string
        can NEVER be encoded to real UTF-8 -- not by sqlite3 (which
        stores TEXT columns as UTF-8 internally), and not by anything
        else that writes it out. Without this, ONE such file anywhere on
        the drive would raise UnicodeEncodeError deep inside a batch
        INSERT and silently abort the ENTIRE indexing run (all files
        queued in that batch, not just the one bad file). Replacing the
        unencodable character with U+FFFD keeps the file discoverable
        (searchable by the rest of its name/content) instead of crashing
        everything. Applied to any path/content string right before it's
        queued for a SQLite INSERT -- see indexing_worker."""
        if not isinstance(s, str):
            return s
        return s.encode("utf-8", errors="replace").decode("utf-8")

    def get_file_content(self, filepath, force_ocr=False):
        if not filepath: return ""
        ext = os.path.splitext(filepath)[1].lower()
        text = ""
        try:
            print(f"[Extract] {ext}  {filepath}", flush=True)
            # v-fix (BUG: "Add file" / Update DB đọc file .pdf/.docx/... ra
            # toàn ký tự rác kiểu header "%PDF-1.7..." dù nội dung thật sự
            # rất dài): is_text_file() chỉ quyết định "đây có phải text
            # file không" bằng cách kiểm tra 1024 BYTE ĐẦU TIÊN có byte
            # \0 hay không -- với PDF/DOCX/XLSX/PPTX... (định dạng nhị
            # phân/ZIP), phần ĐẦU file không PHẢI LÚC NÀO cũng chứa byte
            # \0 trong 1024 byte đầu (tuỳ file cụ thể), nên is_text_file()
            # đôi khi trả về True NHẦM cho những file này. Vì check này
            # đứng TRƯỚC "elif ext == '.pdf':" bên dưới, khi nó trả về
            # True nhầm, file bị đọc bằng open(..., 'r', encoding='utf-8')
            # (chế độ text thô) thay vì qua pypdf/docx/openpyxl -- kết quả
            # chỉ là vài ký tự đọc được đầu tiên của header nhị phân, còn
            # lại toàn ký tự rác bị errors='ignore' bỏ qua, khiến AI Chat
            # tưởng file "không có nội dung" dù file có hàng chục trang
            # chữ thật. Giờ: những đuôi file ĐÃ CÓ handler riêng bên dưới
            # (_DEDICATED_HANDLER_EXTS) LUÔN đi qua đúng handler của nó,
            # is_text_file() chỉ còn dùng làm phương án CUỐI CÙNG cho các
            # đuôi file KHÔNG nằm trong danh sách đó (ví dụ .txt/.log/.py/
            # đuôi lạ chưa biết) -- không còn "cướp" mất các đuôi đã có
            # cách đọc đúng riêng nữa.
            if ext not in _DEDICATED_HANDLER_EXTS and is_text_file(filepath):
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    text = f.read(MAX_CHARS_TO_INDEX)
            elif ext == '.pdf':
                # v-fix (AI Chat báo "không có" dù nội dung thật sự nằm
                # trong file): this used to hard-cap extraction at the
                # FIRST 10 PAGES of every PDF, no matter how long the
                # document actually is. A topic covered further in (like
                # "Inertia Relief Modes" deep inside a training manual)
                # was never even READ off disk, let alone stored in
                # content_store -- so no amount of tuning
                # ONLINE_AI_SNIPPET_CHARS/ONLINE_AI_TOP_N_FILES/
                # ONLINE_AI_MAX_CONTEXT_FILES in AI Chat could ever surface
                # it; those only control how much of the ALREADY-INDEXED
                # text gets used, not how much of the PDF gets indexed in
                # the first place. Now extracts page by page until hitting
                # the same MAX_CHARS_TO_INDEX budget every other file type
                # already uses (instead of a fixed page count), so a
                # dense/long document isn't cut off arbitrarily early,
                # while still bounding extraction time for huge PDFs.
                import pypdf
                with open(filepath, 'rb') as f:
                    reader = pypdf.PdfReader(f)
                    for page in reader.pages:
                        text += (page.extract_text() or "") + " "
                        if len(text) >= MAX_CHARS_TO_INDEX:
                            break
            elif ext == '.docx':
                doc = docx.Document(filepath)
                text = " ".join([p.text for p in doc.paragraphs])
            elif ext == '.doc':
                def _open_doc():
                    import win32com.client
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    # v-new: DisplayAlerts=0 (wdAlertsNone) tắt bớt các hộp
                    # thoại cảnh báo không quan trọng của Word; kèm truyền
                    # một PasswordDocument SAI CỐ Ý -- nếu file thật sự có
                    # mật khẩu, Word thử mật khẩu này, thất bại, và ném lỗi
                    # ngay lập tức thay vì HIỆN HỘP THOẠI xin nhập mật khẩu
                    # (không ảnh hưởng file KHÔNG có mật khẩu -- tham số
                    # PasswordDocument chỉ được Word dùng khi file thực sự
                    # yêu cầu mật khẩu). _open_office_doc_with_timeout() bên
                    # ngoài là lưới an toàn cuối, phòng khi cách này không
                    # chặn được hộp thoại trên một số phiên bản Office.
                    try:
                        word.DisplayAlerts = 0
                    except Exception:
                        pass
                    try:
                        doc = word.Documents.Open(filepath, ReadOnly=True, Visible=False,
                                                   PasswordDocument="\x01")
                    except Exception:
                        # v-fix: PasswordDocument là tham số chuẩn/tài liệu
                        # hoá đầy đủ của Word (khác PowerPoint) nên hầu như
                        # luôn được chấp nhận -- nhưng vẫn thử lại một lần
                        # không kèm nó để không bỏ qua nhầm file HOÀN TOÀN
                        # BÌNH THƯỜNG chỉ vì một lỗi khác không liên quan
                        # tới mật khẩu. Nếu file thật sự có mật khẩu, lần
                        # thử lại này có thể hiện hộp thoại -- đó là lúc
                        # _open_office_doc_with_timeout() (timeout 12s) cắt
                        # ngang, không để treo.
                        try:
                            doc = word.Documents.Open(filepath, ReadOnly=True, Visible=False)
                        except Exception:
                            try:
                                word.Quit()
                            except Exception:
                                pass
                            raise
                    txt = doc.Range().Text
                    doc.Close(False)
                    word.Quit()
                    return txt
                text = self._open_office_doc_with_timeout(_open_doc, f".doc ({os.path.basename(filepath)})")
            elif ext in ['.xlsx', '.xls', '.csv']:
                if ext == '.xls':
                    import pandas as pd
                    df = pd.read_excel(filepath, engine='xlrd')
                    text = df.head(100).to_string(index=False)
                else:
                    df = pd.read_excel(filepath) if ext != '.csv' else pd.read_csv(filepath)
                    text = df.head(100).to_string(index=False)
            elif ext == '.pptx':
                prs = Presentation(filepath)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            text += shape.text + " "
            elif ext == '.ppt':
                def _open_ppt():
                    import win32com.client
                    app = win32com.client.Dispatch("PowerPoint.Application")
                    # v-new: cùng kỹ thuật như .doc ở trên -- PowerPoint's
                    # Presentations.Open cũng nhận tham số Password (đã xác
                    # nhận hoạt động qua nhiều bản Office 2010/2016/365);
                    # truyền một mật khẩu SAI CỐ Ý khiến PowerPoint báo lỗi
                    # ngay nếu file thật sự có mật khẩu, thay vì hiện hộp
                    # thoại xin nhập mật khẩu treo cả luồng Update DB. File
                    # KHÔNG có mật khẩu không bị ảnh hưởng gì.
                    try:
                        app.DisplayAlerts = 1  # ppAlertsNone
                    except Exception:
                        pass
                    try:
                        pres = app.Presentations.Open(filepath, ReadOnly=True, Untitled=False,
                                                        WithWindow=False, Password="\x01")
                    except Exception:
                        # v-fix: một số phiên bản PowerPoint không nhận
                        # tham số Password đúng như vậy qua late-bound COM
                        # (TypeError/pywintypes.com_error "Unknown name") --
                        # thử lại một lần KHÔNG kèm Password thay vì bỏ qua
                        # luôn cả những file HOÀN TOÀN BÌNH THƯỜNG không có
                        # mật khẩu. Nếu file thật sự có mật khẩu, lần thử
                        # lại này có thể vẫn hiện hộp thoại -- nhưng đó là
                        # đúng lúc _open_office_doc_with_timeout() bên
                        # ngoài (timeout 12s) sẽ cắt ngang, không để treo.
                        try:
                            pres = app.Presentations.Open(filepath, ReadOnly=True, Untitled=False,
                                                            WithWindow=False)
                        except Exception:
                            try:
                                app.Quit()
                            except Exception:
                                pass
                            raise
                    txt_parts = []
                    for slide in pres.Slides:
                        for shape in slide.Shapes:
                            if hasattr(shape, "TextFrame") and shape.TextFrame.HasText:
                                txt_parts.append(shape.TextFrame.TextRange.Text)
                    pres.Close()
                    app.Quit()
                    return " ".join(txt_parts)
                text = self._open_office_doc_with_timeout(_open_ppt, f".ppt ({os.path.basename(filepath)})")
            elif ext == '.one':
                # Direct binary extraction — works without OneNote running.
                # .one files store text as UTF-16-LE strings interspersed in binary data.
                try:
                    with open(filepath, 'rb') as f:
                        raw = f.read(MAX_CHARS_TO_INDEX * 8)
                    import re as _re
                    # Extract UTF-16-LE runs of printable chars (min 3 chars = 6 bytes)
                    utf16_chunks = _re.findall(rb'(?:[\x20-\x7e\x00][\x00]){3,}', raw)
                    parts = []
                    for chunk in utf16_chunks:
                        try:
                            s = chunk.decode('utf-16-le', errors='ignore').strip()
                            if len(s) >= 3 and not all(c in ' \t\n\r' for c in s):
                                parts.append(s)
                        except:
                            pass
                    # Also extract plain ASCII runs (page titles, tags often ASCII)
                    ascii_chunks = _re.findall(b'[\x20-\x7e]{4,}', raw)
                    for chunk in ascii_chunks:
                        try:
                            s = chunk.decode('ascii', errors='ignore').strip()
                            if len(s) >= 4:
                                parts.append(s)
                        except:
                            pass
                    text = " ".join(parts)
                except:
                    pass
            elif ext in _OCR_IMAGE_EXTS:
                # v-fix: OCR_ENABLED is the bulk-indexing ("Update DB" dialog)
                # toggle — it's meant to gate whether OCR runs across an
                # entire drive scan (slow, opt-in). Search Files(img) is a
                # single explicit user action on one picked image, where OCR
                # not running at all defeats the whole point of the feature
                # — so force_ocr bypasses that toggle for this call path.
                if OCR_ENABLED or force_ocr:
                    text = _run_ocr(filepath)
            elif ext == '.msg':
                try:
                    import win32com.client
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    ns = outlook.GetNamespace("MAPI")
                    msg_obj = ns.OpenSharedItem(os.path.abspath(filepath))
                    text = f"{msg_obj.Subject or ''} {msg_obj.Body or ''} {msg_obj.SenderName or ''} {msg_obj.SenderEmailAddress or ''}"
                    msg_obj.Close(0)
                except:
                    try:
                        # Fallback: read raw bytes and extract printable strings
                        with open(filepath, 'rb') as f:
                            raw = f.read(MAX_CHARS_TO_INDEX * 4)
                        import re as _re
                        # Extract UTF-16-LE strings (common in .msg) — catches full-width chars like ：
                        utf16_strings = []
                        for s in _re.findall(b'(?:[\x20-\x7e\x00-\xff]\x00){4,}', raw):
                            try:
                                decoded = s.decode('utf-16-le', errors='ignore').strip()
                                if decoded: utf16_strings.append(decoded)
                            except: pass
                        # Extract UTF-8 chunks — catches CJK, full-width punctuation ：＜＞ etc.
                        utf8_strings = []
                        try:
                            utf8_text = raw.decode('utf-8', errors='ignore')
                            import re as _re2
                            for chunk in _re2.findall(r'[\x20-\x7e\u3000-\u9fff\uff00-\uffef\u4e00-\u9fff]{3,}', utf8_text):
                                utf8_strings.append(chunk)
                        except: pass
                        ascii_strings = [s.decode('ascii', errors='ignore')
                                         for s in _re.findall(b'[\x20-\x7e]{4,}', raw)]
                        text = " ".join(utf16_strings + utf8_strings + ascii_strings)
                    except:
                        pass
        except: pass
        return text.strip()[:MAX_CHARS_TO_INDEX]

    def _preload_semantic_model(self):
        """Load AI model in background at startup — model ready immediately
        when needed. Also preloads OCR readers (see the `finally` block)
        AFTER this finishes, in this SAME thread/sequence rather than a
        separate concurrent thread -- two threads both trying to
        initialize a CUDA context at the same time was causing OCR's GPU
        init to fail ("Cannot copy out of meta tensor") and permanently
        fall back to CPU for the rest of the session. Running them one
        after another in the same thread avoids that race entirely while
        still keeping both off the main UI thread."""
        try:
            table = _semantic_table_for()
            try:
                conn = sqlite3.connect(DB_FILE, timeout=5)
                c = conn.cursor()
                c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table,))
                has_table = c.fetchone()[0] > 0
                count = 0
                if has_table:
                    c.execute(f"SELECT count(*) FROM {table}")
                    count = c.fetchone()[0]
                conn.close()
            except Exception as _e:
                print(f"[Semantic] Startup check failed: {_e}")
                return
            if count == 0:
                print(f"[Semantic] No {table} data — run --update data to build (model: {_sem_model_key})")
                return
            print(f"[Semantic] Loading model at startup ({count} indexed docs, {_sem_model_key})...")
            if _load_semantic_model():
                print("[Semantic] Model ready!")
                # Update hybrid status label in UI — just mark AI as ready, not active
                def _update_lbl():
                    if hasattr(self, '_hybrid_status_lbl'):
                        try:
                            self._hybrid_status_lbl.config(
                                text="📊 BM25  (AI ready)", fg="#7ec8e3")
                        except Exception: pass
                self.root.after(0, _update_lbl)
            else:
                print("[Semantic] Model load FAILED — check path and packages")
        finally:
            _load_ocr_readers()

    def _ai_fully_indexed(self, model_key=None):
        """True when the given model's (or, if omitted, the currently
        selected model's) embedding table covers every row currently in
        content_store — i.e. AI Search data is completely up to date, not
        just BM25/content. Used to decide whether the ramp light should read
        Green (BM25 ready, AI pending) or Blue (BM25 + AI both ready). Safe
        to call from a background thread — opens its own short-lived
        connection rather than touching self.db_conn."""
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            c = conn.cursor()
            c.execute("SELECT count(*) FROM content_store")
            total = c.fetchone()[0]
            if total == 0:
                conn.close()
                return False
            table = _semantic_table_for(model_key)
            c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table,))
            if not c.fetchone()[0]:
                conn.close()
                return False
            c.execute(f"SELECT count(*) FROM {table}")
            sem_count = c.fetchone()[0]
            conn.close()
            return sem_count >= total
        except Exception:
            return False

    _RAMP_BLINK_COLORS = {
        "yellow": ("#ffcc00", "#6b5200"),  # BM25/content indexing stage
        "green":  ("#4caf50", "#1f4623"),  # AI/semantic embedding stage
        "blue":   ("#2196f3", "#0d3d5c"),  # v-outlook: Outlook mail indexing (only used when the ramp is already solid Blue — see _start_outlook_only_update)
    }

    def _ramp_blink_start(self, color="yellow"):
        """Start blinking the ramp-light while an Update DB run is actively
        working. color="yellow" for the BM25/content stage, color="green"
        for the AI/semantic-embedding stage that follows it — lets you tell
        the two stages apart at a glance instead of both looking the same.
        Call _ramp_blink_stop() once all work finishes — success snaps to a
        solid Green/Blue (see indexing_worker), failure snaps to solid Red.

        Calling this again with a DIFFERENT color while already blinking
        (e.g. switching from the yellow BM25 stage straight into the green
        AI stage) just re-colors the ongoing blink in place, without
        stopping/restarting the loop — this is what keeps the light
        continuously blinking across the stage transition instead of
        pausing solid in between.

        Idempotent by design: this gets called from more than one place for
        the same run (the button/command handler AND indexing_worker's own
        startup), and without the guard below each call would spin up its
        own independent self-rescheduling after() loop — two loops toggling
        the same shared _ramp_blink_on flag in close succession effectively
        cancel each other out, which is why the dot used to look static
        instead of blinking."""
        self._ramp_blink_on_color, self._ramp_blink_off_color = \
            self._RAMP_BLINK_COLORS.get(color, self._RAMP_BLINK_COLORS["yellow"])
        if getattr(self, "_ramp_blinking", False):
            return  # already blinking -- color above takes effect on the next tick
        self._ramp_blinking = True
        self._ramp_blink_on = True
        self._ramp_blink_tick()

    def _ramp_blink_tick(self):
        if not getattr(self, "_ramp_blinking", False):
            return
        try:
            on_c = getattr(self, "_ramp_blink_on_color", "#ffcc00")
            off_c = getattr(self, "_ramp_blink_off_color", "#6b5200")
            self.status_label.config(fg=on_c if self._ramp_blink_on else off_c)
        except Exception:
            pass
        self._ramp_blink_on = not self._ramp_blink_on
        try:
            self.root.after(500, self._ramp_blink_tick)
        except Exception:
            pass

    def _ramp_blink_stop(self, final_fg=None):
        """Stop blinking. If final_fg is given, snap the ramp to that solid
        color; otherwise leave whatever color the last tick painted."""
        self._ramp_blinking = False
        if final_fg is not None:
            try:
                self.status_label.config(fg=final_fg)
            except Exception:
                pass

    def _begin_update_ramp(self, text):
        """Called once at the very start of indexing_worker (from the main
        thread via root.after) — sets the Update DB button/status text and
        starts the Yellow blink that runs through the whole BM25 content
        stage."""
        self._set_index_status(text, "#ffcc00")
        self._ramp_blink_start()

    def _sync_ai_adv_lock(self):
        """Grey-out (disable) the Advanced button while the ramp light is Red
        or Yellow (DB content not fully indexed yet); re-enabled the moment
        it turns Green or Blue.

        The AI Search button and the Jina/BGE model dropdown are stricter
        still: they require the ramp to be solid Blue (BM25 content AND the
        currently selected AI model's embeddings both fully up to date).
        Green is NOT enough for AI Search — Green can mean BM25 is ready but
        AI embeddings are still being built in the background (e.g. "AI
        2/2: 20%..."), and AI Search/the model dropdown must stay greyed out
        through all of that, only unlocking once the ramp actually reaches
        Blue. Safe to call any time -- widgets may not exist yet (idle mode,
        before results are shown) so every lookup is guarded."""
        try:
            _fg = self.status_label.cget("fg")
            # v-outlook: treat the on/off shades of a blink as equivalent —
            # both variants of Green (BM25-ready/AI-in-progress) and Blue
            # (fully ready) should read as ready/blue throughout the whole
            # blink cycle, not just during the brighter "on" tick. Without
            # this, e.g. the Outlook-mail blue blink (or the existing AI-
            # stage green blink) would make AI Search/Advanced flicker
            # disabled every ~500ms on the "dim" tick even though nothing
            # about real DB readiness changed at all.
            _READY_FGS = {"#4caf50", "#1f4623", "#2196f3", "#0d3d5c"}
            _BLUE_FGS = {"#2196f3", "#0d3d5c"}
            is_ready = _fg in _READY_FGS      # Green or Blue (either blink shade) = BM25 content ready
            is_blue = _fg in _BLUE_FGS        # Blue (either blink shade) = BM25 + AI both fully ready
        except Exception:
            is_ready = False
            is_blue = False
        updating = bool(getattr(self, "_update_db_running", False))
        adv_state = "normal" if is_ready else "disabled"
        ai_ready = is_blue and not updating and ENABLE_AI_SEARCH_FEATURE
        ai_state = "normal" if ai_ready else "disabled"
        combo_state = "readonly" if ai_ready else "disabled"
        try:
            _adv = getattr(self, "_adv_search_btn", None)
            if _adv and _adv.winfo_exists():
                _adv.config(state=adv_state)
        except Exception:
            pass
        try:
            _ai = getattr(self, "_ai_search_btn", None)
            if _ai and _ai.winfo_exists():
                _ai.config(state=ai_state)
        except Exception:
            pass
        try:
            _combo = getattr(self, "_ai_model_combo", None)
            if _combo and _combo.winfo_exists():
                _combo.config(state=combo_state)
        except Exception:
            pass

    def _refresh_update_db_btn_lock(self):
        """v-outlook/v-onenote: keep the Update DB button greyed out while
        EITHER a file-DB update OR a mail/notes-only update (Outlook
        and/or OneNote — see _start_mail_notes_update) is running — any of
        these means "don't let the user start another update run right
        now". Safe to call any time; the button may not exist yet (idle
        mode)."""
        try:
            btn = getattr(self, "_update_db_btn", None)
            if not (btn and btn.winfo_exists()):
                return
            any_running = bool(getattr(self, "_update_db_running", False)) or \
                           bool(getattr(self, "_mailnotes_update_running", False))
            if any_running:
                text = self._last_index_status_text or "Updating..."
                if getattr(self, "_mailnotes_update_running", False) and not getattr(self, "_update_db_running", False):
                    # Outlook's stage text takes priority while it's still
                    # running; once it clears (stage finished), OneNote's
                    # takes over — see _start_mail_notes_update.
                    text = (getattr(self, "_outlook_progress_text", None) or
                            getattr(self, "_onenote_progress_text", None) or
                            "Updating...")
                btn.config(state="disabled", text=text, fg="#ffcc00", bg="#2a2a1a")
            else:
                btn.config(state="normal", text="Update DB", fg="#33363c", bg="#e6e8ec")
        except Exception:
            pass

    def _set_index_status(self, text, fg, done=False, error=False):
        """Central place indexing_worker() reports progress through — updates
        the small status dot on the root searchbox AND the Update DB button
        (if the Results window happens to be open right now), so whichever
        one the user is currently looking at shows live progress.
        done/error also resets self._update_db_running and re-enables the button.

        The ramp dot itself only ever shows the "●" glyph and changes color
        (grey/red/yellow/green) — it never swaps to the progress text (that
        would duplicate what the Update DB button/tooltip already say and
        made the dot visually jump between a dot and a sentence). The
        Update DB button, on the other hand, does show the live text.

        v9.11 fix: also remember `text` in self._last_index_status_text.
        The search box (and its Update DB button) gets torn down and
        recreated when minimized/reopened — without remembering the last
        real progress string here, the freshly-recreated button had no way
        to know we were mid-run at "AI 2/2: 28%" and fell back to a bare
        "Updating..." with no percentage, discarding progress info that
        was still perfectly valid, just not visible anymore."""
        if not done and not error:
            self._last_index_status_text = text
        try:
            self.status_label.config(text="●", fg=fg)
        except Exception:
            pass
        self._sync_ai_adv_lock()
        if done or error:
            self._update_db_running = False
            self._last_index_status_text = None
        try:
            btn = self._update_db_btn
            if btn and btn.winfo_exists():
                if done:
                    btn.config(state="normal", text="Update DB", fg="#33363c", bg="#e6e8ec")
                elif error:
                    btn.config(state="normal", text="Update DB (error)", fg="#ff5555", bg="#3a1a1a")
                else:
                    btn.config(state="disabled", text=text, fg="#ffcc00", bg="#2a2a1a")
        except Exception:
            pass

    _MAILNOTES_STALL_WARN_SEC = 15 * 60  # v-watchdog: warn once after this long with zero progress

    def _check_mailnotes_watchdog(self):
        """Self-rescheduling check (every 60s) while an Outlook/OneNote
        indexing run is active: if progress_cb hasn't fired in over
        _MAILNOTES_STALL_WARN_SEC, warn once. This does NOT and CANNOT
        stop/unstick a genuinely hung COM call -- Python has no way to
        interrupt a blocked win32 call from another thread -- it only
        surfaces the situation promptly instead of the user discovering
        it themselves after an hour or two of silence. If this fires,
        the practical next step is to close this app and Outlook and
        restart the update (already-indexed items are safe -- both
        indexers are incremental)."""
        if not self._mailnotes_update_running:
            return  # run finished (or was never one) -- stop rescheduling
        stalled_for = time.time() - getattr(self, "_mailnotes_last_progress_ts", time.time())
        if stalled_for > self._MAILNOTES_STALL_WARN_SEC and not getattr(self, "_mailnotes_stall_warned", False):
            self._mailnotes_stall_warned = True
            messagebox.showwarning(
                "Outlook/OneNote update",
                f"No progress in over {int(stalled_for // 60)} minutes.\n\n"
                "This usually means Outlook's COM automation has stalled on a "
                "specific item or folder (a known Outlook automation issue, "
                "not something this app can force past). Already-indexed "
                "items are safe.\n\n"
                "If it doesn't recover on its own, close this app and Outlook, "
                "reopen Outlook, then run Update again — it will resume "
                "where it left off.")
        self.root.after(60_000, self._check_mailnotes_watchdog)

    def _start_mail_notes_update(self, do_outlook=True, do_onenote=True, on_complete=None):
        """Index Outlook mail and/or OneNote pages — completely independent
        of DB_FILE/db_conn/the ramp light (see outlook_search.py's and
        onenote_search.py's module docstrings for why each has its own db
        file). Shared by the Update DB dialog's Outlook/OneNote checkboxes
        and the standalone '--update outlook' / '--update onenote'
        Searchbox commands.

        do_outlook / do_onenote: which stage(s) to run. If both are True,
        Outlook always runs first, then OneNote — sequential, not
        parallel, since both go through COM automation and running two
        different Office apps' COM servers at once is more likely to
        cause contention than to save real time.

        on_complete: optional zero-arg callback fired once this whole
        stage is done (after the "indexed X, skipped Y" summary popup, if
        any) — used to chain a file-DB update right after, for the
        "update everything, unattended" option in the dialog. Not fired
        if this call was rejected outright (already running, module
        unavailable, etc.) — those return before starting the worker at
        all, so there's nothing to chain onto.

        v-fix: status_label is the ramp *light* — it should only ever show
        "●" and change/blink COLOR (same as the file-DB flow's _set_index_status,
        which also never touches the label's text, only its fg). Progress
        TEXT belongs on the Update DB button instead, exactly like the
        file-DB flow already does it (see _refresh_update_db_btn_lock).
        _sync_ai_adv_lock() reads the ramp's color (blink-shade-tolerant,
        see its _READY_FGS/_BLUE_FGS) to decide whether AI Search/model
        dropdown should be enabled. If the ramp is already solid Blue when
        this starts, we blink it (on/off Blue, both still read as "ready")
        purely for visual feedback — the real DB state hasn't changed, so
        this is safe. If the ramp ISN'T Blue (DB/AI not fully built yet), we
        deliberately do NOT force a blue blink — that would misrepresent
        real readiness.
        """
        if do_outlook and not OUTLOOK_SEARCH_AVAILABLE:
            messagebox.showwarning("Outlook", "outlook_search.py not found next to this script.")
            return
        if do_onenote and not ONENOTE_SEARCH_AVAILABLE:
            messagebox.showwarning("OneNote", "onenote_search.py not found next to this script.")
            return
        if not do_outlook and not do_onenote:
            return
        if self._mailnotes_update_running or self._update_db_running:
            messagebox.showinfo("Update", "An update is already running.")
            return
        self._mailnotes_update_running = True
        self._mailnotes_last_progress_ts = time.time()  # v-watchdog: see _check_mailnotes_watchdog
        self._mailnotes_stall_warned = False
        if do_outlook:
            self._outlook_progress_text = "Indexing Outlook mail..."
        if do_onenote:
            self._onenote_progress_text = "Indexing OneNote..." if not do_outlook else None
        self._refresh_update_db_btn_lock()
        self._check_mailnotes_watchdog()  # kicks off its own self-rescheduling loop
        _ramp_fg = self.status_label.cget("fg")
        _was_blue = (_ramp_fg == "#2196f3")
        if _was_blue:
            self._ramp_blink_start(color="blue")

        def _worker():
            o_indexed = o_skipped = 0
            o_err = None
            n_indexed = n_skipped = n_locked = 0
            n_err = None

            if do_outlook:
                def _prog_o(done, total, subj):
                    self._outlook_progress_text = f"Indexing Outlook mail... {done}/{total}"
                    self._mailnotes_last_progress_ts = time.time()
                    self._mailnotes_stall_warned = False
                    self.root.after(0, self._refresh_update_db_btn_lock)
                o_indexed, o_skipped, o_err = outlook_search.index_outlook_mail(progress_cb=_prog_o)
                self._outlook_progress_text = None

            if do_onenote:
                # Runs regardless of whether the Outlook stage succeeded —
                # one failing shouldn't block the other from being attempted.
                def _prog_n(done, total, subj):
                    self._onenote_progress_text = f"Indexing OneNote... {done}/{total}"
                    self._mailnotes_last_progress_ts = time.time()
                    self._mailnotes_stall_warned = False
                    self.root.after(0, self._refresh_update_db_btn_lock)
                n_indexed, n_skipped, n_locked, n_err = onenote_search.index_onenote(progress_cb=_prog_n)
                self._onenote_progress_text = None

            def _done():
                self._mailnotes_update_running = False
                self._outlook_progress_text = None
                self._onenote_progress_text = None
                if _was_blue:
                    self._ramp_blink_stop(final_fg="#2196f3")  # snap back to solid Blue
                self._refresh_update_db_btn_lock()
                msgs = []
                if do_outlook:
                    if o_err:
                        msgs.append(f"Outlook: {o_err}")
                    else:
                        print(f"[Outlook] {o_indexed} mail indexed, {o_skipped} unchanged.")
                if do_onenote:
                    if n_err:
                        msgs.append(f"OneNote: {n_err}")
                    else:
                        print(f"[OneNote] {n_indexed} page indexed, {n_skipped} unchanged, {n_locked} locked/unreadable.")
                if msgs:
                    messagebox.showwarning("Update", "\n".join(msgs))
                if on_complete:
                    on_complete()
            self.root.after(0, _done)
        threading.Thread(target=_worker, daemon=True).start()

    def _start_outlook_only_update(self):
        """Backward-compatible wrapper — Outlook only. See
        _start_mail_notes_update for the real implementation."""
        self._start_mail_notes_update(do_outlook=True, do_onenote=False)

    def _start_onenote_only_update(self):
        """OneNote only. See _start_mail_notes_update for the real
        implementation."""
        self._start_mail_notes_update(do_outlook=False, do_onenote=True)

    def _show_update_db_summary(self, started_at, selected_tiers, selected_models, ocr_enabled, force_reindex):
        """v-new: small popup shown once an Update DB run finishes -- when
        it started/ended, how long it took, and exactly which options were
        in effect (tiers, AI models, OCR, Force re-index) so there's no
        need to scroll back through console output to check what a run
        actually covered."""
        ended_at = datetime.now()
        total_seconds = int((ended_at - started_at).total_seconds())
        hh, rem = divmod(total_seconds, 3600)
        mm, ss = divmod(rem, 60)
        dur_str = f"{hh}h {mm}m {ss}s" if hh else (f"{mm}m {ss}s" if mm else f"{ss}s")

        _tier_names = {0: "Tier 1 (Office/PDF)", 1: "Tier 2", 2: "Tier 3", 3: "Tier 4"}
        if selected_tiers is None:
            tiers_str = "All (Tier 1-4)"
        else:
            tiers_str = ", ".join(_tier_names.get(t, f"Tier {t + 1}") for t in sorted(selected_tiers))

        if selected_models is None:
            models_str = "All (" + ", ".join(v.get("label", k) for k, v in SEMANTIC_MODELS.items()) + ")"
        elif len(selected_models) == 0:
            models_str = "None (Tier-only -- AI Search not updated this run)"
        else:
            models_str = ", ".join(SEMANTIC_MODELS.get(k, {}).get("label", k) for k in selected_models)

        msg = (
            f"Update DB completed.\n\n"
            f"Started:  {started_at.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Finished: {ended_at.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Duration: {dur_str}\n\n"
            f"Tiers: {tiers_str}\n"
            f"AI models: {models_str}\n"
            f"OCR images: {'Yes' if ocr_enabled else 'No'}\n"
            f"Force re-index: {'Yes' if force_reindex else 'No'}"
        )
        try:
            messagebox.showinfo("Update DB", msg, parent=self.root)
        except Exception:
            messagebox.showinfo("Update DB", msg)

    def indexing_worker(self, selected_tiers=None, selected_models=None, ocr_enabled=False, force_reindex=False):
        """selected_tiers: None = all tiers (1-4, default/backward-compatible).
        Otherwise a set of 0-indexed tier numbers (0=Tier1 ... 3=Tier4) — only
        files in these tiers get their CONTENT extracted/embedded. Filename
        metadata (files table, used by filename search) is still built for
        every file regardless of tier filter — the tier filter only limits
        the (slow) content-extraction step.
        selected_models: None = build AI embeddings for every model in
        SEMANTIC_MODELS (default/backward-compatible). An empty list/set =
        build NO AI embeddings this run (Tier-only update — BM25/content
        only, AI stage entirely skipped). Otherwise a list/set of model_key
        strings — only these get embeddings built this run, skipping the
        others entirely (saves a lot of time if the user only cares about
        one model, e.g. just Jina-v3).
        ocr_enabled: (v9.13) False by default — when True, image files
        (.jpg/.png/...) also get their content extracted via OCR
        (EasyOCR), same as Office/PDF/text files. Off by default because
        OCR is noticeably slower than the other extraction methods and
        downloads its own model weights on first use — an explicit opt-in
        via the Update DB dialog's "OCR images" checkbox.
        force_reindex: (v-new) False by default — a file whose mtime hasn't
        changed since it was last indexed is normally SKIPPED entirely
        (existing content_store/content_index row kept as-is, see the
        mtime check below), which is what makes routine Update DB runs
        fast. When True, that skip is bypassed and every file's content is
        re-extracted regardless of mtime -- e.g. after a code fix like the
        PDF page-count cap changing what gets read out of a PDF, so
        already-indexed files actually pick up the new extraction instead
        of being silently left with their old, incomplete content forever
        (their mtime on disk never changed, so they'd otherwise never be
        touched again). Opt-in via the Update DB dialog's "Force re-index"
        checkbox since it makes the run as slow as a first-time scan."""
        global OCR_ENABLED
        OCR_ENABLED = bool(ocr_enabled)
        # v-new: for the "Update DB completed" summary dialog shown at the
        # end (see _finish below) -- when this run started, plus a copy of
        # the options it was called with (captured here since they're
        # ordinary function args, cheapest way to carry them through to
        # the end of a long function).
        _run_started_at = datetime.now()
        try:
            self._update_db_running = True
            self.root.after(0, lambda: self._begin_update_ramp("Updating..."))
            # Close any existing persistent connection before rebuilding DB
            if self.db_conn:
                try: self.db_conn.close()
                except: pass
                self.db_conn = None
            conn = sqlite3.connect(DB_FILE, timeout=30)
            c = conn.cursor()
            # v10.19 FIX: this connection used to have no timeout at all
            # (defaulted to sqlite3's built-in 5s) and never set WAL mode
            # itself -- if self.db_conn (the persistent realtime-search
            # reader, opened elsewhere with WAL already on) hadn't fully
            # released its file handle yet (e.g. Windows still finishing a
            # close() from a background search thread at the exact moment
            # "Update DB" was clicked), the very first CREATE TABLE below
            # could hit "database is locked" and abort the whole indexing
            # run. WAL mode lets one writer and any number of readers work
            # at the same time without blocking each other -- setting it
            # here too (harmless/no-op if the file is already in WAL mode)
            # plus a generous 30s timeout makes this connection wait out
            # any brief contention instead of failing immediately.
            c.execute("PRAGMA journal_mode=WAL")
            c.execute("PRAGMA busy_timeout=30000")
            c.execute("CREATE TABLE IF NOT EXISTS history (id INTEGER PRIMARY KEY, query TEXT, date TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS files (type TEXT, name TEXT, path TEXT, size INTEGER)")
            c.execute("CREATE VIRTUAL TABLE IF NOT EXISTS content_index USING fts5(path, content, tokenize='trigram case_sensitive 0')")
            c.execute("CREATE TABLE IF NOT EXISTS content_store (path TEXT PRIMARY KEY, content TEXT, mtime REAL)")
            # v1.5: checkpoint table — tracks resume state across interrupted runs.
            # stage: 'scan' | 'content' | 'embed'
            # model_key: NULL for scan/content stages, model name for embed stage
            # last_path: last successfully processed path (resume point)
            c.execute("""CREATE TABLE IF NOT EXISTS update_checkpoint (
                            stage     TEXT,
                            model_key TEXT,
                            last_path TEXT,
                            done_count INTEGER,
                            total_count INTEGER,
                            updated_at TEXT,
                            PRIMARY KEY (stage, model_key)
                         )""")
            conn.commit()

            def _save_checkpoint(stage, model_key, last_path, done_count, total_count):
                try:
                    c.execute("""INSERT INTO update_checkpoint
                                 (stage, model_key, last_path, done_count, total_count, updated_at)
                                 VALUES (?,?,?,?,?,?)
                                 ON CONFLICT(stage, model_key) DO UPDATE SET
                                 last_path=excluded.last_path, done_count=excluded.done_count,
                                 total_count=excluded.total_count, updated_at=excluded.updated_at""",
                              (stage, model_key or '', last_path, done_count, total_count,
                               datetime.now().isoformat()))
                    conn.commit()
                except Exception:
                    pass

            def _get_checkpoint(stage, model_key):
                try:
                    c.execute("SELECT last_path, done_count, total_count FROM update_checkpoint WHERE stage=? AND model_key=?",
                              (stage, model_key or ''))
                    row = c.fetchone()
                    return row if row else (None, 0, 0)
                except Exception:
                    return (None, 0, 0)

            def _clear_checkpoint(stage, model_key=None):
                try:
                    if model_key is None:
                        c.execute("DELETE FROM update_checkpoint WHERE stage=?", (stage,))
                    else:
                        c.execute("DELETE FROM update_checkpoint WHERE stage=? AND model_key=?", (stage, model_key))
                    conn.commit()
                except Exception:
                    pass

            # ── STAGE 1: File scan ────────────────────────────────────────────
            # File scan is fast (minutes) — always rerun fully to catch new/deleted/moved
            # files. Not the bottleneck, so no resume needed here.
            c.execute("DROP TABLE IF EXISTS files")
            c.execute("CREATE TABLE files (type TEXT, name TEXT, path TEXT, size INTEGER)")
            c.execute("DROP TABLE IF EXISTS files_temp")
            c.execute("CREATE TABLE files_temp (type TEXT, name TEXT, path TEXT, size INTEGER)")
            # content_index (FTS5) + content_store are content caches — do NOT drop them.
            # We diff against existing mtime to skip unchanged files (huge time save on resume).
            c.execute("CREATE VIRTUAL TABLE IF NOT EXISTS content_index USING fts5(path, content, tokenize='trigram case_sensitive 0')")
            c.execute("CREATE TABLE IF NOT EXISTS content_store (path TEXT PRIMARY KEY, content TEXT, mtime REAL)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_content_store_path ON content_store(path)")
            # v1.5: migrate old DBs (pre-checkpoint) that lack the mtime column
            c.execute("PRAGMA table_info(content_store)")
            _cs_cols = [row[1] for row in c.fetchall()]
            if 'mtime' not in _cs_cols:
                c.execute("ALTER TABLE content_store ADD COLUMN mtime REAL")
                conn.commit()
                print("[Migration] Added mtime column to content_store (old DB detected)")

            # Load existing content_store mtimes — used to skip re-extracting unchanged files
            c.execute("SELECT path, mtime FROM content_store")
            _existing_mtime = {r[0]: r[1] for r in c.fetchall()}
            # v1.5: semantic tables are now resumed/appended, never dropped here.
            # Table creation + resume logic happens later in the embedding loop below.
            
            batch_files = []; batch_content = []
            _pending_extract = []  # v2.5: (tier, path, mtime) — extracted AFTER full walk, sorted by tier
            _scan_total = 0; _scan_skipped = 0; _scan_new = 0
            _seen_paths = set()  # v1.5: deduplicate — prevents double-extract if drives overlap

            # v1.5: known text extensions — skip is_text_file() open() call for these
            # is_text_file() reads 1024 bytes from every unknown file → huge bottleneck
            # with millions of files. Whitelist covers 99% of indexable text files.
            # v2.5: trimmed compiled-language source extensions (.c/.h/.cpp/.java/.cs/
            # .go/.rs/.php/.sh) — these are almost always installer/SDK payload noise
            # (Abaqus, CATIA, etc.) in an engineering file share, not content users
            # actually search for. Add them back below if you do want them indexed.
            _TEXT_EXTS = {
                '.txt', '.md', '.rst', '.csv', '.log', '.ini', '.cfg', '.conf',
                '.json', '.xml', '.yaml', '.yml', '.toml', '.html', '.htm',
                '.py', '.js', '.ts', '.css', '.sql', '.bat', '.ps1',
                '.tsv', '.spck', '.env',
            }
            # .env / .env.local / .env.production etc. are dotfiles — Python's
            # splitext() treats the leading dot as the filename, not an extension
            # (splitext(".env") == (".env", "")), so they'd never match _TEXT_EXTS
            # by extension. Match by basename prefix instead.
            # NOTE: .env files commonly hold API keys/credentials — indexing their
            # content means those secrets become searchable/readable via the app.
            # Keep this only if that's an acceptable tradeoff in your environment.
            _ENV_FILE_PREFIX = ".env"
            # Extensions that are definitively binary — never open, never extract
            _BINARY_EXTS = {
                '.exe', '.dll', '.lib', '.obj', '.pyc', '.bin',
                '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.ico', '.svg', '.webp',
                '.zip', '.7z', '.rar', '.tar', '.gz', '.bz2',
                '.mp3', '.mp4', '.avi', '.mov', '.mkv', '.wav', '.flac',
                '.db', '.sqlite', '.ldb', '.sst',
                '.lock', '.tmp', '.bak', '.cache',
            }

            # v1.5: Whitelist-based scan — much faster than blacklisting system folders.
            # Strategy: scan C:\Users\<current_user> + all non-C drives (D, E, F...) fully.
            # This covers 99% of user documents while skipping Windows/app install folders.
            import getpass as _getpass
            _username = _getpass.getuser()
            _scan_roots = []
            # Always include current user's home folder on C:\
            _user_home = os.path.join("C:\\", "Users", _username)
            if os.path.exists(_user_home):
                _scan_roots.append(_user_home)
            # Include all non-C drives fully
            for _dl in string.ascii_uppercase:
                if _dl == "C": continue
                _dp = f"{_dl}:\\"
                if os.path.exists(_dp):
                    _scan_roots.append(_dp)
            print(f"[Scan] Roots: {_scan_roots}")

            for d in _scan_roots:
                for root_dir, dirs, files in os.walk(d):
                    def _should_skip(name):
                        nl = name.lower()
                        if nl in SKIP_FOLDERS: return True
                        if nl.startswith("python") and len(nl) <= 12: return True
                        return False
                    dirs[:] = [dd for dd in dirs
                               if not _should_skip(dd)
                               and not _looks_like_model_repo(os.path.join(root_dir, dd))]
                    for item in (dirs + files):
                        try:
                            full_p = os.path.normpath(os.path.join(root_dir, item))
                            if full_p in _seen_paths:
                                continue
                            _seen_paths.add(full_p)
                            is_dir = item in dirs
                            
                            f_size = os.path.getsize(full_p) if not is_dir else 0
                            batch_files.append(("Folder" if is_dir else "File",
                                                 self._sanitize_utf8(item),
                                                 self._sanitize_utf8(full_p), f_size))
                            
                            if not is_dir:
                                f_ext = os.path.splitext(full_p)[1].lower()
                                # Skip sensitive files (IT security / DLP)
                                f_basename = os.path.basename(full_p)
                                if _SENSITIVE_FILE_RE.search(f_basename):
                                    continue
                                is_allowed = False
                                # Priority 1: known Office/PDF — always extract
                                if f_ext in {'.pdf', '.docx', '.doc', '.xlsx', '.xls',
                                             '.csv', '.pptx', '.ppt', '.one', '.msg'}:
                                    is_allowed = True
                                # Priority 2: known text extensions — no need to open file
                                elif f_ext in _TEXT_EXTS:
                                    if f_size < 2 * 1024 * 1024:
                                        is_allowed = True
                                # Priority 3: .env / .env.local / .env.* dotfiles
                                elif f_basename.lower().startswith(_ENV_FILE_PREFIX):
                                    if f_size < 2 * 1024 * 1024:
                                        is_allowed = True
                                # Priority 4 (v9.13): images, only when OCR is
                                # enabled for this run (Update DB "OCR images"
                                # checkbox) — off by default since OCR is much
                                # slower than the other extraction methods.
                                # Capped at 15MB: legitimate screenshots/scans
                                # are almost always well under this; anything
                                # bigger is more likely a huge raw photo/scan
                                # where OCR would be slow for little benefit.
                                elif f_ext in _OCR_IMAGE_EXTS and OCR_ENABLED:
                                    if f_size < 15 * 1024 * 1024:
                                        is_allowed = True
                                # v2.5: removed the old "unknown extension -> open file
                                # and sniff for text" fallback. It was extracting a lot
                                # of installer/SDK metadata (.catnls, .clsid, .iid,
                                # .intinfo, .tmw, ...) that just happens to be plain-text
                                # formatted, plus it opened every single unrecognized
                                # file on the drive (slow). Anything not explicitly
                                # whitelisted above is now skipped outright.
                                if is_allowed:
                                    _scan_total += 1
                                    # v1.5: skip cloud-only (not-yet-downloaded) OneDrive/SharePoint files
                                    # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000 means cloud placeholder
                                    try:
                                        _attrs = os.stat(full_p).st_file_attributes
                                        if _attrs & 0x400000:  # cloud-only placeholder
                                            _scan_skipped += 1
                                            continue
                                    except Exception:
                                        pass
                                    # v1.5: skip re-extracting content if file unchanged since last run
                                    # v-new: ...unless "Force re-index" was checked -- see
                                    # indexing_worker's force_reindex docstring.
                                    try:
                                        _cur_mtime = os.path.getmtime(full_p)
                                    except Exception:
                                        _cur_mtime = None
                                    _prev_mtime = _existing_mtime.get(full_p)
                                    if (not force_reindex and _prev_mtime is not None
                                            and _cur_mtime is not None and abs(_prev_mtime - _cur_mtime) < 1.0):
                                        _scan_skipped += 1
                                        continue  # content unchanged — keep existing content_store/content_index row
                                    # v2.5: DON'T read/extract content here — just queue it
                                    # (tier, path, mtime). Content is extracted AFTER the
                                    # full walk across ALL drives finishes, sorted by tier
                                    # (see _ext_tier) — so tier-1 (office/pdf) content is
                                    # fully indexed across the WHOLE filesystem before
                                    # tier-2 starts, then tier-3, then tier-4. If the run
                                    # gets interrupted (Ctrl+C, power loss, etc.), whatever
                                    # was already committed is always the highest-value
                                    # content first, never a random alphabetical slice.
                                    # A re-run resumes naturally: unchanged files are
                                    # already skipped above by the mtime check, so only
                                    # genuinely new/changed files get queued again.
                                    _tier = 2 if (f_ext == '' and f_basename.lower().startswith(_ENV_FILE_PREFIX)) \
                                            else self._ext_tier(f_ext)
                                    # v9.13: images don't belong to any of the
                                    # 4 Tiers (they'd fall through to an
                                    # "unlisted" tier index that no Tier
                                    # checkbox can ever select), so gate them
                                    # ONLY by the OCR_ENABLED checkbox
                                    # (already checked above via is_allowed),
                                    # not by Tier selection at all.
                                    if f_ext in _OCR_IMAGE_EXTS:
                                        pass
                                    elif selected_tiers is not None and _tier not in selected_tiers:
                                        continue  # tier not requested — skip content extraction
                                    _pending_extract.append((_tier, full_p, _cur_mtime))
                            if len(batch_files) >= 2000:
                                c.executemany("INSERT INTO files_temp VALUES (?,?,?,?)", batch_files); batch_files = []
                        except: continue
            if batch_files: c.executemany("INSERT INTO files_temp VALUES (?,?,?,?)", batch_files)
            conn.commit()

            # ── v2.5: extract content in tier order ──────────────────────────
            # Tier 1: office/pdf | Tier 2: msg/txt/log/one | Tier 3: scripts/
            # config/spck | Tier 4: markup/misc. Stable sort keeps files within
            # the same tier in their original (discovery) order.
            _pending_extract.sort(key=lambda x: x[0])
            _tier_labels = {0: "Tier1 office/pdf", 1: "Tier2 msg/txt/log/one",
                            2: "Tier3 scripts/config", 3: "Tier4 markup/misc"}
            _tier_counts = {}
            for _t, _p, _m in _pending_extract:
                _tier_counts[_t] = _tier_counts.get(_t, 0) + 1
            print(f"[Scan] Tier filter: {'ALL (1-4)' if selected_tiers is None else ','.join(str(t+1) for t in sorted(selected_tiers))}")
            print(f"[Scan] {len(_pending_extract)} files queued for extraction — "
                  + ", ".join(f"{_tier_labels.get(t, f'Tier{t+1}')}: {n}" for t, n in sorted(_tier_counts.items())))

            for _tier, full_p, _cur_mtime in _pending_extract:
                try:
                    content = self.get_file_content(full_p)
                    if content:
                        # v10.19 FIX: sanitize both path and content right
                        # before they're queued for the SQLite INSERT below
                        # -- see _sanitize_utf8 for why (a single file with
                        # a lone-surrogate path, or extracted text that
                        # somehow ends up with one, would otherwise crash
                        # this WHOLE batch's INSERT and abort the run).
                        batch_content.append((self._sanitize_utf8(full_p),
                                               self._sanitize_utf8(content),
                                               _cur_mtime))
                        _scan_new += 1
                except Exception:
                    continue
                if len(batch_content) >= 100:
                    # Remove stale rows before insert (file changed → re-index FTS + store)
                    c.executemany("DELETE FROM content_index WHERE path=?", [(bc[0],) for bc in batch_content])
                    c.executemany("INSERT INTO content_index (path, content) VALUES (?,?)",
                                  [(bc[0], bc[1]) for bc in batch_content])
                    c.executemany("INSERT OR REPLACE INTO content_store VALUES (?,?,?)", batch_content)
                    conn.commit()
                    batch_content = []
            if batch_content:
                c.executemany("DELETE FROM content_index WHERE path=?", [(bc[0],) for bc in batch_content])
                c.executemany("INSERT INTO content_index (path, content) VALUES (?,?)",
                              [(bc[0], bc[1]) for bc in batch_content])
                c.executemany("INSERT OR REPLACE INTO content_store VALUES (?,?,?)", batch_content)
                conn.commit()
            print(f"[Scan] {_scan_total} files checked, {_scan_skipped} unchanged (skipped), {_scan_new} new/changed (re-extracted)")
            c.execute("DELETE FROM files"); c.execute("INSERT INTO files SELECT * FROM files_temp")
            c.execute("DROP TABLE IF EXISTS files_temp")
            c.execute("CREATE INDEX IF NOT EXISTS idx_files_path ON files(path)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_files_name ON files(name)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_files_type ON files(type)")
            conn.commit()

            # ── BM25/content indexing stage complete ─────────────────────────
            # files + content_store + content_index are all rebuilt and usable
            # right now, so Search and Advanced become usable immediately even
            # though the AI embedding pass below can still take a long time.
            # If an AI embedding pass IS about to run, keep the ramp blinking
            # but switch it from Yellow to Green so it's visually obvious
            # which stage is active. If there's no AI pass this run
            # (selected_models == [] , a deliberate "Tier-only" update), stop
            # blinking and go solid Green instead. The ramp only advances to
            # Blue once AI embeddings finish too (see the end of this
            # function / _ai_fully_indexed()).
            _will_run_ai_phase = (selected_models is None) or (len(selected_models) > 0)
            if _will_run_ai_phase:
                self.root.after(0, lambda: (self._ramp_blink_start(color="green"), self._sync_ai_adv_lock()))
            else:
                self.root.after(0, lambda: (self._ramp_blink_stop(final_fg="#4caf50"), self._sync_ai_adv_lock()))

            # v1.5: track paths whose content was just re-extracted (new/changed files).
            # Their OLD embeddings (if any, from previous run) are now stale and must
            # be purged from every semantic table so they get re-embedded below.
            _changed_paths = set()
            c.execute("SELECT path, mtime FROM content_store")
            for _p, _m in c.fetchall():
                _prev = _existing_mtime.get(_p)
                # v-new: with Force re-index, disk mtime for an unchanged
                # file never actually differs from before, so this check
                # alone would miss it -- content_store's TEXT was
                # re-extracted (possibly deeper/different now), but the
                # semantic embeddings built from the OLD text would
                # otherwise never get invalidated/rebuilt. Force always
                # counts every currently-stored file as "changed" here too.
                if force_reindex or _prev is None or (_m is not None and abs(_prev - _m) >= 1.0):
                    _changed_paths.add(_p)
            if _changed_paths:
                for _mk in SEMANTIC_MODELS.keys():
                    _tbl = _semantic_table_for(_mk)
                    try:
                        c.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?", (_tbl,))
                        if c.fetchone():
                            c.executemany(f"DELETE FROM {_tbl} WHERE path=?", [(p,) for p in _changed_paths])
                    except Exception:
                        pass
                conn.commit()
                print(f"[Semantic] Invalidated stale embeddings for {len(_changed_paths)} changed file(s)")

            # ── Build semantic embeddings for ALL models in one pass ────────
            import numpy as np
            c2 = conn.cursor()
            # v1.5b: only pull path + first 300 chars (SUBSTR in SQL) — avoids loading
            # full file content (can be MBs per row) into RAM for hundreds of thousands
            # of rows. This was causing RAM to fill up (86%+) and trigger Windows page
            # file swapping, which is far slower than CPU-bound embedding itself.
            c2.execute("SELECT path, SUBSTR(content, 1, 300) FROM content_store")
            rows = c2.fetchall()
            total_rows = len(rows)
            rows_by_path = {r[0]: r[1] for r in rows}  # dict for O(1) lookup, replaces rows list

            all_model_keys = list(SEMANTIC_MODELS.keys())
            # None => every model (backward-compatible default). An empty
            # list/set is a deliberate "Tier-only, no AI this run" choice —
            # must be checked with `is not None`, not truthiness, or an
            # explicit empty selection would be silently treated the same
            # as None and build every model anyway.
            if selected_models is not None:
                all_model_keys = [k for k in all_model_keys if k in selected_models]
            n_models = len(all_model_keys)
            for model_idx, model_key in enumerate(all_model_keys, 1):
                try:
                    label = SEMANTIC_MODELS[model_key]["label"]
                    print(f"[Semantic] Building index for model {model_idx}/{n_models}: {model_key} ({label})")
                    self.root.after(0, lambda mi=model_idx, mn=n_models, lbl=label:
                        self._set_index_status(f"AI {mi}/{mn}...", "#4caf50"))

                    if not _load_semantic_model(model_key):
                        print(f"[Semantic] Skipping {model_key} — model load failed")
                        continue

                    sem_table = _semantic_table_for(model_key)
                    # v1.5: DO NOT drop table — resume by skipping paths already embedded.
                    # If table doesn't exist yet, create fresh (first run for this model).
                    c2.execute(f"CREATE TABLE IF NOT EXISTS {sem_table} (path TEXT PRIMARY KEY, snippet TEXT, embedding BLOB)")
                    conn.commit()

                    # Load paths already embedded for this model — these are skipped.
                    c2.execute(f"SELECT path FROM {sem_table}")
                    _already_done = set(r[0] for r in c2.fetchall())
                    # v1.5b: iterate keys directly instead of filtering a full copy of `rows`
                    remaining_paths = [p for p in rows_by_path.keys() if p not in _already_done]
                    skipped_count = total_rows - len(remaining_paths)
                    if skipped_count > 0:
                        print(f"[Semantic] {model_key}: resuming — {skipped_count}/{total_rows} already embedded, "
                              f"{len(remaining_paths)} remaining")

                    if not remaining_paths:
                        print(f"[Semantic] {model_key}: nothing to do — already complete ({total_rows} docs)")
                        continue

                    SEM_BATCH = 8 if model_key == "bge_gemma2" else 32
                    sem_buf = []
                    total_done = skipped_count

                    for i in range(0, len(remaining_paths), SEM_BATCH):
                        paths = remaining_paths[i:i+SEM_BATCH]
                        # v9.8 fix: prepend the filename to the text that gets
                        # embedded, not just the extracted file content. Before
                        # this, a file like "Switch-model.zip" — a format we
                        # never extract text from (.zip is in _BINARY_EXTS) —
                        # got embedded from an EMPTY content string, producing
                        # a near-meaningless vector totally disconnected from
                        # its obviously-relevant filename. Even for text-bearing
                        # formats (.spck, etc.) whose content is mostly numeric/
                        # config data rather than natural language, the filename
                        # is often the single strongest topical signal available
                        # and was previously being thrown away entirely.
                        clean_snippets = [
                            (os.path.basename(p) + " " + (rows_by_path[p] or "")).replace("\n", " ")
                            for p in paths
                        ]
                        vecs = _encode_passages(clean_snippets, model_key, batch_size=SEM_BATCH)
                        for path, snippet, vec in zip(paths, clean_snippets, vecs):
                            sem_buf.append((path, snippet, vec.astype("float32").tobytes()))
                        total_done += len(paths)
                        if len(sem_buf) >= 500:
                            c2.executemany(f"INSERT OR REPLACE INTO {sem_table} VALUES (?,?,?)", sem_buf)
                            conn.commit(); sem_buf = []
                            # v1.5: checkpoint — survives interrupt; next run resumes from here
                            _save_checkpoint('embed', model_key, paths[-1], total_done, total_rows)
                        if total_done % (SEM_BATCH * 10) == 0 or total_done == total_rows:
                            _pct = int(total_done * 100 / max(total_rows, 1))
                            self.root.after(0, lambda p=_pct, mi=model_idx, mn=n_models:
                                self._set_index_status(f"AI {mi}/{mn}: {p}%", "#4caf50"))
                    if sem_buf:
                        c2.executemany(f"INSERT OR REPLACE INTO {sem_table} VALUES (?,?,?)", sem_buf)
                        conn.commit()
                        _save_checkpoint('embed', model_key, paths[-1] if paths else '', total_done, total_rows)
                    print(f"[Semantic] {model_key}: indexed {total_done}/{total_rows} docs into {sem_table}")
                    _clear_checkpoint('embed', model_key)  # model fully done — checkpoint no longer needed

                except Exception as _se:
                    print(f"[Semantic] Embedding build failed for {model_key}: {_se}")
                    print(f"[Semantic] {model_key}: progress saved — next --update data run will resume from here")
            # ────────────────────────────────────────────────────────────────

            conn.commit()
            conn.close()
            import time; time.sleep(0.5)  # let WAL checkpoint flush before reopening

            # ── Outlook mail indexing (v-outlook) ───────────────────────────
            # Runs as part of the same "Update DB" click so there's only one
            # button for the user to remember — own db file, own try/except,
            # so if Outlook isn't installed/running this just prints a note
            # and the rest of Update DB continues unaffected.
            if OUTLOOK_SEARCH_AVAILABLE:
                try:
                    try:
                        self.root.after(0, lambda: self.placeholder.config(text="Indexing Outlook mail..."))
                    except Exception:
                        pass
                    def _outlook_progress(done, total, subj):
                        print(f"[Outlook] {done}/{total}: {subj}")
                    o_indexed, o_skipped, o_err = outlook_search.index_outlook_mail(progress_cb=_outlook_progress)
                    if o_err:
                        print(f"[Outlook] Indexing skipped: {o_err}")
                    else:
                        print(f"[Outlook] Indexed {o_indexed} mail, skipped {o_skipped} unchanged.")
                except Exception as _oe:
                    print(f"[Outlook] Indexing failed: {_oe}")
            # ─────────────────────────────────────────────────────────────

            # ── OneNote indexing (v-onenote) ────────────────────────────────
            # Same pattern as Outlook above — own db file, own try/except,
            # runs as part of the same "Update DB" click.
            if ONENOTE_SEARCH_AVAILABLE:
                try:
                    try:
                        self.root.after(0, lambda: self.placeholder.config(text="Indexing OneNote..."))
                    except Exception:
                        pass
                    def _onenote_progress(done, total, subj):
                        print(f"[OneNote] {done}/{total}: {subj}")
                    n_indexed, n_skipped, n_locked, n_err = onenote_search.index_onenote(progress_cb=_onenote_progress)
                    if n_err:
                        print(f"[OneNote] Indexing skipped: {n_err}")
                    else:
                        print(f"[OneNote] Indexed {n_indexed} page, skipped {n_skipped} unchanged, {n_locked} locked/unreadable.")
                except Exception as _ne:
                    print(f"[OneNote] Indexing failed: {_ne}")
            # ─────────────────────────────────────────────────────────────

            # ⚠️ VACUUM removed: 44GB DB needs ~88GB free disk + 30min → causes RED LIGHT
            # v3.5: Blue only once the currently selected AI model's embeddings
            # actually cover every content_store row (they might not, e.g. if
            # _load_semantic_model() failed above) — otherwise stay Green.
            # v7.10 FIX: previously the ramp only turned Blue (and AI Search
            # unlocked) if the model happening to be selected in the dropdown
            # was fully indexed -- if that model's build failed (e.g. Jina-v3
            # here: "'XLMRobertaLoRA' object has no attribute
            # 'all_tied_weights_keys'") but a DIFFERENT model in the same run
            # succeeded (e.g. BGE-Gemma2), the ramp stayed stuck on Green and
            # AI Search stayed greyed out even though a usable AI index
            # existed. Now: if the currently selected model isn't fully
            # indexed, check the other models that were part of this run and
            # auto-switch to the first one that IS fully indexed.
            global _sem_model_key
            _cur_key = self.ai_model_var.get() if self.ai_model_var.get() in SEMANTIC_MODELS else DEFAULT_SEMANTIC_MODEL
            if not self._ai_fully_indexed(_cur_key):
                for _mk in all_model_keys:
                    if _mk != _cur_key and self._ai_fully_indexed(_mk):
                        print(f"[Semantic] '{_cur_key}' isn't fully indexed but '{_mk}' is -- "
                              f"switching the AI model selection to '{_mk}'")
                        _cur_key = _mk
                        _sem_model_key = _mk
                        self.ai_model_var.set(_mk)

                        def _update_combo_display(k=_mk):
                            try:
                                _combo = getattr(self, "_ai_model_combo", None)
                                if _combo and _combo.winfo_exists():
                                    _combo.set(SEMANTIC_MODELS[k]["label"])
                            except Exception:
                                pass
                        self.root.after(0, _update_combo_display)
                        break
            _final_fg = "#2196f3" if self._ai_fully_indexed(_cur_key) else "#4caf50"
            print(f"[Semantic] Indexing finished — final ramp color: "
                  f"{'BLUE (AI fully indexed)' if _final_fg == '#2196f3' else 'GREEN (AI not fully indexed)'}")

            def _finish(fg=_final_fg):
                # v5.3 fix: each step is now guarded independently. Before, this
                # was a single list-expression -- if any ONE call raised, the
                # remaining calls (including _set_index_status(done=True), the
                # one that actually flips the ramp to Blue) silently never ran,
                # leaving the button stuck on "AI x/x..." / Green until restart.
                try: self._open_db_conn()   # reopen persistent search connection
                except Exception as _e1: print(f"[Semantic] _open_db_conn failed: {_e1}")
                try: self._ramp_blink_stop()
                except Exception as _e2: print(f"[Semantic] _ramp_blink_stop failed: {_e2}")
                try: self._set_index_status("●", fg, done=True)
                except Exception as _e3: print(f"[Semantic] _set_index_status failed: {_e3}")
                try: self.placeholder.config(text=READY_PH)
                except Exception as _e4: print(f"[Semantic] placeholder update failed: {_e4}")
                try: self._sync_ai_adv_lock()  # belt-and-suspenders re-sync
                except Exception as _e5: print(f"[Semantic] _sync_ai_adv_lock failed: {_e5}")
                # v-new: "Update DB completed" summary popup -- what ran,
                # when it started/ended, how long it took. Shown last, after
                # everything above has already unlocked the UI, so the popup
                # being modal doesn't block the ramp/status from updating.
                try:
                    self._show_update_db_summary(
                        _run_started_at, selected_tiers, selected_models, ocr_enabled, force_reindex)
                except Exception as _e6:
                    print(f"[Semantic] _show_update_db_summary failed: {_e6}")

            self.root.after(0, _finish)
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            print(f"Indexing Error:\n{err}")
            try:
                log_path = os.path.join(os.path.dirname(os.path.abspath(DB_FILE)), "search_error.log")
                with open(log_path, "a", encoding="utf-8", errors="replace") as lf:
                    from datetime import datetime as _dt
                    lf.write(f"\n[{_dt.now()}] INDEXING ERROR:\n{err}\n")
            except: pass
            self.root.after(0, lambda: (self._ramp_blink_stop(), self._set_index_status("●", "#ff0000", error=True)))

    def on_key_release(self, event):
        if event.keysym in ["Up", "Down", "Return", "Escape", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"]: return
        # v-fix (AI Chat phân tích quá sớm): mark "user is actively typing"
        # right now -- _auto_ai_chat_after_search checks this before firing
        # so it waits for a real pause instead of reacting to the very
        # first settled word of a multi-word JP query.
        self._last_keystroke_ts = time.time()
        q = self.entry_var.get().strip()

        if len(q) < 1:
            # v2.7: user request — clearing the box to retype shouldn't shrink the
            # results window back down. Just cancel any pending debounced search and
            # leave the last results on screen; only the ✕ close button (or a special
            # command below) should shrink the window back to the small search box.
            if self.search_timer:
                self.root.after_cancel(self.search_timer)
                self.search_timer = None
            # v9.14: box is empty — clear the dedupe tracker so that if the
            # user retypes the exact same text they had before, it isn't
            # mistaken for a no-op repeat of the old (already-cleared) query.
            self._last_processed_query = None
            return

        _CMD_LIST = ["--update data", "--update outlook", "--update onenote", "--exit", "--quit"]
        ql = q.lower()

        # Cancel any pending "close-on-command" check from a previous keystroke --
        # if we're here, the user just typed/changed something, so an earlier
        # scheduled check is stale and must not fire on top of the new text.
        if getattr(self, '_cmd_close_timer', None):
            try: self.root.after_cancel(self._cmd_close_timer)
            except Exception: pass
            self._cmd_close_timer = None

        # v3.1 FIX: the box starting with "-" (e.g. "-", "--", "--h", "--u"...)
        # no longer closes the results window / clears the box on the spot --
        # that made "--" itself vanish and forced retyping "--help" in the
        # shrunken box. Instead treat it as "possibly typing a command": don't
        # search, and only AFTER the user pauses (400ms with no further
        # keystroke) do we check whether the finished text is a real command
        # and close the results window then. If the user keeps typing, this
        # check keeps getting cancelled/rescheduled above, so it never fires
        # mid-typing.
        if ql.startswith("-"):
            if self.search_timer:
                self.root.after_cancel(self.search_timer)
                self.search_timer = None
            self._cmd_close_timer = self.root.after(
                400, lambda snapshot=q: self._maybe_close_for_cmd(snapshot))
            return

        # v9.14 fix: KeyRelease also fires for shortcuts that don't change the
        # text at all — Ctrl+A (select all), Ctrl+C (copy), cursor movement
        # with Home/End/Left/Right, etc. Their keysym is the letter/arrow key
        # itself (e.g. "a", "c", "Left"), not "Control_L", so the ignore-list
        # above never caught them, and they silently re-ran the full search
        # for UNCHANGED text. When AI Search was active, that redundant
        # re-search overwrote the AI-merged results already on screen with
        # plain BM25/MFT results (the results appeared to "vanish" even
        # though the AI Search button still showed as active), forcing the
        # user to click AI Search again to restore them. Skip entirely when
        # the text hasn't actually changed since the last run.
        if q == getattr(self, "_last_processed_query", None):
            return
        self._last_processed_query = q

        self.current_search_id += 1
        search_version = self.current_search_id
        if self.search_timer: self.root.after_cancel(self.search_timer)
        self.root.update_idletasks()
        box_x, box_y, box_h = self.root.winfo_x(), self.root.winfo_y(), self._search_bar_box_h()

        # v5.8: save to Search History automatically once the user pauses on
        # this text for 10s without pressing Enter (see _maybe_save_hist_idle).
        # v-fix: cancel+reschedule happen together, right here, only on a
        # genuine text change (we've already passed the "q unchanged" early
        # return above by this point). Previously the cancel ran at the very
        # top of on_key_release unconditionally, so a no-op KeyRelease after
        # the user finished typing (arrow keys, Home/End, Ctrl+A, clicking
        # back into the box, etc.) killed the pending countdown without ever
        # restarting it — the idle-save then silently never fired again for
        # that search, for ANY query.
        if getattr(self, '_hist_idle_timer', None):
            try: self.root.after_cancel(self._hist_idle_timer)
            except Exception: pass
        self._hist_idle_timer = self.root.after(
            10000, lambda snapshot=q: self._maybe_save_hist_idle(snapshot))

        # v2.3: Always launch MFT scan (realtime, no DB needed) for File/Folder Name tabs
        threading.Thread(
            target=self._mft_scan_search,
            args=(q, search_version, box_x, box_y, box_h), daemon=True
        ).start()

        # v2.3: Also launch BM25 (File Content tab) if DB is available — blend results
        if self.db_conn is not None:
            self.search_timer = self.root.after(350, lambda: threading.Thread(
                target=self._smart_search_realtime,
                args=(q, search_version, box_x, box_y, box_h, False), daemon=True
            ).start())

    def _search_bar_box_h(self):
        """v-fix: height of just the persistent top search-bar row
        (self.bg_f, fixed height=35 -- see __init__), NOT the whole window.
        self.root IS the results window once results are showing (see
        show_results: "results now render inside THIS SAME window
        (self.root)") -- so self.root.winfo_height() at that point returns
        the height of the entire (tall, ~930px) results window instead of
        just the search bar. bg_f's height never changes (idle vs results
        mode), so using it here always gives the right offset regardless
        of whether results are currently showing."""
        try:
            h = self.bg_f.winfo_height()
            return h if h > 1 else 35
        except Exception:
            return 35

    def _rerun_current_search(self):
        """v7.10: re-run the search currently in the box with the SAME text
        -- used when a toggle that changes matching behavior (e.g. "Whole
        word") flips, so results refresh immediately without the user having
        to retype/re-trigger the query themselves."""
        q = self.entry_var.get().strip()
        if not q or q.startswith("-"):
            return
        self.current_search_id += 1
        search_version = self.current_search_id
        if self.search_timer:
            try: self.root.after_cancel(self.search_timer)
            except Exception: pass
        self.root.update_idletasks()
        box_x, box_y, box_h = self.root.winfo_x(), self.root.winfo_y(), self._search_bar_box_h()
        threading.Thread(
            target=self._mft_scan_search,
            args=(q, search_version, box_x, box_y, box_h), daemon=True
        ).start()
        if self.db_conn is not None:
            threading.Thread(
                target=self._smart_search_realtime,
                args=(q, search_version, box_x, box_y, box_h, False), daemon=True
            ).start()

    def _maybe_close_for_cmd(self, snapshot):
        """Fired ~400ms after the user stops typing a string starting with '-'.
        Only closes the results window if the box still holds exactly the same
        text (i.e. the user has genuinely paused / finished) AND that text is
        one of the recognized commands. Actually running the command (help
        popup, --update data, exit, etc.) still happens on Enter via
        handle_action() as before -- this only tidies up the results window
        so it isn't sitting open behind a command about to run."""
        self._cmd_close_timer = None
        if self.entry_var.get().strip() != snapshot:
            return  # user kept typing / edited since — stale check, ignore
        ql = snapshot.lower()
        _CMD_LIST = ["--update data", "--update outlook", "--update onenote", "--exit", "--quit"]
        if ql in _CMD_LIST or ql.startswith("--update data"):
            if self.active_result_win and tk.Toplevel.winfo_exists(self.active_result_win):
                self._close_results_window()
                self.entry_var.set(snapshot)  # restore the command text the user typed
                self.entry.icursor("end")

    # ── v2.3: MFT / os.walk realtime scan ────────────────────────────────────
    def _mft_scan_search(self, q, sid, box_x, box_y, box_h):
        """Realtime filename scan — no DB needed. Streams results into File Name
        and Folder Name tabs as they are found. If DB is also available, BM25
        runs in parallel and populates File Content tab independently."""
        try:
            # Open result window if needed — guarded by a lock so a burst of
            # KeyRelease events (e.g. Japanese/CJK IME composition) can't race
            # each other into creating two separate result windows.
            with self._win_create_lock:
                win_missing = not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win)
                already_opening = self._opening_result_win
                if win_missing and not already_opening:
                    self._opening_result_win = True

            if win_missing and not already_opening:
                try:
                    self.root.after(0, lambda: self.show_results(
                        [], [], q, False, box_x, box_y, box_h, sid))
                    time.sleep(0.15)
                finally:
                    self._opening_result_win = False
            elif win_missing and already_opening:
                # Another thread is already opening the window — wait for it
                # to appear instead of opening a second one.
                for _ in range(30):
                    if self.active_result_win and tk.Toplevel.winfo_exists(self.active_result_win):
                        break
                    time.sleep(0.02)
            else:
                # v2.3 fix: window already open from a previous query — MFT streams
                # results into it but never touched the title, so it stayed stale.
                def _set_title():
                    try:
                        if sid != self.current_search_id: return
                        if self.active_result_win and tk.Toplevel.winfo_exists(self.active_result_win):
                            self.active_result_win.title(f"Results: {q}")
                    except Exception: pass
                self.root.after(0, _set_title)

            # Clear File Name + Folder Name trees (not Content — BM25 handles that)
            def _clear():
                try:
                    if hasattr(self, 'tree_f')   and self.tree_f.winfo_exists():
                        self.tree_f.delete(*self.tree_f.get_children())
                    if hasattr(self, 'tree_fol') and self.tree_fol.winfo_exists():
                        self.tree_fol.delete(*self.tree_fol.get_children())
                except Exception: pass
            self.root.after(0, _clear)
            time.sleep(0.05)

            import getpass
            _user_home = os.path.join("C:\\", "Users", getpass.getuser())
            scan_roots = [_user_home] if os.path.exists(_user_home) else []
            for letter in string.ascii_uppercase:
                if letter == "C": continue
                dp = f"{letter}:\\"
                if os.path.exists(dp): scan_roots.append(dp)

            # AND match: each whitespace-separated keyword must appear
            # SOMEWHERE in the name (not required to be contiguous) — this is
            # what the app's own help text promises ("space = AND search").
            # We also track an OR match (any keyword, but not full AND) in
            # parallel. AND results always render first; OR-only results are
            # appended BELOW them at the end of the scan (row[5] = 0 for AND,
            # 1 for OR — see _sort_priority) — applied independently for File
            # Name and Folder Name, so one tab having enough AND matches
            # doesn't suppress the other tab's OR fallback.
            keywords = [k for k in _split_query_tokens(q.lower()) if k]
            if not keywords:
                return
            # v5.8: drop filler words ("to", "relevant", ...) before AND/OR
            # matching -- see STOPWORDS comment above for why this matters.
            keywords = _strip_stopwords(keywords)
            multi_kw = len(keywords) > 1
            ww = self.whole_word_var.get()  # v7.10: "Whole word" toggle
            file_idx    = [0]
            folder_idx  = [0]
            or_file_idx   = [0]
            or_folder_idx = [0]
            MAX = 500
            # Hard time budget: for a genuinely rare query (or on huge drives)
            # neither MAX may ever be reached, so without a cap we'd walk the
            # entire filesystem before showing anything. Bound the worst case
            # latency instead of scanning forever.
            #
            # v2.5 fix: drives are scanned sequentially (C-home, then D, E,
            # F...). A single slow/large drive (esp. a network share) could
            # eat the ENTIRE global budget, so later drives (e.g. F:) never
            # got scanned at all — even though they exist and the BM25/DB
            # index (built offline, no live scanning needed) still has them.
            # Give each root a fair per-root time slice too, so a slow drive
            # gets cut off and moves on instead of starving the rest.
            # v7.10: bumped from 3.0s -- a query whose matches sit deep inside
            # a large drive (e.g. "adas" living several folders down on a big
            # D:/E: drive full of engineering data) could get cut off by
            # os.walk's per-root time slice before ever reaching those paths,
            # while a query whose matches happen to sit shallower (or on a
            # smaller/faster drive) found plenty within the same budget. This
            # isn't about the keyword itself, just where on disk it happens to
            # live vs. how much of the tree got walked before the clock ran
            # out. 10s is still a soft cap (not exhaustive on huge drives) --
            # for guaranteed complete results regardless of location/depth,
            # run --update data once so search uses the full DB index instead
            # of this live best-effort scan.
            TIME_BUDGET_SEC = 12.0
            PER_ROOT_BUDGET_SEC = max(2.0, TIME_BUDGET_SEC / max(1, len(scan_roots)))
            t_start = time.time()
            self._mft_file_res   = []
            self._mft_folder_res = []
            or_file_res   = []   # OR-only matches, appended below AND at the end
            or_folder_res = []
            timed_out = False

            for root_dir in scan_roots:
                if sid != self.current_search_id: return
                if timed_out: break
                t_root_start = time.time()
                for root, dirs, files in os.walk(root_dir):
                    if sid != self.current_search_id: return
                    if time.time() - t_start > TIME_BUDGET_SEC:
                        timed_out = True
                        break
                    if time.time() - t_root_start > PER_ROOT_BUDGET_SEC:
                        break   # move on to the next drive, don't starve it
                    dirs[:] = [d for d in dirs
                                if d.lower() not in SKIP_FOLDERS
                                and not (d.lower().startswith("python") and len(d) <= 12)]

                    # ── Folders ───────────────────────────────────────────
                    for d in dirs:
                        dl = d.lower()
                        # v-fix (Vấn đề: "鉄道不整" 0 kết quả ở Folder Name
                        # dù "鉄道"/"不整" riêng lẻ có): dùng
                        # _kw_matches_with_glued_fallback thay vì
                        # _kw_matches thẳng -- xem docstring hàm đó.
                        is_and = all(_kw_matches_with_glued_fallback(kw, dl, ww) for kw in keywords)
                        if is_and and folder_idx[0] < MAX:
                            full_p = os.path.normpath(os.path.join(root, d))
                            mt = get_live_mtime(full_p)
                            folder_idx[0] += 1
                            row = ("Folder", d, full_p, 0, mt, 0)
                            self._mft_folder_res.append(row)
                            self._schedule_mft_render("fol", sid)
                        elif multi_kw and or_folder_idx[0] < MAX:
                            # v5.8: track how many keywords matched (not just
                            # any/none) so results matching MORE keywords can
                            # be ranked above single-keyword-only matches
                            # instead of sitting at the same priority.
                            m_count = sum(1 for kw in keywords if _kw_matches_with_glued_fallback(kw, dl, ww))
                            if m_count > 0:
                                full_p = os.path.normpath(os.path.join(root, d))
                                mt = get_live_mtime(full_p)
                                or_folder_idx[0] += 1
                                or_folder_res.append(("Folder", d, full_p, 0, mt, 1, m_count))

                    # ── Files ─────────────────────────────────────────────
                    for f in files:
                        fl = f.lower()
                        is_and = all(_kw_matches_with_glued_fallback(kw, fl, ww) for kw in keywords)
                        if is_and and file_idx[0] < MAX:
                            full_p = os.path.normpath(os.path.join(root, f))
                            try: sz = os.path.getsize(full_p)
                            except: sz = 0
                            mt = get_live_mtime(full_p)
                            file_idx[0] += 1
                            row = ("File", f, full_p, sz, mt, 0)
                            self._mft_file_res.append(row)
                            self._schedule_mft_render("f", sid)
                        elif multi_kw and or_file_idx[0] < MAX:
                            m_count = sum(1 for kw in keywords if _kw_matches_with_glued_fallback(kw, fl, ww))
                            if m_count > 0:
                                full_p = os.path.normpath(os.path.join(root, f))
                                try: sz = os.path.getsize(full_p)
                                except: sz = 0
                                mt = get_live_mtime(full_p)
                                or_file_idx[0] += 1
                                or_file_res.append(("File", f, full_p, sz, mt, 1, m_count))

                    # Stop early once we've got plenty of AND results AND
                    # plenty of OR fallback data too (no point walking the
                    # whole disk just to keep collecting more of either pool).
                    #
                    # v7.4 FIX: the old condition put "not multi_kw" at the
                    # top level of the OR, e.g. (file_idx>=MAX or not multi_kw
                    # or or_file_idx>=MAX). For a single-keyword query
                    # (multi_kw=False), "not multi_kw" is always True, which
                    # made the WHOLE clause True regardless of file_idx --
                    # the scan broke out after the very FIRST os.walk
                    # directory, before ever descending into subfolders.
                    # That's why a single generic keyword (e.g. "adas") could
                    # return 0 results while a 2-keyword query (multi_kw=True,
                    # where this bug doesn't trigger) correctly walked the
                    # whole tree and found plenty. The AND-count requirement
                    # must always hold; only the OR-count requirement should
                    # be skipped when there's just one keyword.
                    if (file_idx[0] >= MAX and (or_file_idx[0] >= MAX or not multi_kw)) and \
                       (folder_idx[0] >= MAX and (or_folder_idx[0] >= MAX or not multi_kw)):
                        break

            # ── Always blend: AND results on top, OR-only results appended
            # below — independently for File Name and Folder Name tabs.
            # (Previously this only kicked in when AND was completely empty,
            # so "simpack realtime" with exactly 1 AND hit never got the OR
            # results appended below it.)
            # v5.8: OR-only rows are sorted by match_count (desc) before
            # being appended, so e.g. a file matching 2 of 3 keywords shows
            # above one matching only 1 of 3 -- then the temporary 7th
            # "match_count" field is stripped back off (rows are 6-tuples
            # everywhere else, incl. _sort_priority/_fill_files/_fill_folders).
            if multi_kw:
                or_file_res.sort(key=lambda r: r[6], reverse=True)
                or_folder_res.sort(key=lambda r: r[6], reverse=True)
                self._mft_file_res   = self._mft_file_res   + [r[:6] for r in or_file_res]
                self._mft_folder_res = self._mft_folder_res + [r[:6] for r in or_folder_res]

            # Final render once the scan finishes, so the very last batch of
            # rows (which might not have hit the coalescing window) is shown.
            self._schedule_mft_render("f", sid, force=True)
            self._schedule_mft_render("fol", sid, force=True)

        except Exception as e:
            print(f"[MFT] Error: {e}", flush=True)

    # ── v2.4: coalesced MFT render — fixes 4 related bugs at once ───────────
    # 1) stale "#" numbering from old-query callbacks still landing after a
    #    new search started (no sid guard before)
    # 2) MFT rows ignored priority sort (office/pdf/msg first, log/code last)
    # 3) horizontal scrollbar / column width reset on every single insert
    # 4) active ext/size/name filter being bypassed by raw streaming inserts
    # Instead of inserting each found row immediately, we append to the
    # result cache and schedule one debounced full re-render (sorted +
    # filtered) a little later. Multiple finds within the debounce window
    # collapse into a single re-render.
    def _schedule_mft_render(self, which, sid, force=False):
        pending_attr = "_mft_render_pending_f" if which == "f" else "_mft_render_pending_fol"
        if getattr(self, pending_attr) and not force:
            return
        setattr(self, pending_attr, True)
        delay = 0 if force else 120

        def _do():
            setattr(self, pending_attr, False)
            if sid != self.current_search_id:
                return
            if which == "f":
                self._render_mft_file_tree()
            else:
                self._render_mft_folder_tree()
        self.root.after(delay, _do)

    def _merge_mft_with_db(self, mft_rows, is_folder):
        """v7.7 FIX: merge the live MFT-scan rows with the last DB-backed
        result snapshot (_db_rendered_file_res), if that snapshot is for the
        SAME search id, instead of letting whichever thread finishes last
        (DB search vs. live disk scan) blindly overwrite the other's results.

        The DB-backed search scans the full indexed corpus and is treated as
        the authoritative/more complete set; the live scan only walks the
        user's home folder + non-C: drives (see _mft_scan_search), so it can
        legitimately have FEWER matches than the DB even though it finishes
        later. Rather than let that narrower set stomp the fuller one, we
        union both by path (DB rows first, then any MFT-only paths the DB
        doesn't know about -- e.g. very recent files not yet indexed)."""
        if self._db_rendered_sid != self.current_search_id:
            # DB hasn't painted anything for this query (yet, or at all) --
            # nothing to merge with, just show what the live scan found.
            return list(mft_rows)
        db_rows = [r for r in self._db_rendered_file_res
                   if (str(r[0]).lower() == 'folder') == is_folder]
        seen_paths = set()
        merged = []
        for r in db_rows:
            p = r[2]
            if p in seen_paths:
                continue
            seen_paths.add(p)
            # Pad DB rows (4-tuple: type,name,path,size) to the 6-tuple shape
            # MFT rows use (type,name,path,size,mtime,and_or_flag) so the
            # renderers below can treat both uniformly.
            merged.append(r if len(r) >= 6 else (r[0], r[1], r[2], r[3], None, 0))
        for r in mft_rows:
            p = r[2]
            if p in seen_paths:
                continue
            seen_paths.add(p)
            merged.append(r)
        return merged

    def _render_mft_file_tree(self):
        """Full (but cheap) re-render of tree_f from self._mft_file_res,
        merged with any DB-backed results already shown for this same query:
        priority-sorted, current filter applied, scroll/column width kept."""
        try:
            if not hasattr(self, 'tree_f') or not self.tree_f.winfo_exists():
                return
        except Exception:
            return
        tree = self.tree_f
        merged_rows = self._merge_mft_with_db(self._mft_file_res or [], is_folder=False)
        rows = self._sort_priority(merged_rows, 1)
        rows = self._filter_file_rows(rows)

        x0 = self._save_xview(tree)
        widths = self._save_col_widths(tree)
        for item in tree.get_children():
            tree.delete(item)
        for rn, item in enumerate(rows, start=1):
            try:
                path = item[2]
                pf = self._pseudo_path_row_fields(path)
                if pf:
                    name_txt, size_str, mtime_str, ftype_str, loc_str, img = pf
                    tree.insert("", "end", text="", **({"image": img} if img else {}), values=(
                        name_txt, size_str, mtime_str, ftype_str, loc_str, path))
                    continue
                img = get_tree_icon_image(path, is_folder=False)
                name_txt = item[1] if img else get_file_icon(path) + item[1]
                mt = item[4] if len(item) > 4 and item[4] is not None else get_live_mtime(path)
                tree.insert("", "end", text="", **({"image": img} if img else {}), values=(
                    name_txt, format_size(item[3]), mt,
                    get_file_type(path), os.path.dirname(path), path))
            except Exception:
                pass
        self._restore_col_widths(tree, widths)
        self._restore_xview(tree, x0)
        # v9.16 fix: the pane header ("📁 Default (N)") was only ever set once,
        # at the moment the AI/Advanced split was opened, from the DB-only
        # BM25 count -- it never accounted for rows the live MFT scan streams
        # in afterward (merged above into `rows`), so it could show e.g.
        # "Default (0)" while the pane visibly had dozens of real matches.
        # Refresh it here with the actual final row count every time this
        # tree is (re)rendered.
        self._update_pane_label(tree, len(rows), len(rows))

    def _render_mft_folder_tree(self):
        """Full (but cheap) re-render of tree_fol from self._mft_folder_res,
        merged with any DB-backed results already shown for this same query:
        current filter applied, scroll/column width kept."""
        try:
            if not hasattr(self, 'tree_fol') or not self.tree_fol.winfo_exists():
                return
        except Exception:
            return
        tree = self.tree_fol
        merged_rows = self._merge_mft_with_db(self._mft_folder_res or [], is_folder=True)
        rows = self._sort_priority(merged_rows, 1)
        name_needle = _norm_txt(self.name_filter_var.get().strip().lower())
        if name_needle:
            rows = [r for r in rows if name_needle in _norm_txt(os.path.basename(r[2]).lower())]

        x0 = self._save_xview(tree)
        widths = self._save_col_widths(tree)
        for item in tree.get_children():
            tree.delete(item)
        for rn, item in enumerate(rows, start=1):
            try:
                path = item[2]
                img = get_tree_icon_image(path, is_folder=True)
                name_txt = item[1] if img else "📁 " + item[1]
                mt = item[4] if len(item) > 4 and item[4] is not None else get_live_mtime(path)
                tree.insert("", "end", text="", **({"image": img} if img else {}), values=(name_txt, mt, path, path))
            except Exception:
                pass
        self._restore_col_widths(tree, widths)
        self._restore_xview(tree, x0)
        # v9.16 fix: see matching note in _render_mft_file_tree above.
        self._update_pane_label(tree, len(rows), len(rows))

    def _filter_file_rows(self, rows):
        """Apply the File Name tab's active size/ext/name filters to a row list.
        Shared by the MFT live renderer and the manual filter re-apply so both
        paths always agree — this is what fixes the 'MFT stream ignores the
        active filter' inconsistency."""
        op, size_bytes = self._get_size_filter_bytes()
        ext_filter  = self._get_ext_filter()
        name_needle = _norm_txt(self.name_filter_var.get().strip().lower())
        if op is None and ext_filter is None and not name_needle:
            return rows
        cmp_fn = {'>':  lambda a, b: a >  b, '>=': lambda a, b: a >= b,
                  '<':  lambda a, b: a <  b, '<=': lambda a, b: a <= b,
                  '=':  lambda a, b: a == b}.get(op) if op else None
        result = []
        for item in rows:
            if cmp_fn is not None:
                if not cmp_fn(get_live_size(item[2], item[3]), size_bytes):
                    continue
            if ext_filter is not None:
                _, ext = os.path.splitext(item[2])
                if ext.lower() not in ext_filter:
                    continue
            if name_needle:
                if name_needle not in _norm_txt(os.path.basename(item[2]).lower()):
                    continue
            result.append(item)
        return result

    # ── shared helpers: preserve scroll position + column widths across a
    #    full clear+rebuild of a Treeview (used by MFT render and filters) ──
    def _save_xview(self, tree):
        try:
            return tree.xview()[0]
        except Exception:
            return 0.0

    def _restore_xview(self, tree, x0):
        if x0 and x0 > 0.0:
            try:
                tree.xview_moveto(x0)
            except Exception:
                pass

    def _save_yview(self, tree):
        """v-new: vertical scroll position — the filter re-apply (_apply_
        size_filter_to_tree) always clears+rebuilds the tree, which used to
        silently jump the view back to row 0 every time (e.g. just clicking
        the Name filter's '✕' with nothing even typed), which looks exactly
        like the sort order changed even when it didn't."""
        try:
            return tree.yview()[0]
        except Exception:
            return 0.0

    def _restore_yview(self, tree, y0):
        if y0 and y0 > 0.0:
            try:
                tree.yview_moveto(y0)
            except Exception:
                pass

    def _save_col_widths(self, tree):
        widths = {}
        try:
            for col in tree["columns"]:
                widths[col] = tree.column(col, "width")
        except Exception:
            pass
        return widths

    def _restore_col_widths(self, tree, widths):
        for col, w in widths.items():
            try:
                tree.column(col, width=w)
            except Exception:
                pass
    # ─────────────────────────────────────────────────────────────────────────

    def _semantic_search(self, query, top_k=SEMANTIC_TOP_K, threshold=None, model_key=None, model_obj=None):
        """Return list of (path, score_pct) using cosine similarity against the embedding
        table of the given model (or the globally-selected one if model_key is None).
        The 2 models (jina_v3/bge_gemma2) have different embedding dimensions/vector
        spaces so each uses a separate table, and (v9.3) each can have its own
        similarity threshold -- see SEMANTIC_THRESHOLD_BY_MODEL near the top of the file.

        model_obj: v-new — pass an already-loaded model object (e.g. the
        secondary blend slot, _sem_model2) to encode with it directly,
        instead of going through the primary-slot load/swap machinery.
        Leave None for the normal single-model behavior (unchanged)."""
        try:
            model_key = model_key or _sem_model_key
            if model_obj is None:
                if not _load_semantic_model(model_key):
                    return []
                model_obj = _sem_model
            if threshold is None:
                threshold = SEMANTIC_THRESHOLD_BY_MODEL.get(model_key, SEMANTIC_THRESHOLD_DEFAULT)
            query = _maybe_restore_diacritics(query)
            import numpy as np
            q_vec = _encode_query_with(model_obj, query, model_key).astype("float32")
            conn = self.db_conn
            if conn is None:
                return []
            c = conn.cursor()
            table = _semantic_table_for(model_key)
            try:
                c.execute(f"SELECT path, embedding FROM {table}")
                rows = c.fetchall()
            except sqlite3.OperationalError:
                # This model's table has not been built yet (--update data not run for this model)
                rows = []
            if not rows:
                return []
            paths = [r[0] for r in rows]
            mat = np.frombuffer(b"".join(r[1] for r in rows), dtype="float32").reshape(len(rows), -1)
            scores = mat @ q_vec   # cosine similarity (vectors normalized)
            top_idx = np.argsort(scores)[::-1][:top_k]
            results = [(paths[i], float(scores[i])) for i in top_idx if scores[i] >= threshold]
            if not results and len(top_idx) > 0:
                # v9.3: everything got filtered out by `threshold`. This is
                # worth logging rather than silently returning [] -- the
                # single SEMANTIC_THRESHOLD (0.25) is shared by both models,
                # but jina_v3 and bge_gemma2 are different architectures
                # with different cosine-similarity score distributions, so
                # the same cutoff can behave very differently between them
                # (e.g. one model still clears 0.25 for a bad/ambiguous
                # query, like non-diacritic Vietnamese, while the other's
                # best score sits just under it and gets filtered to
                # nothing). Print the actual best score reached so this is
                # easy to diagnose instead of just looking "broken".
                _best = float(scores[top_idx[0]])
                print(f"[Semantic] {model_key}: 0 results after threshold filter "
                      f"(best raw score was {_best:.3f}, threshold is {threshold}). "
                      f"If this best score is consistently just under the threshold for "
                      f"this model, consider lowering SEMANTIC_THRESHOLD for it specifically.")
            return results
        except Exception as _e:
            print(f"[Semantic] Search error: {_e}")
            return []

    def _semantic_search_blend(self, query, top_k=SEMANTIC_TOP_K):
        """v-new: replaces the model dropdown with automatic behavior so the
        user never has to know what 'Jina-v3' or 'BGE-Gemma2' even are —
        there is just one 'AI Search' button:
          - 0 models have embeddings built yet -> return [] (caller shows
            the "no AI data, run --update data" status).
          - exactly 1 model has data -> use that one directly, unchanged
            behavior from before.
          - both models have data -> query both and BLEND: scores are
            min-max normalized per-model first (their raw cosine-similarity
            ranges aren't comparable to each other), then combined per path.
            A path both models agree on (found by both) is boosted above a
            path only one model surfaced, since cross-model agreement is a
            stronger relevance signal than either model alone.
        """
        available = [k for k in SEMANTIC_MODELS if _check_semantic_table_count(k) > 0]
        if not available:
            return []
        if len(available) == 1:
            return self._semantic_search(query, top_k=top_k, model_key=available[0])

        # 2+ models available.
        if _dual_resident_allowed():
            # Fast path: this machine has enough spare RAM to keep BOTH
            # models loaded permanently (typically one on GPU, one on CPU —
            # see _pick_device_for_model), so encode the query with both AT
            # THE SAME TIME in 2 threads instead of paying a full
            # unload/reload between them on every blended search.
            results_by_model = {}
            def _run(mk, is_primary):
                try:
                    if is_primary:
                        if not _load_semantic_model(mk):
                            return
                        obj = _sem_model
                    else:
                        if not _load_semantic_model_secondary(mk):
                            return
                        obj = _sem_model2
                    pairs = self._semantic_search(query, top_k=top_k * 2, model_key=mk, model_obj=obj)
                    if pairs:
                        results_by_model[mk] = pairs
                except Exception as _e:
                    print(f"[Semantic] blend thread error ({mk}): {_e}")

            threads = [threading.Thread(target=_run, args=(available[0], True), daemon=True)]
            threads += [threading.Thread(target=_run, args=(mk, False), daemon=True) for mk in available[1:]]
            for t in threads:
                t.start()
            for t in threads:
                t.join()
        else:
            # Fallback: not enough spare RAM to keep both models resident.
            # Same behavior as before -- sequential, each switch reloads a
            # model from disk (slower), but safe on constrained machines.
            results_by_model = {}
            for mk in available:
                pairs = self._semantic_search(query, top_k=top_k * 2, model_key=mk)
                if pairs:
                    results_by_model[mk] = pairs

        per_model = {}
        for mk, pairs in results_by_model.items():
            vals = [s for _, s in pairs]
            lo, hi = min(vals), max(vals)
            span = (hi - lo) or 1.0
            # normalized 0..1 within this model's own returned spread
            per_model[mk] = {p: (s - lo) / span for p, s in pairs}

        if not per_model:
            return []
        if len(per_model) == 1:
            # only one model actually returned anything post-threshold even
            # though both had data -- just use its (un-normalized) raw scores
            only_mk = next(iter(per_model))
            return self._semantic_search(query, top_k=top_k, model_key=only_mk)

        AGREEMENT_BONUS = 0.15
        all_paths = set()
        for d in per_model.values():
            all_paths.update(d.keys())

        blended = []
        for p in all_paths:
            hits = [d[p] for d in per_model.values() if p in d]
            if len(hits) >= 2:
                combined = min(1.0, (sum(hits) / len(hits)) + AGREEMENT_BONUS)
            else:
                combined = hits[0]
            blended.append((p, combined))

        blended.sort(key=lambda x: x[1], reverse=True)
        return blended[:top_k]

    def _ai_search_and_update(self, query):
        """Run semantic search and merge with cached BM25 results for ALL tabs.

        v-simplify: used to toggle back to BM25-only if called again while
        already in AI mode (click AI Search a second time = revert). Per
        user request, AI Search is now one-directional — clicking it again
        just re-runs AI search. To go back to plain BM25 results, close and
        reopen the results window (the user's own stated workflow)."""
        # Strategy: BM25 results always on top (sorted by BM25 score),
        # AI-only results appended below (sorted by semantic score).
        # This respects BM25 precision while still surfacing AI extras.
        BM25_WEIGHT = 0.7
        SEM_WEIGHT  = 0.3

        def _set_btn(text, fg, bg, state="normal"):
            def _do():
                try:
                    if self._ai_search_btn and self._ai_search_btn.winfo_exists():
                        self._ai_search_btn.config(text=text, fg=fg, bg=bg, state=state)
                except Exception: pass
            self.root.after(0, _do)

        def _set_status(text, color):
            def _do():
                try:
                    if hasattr(self, '_hybrid_status_lbl') and self._hybrid_status_lbl.winfo_exists():
                        self._hybrid_status_lbl.config(text=text, fg=color)
                except Exception: pass
            self.root.after(0, _do)

        def _clear_pending():
            # v-fix: only clear if it's still THIS query's pending flag --
            # a newer call for a different query may have set it again in
            # the meantime, and that one isn't ours to clear.
            if getattr(self, "_ai_pending_query", None) == query:
                self._ai_pending_query = None

        _set_btn("⏳ Running...", "#ffcc00", "#2a2a1a", "disabled")
        _set_status("⏳ AI searching...", "#ffcc00")

        try:
            # v-revert: back to using whichever single model the user picked
            # in the dropdown (self._sem_model_key), per user request to
            # keep manual model choice instead of always auto-blending.
            # _semantic_search_blend() is left defined above in case blend
            # mode is wanted again later, just no longer called here.
            if not _load_semantic_model():
                _set_btn("🤖 AI Search", "#7ec8e3", "#1e3a5f")
                _set_status("⚠️ AI model not available", "#ff6666")
                self.root.after(0, _clear_pending)
                return

            sem_pairs = self._semantic_search(query)
            if not sem_pairs:
                _set_btn("🤖 AI Search", "#7ec8e3", "#1e3a5f")
                _set_status("ℹ️ AI: no additional results (run --update data if this is unexpected)", "#888888")
                self.root.after(0, _clear_pending)
                return

            sem_scores = {p: s for p, s in sem_pairs}   # path → 0.0-1.0

            # ── Format keyword → extension filter ────────────────────────
            # Parsed ONCE here so it can be applied to BOTH the File Content
            # merge (A) and the File Name/Folder Name merge (B). Previously
            # this was only computed inside (B), so a query like "tai lieu
            # excel lien quan simpack adas" correctly filtered File Name to
            # 100% .xlsx, but File Content still showed whatever file types
            # scored highest semantically (e.g. 80% PDF/20% PPT) — the
            # "excel" keyword never reached the content merge at all.
            import re as _re_ai
            _FORMAT_MAP = {
                "excel": [".xlsx", ".xls", ".csv"],
                "xlsx":  [".xlsx"],
                "xls":   [".xls"],
                "csv":   [".csv"],
                "word":  [".doc", ".docx"],
                "doc":   [".doc", ".docx"],
                "docx":  [".docx"],
                "pdf":   [".pdf"],
                "ppt":   [".ppt", ".pptx"],
                "pptx":  [".pptx"],
                "powerpoint": [".ppt", ".pptx"],
                "python": [".py"],
                "py":    [".py"],
                "text":  [".txt", ".log"],
                "txt":   [".txt"],
                "log":   [".log"],
                "image": [".png", ".jpg", ".jpeg", ".bmp", ".gif"],
                "png":   [".png"],
                "jpg":   [".jpg", ".jpeg"],
                "zip":   [".zip", ".7z", ".rar"],
                "onenote": [".one"],
                "one":   [".one"],
                "msg":   [".msg"],
                "outlook": [".msg"],
                "email": [".msg"],
            }
            query_lower = query.lower()
            raw_parts = _re_ai.split(r'[\s]+', query_lower)
            clean_parts = [_re_ai.sub(r'[\-_\.\(\)]', '', p) for p in raw_parts]

            ext_filter_set = set()   # extensions to filter by (used in A and B)
            consumed = set()         # indices "used up" by a matched format keyword

            # v9.14 fix: some format keywords are typed as two separate words
            # ("power point" for PowerPoint). The old single-token check never
            # matched "power" or "point" individually against the
            # _FORMAT_MAP, so ext_filter_set stayed EMPTY for such queries —
            # both File Name and File Content tabs then fell back to raw
            # relevance ranking and showed every file type mixed together
            # instead of being restricted to .ppt/.pptx. Check adjacent word
            # pairs first, then fall back to single words.
            for i in range(len(clean_parts) - 1):
                combo = clean_parts[i] + clean_parts[i + 1]
                if combo in _FORMAT_MAP:
                    ext_filter_set.update(_FORMAT_MAP[combo])
                    consumed.add(i); consumed.add(i + 1)
            for i, clean in enumerate(clean_parts):
                if i in consumed:
                    continue
                if clean in _FORMAT_MAP:
                    ext_filter_set.update(_FORMAT_MAP[clean])
                    consumed.add(i)

            # Whatever wasn't consumed as a format keyword becomes filename
            # search tokens (shared with section B below, so both agree on
            # what counts as a "real" name token vs. a format directive).
            name_toks = []
            for i, part in enumerate(raw_parts):
                if i in consumed:
                    continue
                # v5.8: drop filler words here too, otherwise e.g. "to"
                # LIKE-matches every ".toc" file.
                subtoks = [t for t in _re_ai.split(r'[\s\-_\.]+', part)
                           if len(t) >= 2 and t not in STOPWORDS]
                name_toks.extend(subtoks)
            seen_t = set()
            name_toks_clean = []
            for t in name_toks:
                tl = t.lower()
                if tl not in seen_t and not tl.isdigit():
                    seen_t.add(tl)
                    name_toks_clean.append(tl)
            name_toks = name_toks_clean[:8]  # cap at 8 tokens

            # ── A: Merge File Content ─────────────────────────────────────
            # Group 1: BM25 paths → on top, sorted by BM25 score (small AI boost)
            # Group 2: AI-only paths → appended below, sorted by semantic score
            bm25_cont = self._last_bm25_cont_res
            bm25_scores_c = {}
            size_map_c = {}
            for item in bm25_cont:
                path = item[0]; sz = item[1] if len(item) > 1 else 0
                size_map_c[path] = sz
                bm25_scores_c[path] = (item[2] / 99.0) if len(item) > 2 and item[2] else 0.0

            new_paths = [p for p in sem_scores if p not in size_map_c]
            if new_paths and self.db_conn:
                try:
                    ph2 = ",".join("?" * len(new_paths))
                    c2 = self.db_conn.cursor()
                    c2.execute(f"SELECT path, size FROM files WHERE path IN ({ph2})", new_paths)
                    for row in c2.fetchall():
                        size_map_c[row[0]] = row[1]
                except Exception: pass

            # v9.13 fix: if the query contains a format keyword (e.g. "excel"),
            # drop content-merge candidates whose extension doesn't match —
            # otherwise File Content tab ignores the format keyword entirely
            # and shows whatever file types scored highest semantically.
            if ext_filter_set:
                bm25_scores_c = {p: v for p, v in bm25_scores_c.items()
                                  if os.path.splitext(p)[1].lower() in ext_filter_set}
                sem_scores_c = {p: v for p, v in sem_scores.items()
                                 if os.path.splitext(p)[1].lower() in ext_filter_set}
            else:
                sem_scores_c = sem_scores

            bm25_group = []
            for path in bm25_scores_c:
                b = bm25_scores_c[path]
                s = sem_scores_c.get(path, 0.0)
                score = BM25_WEIGHT * b + SEM_WEIGHT * s
                bm25_group.append((path, size_map_c.get(path, 0), score))
            bm25_group.sort(key=lambda x: x[2], reverse=True)

            ai_only_group = []
            for path in sem_scores_c:
                if path not in bm25_scores_c:
                    ai_only_group.append((path, size_map_c.get(path, 0), sem_scores_c[path]))
            ai_only_group.sort(key=lambda x: x[2], reverse=True)

            combined = bm25_group + ai_only_group
            max_c = combined[0][2] if combined else 1.0
            cont_res = [(p, sz, int((sc / max_c) * 99) if max_c > 0 else 0)
                        for p, sz, sc in combined]

            # ── B: File Name + Folder Name — AI-augmented ────────────────
            # Problem: BM25 file_res may be empty when query has special chars/short tokens
            # (e.g. "Python file 3.2_e5-base" → "file" is too short, "3.2_e5-base" has dots)
            # Solution: pull file records for ALL paths that semantic returned, then
            # merge with BM25 file_res (BM25 on top, AI-only extras below).
            bm25_files = self._last_bm25_file_res  # list of (type, name, path, size)

            # v-fix: same issue as the BM25-toggle restore path above — the
            # DB-backed name search can miss files the live MFT scan already
            # found (stale/partial index). Merge those in so AI Search never
            # regresses File Name/Folder Name to fewer results than plain
            # MFT search had.
            mft_fallback = (list(getattr(self, '_mft_file_res', []) or []) +
                            list(getattr(self, '_mft_folder_res', []) or []))
            if mft_fallback:
                seen_p0 = {r[2] for r in bm25_files}
                bm25_files = list(bm25_files) + [r for r in mft_fallback if r[2] not in seen_p0]

            # Collect AI-suggested paths from content semantic results
            sem_paths_all = list(sem_scores.keys())

            # Also look up parent folders and sibling files for semantic paths
            ai_file_extras = []
            if sem_paths_all and self.db_conn:
                try:
                    # NOTE: ext_filter_set AND name_toks were already computed
                    # once above (before section A) and apply here unchanged —
                    # not recomputed, so File Name/Folder Name and File
                    # Content always agree on what's a format keyword vs. a
                    # real name token.

                    # Get file records whose path appears in semantic results
                    ph_sem = ",".join("?" * len(sem_paths_all))
                    c3 = self.db_conn.cursor()
                    c3.execute(
                        f"SELECT type, name, path, size FROM files "
                        f"WHERE path IN ({ph_sem}) AND type != 'Folder' LIMIT 500",
                        sem_paths_all)
                    sem_file_rows = c3.fetchall()

                    # ── Name token search with OR logic ──────────────────
                    # Each token searched independently → union of results
                    # Score = how many tokens match the filename (used for ranking)
                    token_score_map = {}   # path → match_count
                    token_row_map   = {}   # path → row tuple

                    # v9.16 fix: this used to be "name LIKE ? OR path LIKE ?".
                    # Matching the FULL PATH let a keyword that only appears
                    # in some distant ancestor folder (e.g. an installed
                    # tool's directory happening to contain "adas" in its
                    # name) pull in EVERY file under that folder into "AI
                    # Search" results — including totally unrelated generic
                    # files (language/flag icons, etc.) that don't mention
                    # the query at all. Filename-only keeps this fallback
                    # tied to what the query actually asked about.
                    for tok in name_toks:
                        c3.execute(
                            "SELECT type, name, path, size FROM files "
                            "WHERE name LIKE ? AND type != 'Folder' LIMIT 500",
                            (f"%{tok}%",))
                        for row in c3.fetchall():
                            p = row[2]
                            token_score_map[p] = token_score_map.get(p, 0) + 1
                            if p not in token_row_map:
                                token_row_map[p] = row

                    # Apply ext filter if any format keywords were detected
                    if ext_filter_set:
                        token_row_map = {
                            p: r for p, r in token_row_map.items()
                            if os.path.splitext(p)[1].lower() in ext_filter_set
                        }
                        token_score_map = {p: s for p, s in token_score_map.items() if p in token_row_map}
                        # Also filter sem_file_rows by ext
                        sem_file_rows = [r for r in sem_file_rows
                                         if os.path.splitext(r[2])[1].lower() in ext_filter_set]

                    # Sort token results by match count (descending) — files matching more tokens rank higher
                    token_rows_sorted = sorted(token_row_map.items(),
                                               key=lambda x: token_score_map.get(x[0], 0),
                                               reverse=True)
                    token_rows = [row for _, row in token_rows_sorted]

                    # ── Folder search ─────────────────────────────────────
                    folder_rows = []
                    for tok in name_toks[:5]:
                        c3.execute(
                            "SELECT type, name, path, size FROM files "
                            "WHERE name LIKE ? AND type = 'Folder' LIMIT 200",
                            (f"%{tok}%",))
                        folder_rows.extend(c3.fetchall())

                    # Deduplicate by path; BM25 file_res paths excluded (they're already on top)
                    seen_paths = {r[2] for r in bm25_files}
                    for row in (sem_file_rows + token_rows):
                        if row[2] not in seen_paths:
                            seen_paths.add(row[2])
                            ai_file_extras.append(row)
                    for row in folder_rows:
                        if row[2] not in seen_paths:
                            seen_paths.add(row[2])
                            ai_file_extras.append(row)
                except Exception as _dbe:
                    print(f"[AI Search] file name DB query error: {_dbe}")

            # Re-rank using BM25 on filename — higher match-count files rise naturally
            import re as _re_ai2
            def _tok(t):
                return [x.lower() for x in _re_ai2.split(r'[\s\W]+', str(t)) if len(x) >= 2]

            def _bm25_rank(rows, q_tok):
                if not rows or not q_tok:
                    return rows
                try:
                    from rank_bm25 import BM25Okapi
                    # v-fix: filename weighted 3x over full path (same idea as
                    # _weighted_fn in _smart_search_realtime). Previously
                    # basename+path were weighted equally, so a file living
                    # in a folder that happens to contain query keywords
                    # (e.g. ".../ADAS/Realtime/UDP/simpack.2024-...log")
                    # could outrank a file whose actual NAME matches the
                    # query better but whose parent folders don't mention
                    # every keyword — e.g. a topically-relevant .pptx titled
                    # "SIMPACK Realtime..." losing to an unrelated .log file
                    # that just happens to sit in a deeply-nested matching
                    # folder path.
                    corpus = [_tok((os.path.basename(r[2]) + " ") * 3 + r[2]) for r in rows]
                    bm25_fn = BM25Okapi(corpus)
                    raw = bm25_fn.get_scores(q_tok)
                    ranked = sorted(zip(rows, raw), key=lambda x: x[1], reverse=True)
                    return [r for r, _ in ranked]
                except Exception:
                    return rows

            q_tok_fn = _tok(query)
            bm25_files_ranked = _bm25_rank(bm25_files, q_tok_fn)  # BM25 group re-ranked
            ai_extras_ranked  = _bm25_rank(ai_file_extras, q_tok_fn)  # AI extras ranked by name relevance

            # Combine: BM25 on top, AI extras below
            file_res_new = bm25_files_ranked + ai_extras_ranked

            only_files_new   = [r for r in file_res_new if str(r[0]).lower() != "folder"]
            only_folders_new = [r for r in file_res_new if str(r[0]).lower() == "folder"]

            # Priority sort: office/pdf/msg first, txt/md middle, log/html/code/... last
            cont_res         = self._sort_priority(cont_res, 0)
            only_files_new   = self._sort_priority(only_files_new, 1)

            # ── C: Update all 3 trees on main thread ─────────────────────
            def _update_ui():
                try:
                    if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win):
                        return
                    self._all_content_data = cont_res
                    self._all_files_data   = only_files_new
                    # FIX 3: cache AI-specific results for split pane display
                    self._ai_cont_res  = cont_res
                    self._ai_file_res  = only_files_new + only_folders_new

                    for item in self.tree_c.get_children():   self.tree_c.delete(item)
                    for item in self.tree_f.get_children():   self.tree_f.delete(item)
                    for item in self.tree_fol.get_children(): self.tree_fol.delete(item)

                    for rn, item in enumerate(cont_res, start=1):
                        self._insert_row(self.tree_c, item, rn, 0)
                    for rn, item in enumerate(only_files_new, start=1):
                        self._insert_row(self.tree_f, item, rn, 1)
                    for rn, item in enumerate(only_folders_new, start=1):
                        self._insert_row(self.tree_fol, item, rn, 2)

                    if self.content_filter_count_label:
                        self.content_filter_count_label.config(text=f"{len(cont_res)} files")
                    if self.filter_count_label:
                        self.filter_count_label.config(text=f"{len(only_files_new)} files")
                    self.nb.tab(0, text=" File Name ")
                    self.nb.tab(1, text=" Folder Name ")

                    self._ai_mode_active = True
                    self._ai_active_query = query  # v5.4: remember which query AI mode belongs to
                    _clear_pending()
                    _set_btn("🤖 AI Search", "#7ec8e3", "#1e3a5f")
                    _set_status("🔀 Hybrid: BM25 + AI", "#7ec8e3")
                    # v-new (yêu cầu: label "🤖 AI Search offline" ĐANG XÁM
                    # trong AI Chat phải chuyển bấm được NGAY khi AI Search
                    # vừa chạy xong, không cần chờ câu hỏi mới):
                    self._activate_ai_search_links_for_query(query)
                    # Open AI split window (left/right 60-40)
                    self.root.after(50, self._open_ai_split_win)
                except Exception as _ue:
                    print(f"[AI Search] UI update error: {_ue}")
                    _clear_pending()

            self.root.after(0, _update_ui)

        except Exception as _e:
            print(f"[AI Search] Error: {_e}")
            _set_btn("🤖 AI Search", "#7ec8e3", "#1e3a5f")
            _set_status("⚠️ AI search failed", "#ff6666")
            self.root.after(0, _clear_pending)

    # ══════════════════════════════════════════════════════════════════
    # ONLINE AI CHAT BLOCK (v1.3) — Gemini Flash / GPT-OSS 120B chat panel
    # sitting below the 2 results panes (Default + offline AI Search), same
    # idea as the app_astro-weather chatbox: 1 input field + an up-arrow (↑)
    # send button INSIDE the input field + a small icon-only "New Chat"
    # button next to it. Gemini Flash is the default model; if Gemini hits
    # a rate limit (429/RESOURCE_EXHAUSTED), the chat automatically switches
    # to GPT-OSS 120B (via Groq) for the rest of the session, no manual
    # model picking needed. Answers are still grounded in the REAL file
    # content excerpts from self._last_bm25_cont_res (latest local BM25/
    # FTS5 results) -- same anti-hallucination principle as before.
    #
    # v1.3: the panel switched from a dark theme to a light theme matching
    # the results panes above it, its default height was halved, and it now
    # auto-asks the AI to summarize the fresh results the moment AI Search
    # is turned on (see _auto_ai_chat_after_search) instead of staying
    # empty until the user types something.
    # ══════════════════════════════════════════════════════════════════

    def _draw_new_chat_icon(self, parent, size=18, color=None):
        """Vector "New Chat" icon.

        v-fix (icon nhìn không đẹp, theo yêu cầu): the old version was a
        flat rectangle "speech bubble" with a small triangular tail bolted
        onto one corner, which read as slightly lopsided/unclear at icon
        size. Replaced with a plain rounded-square outline + "+" -- a
        cleaner, more universally-recognized "start new" glyph, and it now
        visually pairs with the rounded-square send button below (see
        _CHAT_SEND_BG / the send button wrapper in _build_online_chat_panel)
        instead of clashing stylistically with it.

        v-new (theo yêu cầu: nút New Chat nên bị greyout trong lúc
        "Thinking..."): trước đây icon này KHÔNG có cách nào để disable --
        bấm được ngay cả khi 1 request đang chạy, dẫn tới race giống hệt
        vấn đề send/EN-VI-JP đã sửa ở _set_chat_busy (click "vô hình" giữa
        chừng một câu trả lời). Theo đúng mẫu của _draw_send_button /
        _set_send_btn_enabled (Canvas không hỗ trợ "-state" đáng tin cậy),
        lưu lại id của outline + 2 nét "+" và dùng 1 cờ Python
        (self._chat_newchat_enabled) để chặn click handler, dim màu outline
        khi disabled thay vì đổi state widget."""
        color = color or "#1a5fb4"
        disabled_color = "#b5b5b5"
        cv = tk.Canvas(parent, width=size, height=size, bg=BG_COLOR, highlightthickness=0)
        lw = max(1, round(size * 0.12))
        pad = size * 0.09
        r = size * 0.24
        pts = [pad + r, pad, size - pad - r, pad, size - pad, pad,
               size - pad, pad + r, size - pad, size - pad - r, size - pad, size - pad,
               size - pad - r, size - pad, pad + r, size - pad, pad, size - pad,
               pad, size - pad - r, pad, pad + r, pad, pad]
        outline_id = cv.create_polygon(pts, outline=color, width=lw, fill="", smooth=True, joinstyle=tk.ROUND)
        # "+" centered inside = start a new conversation
        cx, cy = size / 2, size / 2
        pl = size * 0.17
        h_id = cv.create_line(cx - pl, cy, cx + pl, cy, fill=color, width=lw, capstyle=tk.ROUND)
        v_id = cv.create_line(cx, cy - pl, cx, cy + pl, fill=color, width=lw, capstyle=tk.ROUND)
        cv._newchat_color = color
        cv._newchat_disabled_color = disabled_color
        cv._newchat_item_ids = (outline_id, h_id, v_id)
        return cv

    def _set_new_chat_btn_enabled(self, enabled):
        """Enable/disable the vector "New Chat" icon drawn by
        _draw_new_chat_icon -- same technique as _set_send_btn_enabled
        (see that method's docstring)."""
        try:
            cv = getattr(self, "_chat_new_chat_btn", None)
            if cv is None or not cv.winfo_exists():
                return
            self._chat_newchat_enabled = enabled
            fill = cv._newchat_color if enabled else cv._newchat_disabled_color
            outline_id, h_id, v_id = cv._newchat_item_ids
            cv.itemconfig(outline_id, outline=fill)
            cv.itemconfig(h_id, fill=fill)
            cv.itemconfig(v_id, fill=fill)
            cv.config(cursor="hand2" if enabled else "arrow")
        except Exception:
            pass

    def _rounded_rect_points(self, x0, y0, x1, y1, r):
        """Point list for a rounded-rectangle create_polygon(..., smooth=True)
        -- shared helper so any rounded-square UI element (send button, New
        Chat icon, future ones) draws corners the same way."""
        return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r,
                x1, y1 - r, x1, y1, x1 - r, y1, x0 + r, y1,
                x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]

    def _draw_send_button(self, parent, size=25):
        """Vector send button -- a rounded-square filled with _CHAT_SEND_BG
        and a bold upward arrow, drawn on a Canvas.

        v-fix (theo yêu cầu -- vuông bo góc + mũi tên giống Claude): a
        plain tk.Button can be forced into a square (see the old
        send_wrap Frame trick) but classic Tk buttons can't have rounded
        corners on Windows without OS theming, and the "↑" glyph's exact
        look depends on the system font. Drawing it as vector shapes on a
        Canvas gives full control over both: a real rounded-rect
        background, and a clean solid triangular arrowhead + stem that
        looks the same on every machine, matching Claude's send button
        instead of whatever "↑" happens to render as locally.

        Enabling/disabling is handled via _set_send_btn_enabled() rather
        than the usual widget "state" option -- Canvas doesn't reliably
        support disabling its own click bindings through "-state" the way
        Button does, so a plain Python flag (self._chat_send_enabled)
        gates the click handler instead, with the fill color dimmed to
        show the disabled look."""
        cv = tk.Canvas(parent, width=size, height=size, bg=ENTRY_BG,
                        highlightthickness=0, cursor="hand2")
        r = size * 0.30
        bg_id = cv.create_polygon(self._rounded_rect_points(1, 1, size - 1, size - 1, r),
                                   smooth=True, fill=_CHAT_SEND_BG_DEFAULT, outline="")
        cx, cy = size / 2, size / 2
        aw, ah = size * 0.20, size * 0.17
        sw, sh = size * 0.09, size * 0.19
        head_id = cv.create_polygon(cx, cy - ah * 1.35, cx - aw, cy + ah * 0.25, cx + aw, cy + ah * 0.25,
                                     fill="#ffffff", outline="#ffffff", joinstyle=tk.ROUND)
        stem_id = cv.create_rectangle(cx - sw, cy + ah * 0.05, cx + sw, cy + ah * 0.05 + sh,
                                       fill="#ffffff", outline="#ffffff")
        self._chat_send_bg_id = bg_id
        self._chat_send_arrow_ids = (head_id, stem_id)
        self._chat_send_enabled = True
        cv.bind("<Button-1>", lambda e: self._send_online_chat_msg() if self._chat_send_enabled else None)
        cv.bind("<Enter>", lambda e: cv.itemconfig(bg_id, fill=_CHAT_SEND_BG_HOVER) if self._chat_send_enabled else None)
        cv.bind("<Leave>", lambda e: cv.itemconfig(bg_id, fill=_CHAT_SEND_BG_DEFAULT) if self._chat_send_enabled else None)
        return cv

    def _set_send_btn_enabled(self, enabled):
        """Enable/disable the vector send button drawn by _draw_send_button
        -- see that method's docstring for why this replaces
        .config(state=...) for this particular widget."""
        try:
            cv = self._chat_send_btn
            if cv is None or not cv.winfo_exists():
                return
            self._chat_send_enabled = enabled
            cv.itemconfig(self._chat_send_bg_id,
                           fill=_CHAT_SEND_BG_DEFAULT if enabled else _CHAT_SEND_BG_DISABLED)
            cv.config(cursor="hand2" if enabled else "arrow")
        except Exception:
            pass

    # ── "+" Add file button (v-new, theo yêu cầu) ──────────────────────
    # Người dùng đính kèm 1 file thật trên máy (PDF/DOCX/PPTX/XLSX/TXT...,
    # tối đa 2MB) để AI Chat LUÔN đọc trọn nội dung file đó, không phụ
    # thuộc file đó có lọt vào top-N của BM25 hay không -- xem
    # CHAT_ATTACH_MAX_BYTES/CHAT_ATTACH_ALLOWED_EXT và cách
    # self._chat_attached_file được ghép vào context trong
    # _online_chat_worker.
    def _draw_attach_icon(self, parent, size=16, color=None):
        """Vector "+" icon in a circle -- deliberately a different shape
        (circle vs. New Chat's rounded square) so the two buttons never
        get mixed up at a glance even though both use a plain "+" glyph."""
        color = color or "#1a5fb4"
        cv = tk.Canvas(parent, width=size, height=size, bg=ENTRY_BG, highlightthickness=0)
        lw = max(1, round(size * 0.12))
        pad = size * 0.09
        oval_id = cv.create_oval(pad, pad, size - pad, size - pad, outline=color, width=lw)
        cx, cy = size / 2, size / 2
        pl = size * 0.18
        h_id = cv.create_line(cx - pl, cy, cx + pl, cy, fill=color, width=lw, capstyle=tk.ROUND)
        v_id = cv.create_line(cx, cy - pl, cx, cy + pl, fill=color, width=lw, capstyle=tk.ROUND)
        cv._attach_color = color
        cv._attach_disabled_color = "#b5b5b5"
        cv._attach_item_ids = (oval_id, h_id, v_id)
        return cv

    def _set_attach_btn_enabled(self, enabled):
        try:
            cv = getattr(self, "_chat_attach_btn", None)
            if cv is None or not cv.winfo_exists():
                return
            self._chat_attach_enabled = enabled
            fill = cv._attach_color if enabled else cv._attach_disabled_color
            oval_id, h_id, v_id = cv._attach_item_ids
            cv.itemconfig(oval_id, outline=fill)
            cv.itemconfig(h_id, fill=fill)
            cv.itemconfig(v_id, fill=fill)
            cv.config(cursor="hand2" if enabled else "arrow")
        except Exception:
            pass

    def _pick_attach_file(self):
        """"+" Add file click handler -- opens a native file picker, checks
        the 2MB cap up front (before spending time extracting anything),
        then extracts content in a background thread (reusing
        get_file_content -- the exact same extractor already used to index
        pdf/docx/pptx/xlsx/... for BM25, so every format it already
        supports "just works" here too)."""
        if getattr(self, "_chat_busy", False) or not getattr(self, "_chat_attach_enabled", True):
            return
        exts = " ".join(f"*{e}" for e in sorted(CHAT_ATTACH_ALLOWED_EXT))
        path = filedialog.askopenfilename(
            title="Add file to AI Chat",
            filetypes=[("Supported files", exts), ("All files", "*.*")])
        if not path:
            return
        self._attach_local_file(path)

    def _attach_local_file(self, path, _dlg_title="Add file"):
        """v-new (refactor cho yêu cầu 3 -- drag & drop file vào chatbox):
        logic kiểm tra (size <= CHAT_ATTACH_MAX_BYTES, đuôi file có trong
        CHAT_ATTACH_ALLOWED_EXT hay là text file) + đính kèm, tách riêng
        khỏi _pick_attach_file() để DÙNG CHUNG cho cả 2 đường: chọn qua
        dialog "+" (askopenfilename) VÀ kéo-thả file từ ngoài vào
        (_on_chat_files_dropped) -- cùng 1 quy tắc <5MB, cùng danh sách
        định dạng hỗ trợ, không viết trùng 2 lần."""
        if getattr(self, "_chat_busy", False) or not getattr(self, "_chat_attach_enabled", True):
            return
        try:
            size = os.path.getsize(path)
        except Exception as _e:
            messagebox.showwarning(_dlg_title, f"Không đọc được file:\n{_e}")
            return
        if size > CHAT_ATTACH_MAX_BYTES:
            messagebox.showwarning(
                _dlg_title,
                f"File quá lớn ({size/1024/1024:.1f} MB) -- giới hạn hiện tại là "
                f"{CHAT_ATTACH_MAX_BYTES/1024/1024:.0f} MB.")
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in CHAT_ATTACH_ALLOWED_EXT and not is_text_file(path):
            messagebox.showwarning(_dlg_title, f"Định dạng '{ext}' chưa được hỗ trợ.")
            return
        # v-new (yêu cầu 4): ảnh đi theo đường riêng (đọc raw bytes để gửi
        # multimodal cho Gemini) -- không qua get_file_content (chỉ
        # extract TEXT, không có ý nghĩa với ảnh).
        if ext in CHAT_IMAGE_EXT:
            self._attach_image_file(path)
            return
        self._set_attach_btn_enabled(False)
        self._attach_chip_var.set(f"📎 Đang đọc {os.path.basename(path)}...")
        self._update_attach_chip_ui(show="loading")
        threading.Thread(target=self._attach_file_worker, args=(path,), daemon=True).start()

    def _attach_image_file(self, path):
        """v-new (yêu cầu 4 -- dán/chọn ảnh vào AI Chat): đọc raw bytes của
        1 file ảnh (từ nút "+" HOẶC từ clipboard, xem
        _try_paste_image_from_clipboard) và lưu vào self._chat_attached_file
        với cờ is_image=True. Không có bước "extract text" nào ở đây --
        _online_chat_worker sẽ tự nhận ra is_image=True và gửi image_bytes
        đi dưới dạng inline_data cho Gemini (xem _call_online_ai_chat);
        với GPT-OSS/Qwen (không hỗ trợ vision), nó chỉ chèn 1 ghi chú thay
        vì gửi ảnh đi vô ích -- xem CHAT_IMAGE_NOT_SUPPORTED_NOTE."""
        try:
            size = os.path.getsize(path)
            if size > CHAT_ATTACH_MAX_BYTES:
                messagebox.showwarning(
                    "Add file",
                    f"Ảnh quá lớn ({size/1024/1024:.1f} MB) -- giới hạn hiện tại là "
                    f"{CHAT_ATTACH_MAX_BYTES/1024/1024:.0f} MB.")
                return
            with open(path, "rb") as f:
                raw = f.read()
        except Exception as _e:
            messagebox.showwarning("Add file", f"Không đọc được ảnh:\n{_e}")
            return
        ext = os.path.splitext(path)[1].lower()
        mime = CHAT_IMAGE_MIME_BY_EXT.get(ext, "image/png")
        self._chat_attached_file = {
            "path": path,
            "text": "[Hình ảnh đính kèm -- xem qua Gemini Flash (vision)]",
            "is_image": True,
            "image_bytes": raw,
            "image_mime": mime,
        }
        self._update_attach_chip_ui(show="attached")

    def _try_paste_image_from_clipboard(self):
        """v-new (yêu cầu 4 -- dán ảnh trực tiếp vào ô chat bằng Ctrl+V,
        thay cho nút "New Chat" vừa bị gỡ): kiểm tra clipboard hiện có
        đang giữ 1 ẢNH (ảnh copy từ trình duyệt/Snipping Tool/Paint/...)
        thay vì text hay không, dùng PIL.ImageGrab.grabclipboard().

        v-note (giới hạn nền tảng): ImageGrab.grabclipboard() hỗ trợ đầy
        đủ trên Windows và macOS. Trên Linux/X11-Wayland, PIL không đọc
        được ảnh raw trong clipboard hệ thống (chỉ đọc được đường dẫn file
        nếu có) -- gọi hàm này trên Linux thường trả về None hoặc list
        đường dẫn file, hàm tự return False và để paste-text mặc định của
        Tk chạy tiếp như trước, không có gì bị vỡ.

        Trả về True nếu đã xử lý xong 1 ảnh (nơi gọi nên return "break" để
        chặn Tk tự paste text/đường dẫn file vào ô nhập); False nếu
        clipboard không có ảnh (nơi gọi nên để hành vi paste mặc định của
        Tk chạy tiếp bình thường)."""
        if getattr(self, "_chat_busy", False):
            return False
        try:
            from PIL import ImageGrab, Image
        except Exception:
            return False
        try:
            grabbed = ImageGrab.grabclipboard()
        except Exception:
            return False
        if grabbed is None:
            return False
        # v-note: trên Windows, copy 1 file ảnh (chứ không phải ảnh trong
        # bộ nhớ, vd copy trong File Explorer) khiến grabclipboard() trả
        # về list đường dẫn thay vì đối tượng Image -- xử lý luôn trường
        # hợp này cho tiện, mở file đó lên như đang chọn qua nút "+".
        if isinstance(grabbed, list):
            for p in grabbed:
                if os.path.splitext(p)[1].lower() in CHAT_IMAGE_EXT and os.path.isfile(p):
                    self._attach_image_file(p)
                    return True
            return False
        if not isinstance(grabbed, Image.Image):
            return False
        try:
            tmp_dir = os.path.join(tempfile.gettempdir(), "smart_search_chat_paste")
            os.makedirs(tmp_dir, exist_ok=True)
            path = os.path.join(tmp_dir, f"pasted_{int(time.time()*1000)}.png")
            # v-note: chuẩn hoá về RGB trước khi lưu PNG -- ảnh copy từ 1
            # số nguồn (vd screenshot có kênh alpha lạ, hoặc chế độ palette
            # "P") thỉnh thoảng lưu PNG lỗi màu nếu giữ nguyên mode gốc.
            if grabbed.mode not in ("RGB", "RGBA"):
                grabbed = grabbed.convert("RGB")
            grabbed.save(path, "PNG")
        except Exception as _e:
            print(f"[Chat] paste-image: could not save clipboard image: {_e}")
            return False
        self._attach_image_file(path)
        return True

    def _attach_file_worker(self, path):
        try:
            text = self.get_file_content(path) or ""
        except Exception as _e:
            text = ""
            print(f"[Add file] extract failed for {path!r}: {_e}")

        def _done():
            self._set_attach_btn_enabled(True)
            if not text.strip():
                messagebox.showwarning(
                    "Add file",
                    f"Không trích xuất được nội dung từ:\n{os.path.basename(path)}")
                self._chat_attached_file = None
                self._update_attach_chip_ui(show="none")
                return
            self._chat_attached_file = {"path": path, "text": text}
            self._update_attach_chip_ui(show="attached")
        self.root.after(0, _done)

    def _remove_attached_file(self):
        self._chat_attached_file = None
        self._update_attach_chip_ui(show="none")

    def _update_attach_chip_ui(self, show=None):
        """Redraw the small "📎 filename ×" chip above the chat input row.
        show=None re-derives from self._chat_attached_file's current
        state (used e.g. right after "New Chat" clears it)."""
        try:
            row = getattr(self, "_attach_chip_row", None)
            if row is None or not row.winfo_exists():
                return
            for w in row.winfo_children():
                w.destroy()
            if show is None:
                show = "attached" if getattr(self, "_chat_attached_file", None) else "none"
            _input_row = getattr(self, "_chat_input_row", None)
            _pack_kwargs = dict(fill="x", padx=(28, 8), pady=(0, 2))
            if _input_row is not None and _input_row.winfo_exists():
                _pack_kwargs["before"] = _input_row
            if show == "loading":
                tk.Label(row, textvariable=self._attach_chip_var, bg=BG_COLOR,
                         fg=PLACE_COLOR, font=("Segoe UI", 8, "italic")).pack(side="left", padx=(2, 0))
                row.pack(**_pack_kwargs)
            elif show == "attached" and getattr(self, "_chat_attached_file", None):
                name = os.path.basename(self._chat_attached_file["path"])
                if len(name) > 40:
                    name = name[:37] + "..."
                chip = tk.Frame(row, bg="#e8eef7", highlightbackground="#8a8a8a", highlightthickness=1)
                tk.Label(chip, text=f"📎 {name}", bg="#e8eef7", fg=TEXT_COLOR,
                         font=("Segoe UI", 8)).pack(side="left", padx=(6, 2), pady=1)
                close_lbl = tk.Label(chip, text="✕", bg="#e8eef7", fg="#c0392b",
                                      font=("Segoe UI", 8, "bold"), cursor="hand2")
                close_lbl.pack(side="left", padx=(2, 6), pady=1)
                close_lbl.bind("<Button-1>", lambda e: self._remove_attached_file())
                chip.pack(side="left", padx=(2, 0))
                row.pack(**_pack_kwargs)
            else:
                row.pack_forget()
        except Exception:
            pass

    def _toggle_chat_placeholder(self, *args):
        """Show/hide the "Ask everything..." placeholder in the chat entry
        depending on whether it currently has text -- same overlay-Label
        technique as the main search box's placeholder (toggle_placeholder).

        v-fix (Shift+Enter): self._chat_entry is now a tk.Text, not an
        Entry+StringVar -- read its content directly instead of the old
        (now removed) self._chat_entry_var."""
        try:
            has_text = bool(self._chat_entry.get("1.0", "end-1c").strip())
            if has_text:
                self._chat_placeholder.place_forget()
            else:
                self._chat_placeholder.place(x=6, rely=0.5, anchor="w")
        except Exception:
            pass

    def _on_chat_entry_return(self, event=None):
        """v-fix (Shift+Enter = Enter, gửi luôn thay vì xuống dòng): plain
        Enter (no Shift) sends the message. Returning "break" stops Text's
        own default binding from also inserting a newline right after (Tk
        runs widget-instance bindings AND the class-default binding for the
        same event unless "break" is returned)."""
        self._send_online_chat_msg()
        return "break"

    def _on_chat_entry_modified(self, event=None):
        """Fires on every edit to the chat Text box (Tk's <<Modified>>
        virtual event) -- must reset the widget's internal modified flag
        each time or this stops firing again. Keeps the placeholder and
        the auto-grow height in sync with the CURRENT content, replacing
        the old StringVar trace_add("write", ...) which doesn't exist for
        a Text widget."""
        try:
            self._chat_entry.edit_modified(False)
        except Exception:
            pass
        self._toggle_chat_placeholder()
        self._autosize_chat_entry()

    def _autosize_chat_entry(self):
        """v-new: grow the chat Text box (up to self._chat_entry_max_lines,
        default 5) as the user types more lines with Shift+Enter, instead
        of staying stuck at a fixed 1-line height with the rest of the
        text scrolled out of view."""
        try:
            n_lines = int(self._chat_entry.index("end-1c").split(".")[0])
            n_lines = max(1, min(n_lines, getattr(self, "_chat_entry_max_lines", 5)))
            if int(self._chat_entry.cget("height")) != n_lines:
                self._chat_entry.config(height=n_lines)
        except Exception:
            pass

    def _build_online_chat_panel(self, parent):
        """Build (once) the online AI chat panel, placed below the results
        Notebook. This panel is SHARED across all 3 tabs (File Content/File
        Name/Folder Name), not duplicated per tab -- per the "1 panel
        below" requirement."""
        if self._online_chat_panel is not None:
            try:
                if self._online_chat_panel.winfo_exists():
                    return self._online_chat_panel
            except Exception:
                pass

        _CHAT_BORDER = "#d5d7db"
        _CHAT_ACCENT = "#1a5fb4"

        panel = tk.Frame(parent, bg=BG_COLOR, height=57,
                          highlightbackground=_CHAT_BORDER, highlightthickness=1)
        panel.grid_propagate(False)

        # ── Header row -- v1.6: dropped the "(Online)" suffix and the
        # ── Gemini/GPT-OSS model-name status label per user request; the
        # ── panel is shorter now too, so header padding is tighter. ──────
        hdr = tk.Frame(panel, bg=BG_COLOR)
        hdr.pack(fill="x", padx=8, pady=(2, 0))
        # v-fix (icon khó hiểu ở size nhỏ): 🌐 (globe) rendered as a plain
        # dot at this font size on some systems, with nothing "AI" about it
        # even when it did render. ✨ reads clearly at small sizes and is
        # a more common "AI" visual shorthand.
        tk.Label(hdr, text="✨ AI Chat", bg=BG_COLOR, fg=_CHAT_ACCENT,
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        # ── Model picker: Gemini Flash / GPT-OSS 120B -- v-new: lets the
        # ── user manually pick which online model AI Chat uses, so they
        # ── can spread usage across both and save on whichever one's
        # ── rate limit they're closer to, instead of always starting on
        # ── Gemini and only ever moving off it automatically on a 429.
        # ── The automatic rate-limit fallback (see _online_chat_worker)
        # ── still applies on top of whatever's picked here. ───────────
        _MODEL_SHORT_LABELS = MODEL_SHORT_LABELS
        model_row = tk.Frame(hdr, bg=BG_COLOR)
        model_row.pack(side="left", padx=(10, 0))
        _model_keys = list(ONLINE_AI_MODELS.keys())
        _model_display = [_MODEL_SHORT_LABELS.get(k, ONLINE_AI_MODELS[k]["label"]) for k in _model_keys]
        self._chat_model_var = tk.StringVar()
        _init_key = getattr(self, "_chat_preferred_model", DEFAULT_ONLINE_MODEL)
        self._chat_model_var.set(_MODEL_SHORT_LABELS.get(_init_key, _model_display[0]))

        def _on_model_pick(event=None):
            chosen_label = self._chat_model_var.get()
            for k in _model_keys:
                if _MODEL_SHORT_LABELS.get(k, ONLINE_AI_MODELS[k]["label"]) == chosen_label:
                    self._chat_preferred_model = k
                    self._chat_active_model = k
                    # v-new (yêu cầu: nhớ Chat AI model đã chọn qua configure.ini):
                    _config_set("Models", "chat_ai_model", k)
                    # v-fix (chọn lại Gemini thủ công vẫn tự bật lại
                    # GPT-OSS): cờ "rate-limited" set bởi lần fallback tự
                    # động trước đó không tự xoá cho tới khi có 1 lần gọi
                    # THÀNH CÔNG -- nhưng khi user chủ động chọn lại model
                    # này, đó chính là tín hiệu "tôi muốn thử lại ngay bây
                    # giờ" (có thể quota đã reset qua ngày mới, hoặc lần
                    # rate-limit trước chỉ là tạm thời). Xoá cờ để lượt
                    # chat kế tiếp thực sự GỌI model vừa chọn, thay vì bị
                    # _model_usable() âm thầm từ chối rồi bật lại model kia.
                    self._mark_model_rate_limited(k, False)
                    break
            try:
                self._chat_model_cbo.selection_clear()
            except Exception:
                pass

        self._chat_model_cbo = ttk.Combobox(
            model_row, textvariable=self._chat_model_var, values=_model_display,
            state="readonly", width=13, font=("Segoe UI", 7), height=len(_model_display))
        self._chat_model_cbo.pack(side="left")
        self._chat_model_cbo.bind("<<ComboboxSelected>>", _on_model_pick)
        # v-fix (theo yêu cầu): bỏ chú thích tooltip cho combobox model

        # ── Language toggle: VI / EN / JP -- switches both the auto-generated
        # ── summary question and the language the AI is instructed to answer
        # ── in (see CHAT_LANG_OPTIONS / _online_chat_worker). ─────────────
        lang_row = tk.Frame(hdr, bg=BG_COLOR)
        lang_row.pack(side="left", padx=(8, 0))
        self._chat_lang_btns = {}

        def _select_chat_lang(code):
            # v-fix (Vấn đề: bấm EN/JP không thấy dịch gì cả): trước đây
            # nếu bấm đúng lúc panel đang "Thinking..." (ví dụ auto AI
            # Chat vừa tự hỏi xong, đang chờ trả lời -- lúc người dùng
            # thường bấm đổi ngôn ngữ NGAY khi thấy panel) thì
            # _translate_current_chat() bên dưới âm thầm return ngay dòng
            # đầu (guard "if self._chat_busy: return") vì busy=True lúc
            # đó -- nút vẫn đổi màu như đã chọn (dễ hiểu lầm là đã bấm
            # thành công) nhưng bản dịch thì không bao giờ chạy, và không
            # có gì tự kích hoạt lại khi hết busy. Cùng lúc, mọi exception
            # bất ngờ trong hàm này trước đây cũng bị nuốt hoàn toàn (Tk
            # chỉ in traceback ra console, dễ bị bỏ qua khi chạy .exe).
            # Fix: khoá hẳn 3 nút VI/EN/JP trong lúc busy (giống nút Send)
            # ở _set_chat_busy bên dưới, để không còn cú click "vô hình";
            # và bọc try/except in traceback rõ ràng ở đây để lần sau nếu
            # vẫn còn lỗi, nó sẽ hiện thẳng trên console thay vì im lặng.
            try:
                prev_code = getattr(self, "_chat_lang", CHAT_DEFAULT_LANG)
                self._chat_lang = code
                for c, b in self._chat_lang_btns.items():
                    if c == code:
                        b.config(bg=_CHAT_ACCENT, fg="#ffffff", relief="sunken")
                    else:
                        b.config(bg="#e4e6ea", fg=TEXT_COLOR, relief="raised")
                # v-new (yêu cầu: nhớ ngôn ngữ VI/EN/JP qua các lần mở app):
                # lưu ngay lúc user bấm, KHÔNG đợi tới lúc đóng app -- nếu
                # app bị tắt đột ngột (crash, Task Manager, mất điện...)
                # lựa chọn gần nhất vẫn không bị mất. Chỉ ghi file khi thực
                # sự đổi ngôn ngữ (code != prev_code), để không ghi ổ đĩa
                # vô ích mỗi lần user bấm lại đúng ngôn ngữ đang chọn.
                if code != prev_code:
                    _config_set("General", "language", code)
                # v-new: if there's a conversation already on screen and the
                # language actually changed, translate what's ALREADY shown
                # instead of only affecting messages sent from now on.
                if code != prev_code and self._chat_history:
                    self._translate_current_chat(code)
            except Exception:
                import traceback
                print(f"[AI Chat] _select_chat_lang({code!r}) FAILED:")
                traceback.print_exc()

        for code in ("VI", "EN", "JP"):
            btn = tk.Button(lang_row, text=code, font=("Segoe UI", 7, "bold"),
                             bd=1, padx=4, pady=0, cursor="hand2",
                             command=lambda c=code: _select_chat_lang(c))
            btn.pack(side="left", padx=1)
            self._chat_lang_btns[code] = btn
        _select_chat_lang(getattr(self, "_chat_lang", CHAT_DEFAULT_LANG))

        # ── ◀/▶ session preview nav -- page through OLDER saved AI Chat
        # ── sessions for whatever keyword is currently in the Searchbox,
        # ── without leaving this panel (see _chat_preview_prev/_next/
        # ── _render_chat_preview). Sits right next to VI/EN/JP per request.
        nav_row = tk.Frame(hdr, bg=BG_COLOR)
        nav_row.pack(side="left", padx=(10, 0))
        self._chat_nav_prev_btn = tk.Button(
            nav_row, text="◀", font=("Segoe UI", 8, "bold"), bd=0,
            bg=BG_COLOR, fg="#c3c6cb", activebackground=BG_COLOR,
            cursor="hand2", command=self._chat_preview_prev)
        self._chat_nav_prev_btn.pack(side="left")
        # v-fix (theo yêu cầu): bỏ chú thích tooltip cho nút ◀
        self._chat_nav_lbl = tk.Label(nav_row, text="", bg=BG_COLOR, fg=PLACE_COLOR,
                                       font=("Segoe UI", 7))
        self._chat_nav_lbl.pack(side="left", padx=2)
        self._chat_nav_next_btn = tk.Button(
            nav_row, text="▶", font=("Segoe UI", 8, "bold"), bd=0,
            bg=BG_COLOR, fg="#c3c6cb", activebackground=BG_COLOR,
            cursor="hand2", command=self._chat_preview_next)
        self._chat_nav_next_btn.pack(side="left")
        # v-fix (theo yêu cầu): bỏ chú thích tooltip cho nút ▶
        self._update_chat_preview_nav_ui()

        # ── v-fix: nút "Online" toggle đã bị BỎ theo yêu cầu -- giờ AI Chat
        # ── luôn có thể tham khảo internet khi cần (self._chat_online_enabled
        # ── mặc định True, xem __init__), không cần user tự bật. Nút
        # ── "Update API" (nếu thiếu/lỗi API key) được vẽ thay vào chỗ này --
        # ── xem _build_update_api_btn. ─────────────────────────────────────
        api_row = tk.Frame(hdr, bg=BG_COLOR)
        api_row.pack(side="left", padx=(10, 0))
        self._update_api_btn = tk.Button(
            api_row, text="🔑 Update API", font=("Segoe UI", 7, "bold"),
            bd=1, padx=5, pady=0, cursor="hand2",
            command=self._open_update_api_dialog)
        self._update_api_btn.pack(side="left")
        self._refresh_update_api_btn_ui()

        # ── Conversation area (read-only, scrollable) ──────────────────────
        body = tk.Frame(panel, bg=BG_COLOR)
        body.pack(fill="both", expand=True, padx=8, pady=(1, 1))
        sb = ttk.Scrollbar(body)
        sb.pack(side="right", fill="y")
        chat_txt = tk.Text(body, bg=ENTRY_BG, fg=TEXT_COLOR, font=("Segoe UI", 9),
                            wrap="word", bd=0, padx=8, pady=2, state="disabled",
                            yscrollcommand=sb.set)
        chat_txt.pack(fill="both", expand=True)
        sb.config(command=chat_txt.yview)
        # v1.4: the "Bạn"/"AI" text labels were replaced by a person/robot
        # icon (see _append_chat_line). v1.6: the "user" tag now also colors
        # the actual message TEXT blue, not just the icon, so the person's
        # own messages are visually distinct from the AI's answers, which
        # always stay near-black (TEXT_COLOR) regardless of tag.
        chat_txt.tag_config("user", foreground="#0b5fa5", font=("Segoe UI", 9))
        chat_txt.tag_config("user_icon", foreground="#0b5fa5", font=("Segoe UI", 10, "bold"))
        chat_txt.tag_config("assistant_icon", foreground=TEXT_COLOR, font=("Segoe UI", 10, "bold"))
        chat_txt.tag_config("error", foreground="#c0392b", font=("Segoe UI", 9, "bold"))
        chat_txt.tag_config("body", foreground=TEXT_COLOR, font=("Segoe UI", 9))
        # v1.7: right-click "Copy" menu -- selecting text already worked,
        # but right-clicking did nothing, forcing the user to reach for
        # Ctrl+C instead. Text stays read-only (state stays "disabled" for
        # editing); <<Copy>> only reads the current selection so it works
        # fine regardless.
        add_only_copy_menu(chat_txt)
        self._chat_txt = chat_txt

        # v-new (nút "+" Add file): chip "📎 filename ×" hiện ngay TRÊN
        # input_row khi có file đính kèm (hoặc trong lúc đang đọc file) --
        # ẩn hoàn toàn (pack_forget) khi không có gì để không chiếm chỗ.
        self._attach_chip_var = tk.StringVar()
        self._attach_chip_row = tk.Frame(panel, bg=BG_COLOR)

        # ── Input row: "New Chat" icon + input box with the send (↑) ──────
        # ── button embedded INSIDE it, like modern AI chat UIs. Extra ──
        # ── right-hand padding keeps the send button clear of the resize ──
        # ── grip that sits in the window's bottom-right corner. ──────────
        input_row = tk.Frame(panel, bg=BG_COLOR)
        input_row.pack(fill="x", padx=(8, 24), pady=(0, 3))
        self._chat_input_row = input_row

        # v-new (theo yêu cầu: đổi vị trí nút "New Chat" và nút "Add file"
        # cho nhau): "+" Add file giờ nằm ở ngoài cùng bên trái (trước ô
        # nhập, chỗ New Chat từng đứng); New Chat giờ nằm bên trong "pill"
        # cạnh nút Enter/mũi tên gửi (chỗ Add file từng đứng). Chỉ đổi vị
        # trí (parent widget + pack side) -- toàn bộ logic/biến/tooltip giữ
        # nguyên như cũ.
        self._chat_attach_btn = self._draw_attach_icon(input_row, size=18, color=_CHAT_ACCENT)
        self._chat_attach_enabled = True
        self._chat_attach_btn.bind(
            "<Button-1>", lambda e: self._pick_attach_file() if self._chat_attach_enabled else None)
        self._chat_attach_btn.pack(side="left", padx=(0, 8))
        add_tooltip(self._chat_attach_btn, "Add file(<5MB)")

        # entry_box acts as a single bordered "pill" -- same border color and
        # roughly the same height as the program's main Searchbox -- so the
        # Entry and the ↑ Send button visually merge into one input field,
        # instead of two separate widgets sitting side by side.
        entry_box = tk.Frame(input_row, bg=ENTRY_BG,
                              highlightbackground="#8a8a8a", highlightthickness=1)
        entry_box.pack(side="left", fill="x", expand=True)

        # v-fix (Shift+Enter không xuống dòng): Entry -> Text bên dưới,
        # không còn dùng StringVar nữa (self._chat_entry_var đã bị xoá).
        chat_entry = tk.Text(entry_box, font=("Segoe UI", 10), bg=ENTRY_BG, fg=TEXT_COLOR,
                              insertbackground=TEXT_COLOR, bd=0, relief="flat",
                              height=1, wrap="word", undo=True)
        # v-fix (Vấn đề: dòng đầu tiên/con trỏ nằm cao hơn giữa khung, chứ
        # không ở giữa): trước đây fill="both" ép chat_entry (Text) DÃN
        # HẾT chiều cao của entry_box (khung này cao bằng nút Gửi 25px) --
        # nhưng Text (khác Entry) không tự canh giữa nội dung theo chiều
        # dọc, nó luôn vẽ dòng đầu tiên dính SÁT MÉP TRÊN của khung đã bị
        # dãn ra đó. Đổi fill="x" (bỏ "y"): Text giữ đúng chiều cao tự
        # nhiên của 1 dòng, và pack tự canh giữa nó theo chiều dọc trong
        # khung entry_box (anchor mặc định của pack là "center") -- y hệt
        # cách Entry cũ từng canh giữa. Khi gõ nhiều dòng (Shift+Enter),
        # Text lớn dần lên và tự chiếm hết chỗ, không cần "y" nữa.
        chat_entry.pack(side="left", fill="x", expand=True, ipady=1, padx=(8, 2))
        # v-fix (Vấn đề: Shift+Enter = Enter, gửi luôn thay vì xuống dòng):
        # tk.Entry (widget cũ) là ô 1 dòng, KHÔNG THỂ chứa nhiều dòng dù có
        # bind gì đi nữa -- đó là giới hạn của bản thân widget, không phải
        # do thiếu 1 dòng bind. Đổi sang tk.Text (vẫn cao 1 dòng lúc rảnh,
        # tự cao thêm khi gõ nhiều dòng -- xem _autosize_chat_entry) mới
        # thực sự chứa được nhiều dòng. Enter thường (không Shift) gửi tin
        # nhắn (return "break" để chặn Text tự chèn newline mặc định);
        # Shift+Enter thì CHỈ bind một no-op rỗng (không "break") để hành
        # vi mặc định của Text (chèn newline) vẫn chạy -- Tk tự ưu tiên
        # binding "khớp chính xác hơn" (<Shift-Return>) so với binding
        # chung <Return> mỗi khi Shift đang được giữ, nên 2 phím này không
        # còn giẫm lên nhau nữa.
        chat_entry.bind("<Return>", self._on_chat_entry_return)
        chat_entry.bind("<Shift-Return>", lambda e: None)
        chat_entry.bind("<<Modified>>", self._on_chat_entry_modified)
        # v-new (yêu cầu 4 -- dán ảnh vào ô chat): chặn TRƯỚC sự kiện
        # <<Paste>> mặc định của Tk -- nếu clipboard đang giữ 1 ảnh, xử lý
        # đính kèm ảnh và return "break" để Tk KHÔNG paste thêm gì vào ô
        # nhập text nữa (tránh dán nhầm 1 đường dẫn file /base64 rác vào
        # khung chat). Nếu clipboard không có ảnh (trường hợp bình thường
        # -- dán text), trả về None/không "break" để hành vi paste-text
        # mặc định của Tk chạy tiếp y như trước, không ảnh hưởng gì.
        # Bind cả <<Paste>> (phát sinh từ menu chuột phải "Paste", xem
        # setup_context_menu) LẪN <Control-v>/<Control-V> (phím tắt gõ
        # trực tiếp) để không phím tắt nào "lọt lưới".
        def _on_chat_entry_paste(_e=None):
            if self._try_paste_image_from_clipboard():
                return "break"
            return None
        chat_entry.bind("<<Paste>>", _on_chat_entry_paste)
        chat_entry.bind("<Control-v>", _on_chat_entry_paste)
        chat_entry.bind("<Control-V>", _on_chat_entry_paste)
        self._chat_entry_max_lines = 5
        self._chat_entry = chat_entry

        # Overlay placeholder -- "Ask everything..." when idle/empty,
        # switched to "Thinking..." while a request is in flight (see
        # _set_chat_busy). v1.6: the Entry itself is never put into Tk's
        # "disabled" state anymore (that used to grey out the whole pill and
        # leave this placeholder's own white background looking like a
        # mismatched box sitting on top of the grey) -- sending is instead
        # blocked by a plain busy-flag check in _send_online_chat_msg, so the
        # entry's look never changes while a request is running.
        self._chat_placeholder = tk.Label(chat_entry, text=CHAT_ASK_PH, fg=PLACE_COLOR,
                                           bg=ENTRY_BG, font=("Segoe UI", 9, "italic"))
        # v-fix (chữ "Ask everything.../Thinking..." bị lệch cao hơn giữa
        # khung): y=3 cố định từng đúng khi chat_entry còn là Entry (Tk tự
        # canh giữa Entry, placeholder overlay lên trên chỉ cần lệch nhẹ).
        # Giờ chat_entry là Text -- chiều cao thật của nó thay đổi (1 dòng
        # lúc rảnh, cao hơn khi Shift+Enter xuống dòng), nên toạ độ y cố
        # định không còn đúng giữa nữa. Dùng rely=0.5 + anchor="w": Tk tự
        # tính lại vị trí canh giữa THEO CHIỀU DỌC mỗi lần, bất kể chiều
        # cao thật của chat_entry lúc đó là bao nhiêu.
        self._chat_placeholder.place(x=6, rely=0.5, anchor="w")
        self._chat_placeholder.bind("<Button-1>", lambda e: self._chat_entry.focus_set())

        # v-fix (nút gửi vuông bo góc + mũi tên giống Claude -- theo yêu
        # cầu): now drawn as vector shapes (see _draw_send_button) instead
        # of a plain tk.Button -- gives real rounded corners and a
        # consistent arrow glyph on every machine, not whatever "↑" looks
        # like in the local system font.
        self._chat_send_btn = self._draw_send_button(entry_box, size=25)
        self._chat_send_btn.pack(side="right", padx=(0, 3), pady=2)
        add_tooltip(self._chat_send_btn, "Send")

        # v-fix (yêu cầu: bỏ hẳn nút 🌐 riêng cạnh Send -- đã có link
        # "🌐 Bấm vào đây..." ngay bên trong nội dung phân tích cục bộ làm
        # đúng việc này rồi, giữ cả 2 là thừa). self._chat_web_search_used
        # vẫn giữ nguyên (đánh dấu đã tra Internet cho câu hỏi hiện tại) --
        # chỉ không còn 1 WIDGET riêng để tự grey-out/enable nữa; trạng
        # thái enable/grey giờ nằm ngay trong nội dung link (xem
        # _online_chat_worker/_insert_chat_text_with_links). Reset cờ vẫn
        # xảy ra ở _send_online_chat_msg/_send_web_search_chat_msg/
        # _new_online_chat như cũ.
        self._chat_web_search_used = False
        self._chat_ai_search_link_used = False  # v-new: tương tự, cho label "🤖 AI Search offline"
        self._chat_ai_search_grey_tags = []  # v-new: registry các tag label xám đang chờ kích hoạt, xem _activate_ai_search_links_for_query

        # v-fix (yêu cầu 4: bỏ nút "New Chat" khỏi giao diện -- user báo hầu
        # như chưa từng bấm tới): KHÔNG tạo/pack icon này nữa. Hàm
        # self._new_online_chat() (logic thật sự) vẫn giữ nguyên và vẫn
        # được các chỗ khác trong code tự gọi khi cần (vd: chuyển sang 1
        # search keyword khác chưa có History -- xem _auto_ai_chat_after_search)
        # -- chỉ riêng NÚT bấm thủ công của user là bị gỡ. self._chat_new_chat_btn
        # cố tình không được set (ở None) -- mọi nơi khác đọc nó qua
        # getattr(self, "_chat_new_chat_btn", None) nên không cần sửa gì
        # thêm, chỉ đơn giản không có nút nào để enable/disable nữa.

        self._chat_attached_file = None
        self._update_attach_chip_ui(show="none")

        # v-new (yêu cầu 3: drag & drop file vào chatbox, giống Add file
        # <5MB): hook cả khung hội thoại (chat_txt) LẪN ô nhập (entry_box)
        # làm "drop target" -- xem _setup_chat_dragdrop().
        self._setup_chat_dragdrop(chat_txt, entry_box)

        self._online_chat_panel = panel
        return panel

    def _setup_chat_dragdrop(self, *drop_widgets):
        """v-fix (Vấn đề: Fatal Python error: PyEval_RestoreThread -- GIL
        released -- crash app hoàn toàn khi thả file vào chatbox trên
        Python 3.12): trước đây dùng thư viện `windnd`, thư viện này tự
        subclass thẳng WndProc của cửa sổ Win32 bằng ctypes để tự bắt
        message WM_DROPFILES, HOÀN TOÀN bỏ qua vòng lặp sự kiện Tcl/Tk.
        Khi Windows gọi callback đó, nó gọi thẳng vào hàm Python mà không
        đảm bảo GIL đã được acquire đúng cách (không gọi
        PyGILState_Ensure) -- trên Python 3.12 điều này crash fatal ở tầng
        interpreter, không thể try/except bắt được.

        Chuyển sang `tkinterdnd2` (pip install tkinterdnd2): thư viện này
        tích hợp đúng qua package Tcl `tkdnd`, xử lý ngay trong event loop
        bình thường của Tk (không có kiểu lỗi GIL này). Dùng
        `TkinterDnD._require(self.root)` để nạp package tkdnd vào đúng
        root Tk() hiện có -- KHÔNG cần đổi root thành TkinterDnD.Tk()
        (tránh phải retrofit lại toàn bộ app).

        Fail SOFT khi thiếu thư viện hoặc không phải Windows: in 1 dòng
        log rồi return, KHÔNG raise/crash -- app vẫn chạy bình thường,
        chỉ đơn giản là chưa kéo-thả được (vẫn còn nút "+" và paste ảnh
        Ctrl+V làm 2 đường thay thế)."""
        if sys.platform != "win32":
            print("[Chat] Drag & drop: chỉ hỗ trợ Windows (tkinterdnd2) -- bỏ qua trên nền tảng này.")
            return
        try:
            from tkinterdnd2 import TkinterDnD, DND_FILES
        except Exception:
            print("[Chat] Drag & drop: chưa cài thư viện 'tkinterdnd2' -- chạy "
                  "`pip install tkinterdnd2` rồi khởi động lại app để bật tính "
                  "năng này. Không ảnh hưởng gì tới các chức năng khác.")
            return
        try:
            TkinterDnD._require(self.root)
        except Exception as _e:
            print(f"[Chat] Drag & drop: không nạp được package tkdnd vào root: {_e}")
            return
        for w in drop_widgets:
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_chat_files_dropped)
            except Exception as _e:
                print(f"[Chat] Drag & drop: không hook được vào widget {w!r}: {_e}")

    def _on_chat_files_dropped(self, event):
        """v-fix: callback của tkinterdnd2 khi user thả file vào
        chat_txt/entry_box -- event.data là 1 chuỗi Tcl-list (mỗi đường
        dẫn cách nhau bởi khoảng trắng, tự bọc {} nếu path có khoảng
        trắng/ký tự đặc biệt bên trong). Dùng self.root.tk.splitlist() để
        Tcl tự parse đúng thay vì tự split(" ") thủ công (sẽ vỡ path có
        khoảng trắng). tkinterdnd2 trả về str (đã là Unicode), không cần
        decode bytes như windnd trước đây.

        Chỉ đính kèm ĐÚNG 1 file (giống nút "+" -- self._chat_attached_file
        hiện chỉ chứa được 1 file tại 1 thời điểm): nếu thả nhiều file
        cùng lúc, lấy file ĐẦU TIÊN hợp lệ trong danh sách, bỏ qua phần
        còn lại."""
        try:
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data] if getattr(event, "data", None) else []
        paths = [p for p in paths if p]
        if not paths:
            return
        # v-note: callback của tkinterdnd2 đã chạy trên main thread của Tk
        # (đúng qua event loop Tcl/Tk bình thường) nên về lý thuyết không
        # bắt buộc phải bọc qua root.after -- nhưng vẫn giữ lại cho nhất
        # quán với mọi callback khác trong file này (và an toàn hơn nếu
        # sau này binding đổi cách gọi).
        self.root.after(0, lambda: self._attach_local_file(paths[0], _dlg_title="Drag & drop"))

    def _refresh_update_api_btn_ui(self):
        """v-fix (theo yêu cầu): nút LUÔN ở trạng thái active/bấm được (state
        luôn "normal") để user có thể mở dialog đổi key bất cứ lúc nào, kể
        cả khi API đang hoạt động bình thường -- trước đây khi API OK nút
        được tô màu xám kiểu "disabled" (dù thật ra vẫn clickable) khiến
        nhiều người tưởng nhầm là nút đã bị khoá. Giờ cả 2 trạng thái đều
        dùng màu "active" bình thường, chỉ đổi text/màu để phân biệt còn
        thiếu key (đỏ/cam) hay đã có key hợp lệ (xanh). Safe to call even
        before the chat panel exists."""
        btn = getattr(self, "_update_api_btn", None)
        if not btn:
            return
        try:
            if not btn.winfo_exists():
                return
        except Exception:
            return
        needs_attention = self._api_needs_attention()
        if needs_attention:
            btn.config(text="🔑 Update API", bg="#5a1a1a", fg="#ff9090",
                       relief="raised", state="normal")
        else:
            btn.config(text="🔑 API OK", bg="#1a4a1a", fg="#90ee90",
                       relief="raised", state="normal")

    def _api_needs_attention(self):
        """True if neither online model currently has a usable key, or the
        last call to either flagged a quota/rate-limit/too-large error --
        see _mark_api_needs_attention (set from _online_chat_worker)."""
        if getattr(self, "_api_attention_flag", False):
            return True
        return not any(_online_ai_available(m) for m in MODEL_FALLBACK_ORDER)

    def _mark_api_needs_attention(self, flag=True):
        """Called from _online_chat_worker whenever a call fails due to a
        missing key, quota/rate-limit (429), or a 'request too large'
        (413/TPM) error -- lights up the Update API button so the user
        notices without having to read the chat error text. Cleared again
        the next time a call succeeds, or when the user saves a new key."""
        self._api_attention_flag = flag
        self.root.after(0, self._refresh_update_api_btn_ui)

    def _mark_model_rate_limited(self, model_choice, limited=True):
        """v-new: remember that `model_choice` ("gemini"/"gptoss") just hit
        a rate-limit/quota error (limited=True) or just answered
        successfully again (limited=False). Every place that picks which
        model to call (start of _online_chat_worker, the Online/web_search
        forced-Gemini block, _translate_current_chat) checks this dict via
        _model_usable() so, once a model is known exhausted, the app keeps
        using the OTHER model for every subsequent message instead of
        re-trying -- and failing on -- the exhausted one each turn. Wiped
        automatically the next time a call to that model actually
        succeeds (quota resets, e.g. next day for Groq's TPD limit)."""
        d = getattr(self, "_chat_model_rate_limited", None)
        if d is None:
            d = self._chat_model_rate_limited = {"gemini": False, "gptoss": False, "qwen": False}
        d[model_choice] = limited

    def _model_usable(self, model_choice):
        """True if `model_choice` has a key configured AND hasn't just been
        marked rate-limited this session (see _mark_model_rate_limited)."""
        if not _online_ai_available(model_choice):
            return False
        return not getattr(self, "_chat_model_rate_limited", {}).get(model_choice, False)

    def _next_alt_model(self, model_choice):
        """v-new (yêu cầu: thêm model thứ 3 Qwen 3.6 27B -- cần 1 chỗ chọn
        "model kế tiếp" chung thay vì khắp nơi tự viết
        `"gptoss" if model_choice == "gemini" else "gemini"`, vốn chỉ đúng
        khi có ĐÚNG 2 model): trả về model KHẢ DỤNG kế tiếp trong
        MODEL_FALLBACK_ORDER (Gemini Flash -> GPT-OSS 120B -> Qwen 3.6
        27B), xoay vòng bắt đầu từ ngay SAU model_choice hiện tại -- bỏ qua
        model_choice và bất kỳ model nào khác đang bị đánh dấu hết quota/
        thiếu key (xem _model_usable). Trả None nếu không còn model nào
        khác dùng được."""
        try:
            start = MODEL_FALLBACK_ORDER.index(model_choice)
        except ValueError:
            start = -1
        n = len(MODEL_FALLBACK_ORDER)
        for offset in range(1, n):
            cand = MODEL_FALLBACK_ORDER[(start + offset) % n]
            if self._model_usable(cand):
                return cand
        return None

    def _sync_chat_model_picker(self):
        """v-new: keep the header dropdown (self._chat_model_var) showing
        whatever self._chat_active_model ACTUALLY is. Without this, the
        dropdown only ever changed when the user manually picked a model
        (see _on_model_pick in _build_online_chat_panel) -- so once the
        rate-limit auto-fallback silently moved _chat_active_model from
        "gemini" to "gptoss" behind the scenes, the dropdown kept showing
        "Gemini Flash" forever after, even though every reply (and the
        "no real web search" footer note) was coming from GPT-OSS. Safe to
        call from a background worker thread -- routed through root.after.
        No-op if the chat panel hasn't been built yet (no _chat_model_var)."""
        var = getattr(self, "_chat_model_var", None)
        if var is None:
            return
        model = getattr(self, "_chat_active_model", DEFAULT_ONLINE_MODEL)
        label = MODEL_SHORT_LABELS.get(model, model)
        self.root.after(0, lambda: var.set(label))

    def _get_key_icon_photo(self):
        """v-new (yêu cầu: icon cửa sổ Update API Key đang là hình cái lông
        vũ mặc định của Tk -- đổi thành hình chìa khoá 🔑 cho đúng ngữ
        cảnh): vẽ 1 icon chìa khoá đơn giản bằng PIL (đã có sẵn trong app,
        dùng cho icon file/thumbnail khác) thay vì cần kèm theo 1 file
        .ico riêng -- tự vẽ 1 lần, cache lại trong self._key_icon_photo để
        dùng lại cho lần mở dialog sau, không vẽ lại mỗi lần."""
        cached = getattr(self, "_key_icon_photo", None)
        if cached is not None:
            return cached
        try:
            from PIL import Image, ImageDraw, ImageTk
            size = 32
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            gold = (255, 193, 7, 255)
            outline = (150, 105, 0, 255)
            # Vòng đầu chìa khoá (hình tròn rỗng ở góc trên-trái)
            d.ellipse([2, 2, 16, 16], fill=gold, outline=outline, width=2)
            d.ellipse([6, 6, 12, 12], fill=(0, 0, 0, 0), outline=outline, width=2)
            # Thân chìa khoá (đường chéo từ vòng đầu xuống góc dưới-phải)
            d.line([13, 13, 27, 27], fill=gold, width=5)
            d.line([13, 13, 27, 27], fill=outline, width=1)
            # 2 răng chìa khoá ở đầu thân
            d.line([22, 22, 26, 18], fill=gold, width=4)
            d.line([25, 25, 29, 21], fill=gold, width=4)
            photo = ImageTk.PhotoImage(img)
            self._key_icon_photo = photo
            return photo
        except Exception as e:
            print(f"[UI] failed to draw key icon: {e}")
            return None

    def _open_update_api_dialog(self):
        """Dialog nhập/sửa GEMINI_API_KEY và GROQ_API_KEY -- lưu vào
        configure.ini (section [APIKeys], xem CONFIG_FILE/
        _save_numbered_keys_to_config) để user tự mở file này sửa tay
        được nếu muốn, và để không phải nhập lại key mỗi lần mở app.

        v-new (yêu cầu: nhiều key/provider + tự động xoay khi hết quota):
        mỗi provider (Gemini/Groq) giờ có THỂ NHIỀU hơn 1 ô nhập key --
        bấm "+" để thêm 1 ô key nữa. Khi key #1 bị 429/hết quota giữa
        chừng, app tự động chuyển sang key #2, #3... của ĐÚNG provider đó
        (xem _rotate_api_key, gọi từ _call_online_ai/_call_online_ai_chat)
        mà không cần user tự mở dialog này đổi key thủ công."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Update API Key")
        dlg.configure(bg=BG_COLOR)
        dlg.resizable(False, False)
        dlg.transient(self.root)  # v-fix (Vấn đề 4): tied to main window's
        # stacking order instead of "-topmost" (which pinned this dialog
        # above EVERY window, including ones the user clicked to AFTER
        # opening it) -- a transient dialog only stays above its owner.
        _key_icon = self._get_key_icon_photo()
        if _key_icon is not None:
            try:
                dlg.iconphoto(False, _key_icon)
            except Exception:
                pass

        status_lbl = tk.Label(dlg, text="", bg=BG_COLOR, fg="#ff9090", font=("Segoe UI", 8))

        def _reflow_dialog():
            """v-fix (yêu cầu: bấm ➕ thì cửa sổ rộng/cao ra nhưng nút
            Save/Delete API Key/Cancel bị khuất mất, không thấy đâu): gốc
            vấn đề là dlg.geometry("WxH+x+y") set 1 kích thước CỐ ĐỊNH
            bằng pixel ngay lúc mở dialog (chỉ tính đúng 1 dòng key/
            provider lúc đó) rồi dlg.resizable(False, False) khoá luôn --
            nên khi ➕ thêm dòng entry mới, layout grid/pack bên trong VẪN
            đẩy các widget xuống thấp hơn như bình thường, nhưng cửa sổ
            THẬT SỰ (kích thước hệ điều hành) không hề giãn theo, nên phần
            dưới (nút Save...) bị cắt mất khỏi vùng nhìn thấy được, không
            phải do ẩn hay lỗi logic. Gọi lại hàm này (recompute
            winfo_reqheight() rồi geometry() lại) mỗi khi thêm/xoá 1 dòng
            key để cửa sổ luôn tự giãn vừa đúng nội dung -- resizable vẫn
            để False vì đây là NGƯỜI GỌI (code) tự resize theo nội dung,
            không phải user kéo tay."""
            dlg.update_idletasks()
            w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            x = rx + max(0, (rw - w) // 2)
            y = ry + max(0, (rh - h) // 2)
            dlg.geometry(f"{w}x{h}+{x}+{y}")

        def _build_provider_section(parent, row_idx, title, link_url, existing_keys):
            """Xây 1 khối (label + link 'lấy key ở đâu' + danh sách ô
            Entry + nút '+') cho 1 provider. Trả về hàm collect_keys()
            đọc lại toàn bộ key user đã nhập trong khối này (bỏ dòng
            rỗng, giữ đúng thứ tự) khi bấm Save."""
            lbl_row = tk.Frame(parent, bg=BG_COLOR)
            lbl_row.grid(row=row_idx, column=0, sticky="w", padx=10, pady=(10, 2))
            tk.Label(lbl_row, text=f"{title} (", bg=BG_COLOR, fg=TEXT_COLOR,
                     font=("Segoe UI", 9)).pack(side="left")
            link = tk.Label(lbl_row, text="...", bg=BG_COLOR, fg="#6ea8fe",
                             font=("Segoe UI", 9, "underline"), cursor="hand2")
            link.pack(side="left")
            link.bind("<Button-1>", lambda e: webbrowser.open(link_url))
            tk.Label(lbl_row, text=")", bg=BG_COLOR, fg=TEXT_COLOR,
                     font=("Segoe UI", 9)).pack(side="left")
            plus_btn = tk.Label(lbl_row, text=" ➕ Add more key", bg=BG_COLOR, fg="#4caf50",
                                 font=("Segoe UI", 9, "bold"), cursor="hand2")
            plus_btn.pack(side="left", padx=(6, 0))

            rows_frame = tk.Frame(parent, bg=BG_COLOR)
            rows_frame.grid(row=row_idx + 1, column=0, sticky="w", padx=10, pady=(0, 4))
            entry_vars = []  # list of (row_frame, tk.StringVar, entry_widget)

            def _add_row(initial_value=""):
                r = tk.Frame(rows_frame, bg=BG_COLOR)
                r.pack(fill="x", pady=1)
                var = tk.StringVar(value=initial_value)
                ent = tk.Entry(r, textvariable=var, width=44,
                                bg=ENTRY_BG, fg=TEXT_COLOR, insertbackground=TEXT_COLOR)
                ent.pack(side="left")
                entry_vars.append((r, var, ent))

                def _remove():
                    if len(entry_vars) <= 1:
                        var.set("")  # last row -- just clear instead of removing entirely
                        return
                    for i, (rr, vv, ee) in enumerate(entry_vars):
                        if vv is var:
                            entry_vars.pop(i)
                            break
                    r.destroy()
                    _reflow_dialog()  # v-fix: cửa sổ nhỏ lại đúng bằng số dòng còn lại

                rm_btn = tk.Label(r, text=" ✕", bg=BG_COLOR, fg="#b0b3b8",
                                   font=("Segoe UI", 9), cursor="hand2")
                rm_btn.pack(side="left", padx=(4, 0))
                rm_btn.bind("<Button-1>", lambda e: _remove())
                return ent

            for k in (existing_keys or [""]):
                _add_row(k)

            def _on_plus_click(e):
                _add_row("")
                _reflow_dialog()  # v-fix: cửa sổ giãn cao thêm đúng 1 dòng, Save/Delete/Cancel luôn ở trong vùng nhìn thấy

            plus_btn.bind("<Button-1>", _on_plus_click)

            def _collect_keys():
                return [v.get().strip() for (_r, v, _e) in entry_vars if v.get().strip()]

            return _collect_keys, (lambda: entry_vars[0][2] if entry_vars else None)

        gem_collect, gem_first_entry = _build_provider_section(
            dlg, 0, "GEMINI_API_KEY", "https://aistudio.google.com/app/api-keys", _GEMINI_API_KEYS)
        groq_collect, _groq_first_entry = _build_provider_section(
            dlg, 2, "GROQ_API_KEY", "https://console.groq.com/keys", _GROQ_API_KEYS)

        status_lbl.grid(row=4, column=0, padx=10, sticky="w")

        def _apply_keys(gemini_keys, groq_keys):
            """Ghi danh sách key mới vào configure.ini + cập nhật toàn bộ
            state trong-bộ-nhớ (danh sách, index đang dùng, key active,
            client SDK đã cache) NGAY LẬP TỨC -- không cần khởi động lại
            app."""
            global _GEMINI_API_KEYS, _GROQ_API_KEYS, _gemini_key_idx, _groq_key_idx
            global ONLINE_GEMINI_API_KEY, ONLINE_GROQ_API_KEY
            _save_numbered_keys_to_config("gemini_key_", gemini_keys)
            _save_numbered_keys_to_config("groq_key_", groq_keys)
            _GEMINI_API_KEYS = list(gemini_keys)
            _GROQ_API_KEYS = list(groq_keys)
            _gemini_key_idx = 0
            _groq_key_idx = 0
            ONLINE_GEMINI_API_KEY = _GEMINI_API_KEYS[0] if _GEMINI_API_KEYS else ""
            ONLINE_GROQ_API_KEY = _GROQ_API_KEYS[0] if _GROQ_API_KEYS else ""
            _reset_online_ai_clients()

        def _save():
            gemini_keys = gem_collect()
            groq_keys = groq_collect()
            if not gemini_keys and not groq_keys:
                status_lbl.config(text="Nhập ít nhất 1 key.")
                return
            _apply_keys(gemini_keys, groq_keys)
            # v-fix (Vấn đề: đổi/thêm API key mới nhưng vẫn switch ngay
            # sang GPT-OSS): xoá cả 3 cờ rate-limited để key vừa lưu
            # được thử lại thật sự trong phiên đang chạy (xem comment gốc
            # ở bản trước đây).
            for _mk in ("gemini", "gptoss", "qwen"):
                self._mark_model_rate_limited(_mk, False)
            self._api_attention_flag = False
            self._refresh_update_api_btn_ui()
            status_lbl.config(text="Đã lưu. Đang áp dụng...", fg="#90ee90")
            self.root.after(400, dlg.destroy)
            # v-fix (Vấn đề 2): auto-run AI analysis for whatever query is
            # currently in the Searchbox right after the key is saved,
            # instead of leaving the chat panel stuck on the "no key"
            # warning until the user manually clicks AI Search again.
            query = self.entry_var.get().strip()
            if query:
                self.root.after(450, lambda: self._auto_ai_chat_after_search(query))

        def _delete():
            _apply_keys([], [])
            _delete_api_keys_from_db(True, True)  # also clear the legacy search_data.db copy, if any
            dlg.destroy()
            self._open_update_api_dialog()  # reopen with everything cleared back to a single empty row each

        btn_row = tk.Frame(dlg, bg=BG_COLOR)
        btn_row.grid(row=5, column=0, pady=10)
        tk.Button(btn_row, text="Save", command=_save, bg="#2a5a2a", fg="white",
                  padx=14).pack(side="left", padx=4)
        tk.Button(btn_row, text="Delete API Key", command=_delete, bg="#5a1a1a",
                  fg="#ff9090", padx=10).pack(side="left", padx=4)
        tk.Button(btn_row, text="Cancel", command=dlg.destroy, padx=14).pack(side="left", padx=4)
        _fe = gem_first_entry()
        if _fe:
            _fe.focus_set()

        # v-fix (Vấn đề 4): center the dialog over the MAIN search window
        # (self.root) instead of leaving it wherever Tk's default placement
        # puts it (top-left-ish, disconnected from the app). Dùng chung
        # _reflow_dialog() với lúc bấm ➕/✕ ở trên, thay vì tính riêng 1
        # lần ở đây -- đảm bảo luôn cùng 1 công thức.
        _reflow_dialog()


    def _chat_model_status_text(self):
        model_key = getattr(self, "_chat_active_model", DEFAULT_ONLINE_MODEL)
        label = ONLINE_AI_MODELS.get(model_key, {}).get("label", model_key)
        return f"({label})"

    def _show_online_chat_panel(self):
        """Mark the online chat panel as active (offline AI Search is on)
        and show it -- but only while the user is actually looking at the
        "File Content" tab (see _sync_online_chat_panel_visibility). Safe
        to call repeatedly / before the widget exists."""
        try:
            if not hasattr(self, "results_frame") or not self.results_frame or \
                    not self.results_frame.winfo_exists():
                return
            self._build_online_chat_panel(self.results_frame)
            self._chat_panel_wanted = True
            self._sync_online_chat_panel_visibility()
        except Exception as _e:
            print(f"[Online Chat] show panel error: {_e}")

    def _hide_online_chat_panel(self):
        self._chat_panel_wanted = False
        try:
            if self._online_chat_panel is not None and self._online_chat_panel.winfo_exists():
                self._online_chat_panel.grid_remove()
        except Exception:
            pass

    def _sync_online_chat_panel_visibility(self):
        """v-fix (AI Chat hiện cả ở File Name / Folder Name -- chỉ nên
        hiện ở File Content): panel chỉ HIỆN khi tab đang xem đúng là
        "File Content"; mọi tab khác (File Name/Folder Name/Help/...)
        đều ẩn panel đi. Vì tab mặc định khi mở kết quả là "File
        Content" (xem show_results/_smart_search_realtime), panel vẫn
        tự hiện ngay khi có kết quả mà không cần user bấm tab. Called
        on every tab switch (see the <<NotebookTabChanged>> binding in
        show_results)."""
        try:
            if not getattr(self, "_chat_panel_wanted", False):
                return
            if self._online_chat_panel is None or not self._online_chat_panel.winfo_exists():
                return
            cur_tab_text = self.nb.tab(self.nb.select(), "text").strip()
            if cur_tab_text == "File Content":
                self._online_chat_panel.grid(row=1, column=0, sticky="nsew")
                grip = getattr(self, "_resize_grip", None)
                if grip is not None and grip.winfo_exists():
                    grip.lift()
            else:
                self._online_chat_panel.grid_remove()
        except Exception:
            pass

    def _append_chat_line(self, role, text, is_error=False):
        """role: 'user' | 'assistant'. Append one line to the chat Text
        widget, auto-scrolling to the bottom. The old "Bạn:"/"AI:" text
        labels are now a person/robot icon instead (v1.4). v1.6: the
        user's own message text is now colored blue too (not just the
        icon), so it's visually distinct from the AI's (near-black)
        answers. v1.9: for assistant answers, each file path inside the
        "Nguồn:" list (see _append_citation_filenames) is rendered as a
        clickable link that opens the file directly."""
        try:
            txt = self._chat_txt
            txt.config(state="normal")
            if txt.index("end-1c") != "1.0":
                txt.insert("end", "\n\n")
            who_icon = "👤" if role == "user" else "🤖"
            icon_tag = "error" if is_error else (f"{role}_icon" if role == "user" else "assistant_icon")
            body_tag = "error" if is_error else ("user" if role == "user" else "body")
            txt.insert("end", f"{who_icon} ", icon_tag)
            if role == "assistant" and not is_error:
                self._insert_chat_text_with_links(txt, text, body_tag)
            else:
                txt.insert("end", text, body_tag)
            txt.config(state="disabled")
            txt.see("end")
        except Exception as _e:
            print(f"[Online Chat] append line error: {_e}")

    def _start_chat_stream(self):
        """v-new (streaming): insert the 🤖 icon and drop a Text mark right
        after it -- _append_chat_stream_chunk keeps inserting raw model
        text at that mark as it arrives (visual "typing" effect, like
        ChatGPT/web AI UIs), and once the full answer is ready,
        _finish_chat_stream wipes everything from the mark onward and
        re-inserts the FULLY POST-PROCESSED final text (citations
        renumbered, "Nguồn:" list appended, internet sources appended,
        etc -- see _online_chat_worker) via _insert_chat_text_with_links,
        exactly as a non-streamed reply would look. Streaming only ever
        shows the raw, not-yet-processed text -- it exists purely for
        perceived speed, not as the final rendering."""
        try:
            txt = self._chat_txt
            txt.config(state="normal")
            if txt.index("end-1c") != "1.0":
                txt.insert("end", "\n\n")
            txt.insert("end", "🤖 ", "assistant_icon")
            txt.mark_set("chat_stream_start", "end-1c")
            txt.mark_gravity("chat_stream_start", "left")
            txt.config(state="disabled")
            txt.see("end")
        except Exception as _e:
            print(f"[Online Chat] start stream error: {_e}")

    def _append_chat_stream_chunk(self, chunk_text):
        """v-new (streaming): append one raw text delta as it arrives from
        the model. Safe no-op if _start_chat_stream was never called this
        turn (e.g. called out of order) -- falls through silently rather
        than raising into the API call thread's after() callback."""
        if not chunk_text:
            return
        try:
            txt = self._chat_txt
            if "chat_stream_start" not in txt.mark_names():
                return
            txt.config(state="normal")
            txt.insert("end", chunk_text, "body")
            txt.config(state="disabled")
            txt.see("end")
        except Exception as _e:
            print(f"[Online Chat] stream chunk error: {_e}")

    def _reset_chat_stream(self):
        """v-new (streaming): discard whatever raw text has been streamed
        into the current bubble so far, back to right after the 🤖 icon --
        used when an in-flight attempt fails partway through a stream
        (rare -- e.g. connection drop mid-answer) and _call_online_ai_chat
        is about to retry with a different model/config, so the user never
        sees 2 partial, unrelated answers glued together end to end."""
        try:
            txt = self._chat_txt
            if "chat_stream_start" not in txt.mark_names():
                return
            txt.config(state="normal")
            txt.delete("chat_stream_start", "end")
            txt.config(state="disabled")
        except Exception as _e:
            print(f"[Online Chat] reset stream error: {_e}")

    def _finish_chat_stream(self, final_text, is_error=False):
        """v-new (streaming): replace the raw streamed text with the fully
        post-processed final answer. If no stream was ever started this
        turn (e.g. streaming wasn't used, or _start_chat_stream silently
        failed), falls back to the normal non-streaming _append_chat_line
        so the answer is never lost."""
        try:
            txt = self._chat_txt
            txt.config(state="normal")
            if "chat_stream_start" in txt.mark_names():
                txt.delete("chat_stream_start", "end")
                if is_error:
                    txt.insert("end", final_text, "error")
                else:
                    self._insert_chat_text_with_links(txt, final_text, "body")
                txt.config(state="disabled")
                txt.see("end")
            else:
                txt.config(state="disabled")
                self._append_chat_line("assistant", final_text, is_error=is_error)
        except Exception as _e:
            print(f"[Online Chat] finish stream error: {_e}")

    def _activate_ai_search_link_tag(self, txt, tag_name):
        """v-new: style + gắn click cho 1 tag label "🤖 AI Search offline"
        thành trạng thái BẤM ĐƯỢC (xanh, gạch chân, gọi _send_ai_search_chat_msg).
        Dùng chung cho cả 2 trường hợp: (a) label được render active NGAY
        từ đầu (🤖 AI Search đã chạy xong trước khi câu trả lời này được
        sinh ra), và (b) 1 label ĐANG XÁM được kích hoạt LẠI về sau, ngay
        khi 🤖 AI Search vừa chạy xong (xem _activate_ai_search_links_for_query)."""
        try:
            txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)

            def _on_ai_search_link_click(e, _tag=tag_name, _txt=txt):
                self._send_ai_search_chat_msg()
                try:
                    _txt.tag_config(_tag, foreground="#8a8d91", underline=False)
                    _txt.tag_unbind(_tag, "<Button-1>")
                except Exception:
                    pass

            txt.tag_bind(tag_name, "<Button-1>", _on_ai_search_link_click)
            txt.tag_bind(tag_name, "<Enter>", lambda e: txt.config(cursor="hand2"))
            txt.tag_bind(tag_name, "<Leave>", lambda e: txt.config(cursor=""))
        except Exception:
            pass

    def _activate_ai_search_links_for_query(self, query):
        """v-new (yêu cầu: bấm 🤖 AI Search (nút ngoài) xong thì label
        "🤖 AI Search offline" ĐANG XÁM trong AI Chat phải chuyển bấm được
        NGAY, không cần user hỏi thêm câu nào mới): gọi ngay sau khi 🤖 AI
        Search (self._ai_search_and_update) chạy xong cho `query` -- tìm
        lại mọi label xám đã đăng ký (self._chat_ai_search_grey_tags,
        xem _insert_chat_text_with_links) thuộc ĐÚNG `query` này, chuyển
        chúng sang trạng thái bấm được, rồi bỏ khỏi danh sách chờ (đã kích
        hoạt xong, không cần theo dõi nữa)."""
        pending = getattr(self, "_chat_ai_search_grey_tags", None)
        if not pending:
            return
        remaining = []
        for entry in pending:
            txt = entry.get("txt")
            tag = entry.get("tag")
            try:
                widget_alive = bool(txt) and txt.winfo_exists()
            except Exception:
                widget_alive = False
            if not widget_alive:
                continue  # widget gone (chat panel closed/rebuilt) -- drop silently
            if entry.get("query") != query:
                remaining.append(entry)  # not this query yet -- keep waiting
                continue
            self._activate_ai_search_link_tag(txt, tag)
        self._chat_ai_search_grey_tags = remaining

    def _insert_chat_text_with_links(self, txt, text, body_tag):
        """Insert an AI answer, turning each "[N] name  —  path" line of
        the trailing "Nguồn:" block into a clickable link (underlined,
        opens the file on click) -- everything else is inserted as plain
        text with body_tag.

        v-new (theo yêu cầu: link internet trong mục "Thông tin bổ sung
        (internet)" phải bấm mở được luôn): bất kỳ dòng nào KHÔNG khớp
        pattern "[N] name — path" ở trên nhưng có chứa 1 URL http(s)://
        (ví dụ dòng nguồn internet do _append_web_sources thêm vào) cũng
        được gạch dưới + click-được, mở bằng trình duyệt mặc định qua
        webbrowser.open() thay vì os.startfile()."""
        citation_line_re = re.compile(r"^(\[\d+\]\s+.+?)\s+—\s+(.+)$")
        url_re = re.compile(r"https?://[^\s<>\"')]+")
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if i > 0:
                txt.insert("end", "\n")
            # v-new (yêu cầu: dòng "Click here for more information from
            # the Internet" ở cuối phần phân tích cục bộ phải bấm được,
            # gọi đúng hành động như nút 🌐): nhận diện qua
            # _CHAT_INTERNET_LINK_MARKER (2 ký tự vô hình bọc quanh dòng
            # -- xem nơi khai báo) thay vì so khớp câu chữ, để không phụ
            # thuộc ngôn ngữ hiển thị. Marker bị bóc ra, KHÔNG hiển thị.
            if line.startswith(_CHAT_INTERNET_LINK_MARKER):
                link_label = line[len(_CHAT_INTERNET_LINK_MARKER):]
                self._chat_link_seq = getattr(self, "_chat_link_seq", 0) + 1
                tag_name = f"chatlink_{self._chat_link_seq}"
                txt.insert("end", link_label, (body_tag, tag_name))
                txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)

                def _on_internet_link_click(e, _tag=tag_name, _txt=txt):
                    # v-new: nếu 🌐 đã được dùng RỒI (bấm nút 🌐 riêng, hoặc
                    # bấm chính link này ở 1 tin nhắn khác) thì
                    # _send_web_search_chat_msg() tự no-op -- nhưng vẫn cần
                    # greyout NGAY link này để không trông như còn bấm được.
                    self._send_web_search_chat_msg()
                    try:
                        _txt.tag_config(_tag, foreground="#8a8d91", underline=False)
                        _txt.tag_unbind(_tag, "<Button-1>")
                    except Exception:
                        pass

                txt.tag_bind(tag_name, "<Button-1>", _on_internet_link_click)
                txt.tag_bind(tag_name, "<Enter>", lambda e: txt.config(cursor="hand2"))
                txt.tag_bind(tag_name, "<Leave>", lambda e: txt.config(cursor=""))
                continue
            # v-new (label "🤖 ... AI Search offline"): giống hệt cơ chế
            # link Internet ở trên, CỘNG THÊM 1 ký tự trạng thái ngay sau
            # marker ('1'=bấm được/xanh, '0'=chưa đủ điều kiện/xám, không
            # gắn click) -- xem _online_chat_worker, nơi quyết định ký tự
            # này lúc sinh câu trả lời.
            if line.startswith(_CHAT_AISEARCH_LINK_MARKER):
                _rest = line[len(_CHAT_AISEARCH_LINK_MARKER):]
                _enabled = _rest[:1] == "1"
                link_label = _rest[1:]
                if _enabled:
                    self._chat_link_seq = getattr(self, "_chat_link_seq", 0) + 1
                    tag_name = f"chatlink_{self._chat_link_seq}"
                    txt.insert("end", link_label, (body_tag, tag_name))
                    self._activate_ai_search_link_tag(txt, tag_name)
                else:
                    # v-fix (yêu cầu: bấm 🤖 AI Search RỒI mà label vẫn xám
                    # -- lẽ ra phải active ngay): trạng thái xám/xanh trước
                    # đây chỉ được TÍNH 1 LẦN DUY NHẤT lúc _online_chat_worker
                    # sinh ra câu trả lời -- nếu tại thời điểm đó user CHƯA
                    # bấm 🤖 AI Search, label bị "đóng băng" xám vĩnh viễn dù
                    # sau đó user có bấm 🤖 AI Search thật hay không, vì
                    # không có gì kích hoạt lại đoạn text ĐÃ render xong.
                    # Giờ đăng ký tag xám này vào self._chat_ai_search_grey_tags
                    # kèm theo keyword nó thuộc về -- ngay khi 🤖 AI Search
                    # (nút ngoài) chạy XONG cho ĐÚNG keyword đó,
                    # _activate_ai_search_links_for_query() sẽ tìm lại tag
                    # này và chuyển nó sang trạng thái bấm được NGAY LẬP TỨC,
                    # không cần user gửi thêm câu hỏi mới hay chờ render lại.
                    self._chat_link_seq = getattr(self, "_chat_link_seq", 0) + 1
                    grey_tag = f"chatlink_grey_{self._chat_link_seq}"
                    txt.insert("end", link_label, (body_tag, grey_tag))
                    txt.tag_config(grey_tag, foreground="#8a8d91")
                    _pending = getattr(self, "_chat_ai_search_grey_tags", None)
                    if _pending is None:
                        _pending = self._chat_ai_search_grey_tags = []
                    _pending.append({
                        "txt": txt, "tag": grey_tag,
                        "query": getattr(self, "_chat_auto_sent_for", "") or "",
                    })
                continue
            m = citation_line_re.match(line)
            if m:
                label, path = m.group(1), m.group(2)
                self._chat_link_seq = getattr(self, "_chat_link_seq", 0) + 1
                tag_name = f"chatlink_{self._chat_link_seq}"
                if _is_outlook_pseudo_path(path) or _is_onenote_pseudo_path(path):
                    # Ẩn hẳn pseudo-path xấu -- chỉ hiện label thân thiện,
                    # gạch dưới, click bằng đúng path thật (ẩn) phía sau.
                    txt.insert("end", label, (body_tag, tag_name))
                else:
                    txt.insert("end", label + "  —  ", body_tag)
                    txt.insert("end", path, (body_tag, tag_name))
                txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)
                txt.tag_bind(tag_name, "<Button-1>", lambda e, p=path: self._open_path_for_chat(p))
                txt.tag_bind(tag_name, "<Enter>", lambda e: txt.config(cursor="hand2"))
                txt.tag_bind(tag_name, "<Leave>", lambda e: txt.config(cursor=""))
                continue
            url_matches = list(url_re.finditer(line))
            if not url_matches:
                txt.insert("end", line, body_tag)
                continue
            pos = 0
            for um in url_matches:
                if um.start() > pos:
                    txt.insert("end", line[pos:um.start()], body_tag)
                url = um.group(0)
                self._chat_link_seq = getattr(self, "_chat_link_seq", 0) + 1
                tag_name = f"chatlink_{self._chat_link_seq}"
                txt.insert("end", url, (body_tag, tag_name))
                txt.tag_config(tag_name, foreground="#1a5fb4", underline=True)
                txt.tag_bind(tag_name, "<Button-1>", lambda e, u=url: webbrowser.open(u))
                txt.tag_bind(tag_name, "<Enter>", lambda e: txt.config(cursor="hand2"))
                txt.tag_bind(tag_name, "<Leave>", lambda e: txt.config(cursor=""))
                pos = um.end()
            if pos < len(line):
                txt.insert("end", line[pos:], body_tag)

    def _open_path_for_chat(self, path):
        """Open a file path referenced in an AI Chat "Nguồn:" citation.

        v-fix (link Outlook/OneNote trong AI Chat không mở được): this used
        to assume citations were always plain filesystem paths and just
        did os.path.exists()/os.startfile() on them -- but Outlook/OneNote
        hits flow through the SAME File Content pipeline as real files
        (see _make_outlook_pseudo_path/_make_onenote_pseudo_path), so their
        "path" is really a synthetic OUTLOOK::/ONENOTE:: token, not a real
        filesystem location. os.path.exists() on that always fails,
        producing "Không tìm thấy đường dẫn: OUTLOOK::......msg" -- even
        though the exact same item opens fine by double-clicking it in the
        actual search results list, which already knows to route these
        through Outlook/OneNote's COM API instead (see
        _open_outlook_result/_open_onenote_result in show_results). Do the
        same check-and-route here."""
        try:
            p = (path or "").strip()
            if not p:
                messagebox.showerror("AI Chat", "Không tìm thấy đường dẫn.")
                return
            if _is_outlook_pseudo_path(p):
                entry_id, store_id = _parse_outlook_pseudo_path(p)
                if not entry_id:
                    messagebox.showerror("AI Chat", "Không đọc được thông tin email này.")
                    return
                def _do_open_outlook():
                    ok, err = outlook_search.open_outlook_item(entry_id, store_id)
                    if not ok:
                        print(f"[Outlook] Could not open item: {err}")
                        self.root.after(0, lambda: messagebox.showerror(
                            "Outlook", f"Không mở được email này trong Outlook:\n{err}"))
                threading.Thread(target=_do_open_outlook, daemon=True).start()
                return
            if _is_onenote_pseudo_path(p):
                page_id = _parse_onenote_pseudo_path(p)
                if not page_id:
                    messagebox.showerror("AI Chat", "Không đọc được thông tin trang OneNote này.")
                    return
                def _do_open_onenote():
                    ok, err = onenote_search.open_onenote_page(page_id)
                    if not ok:
                        print(f"[OneNote] Could not open page: {err}")
                        self.root.after(0, lambda: messagebox.showerror(
                            "OneNote", f"Không mở được trang này trong OneNote:\n{err}"))
                threading.Thread(target=_do_open_onenote, daemon=True).start()
                return
            if not os.path.exists(p):
                messagebox.showerror("AI Chat", f"Không tìm thấy đường dẫn:\n{p}")
                return
            try:
                os.startfile(p)
            except Exception:
                subprocess.Popen(f'cmd /c start "" "{p}"', shell=True)
        except Exception as _e:
            print(f"[AI Chat] open path error: {_e}")
            try:
                messagebox.showerror("AI Chat", f"Không mở được:\n{path}\n\n{_e}")
            except Exception:
                pass

    def _safe_focus_chat_entry(self):
        """Focus the AI Chat entry -- but NEVER steal keyboard focus away
        from the main Search box (self.entry) while the user is actively
        typing there.

        v-fix (JP IME 変換 bug): _set_chat_busy(False) and _new_online_chat()
        both used to call self._chat_entry.focus_set() unconditionally.
        Those two fire constantly from the AUTO chat summary path
        (_auto_ai_chat_after_search), completely independent of whether the
        user ever touched the Chat panel -- e.g. every time a background AI
        reply lands, or every time the debounced search settles on a "new"
        keyword. When typing Japanese, committing an IME conversion (Space/
        henkan then Enter to kakutei, e.g. らいせんす -> ライセンス) can
        land right when one of these fires, silently yanking OS keyboard
        focus over to the Chat entry -- so the NEXT characters the user
        types (e.g. さーばー, meant to continue the Search box) landed in
        the AI Chat box instead. Only move focus to Chat now if the Search
        entry does NOT currently have it; if the user is mid-typing in
        Search, leave focus alone."""
        try:
            if self.root.focus_get() is self.entry:
                return
            self._chat_entry.focus_set()
        except Exception:
            pass

    def _set_chat_busy(self, busy):
        """v1.6: no longer puts the Entry itself into Tk's "disabled" state
        -- that used to grey out the whole input pill and leave the
        placeholder's white background looking like a mismatched box on
        top of it. The entry now stays visually unchanged; only the send
        button is disabled and the placeholder swaps between "Ask
        everything..." (idle) and "Thinking..." (waiting on the AI).
        Sending itself is blocked via self._chat_busy, checked in
        _send_online_chat_msg / the Enter-key handler.

        v-fix (Vấn đề: bấm EN/JP không thấy dịch): _translate_current_chat
        tự bỏ qua trong lúc busy=True (đang "Thinking...") -- trước đây
        nút VI/EN/JP vẫn bấm được và đổi màu như bình thường ngay lúc đó,
        khiến người dùng tưởng đã bấm thành công nhưng bản dịch âm thầm
        không chạy. Giờ khoá (disable) luôn 3 nút này trong lúc busy,
        y hệt nút Send, để không còn cú click "vô hình" -- người dùng chỉ
        cần đợi "Thinking..." xong rồi bấm lại là dịch chạy bình thường."""
        try:
            self._chat_busy = busy
            self._set_send_btn_enabled(not busy)
            self._set_new_chat_btn_enabled(not busy)
            self._set_attach_btn_enabled(not busy)
            for _b in getattr(self, "_chat_lang_btns", {}).values():
                try:
                    _b.config(state=("disabled" if busy else "normal"))
                except Exception:
                    pass
            self._chat_placeholder.config(text=CHAT_THINKING_PH if busy else self._current_chat_placeholder_text())
            if busy:
                self._chat_placeholder.place(x=6, rely=0.5, anchor="w")
            else:
                self._toggle_chat_placeholder()
                self._safe_focus_chat_entry()
        except Exception:
            pass

    def _get_followup_topic(self):
        """v-new (dùng chung cho cả 🌐 Internet lẫn 🤖 AI Search offline --
        theo yêu cầu: phải lấy đúng câu hỏi FOLLOW-UP gần nhất user vừa gõ
        trong Chatbox làm chủ đề tra cứu, thay vì LUÔN LẶP LẠI keyword gốc
        ở ô Search chính).

        v-fix (yêu cầu: keyword gốc "Simpack realtime" vẫn bị dính vào 1
        câu hỏi follow-up HOÀN TOÀN KHÔNG LIÊN QUAN, ví dụ "iHawk RTOS là
        máy tính như nào?"): bản trước có thêm bước "nếu follow-up chưa
        nhắc tới keyword gốc thì tự nối thêm keyword gốc vào cuối" -- ý
        định là giữ ngữ cảnh cho câu hỏi rút gọn (vd gõ "GPU" thay vì
        "GPGPU"), nhưng lại vô tình nối CẢ keyword gốc vào những câu hỏi
        follow-up hoàn toàn KHÔNG liên quan tới keyword đó nữa, gây nhiễu.
        Không có cách nào phân biệt chắc chắn "rút gọn của keyword cũ" với
        "câu hỏi mới không liên quan" bằng heuristic đơn giản mà không gọi
        thêm 1 lượt AI để phân loại -- nên bỏ hẳn bước tự nối thêm keyword
        này: dùng ĐÚNG NGUYÊN VĂN câu hỏi follow-up user vừa gõ, không
        thêm/bớt gì.

        Cách làm: duyệt ngược self._chat_history tìm tin nhắn "user" GẦN
        NHẤT do user TỰ GÕ. Nếu tin nhắn gần nhất chỉ là dòng auto-summary
        do chính app tự sinh ra lúc vừa search xong, hoặc chính là 1 lượt
        "...information from Internet🌐"/"...(AI Search offline)" TRƯỚC ĐÓ
        (tức user CHƯA từng tự gõ câu hỏi thật nào) -- thì dùng keyword gốc
        ở ô Search chính như cũ."""
        keyword = (getattr(self, "_chat_auto_sent_for", "")
                   or getattr(self, "_last_chat_placeholder_query", "")
                   or self.entry_var.get().strip())
        _auto_msg_prefixes = (
            "Hãy phân tích và tóm tắt",       # VI
            "Please analyze and summarize",   # EN
            "クエリ「",                         # JP
        )
        last_user_text = ""
        for entry in reversed(self._chat_history or []):
            if entry.get("role") == "user":
                last_user_text = (entry.get("text") or "").strip()
                break
        is_auto_or_link = (not last_user_text
                            or last_user_text.startswith(_auto_msg_prefixes)
                            or last_user_text.endswith("information from Internet🌐")
                            or last_user_text.endswith("(AI Search offline)"))
        return keyword if is_auto_or_link else last_user_text

    def _send_web_search_chat_msg(self):
        """v-new (nút 🌐 -- theo yêu cầu: tắt grounding tự động, chỉ tra
        cứu internet khi user CHỦ ĐỘNG bấm): gửi 1 lượt hỏi giả lập
        "<chủ đề>" information from Internet🌐 vào khung chat, ép
        force_web_search=True cho ĐÚNG lượt này (xem _online_chat_worker)
        -- không đụng tới các lượt chat cục bộ khác, không cần user tự
        gõ câu hỏi. v-new: "<chủ đề>" giờ lấy từ _get_followup_topic()
        (câu hỏi follow-up gần nhất nếu có) thay vì luôn luôn là keyword
        gốc ở ô Search chính.

        v-fix (yêu cầu: bỏ hẳn nút 🌐 riêng, vì đã có link "🌐 Bấm vào đây
        ..." ngay trong khung chat làm việc này rồi): hàm này giờ CHỈ còn
        được gọi từ tag click của link đó (_insert_chat_text_with_links)
        -- không còn 1 widget Label 🌐 riêng nào để tự grey-out/enable
        nữa, nên mọi thao tác lên self._chat_web_search_btn đã bị bỏ."""
        if getattr(self, "_chat_busy", False):
            return  # a request is already running -- ignore until it finishes
        # v-new (yêu cầu: mỗi câu hỏi chỉ tra Internet 1 lần -- lần 2 không
        # cần thiết): link trong chat tự grey-out ngay sau khi bấm 1 lần
        # (xem _insert_chat_text_with_links), nhưng vẫn check lại cờ ở đây
        # phòng trường hợp hàm này được gọi bằng cách khác.
        if getattr(self, "_chat_web_search_used", False):
            return
        if getattr(self, "_chat_preview_idx", -1) != -1:
            self._chat_preview_idx = -1
            self._render_chat_preview()
        topic = self._get_followup_topic()
        if not topic:
            return
        msg = f'"{topic}" information from Internet🌐'
        self._append_chat_line("user", msg)
        self._chat_web_search_used = True
        self._set_chat_busy(True)
        session_id = getattr(self, "_chat_session_id", 0)
        chat_query = getattr(self, "_chat_auto_sent_for", "") or topic
        threading.Thread(target=self._online_chat_worker,
                          args=(msg, session_id, chat_query),
                          kwargs={"force_web_search": True}, daemon=True).start()

    def _send_ai_search_chat_msg(self):
        """v-new (yêu cầu: label "🤖 Bấm vào đây để xem thêm thông tin từ
        AI Search offline" -- cho AI Chat phân tích lại dùng luôn kết quả
        Jina/BGE (self._ai_cont_res, phong phú hơn BM25-only) thay vì chỉ
        self._last_bm25_cont_res như thường lệ. Chỉ hoạt động khi nút 🤖
        AI Search Ở NGOÀI đã chạy XONG cho ĐÚNG keyword hiện tại (xem điều
        kiện _ai_mode_active/_ai_active_query trong _online_chat_worker,
        nơi label này được render enable/grey) -- nếu chưa, hàm này no-op
        (label lúc đó cũng không có binding click, nhưng vẫn check phòng
        hờ)."""
        if getattr(self, "_chat_busy", False):
            return
        if getattr(self, "_chat_ai_search_link_used", False):
            return
        if not (getattr(self, "_ai_mode_active", False)
                and getattr(self, "_ai_active_query", None) == (getattr(self, "_chat_auto_sent_for", "") or "")
                and getattr(self, "_ai_cont_res", None)):
            return
        if getattr(self, "_chat_preview_idx", -1) != -1:
            self._chat_preview_idx = -1
            self._render_chat_preview()
        topic = self._get_followup_topic()
        if not topic:
            return
        msg = f'"{topic}" (AI Search offline)'
        self._append_chat_line("user", msg)
        self._chat_ai_search_link_used = True
        self._set_chat_busy(True)
        session_id = getattr(self, "_chat_session_id", 0)
        chat_query = getattr(self, "_chat_auto_sent_for", "") or topic
        threading.Thread(target=self._online_chat_worker,
                          args=(msg, session_id, chat_query),
                          kwargs={"force_web_search": False, "use_ai_search_context": True}, daemon=True).start()

    def _send_online_chat_msg(self):
        """Gửi tin nhắn user gõ trong Chatbox -- Enter (không giữ Shift)
        hoặc bấm nút Send đều gọi hàm này (xem _on_chat_entry_return)."""
        if getattr(self, "_chat_busy", False):
            return  # a request is already running -- ignore Enter/Send until it finishes
        # v-fix (theo yêu cầu): gõ tiếp trong lúc đang xem 1 trang cũ (◀)
        # không còn bị chặn -- vì vẫn là cùng 1 từ khóa, tự động quay lại
        # "live" (hiển thị lại toàn bộ self._chat_history) rồi gửi tiếp vào
        # đó, thay vì phải bấm ▶ trước.
        if getattr(self, "_chat_preview_idx", -1) != -1:
            self._chat_preview_idx = -1
            self._render_chat_preview()
        msg = self._chat_entry.get("1.0", "end-1c").strip()
        if not msg:
            return
        self._chat_entry.delete("1.0", "end")
        self._toggle_chat_placeholder()
        self._autosize_chat_entry()
        self._append_chat_line("user", msg)
        # v-new (yêu cầu: link 🌐/🤖 chỉ cần bấm 1 lần/câu hỏi): 1 câu hỏi
        # MỚI (do user tự gõ) reset lại cả 2 cờ, để 2 link lại dùng được
        # cho câu hỏi này nếu cần -- không bị "kẹt" greyout mãi từ câu hỏi
        # trước.
        self._chat_web_search_used = False
        self._chat_ai_search_link_used = False
        self._set_chat_busy(True)
        # v-fix (Vấn đề 2): snapshot session_id/query RIGHT NOW, at the
        # moment this request is spawned -- not later, inside the worker,
        # after the AI has replied (see _online_chat_worker docstring).
        session_id = getattr(self, "_chat_session_id", 0)
        query = getattr(self, "_chat_auto_sent_for", "") or ""
        threading.Thread(target=self._online_chat_worker,
                          args=(msg, session_id, query), daemon=True).start()

    def _translate_current_chat(self, target_lang):
        """v-new: called when the user clicks VI/EN/JP while a conversation
        is already showing in the live AI Chat panel -- translates the
        FULL conversation already on screen into the newly selected
        language (instead of only affecting messages sent from now on).
        Runs in the background; sending/translating again is blocked via
        the same _chat_busy flag a normal chat request uses."""
        if getattr(self, "_chat_busy", False):
            return
        if not self._chat_history:
            return
        self._set_chat_busy(True)
        try:
            self._chat_placeholder.config(text=CHAT_TRANSLATING_PH.get(target_lang, CHAT_THINKING_PH))
            self._chat_placeholder.place(x=6, rely=0.5, anchor="w")
        except Exception:
            pass
        threading.Thread(target=self._translate_chat_worker,
                          args=(target_lang, list(self._chat_history)), daemon=True).start()

    def _translate_chat_worker(self, target_lang, messages):
        """Background worker for _translate_current_chat().

        v-rewrite (Vấn đề: dịch JP hay bị dịch DỞ DANG -- ví dụ tin nhắn
        của user đã sang tiếng Nhật nhưng câu trả lời của AI vẫn còn tiếng
        Việt): the previous design sent the WHOLE capped batch of messages
        as ONE "reply with a JSON array" request, betting that the model's
        reply would (a) be valid JSON, (b) have EXACTLY the same number of
        elements as the input, in the SAME order, every single time. After
        several rounds of hardening _extract_json_str_array (raw literal
        newlines, stray backslashes from Windows paths, stray '[' in a
        preamble, truncated output for token-hungry languages like
        Japanese...) it became clear the whole approach is fundamentally
        fragile: ANY one of dozens of ways a multi-thousand-token JSON
        array can come back slightly malformed either kills the ENTIRE
        batch, or -- worse, as seen here -- can silently shift/short one
        element, leaving that message's translation just missing while
        everything else looks fine.

        Translating ONE message at a time as PLAIN TEXT (no JSON, no
        array, no length bookkeeping) removes that entire failure class:
        there's nothing to parse, nothing to misalign, and a single
        message failing to translate just leaves THAT one message as-is
        instead of corrupting the whole batch."""
        def _reenable():
            self._set_chat_busy(False)

        model_choice = getattr(self, "_chat_active_model", DEFAULT_ONLINE_MODEL)
        # v-new: skip a model already known rate-limited this session too,
        # not just a missing key -- see _model_usable/_mark_model_rate_limited.
        if not self._model_usable(model_choice):
            alt = self._next_alt_model(model_choice)
            if alt:
                model_choice = alt
                self._chat_active_model = alt
                self._sync_chat_model_picker()
            else:
                # v-fix (Vấn đề 1): no key at all -- can't call the AI to
                # translate anything. But the only thing on screen in this
                # situation is almost always the "no API key" warning
                # itself, a fixed pre-written string (see
                # CHAT_NO_API_KEY_MSG) -- swap it for its translation in
                # the target language directly, no AI call needed, instead
                # of leaving it stuck in whatever language it was first
                # shown in (or worse, showing a confusing "Translate
                # failed" for a message that never needed the AI at all).
                no_key_texts = set(CHAT_NO_API_KEY_MSG.values())
                localized, changed = [], False
                for m in messages:
                    t = m.get("text", "")
                    if t in no_key_texts:
                        localized.append(CHAT_NO_API_KEY_MSG.get(target_lang, t))
                        changed = True
                    else:
                        localized.append(t)
                if changed:
                    def _apply_local():
                        for m, t in zip(self._chat_history, localized):
                            m["text"] = t
                        self._redraw_chat_panel_from_history()
                        self._replace_saved_session(self._chat_history)
                        _reenable()
                    self.root.after(0, _apply_local)
                else:
                    self.root.after(0, _reenable)
                return

        lang_name = CHAT_LANG_NAMES.get(target_lang, target_lang)
        instruction = (
            f"You will be given ONE chat message's text from an internal "
            f"file-search assistant conversation (it may be a user "
            f"question or an assistant answer). Translate ONLY its "
            f"natural-language prose into {lang_name}. Keep every [N] "
            "citation marker (e.g. [1], [2]) EXACTLY as written, and keep "
            "every file name and file path EXACTLY as written (never "
            "translate or alter them) -- only translate a trailing source-"
            f"list heading (e.g. 'Nguồn:', 'Source:', '出典:') into its "
            f"{lang_name} equivalent. Reply with ONLY the translated "
            "message text, nothing else -- no quotes around it, no "
            "markdown code fences, no explanation, no preamble, and NEVER "
            "reformat it as JSON, a function/tool call, or any other "
            "structured/code format -- the reply must always be plain "
            "human-readable prose, exactly like the input."
        )

        # v-fix (dịch cả cuộc hội thoại dài luôn fail): only translate the
        # MOST RECENT messages, not the entire history. Older messages
        # simply keep whatever language they're already in.
        cap = TRANSLATE_MAX_MSGS_GPTOSS if model_choice != "gemini" else TRANSLATE_MAX_MSGS_GEMINI
        cap = min(cap, len(messages))
        to_translate = messages[-cap:] if len(messages) > cap else messages
        offset = max(0, len(self._chat_history) - len(to_translate))

        results = [None] * len(to_translate)   # None = leave this one as-is
        fail_count = 0
        any_hard_error = None
        tried_gptoss_fallback = False
        mc = model_choice

        # v-fix (Vấn đề: bấm JP không dịch câu trả lời, chỉ câu hỏi được
        # dịch, mà không có bất kỳ cảnh báo nào -- trong khi EN thì luôn
        # ổn): nguyên nhân thường là Gemini's RECITATION filter trả về
        # RỖNG khi được yêu cầu dịch một câu trả lời có cấu trúc rõ ràng
        # (heading, danh sách đánh số, nhiều "[N]") sang JP -- CJK dễ bị
        # filter này hiểu nhầm là "tái tạo y nguyên" hơn hẳn EN/VI (xem
        # comment trong _call_online_ai). Code cũ chỉ đổi hẳn sang GPT-OSS
        # MỘT lần cho cả batch rồi bỏ cuộc ngay nếu message đó vẫn fail --
        # không có bước thử lại CHÍNH message đó, và người dùng không thấy
        # bất kỳ dấu hiệu nào là đã fail (bản cũ vẫn hiển thị y nguyên).
        # Sửa: thêm MỘT lượt thử lại với instruction "nới lỏng" (cho phép
        # diễn đạt lại đôi chút thay vì bám sát cấu trúc/văn phong gốc --
        # chính điều này làm giảm khả năng bị filter chặn), áp dụng cho cả
        # trường hợp rỗng/lỗi lẫn trường hợp nghi bị cắt cụt (thiếu [N]).
        #
        # v-fix (Vấn đề: sau khi thêm soft_instruction ở trên, có lúc câu
        # hỏi bị thay bằng nguyên văn '[search: {"query": "keyword"}]' thay
        # vì bản dịch thật): model hiểu câu "phrase things a little
        # differently" là được phép đổi hẳn ĐỊNH DẠNG (thành JSON/tool-call
        # giả) chứ không chỉ đổi CÁCH DIỄN ĐẠT. Siết lại rõ ràng hơn: chỉ
        # được diễn đạt lại CÂU CHỮ, KHÔNG được đổi định dạng.
        soft_instruction = instruction + (
            " If closely mirroring the source text's exact sentence "
            "structure or wording causes you to refuse or return nothing, "
            "it's fine to phrase the SAME plain-prose translation a little "
            "differently -- just make sure the meaning, every [N] marker, "
            "and every file name/path stay exactly the same. This never "
            "means switching to JSON, a tool call, or any structured/code "
            "format -- the output must still be a plain, natural sentence "
            "or paragraph, nothing else."
        )

        # v-fix (Vấn đề: dịch EN bị lỗi 413 "Request too large ... tokens
        # per minute (TPM): Limit 8000, Requested 11324"): tài khoản Groq
        # hiện tại giới hạn CHUNG cho cả prompt + output chỉ 8000 token/
        # phút -- xin generous như Gemini (8192/12288) gần như luôn vượt
        # giới hạn này ngay từ đầu. _call_online_ai() giờ đã tự retry giảm
        # budget khi dính 413 (dùng đúng số Groq trả về), nhưng xin đúng
        # mức khiêm tốn ngay từ đầu cho GPT-OSS vẫn tốt hơn (đỡ tốn 1 lượt
        # gọi chắc chắn fail). JP cần nhiều token đầu ra hơn hẳn cho cùng
        # nội dung khi dùng Gemini (xem docstring của _call_online_ai).
        out_cap = 12288 if target_lang == "JP" else 8192

        def _budget_for(mc_):
            if mc_ != "gemini":  # gptoss or qwen -- both share Groq's small TPM cap
                return 4000
            return min(out_cap, max(4096, int(len(text) * 2.5)))

        def _looks_like_tool_call(s):
            """v-fix (Vấn đề: '[search: {"query": "..."}]' thế chỗ bản
            dịch thật): model thỉnh thoảng trả về một cấu trúc kiểu lệnh/
            JSON thay vì một câu văn dịch bình thường -- chắc chắn KHÔNG
            phải bản dịch hợp lệ, không được áp dụng dù không rỗng/không
            lỗi."""
            s = s.strip()
            if not s:
                return False
            if re.match(r'^\[\s*[a-zA-Z_]+\s*:\s*\{.*\}\s*\]$', s, re.S):
                return True
            if re.match(r'^[\{\[].*[\}\]]$', s, re.S) and s.count('"') >= 4:
                return True
            return False

        def _clean(raw_):
            tt = raw_.strip()
            if tt.startswith("```"):
                tt = re.sub(r"^```[a-zA-Z]*\s*", "", tt)
                tt = re.sub(r"\s*```$", "", tt).strip()
            if len(tt) >= 2 and tt[0] == tt[-1] and tt[0] in ("\"", "'") and tt.count(tt[0]) == 2:
                tt = tt[1:-1].strip()
            return tt

        def _split_off_link_lines(text_):
            """v-new (yêu cầu: dịch VI/EN/JP không được làm hỏng link 🤖/🌐
            ở cuối câu trả lời): 2 dòng link này được bọc bởi marker vô
            hình + đã có sẵn bản dịch cố định cho 3 ngôn ngữ (xem
            CHAT_AI_SEARCH_LINK_TEXT/CHAT_INTERNET_LINK_TEXT) -- KHÔNG được
            gửi cho AI dịch tự do (AI sẽ không biết giữ lại marker vô hình,
            khiến link biến thành text thường không bấm được nữa). Tách
            chúng ra khỏi cuối text TRƯỚC khi dịch, dịch phần prose còn
            lại như bình thường, rồi _rebuild_link_suffix() ghép lại bằng
            đúng bản dịch cố định theo target_lang."""
            ai_state = None
            has_internet = False
            lines_ = text_.split("\n")
            while lines_ and not lines_[-1].strip():
                lines_.pop()
            if lines_ and lines_[-1].startswith(_CHAT_INTERNET_LINK_MARKER):
                has_internet = True
                lines_.pop()
                while lines_ and not lines_[-1].strip():
                    lines_.pop()
            if lines_ and lines_[-1].startswith(_CHAT_AISEARCH_LINK_MARKER):
                _rest = lines_[-1][len(_CHAT_AISEARCH_LINK_MARKER):]
                if _rest[:1] in ("0", "1"):
                    ai_state = _rest[:1]
                lines_.pop()
                while lines_ and not lines_[-1].strip():
                    lines_.pop()
            return "\n".join(lines_), ai_state, has_internet

        def _rebuild_link_suffix(ai_state, has_internet, lang):
            parts = []
            if ai_state is not None:
                label = CHAT_AI_SEARCH_LINK_TEXT.get(lang, CHAT_AI_SEARCH_LINK_TEXT["EN"])
                parts.append(f"{_CHAT_AISEARCH_LINK_MARKER}{ai_state}{label}")
            if has_internet:
                label2 = CHAT_INTERNET_LINK_TEXT.get(lang, CHAT_INTERNET_LINK_TEXT["EN"])
                parts.append(f"{_CHAT_INTERNET_LINK_MARKER}{label2}")
            return ("\n\n" + "\n\n".join(parts)) if parts else ""

        for i, m in enumerate(to_translate):
            text = m.get("text", "")
            if not text.strip():
                continue
            text, _link_ai_state, _link_has_internet = _split_off_link_lines(text)
            if not text.strip():
                # Toàn bộ tin nhắn CHỈ có link (không nên xảy ra trong thực
                # tế -- link luôn đi kèm phần phân tích phía trên) -- vẫn
                # ghép lại suffix đã dịch, không gọi AI dịch 1 chuỗi rỗng.
                results[i] = _rebuild_link_suffix(_link_ai_state, _link_has_internet, target_lang)
                continue

            def _try(mc_, instr_):
                raw_, err_ = _call_online_ai(
                    mc_, instr_, text, max_output_tokens=_budget_for(mc_))
                if err_ or not raw_ or not raw_.strip():
                    return None, (err_ or "empty response")
                tt = _clean(raw_)
                if not tt:
                    return None, "empty response"
                if _looks_like_tool_call(tt):
                    return None, "model returned a JSON/tool-call-shaped reply instead of a translation"
                return tt, None

            t, err = _try(mc, instruction)
            # v-fix ("Translate failed" dù key OK): Gemini trả RỖNG (thường
            # do bộ lọc RECITATION) phổ biến hơn rate-limit -- dịch thuật
            # không cần tính năng riêng của Gemini nên tự chuyển hẳn sang
            # GPT-OSS cho phần còn lại của batch nếu gặp bất kỳ lỗi nào từ
            # Gemini (chỉ chuyển 1 lần, áp dụng cho các tin nhắn còn lại).
            if err and mc == "gemini" and not tried_gptoss_fallback:
                _groq_alt = next((m for m in ("gptoss", "qwen") if _online_ai_available(m)), None)
                if _groq_alt:
                    tried_gptoss_fallback = True
                    mc = _groq_alt
                    self._chat_active_model = _groq_alt
                    self._sync_chat_model_picker()
                    t, err = _try(mc, instruction)
            # Vẫn rỗng/lỗi/sai định dạng sau khi (có thể) đã đổi model --
            # thử lại ĐÚNG model hiện tại một lần nữa với instruction nới
            # lỏng trước khi chấp nhận bỏ cuộc cho message này.
            if err:
                t, err = _try(mc, soft_instruction)
            if err or not t:
                # Leave THIS message untranslated -- doesn't block the rest
                # of the batch. Remember the error only for the edge case
                # where EVERY message in the batch fails (see below), and
                # log it so a partial failure isn't completely invisible.
                any_hard_error = any_hard_error or err or "empty response"
                fail_count += 1
                print(f"[translate] msg {i} ({m.get('role','?')}) -> {target_lang} "
                      f"FAILED after retries via {mc}: {any_hard_error}")
                continue
            # v-fix (Vấn đề: sau khi đổi VI->EN->JP->VI liên tiếp, bấm VI
            # chỉ thấy câu hỏi + DÒNG ĐẦU của câu trả lời chuyển sang tiếng
            # Việt, phần còn lại vẫn tiếng Anh/Nhật): dịch lại một tin nhắn
            # ĐÃ bị cắt cụt (do bug max_output_tokens ở trên, hoặc do
            # provider vẫn cắt dù đã tăng token) rất dễ khiến model bối
            # rối/trả lời càng ngắn hơn -- áp thẳng kết quả đó (m["text"] =
            # t) thì coi như GHI ĐÈ câu trả lời đầy đủ cũ bằng một bản dịch
            # cụt ngủn, đúng triệu chứng đang gặp. Chặn ở đây bằng cách
            # đếm số marker trích dẫn [N] -- bản dịch PHẢI giữ nguyên số
            # lượng [N] như bản gốc (đúng theo instruction ở trên); nếu
            # thiếu, coi như bản dịch KHÔNG đầy đủ và bỏ qua (giữ nguyên
            # text cũ) thay vì áp một bản nửa vời đè lên bản đầy đủ trước
            # đó. Đếm [N] thay vì so độ dài ký tự vì độ dài tự nhiên chênh
            # lệch nhiều giữa VI/EN/JP (không phải dấu hiệu tin cậy).
            # v-fix (Vấn đề: câu trả lời dài luôn bị coi là "bị cắt" dù dịch
            # đủ nghĩa, chỉ câu hỏi ngắn không có [N] mới lọt qua): so RAW
            # COUNT của "[N]" là sai vì khi dịch, model hay viết lại câu và
            # GỘP các lần lặp lại của CÙNG một citation (ví dụ gốc có
            # "...[1]... theo [1]... như [1] đã nêu..." dịch thành "...as
            # noted in [1], the document also states...") -- không mất
            # thông tin gì nhưng số lần xuất hiện thô giảm, khiến guard cũ
            # luôn reject các câu trả lời dài (nhiều [N]). Sửa: so theo TẬP
            # HỢP các citation ID duy nhất -- chỉ coi là lỗi khi có ID biến
            # mất HOÀN TOÀN khỏi bản dịch, không quan tâm số lần lặp lại.
            orig_ids = _extract_citation_ids(text)
            t_ids = _extract_citation_ids(t)
            missing_ids = orig_ids - t_ids
            if len(orig_ids) >= 2 and missing_ids:
                # Thử lại một lần với instruction nới lỏng trước khi bỏ
                # cuộc -- nhiều trường hợp "thiếu [N]" thực chất cũng là
                # do model né tránh lặp lại cấu trúc gốc (recitation-ish),
                # không phải do output bị cắt cụt thật sự.
                t2, err2 = _try(mc, soft_instruction)
                if t2 and not err2 and not (orig_ids - _extract_citation_ids(t2)):
                    t, missing_ids = t2, set()
                if missing_ids:
                    any_hard_error = any_hard_error or "translation looked truncated (missing [N] citation markers) -- kept original"
                    fail_count += 1
                    print(f"[translate] msg {i} ({m.get('role','?')}) -> {target_lang} "
                          f"missing citations {sorted(missing_ids)} after retry -- kept original")
                    continue
            results[i] = t + _rebuild_link_suffix(_link_ai_state, _link_has_internet, target_lang)

        if all(r is None for r in results):
            self.root.after(0, lambda: (
                self._append_chat_line(
                    "assistant", f"⚠️ {any_hard_error or 'Translate failed'}", True),
                _reenable()))
            return

        def _apply():
            for m, t in zip(self._chat_history[offset:], results):
                if isinstance(t, str) and t.strip():
                    m["text"] = t
            self._redraw_chat_panel_from_history()
            self._replace_saved_session(self._chat_history)
            # v-fix (Vấn đề: dịch JP fail một phần thì y nguyên bản cũ mà
            # KHÔNG có dấu hiệu gì -- người dùng tưởng nút JP không hoạt
            # động): trước đây any_hard_error chỉ được hiển thị khi TOÀN
            # BỘ batch fail (xem nhánh "all(r is None ...)" ở trên); còn
            # fail một phần thì hoàn toàn im lặng. Giờ báo rõ số tin nhắn
            # không dịch được để không bị hiểu nhầm là bug/không hoạt động.
            if fail_count:
                try:
                    self._append_chat_line(
                        "assistant",
                        f"⚠️ {fail_count}/{len(to_translate)} tin nhắn không "
                        f"dịch được sang {CHAT_LANG_NAMES.get(target_lang, target_lang)} "
                        f"(đã giữ nguyên bản trước đó) -- {any_hard_error}",
                        True)
                except Exception:
                    pass
            _reenable()
        self.root.after(0, _apply)

    def _redraw_chat_panel_from_history(self):
        """Re-render the live AI Chat Text widget from self._chat_history
        (used after a translation pass replaces every message's text)."""
        try:
            txt = self._chat_txt
            txt.config(state="normal")
            txt.delete("1.0", "end")
            txt.config(state="disabled")
        except Exception:
            return
        for m in self._chat_history:
            self._append_chat_line(m.get("role", "assistant"), m.get("text", ""))

    def _load_cached_chat_session_for_query(self, query):
        """v-new (theo yêu cầu: search trùng keyword đã có trong AI Chat
        History thì không gọi API lại): tìm phiên (session) GẦN NHẤT trong
        HISTORY_DB_FILE đã từng lưu cho ĐÚNG keyword này, trả về
        (session_id, [messages]) nếu có, hoặc None nếu chưa từng có.

        Vì sao chỉ cần check role='assistant' tồn tại là đủ, không cần lọc
        lỗi riêng: _online_chat_worker return SỚM ngay khi gặp lỗi (thiếu
        key, 429/hết quota, mọi model đều rate-limited...) -- xem các
        nhánh `return` trước dòng self._save_chat_hist(...) của nó -- nên
        KHÔNG CÓ đường nào lưu một câu trả lời lỗi vào bảng chat_history
        cả. Mọi dòng role='assistant' có mặt trong DB chắc chắn là một câu
        trả lời từng THÀNH CÔNG thật sự.

        Lấy theo session_id lớn nhất (mới nhất) khớp query -- nếu 1
        keyword có nhiều phiên đã lưu (search lại nhiều lần, hoặc từng
        bấm "New Chat" trên cùng keyword đó), ưu tiên phiên gần đây nhất.

        Trả về None im lặng nếu DB/bảng chưa tồn tại hoặc có lỗi bất kỳ --
        nơi gọi (_auto_ai_chat_after_search) tự rơi về nhánh gọi API như
        cũ, không có gì vỡ."""
        q = (query or "").strip()
        if not q:
            return None
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE)
            c = conn.cursor()
            c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='chat_history'")
            if c.fetchone()[0] == 0:
                conn.close()
                return None
            row = c.execute(
                "SELECT session_id FROM chat_history WHERE query = ? AND role = 'assistant' "
                "AND text != '' ORDER BY session_id DESC, id DESC LIMIT 1", (q,)).fetchone()
            if not row:
                conn.close()
                return None
            session_id = row[0]
            rows = c.execute(
                "SELECT role, text FROM chat_history WHERE session_id = ? ORDER BY id ASC",
                (session_id,)).fetchall()
            conn.close()
            if not rows:
                return None
            return session_id, [{"role": r, "text": t} for r, t in rows]
        except Exception as _e:
            print(f"[Online Chat] cache lookup for query={query!r} failed: {_e}")
            return None

    def _reuse_cached_chat_session(self, query, cached):
        """v-new: hiển thị lại một phiên AI Chat đã lưu sẵn cho keyword
        này (tìm bởi _load_cached_chat_session_for_query) thay vì gọi API
        -- KHÔNG tốn request Gemini/GPT-OSS nào cho lượt search này.

        v-note (vì sao "New Chat" vẫn luôn ra câu trả lời MỚI, không bị
        cache chặn): điểm gọi duy nhất của hàm này nằm trong nhánh
        `prev not in (None, query)` của _auto_ai_chat_after_search -- tức
        là CHỈ chạy đúng lúc app sắp chuyển sang một keyword khác/mới lần
        đầu trong phiên làm việc này. Khi user tự bấm "New Chat" trên
        CÙNG keyword đang xem, self._chat_auto_sent_for không đổi (giữ
        nguyên == query -- xem _new_online_chat) nên lần gọi lại sau đó
        rơi thẳng vào nhánh else/elif, không bao giờ chạm hàm cache này --
        vẫn gọi API bình thường như trước, đúng ý "New Chat = làm lại từ
        đầu".

        Tiếp tục dùng ĐÚNG session_id cũ đã lưu (thay vì tạo session mới)
        -- nếu user gõ tiếp câu hỏi thủ công trong panel, nó nối tiếp vào
        đúng phiên đã lưu này thay vì bắt đầu một phiên rời rạc."""
        session_id, messages = cached
        self._chat_history = list(messages)
        self._chat_citation_paths = []  # phiên cũ nạp lại -- không có snippet context mới của lượt search này để đối chiếu số [N]
        self._chat_session_id = session_id
        self._chat_auto_sent_for = query
        self._set_chat_busy(False)
        self._redraw_chat_panel_from_history()
        # v-fix (theo yêu cầu 2, lượt 2: thêm keyword vào dòng ghi chú cho
        # dễ hiểu hơn -- keyword ở đây chính là `query`, đúng chuỗi user
        # đã gõ vào Searchbox và dùng để tra cứu cache phía trên, xem
        # _load_cached_chat_session_for_query): "🤖 📎 (from AI Chat
        # History)" -> "🤖 📎 (from AI Chat History for "{query}")".
        _note = f'🤖 📎 (from AI Chat History for "{query}")'
        try:
            self._append_chat_line("assistant", _note, False)
        except Exception:
            pass
        try:
            self._notify_chat_history_panel_for(query)
            self._refresh_chat_preview_nav(query, force_live=True)
        except Exception:
            pass

    def _replace_saved_session(self, messages):
        """Overwrite the CURRENT session's rows in HISTORY_DB_FILE with the
        (now translated) messages, so re-opening this keyword later in AI
        Chat History shows the translated version too, not the original."""
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("CREATE TABLE IF NOT EXISTS chat_history "
                      "(id INTEGER PRIMARY KEY, role TEXT, text TEXT, date TEXT, session_id INTEGER DEFAULT 0)")
            for _ddl in ("ALTER TABLE chat_history ADD COLUMN session_id INTEGER DEFAULT 0",
                         "ALTER TABLE chat_history ADD COLUMN query TEXT DEFAULT ''"):
                try:
                    c.execute(_ddl)
                except Exception:
                    pass
            session_id = getattr(self, "_chat_session_id", 0)
            query = getattr(self, "_chat_auto_sent_for", "") or ""
            c.execute("DELETE FROM chat_history WHERE session_id=?", (session_id,))
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for m in messages:
                c.execute("INSERT INTO chat_history (role, text, date, session_id, query) VALUES (?,?,?,?,?)",
                          (m.get("role", "assistant"), m.get("text", ""), now, session_id, query))
            conn.commit(); conn.close()
        except Exception as _e:
            print(f"[Chat History] Replace-on-translate FAILED: {_e}")

    def _notify_chat_history_panel_for(self, query):
        """Tell ChatHistoryPanel (the saved AI Chat History pane) to show
        this query, without requiring the user to manually click the
        matching row in Search History first. Safe to call from a
        background thread (hops to the main thread via root.after) and
        safe to call before anything has actually been saved yet for this
        query -- show_for_query() already handles the empty case, and its
        own 5s auto-refresh (see ChatHistoryPanel._schedule_auto_refresh)
        will pick up the reply once it lands.

        v-fix (History hiện lẫn cả từ khóa cũ lẫn hiện tại): this used to
        call show_for_query(query) unconditionally. It's called twice per
        auto-chat round -- once when the request starts, again once the
        reply is saved a few seconds later (from a background thread). If
        the user has since moved on to a DIFFERENT search in those few
        seconds, that second, delayed call still fires and flips the pane
        back to the now-abandoned query, right after the pane had already
        correctly switched to whatever the user is looking at now -- which
        can look like both queries' content are showing depending on which
        callback lands last. Only actually push the switch if this query
        is still the one currently active in the Searchbox."""
        try:
            chp = getattr(self, "chat_history_panel", None)
            if chp is None:
                return
            def _do():
                try:
                    if self.entry_var.get().strip() == query:
                        chp.show_for_query(query)
                except Exception:
                    pass
            self.root.after(0, _do)
        except Exception:
            pass

    def _update_chat_placeholder(self, query):
        """v-new (theo yêu cầu): placeholder của ô nhập AI Chat đổi từ tĩnh
        "Ask anything..." thành động theo keyword đang search, ví dụ
        "Ask anything about Simpack adas" -- chỉ áp dụng khi ô đang trống
        (nếu người dùng đã gõ dở câu hỏi thì không đụng vào text đó, chỉ
        đổi phần placeholder ẩn phía sau)."""
        try:
            if not query:
                return
            self._last_chat_placeholder_query = query  # v-fix: remembered so
            # _current_chat_placeholder_text() below can restore this text
            # after "Thinking..."/re-enable/New-Chat resets, instead of
            # those spots falling back to the static CHAT_ASK_PH forever.
            self._chat_placeholder.config(text=f"Ask anything about {query}...")
        except Exception:
            pass

    def _current_chat_placeholder_text(self):
        """v-fix (Vấn đề: "Ask anything about {keyword}" hiện đúng lúc vừa
        search xong, nhưng biến mất ngay sau khi AI trả lời/finish "đang
        suy nghĩ..." xong): 3 chỗ khác trong code (bật/tắt "Thinking...",
        re-enable sau khi gửi xong, New Chat) đều reset thẳng về hằng số
        tĩnh CHAT_ASK_PH="Ask anything...", ghi đè mất text động đã set ở
        _update_chat_placeholder. Dùng hàm này ở tất cả những chỗ đó thay
        vì literal CHAT_ASK_PH, để luôn quay lại đúng "Ask anything about
        {keyword hiện tại}" nếu đã biết, chỉ fallback về text tĩnh khi
        chưa từng search lần nào."""
        q = getattr(self, "_last_chat_placeholder_query", "")
        return f"Ask anything about {q}..." if q else CHAT_ASK_PH

    def _auto_ai_chat_after_search(self, query, _deadline=None):
        """Automatically ask the online AI chat to analyze/summarize the
        freshly found search results the moment AI Search runs, instead of
        leaving the chat panel empty until the user types something
        themselves. v1.7: this used to only ever fire once per app session
        -- any query after the very first one was silently ignored because
        the chat already had history. Now, whenever this is called for a
        genuinely NEW query (different from the one the chat was last
        auto-summarized for), it automatically starts a New Chat first
        (clearing the old conversation) and then summarizes the new query,
        so every keyword typed into the Searchbox gets its own fresh
        AI Chat summary.

        v1.10: if the (fast) BM25-only File Content render already happened
        but the follow-up Outlook/OneNote merge is still running (see
        _merge_mail_notes_async), wait for it -- otherwise the AI ends up
        summarizing only the first few documents that were ready early on,
        instead of the complete result set, even though the File Content
        tab itself finishes filling in moments later. Capped at ~6s so a
        stuck COM call (the known OneNote bitness issue) can't block the
        chat forever.

        v-fix (AI Chat phản ứng với từ khóa gõ dở): this only ever gets
        called via _open_ai_split_win(), scheduled ~150ms after a sticky-AI
        re-run finishes for whatever query was in the Searchbox WHEN THAT
        RE-RUN STARTED. The semantic search + rebuild in between easily
        takes longer than the 350ms keystroke debounce, so by the time this
        actually fires, the user has often already typed more (or fixed a
        typo) -- e.g. searching "ddh1" letter by letter kicks off a sticky
        re-run for "dd", which then finishes and fires this well after the
        box already reads "ddh1"; typing "aasd" then quickly correcting to
        "adas" fires this for the stale "aasd". Bail out immediately if the
        query we were asked to summarize no longer matches what's actually
        in the Searchbox right now -- the next genuine settle (the box
        stops changing) will re-fire this correctly for the real, final
        text instead."""
        try:
            if not query:
                return
            if query != self.entry_var.get().strip():
                return  # stale -- user has already typed further / changed the text since this was scheduled
            self._update_chat_placeholder(query)

            # v-fix (AI Chat phân tích quá sớm khi gõ tiếng Nhật nhiều từ):
            # don't actually send the auto-summary until the Search box has
            # been quiet (no KeyRelease) for _CHAT_AUTO_QUIET_SEC. Typing JP
            # in bursts -- e.g. らいせんす, henkan/Enter, ~1s pause, then
            # さーばー -- used to get summarized after JUST "ライセンス"
            # the instant that first word settled (350ms search debounce +
            # 150ms), well before the user typed "サーバー". Waiting for a
            # longer genuine pause here (separate from the short search
            # debounce, which still updates results live as before) lets
            # the query actually finish before AI Chat reacts to it. Self-
            # correcting: every real keystroke bumps _last_keystroke_ts, so
            # each reschedule re-checks against the LATEST typing activity.
            _CHAT_AUTO_QUIET_SEC = 0.8
            _idle_for = time.time() - getattr(self, "_last_keystroke_ts", 0.0)
            if _idle_for < _CHAT_AUTO_QUIET_SEC:
                if _deadline is None:
                    _deadline = time.time() + 6.0  # same overall cap as the other waits below
                if time.time() < _deadline:
                    _wait_ms = int(max(50, (_CHAT_AUTO_QUIET_SEC - _idle_for) * 1000))
                    self.root.after(_wait_ms, lambda: self._auto_ai_chat_after_search(query, _deadline))
                    return
                # deadline blown (rare) -- fall through and summarize whatever's there now

            if getattr(self, "_chat_busy", False):
                return  # a request is already running
            if not (self._last_bm25_cont_res or self._ai_cont_res):
                return  # nothing indexed yet to talk about

            # v-fix (theo yêu cầu: email/note THẬT SỰ chứa từ khóa nhưng vẫn
            # không vào "Nguồn:" -- deadline CHUNG 6s cho cả 2 lượt chờ (gõ
            # chậm lại + merge mail/notes) đôi khi không đủ trên máy có
            # nhiều mail/note, đặc biệt lần tra cứu đầu (COM/sqlite chưa
            # "ấm"). Tăng lên 12s CHỈ cho nhánh chờ merge này -- việc chờ
            # thêm vài giây không chặn UI (đây là async callback, người
            # dùng vẫn gõ/xem kết quả bình thường), và có context ĐẦY ĐỦ ở
            # lượt tóm tắt đầu tiên quan trọng hơn việc tiết kiệm vài giây.
            _MAIL_MERGE_WAIT_SEC = 12.0
            if getattr(self, "_mail_notes_merge_pending", False):
                if _deadline is None:
                    _deadline = time.time() + _MAIL_MERGE_WAIT_SEC
                if time.time() < _deadline:
                    self.root.after(200, lambda: self._auto_ai_chat_after_search(query, _deadline))
                    return
                # v-fix: log rõ khi thật sự bị timeout, để không phải đoán
                # mò lần sau nếu Outlook/OneNote vẫn thiếu trong "Nguồn:".
                print(f"[Online Chat] mail/notes merge still pending after "
                      f"{_MAIL_MERGE_WAIT_SEC:.0f}s -- summarizing WITHOUT "
                      f"them for query={query!r} (they'll still show up on "
                      f"the next follow-up question once the merge finishes)")

            # v-fix (AI Chat trả lời trùng lặp): see _last_auto_chat_attempt
            # in __init__. The "AI Search" button fires this function
            # directly AND indirectly (through _open_ai_split_win once its
            # own embedding search finishes) -- if BOTH attempts slip past
            # the _chat_busy check (e.g. the first reply came back fast
            # enough that busy was already cleared again before the second,
            # delayed trigger fires), the second one would re-send the SAME
            # auto-summary prompt and produce a near-identical duplicate
            # answer. Collapse any 2nd+ attempt for the same query within a
            # few seconds of the first into a no-op.
            _last_attempt = getattr(self, "_last_auto_chat_attempt", None)
            if _last_attempt and _last_attempt[0] == query and (time.time() - _last_attempt[1]) < 5.0:
                return
            self._last_auto_chat_attempt = (query, time.time())

            # v-fix (Vấn đề 3 -- AI chat phản ứng quá nhanh, chỉ đọc được
            # 1-2 file dù có ~100 kết quả): self._last_bm25_cont_res (the
            # full BM25 result set _online_chat_worker slices file context
            # from) is written by _smart_search_realtime for whatever query
            # it was LAST called with -- normally that's already this exact
            # query by the time we get here, since the sticky-AI retrigger
            # that leads to this call only starts AFTER that write. But on
            # a fast follow-up search, sticky AI's own semantic-search leg
            # can still be mid-flight when the NEXT keystroke settles and
            # _smart_search_realtime reruns for a newer query in between --
            # or a manual AI Search click can race a File Content render
            # that hasn't reached that assignment yet. Either way, if
            # self._last_query doesn't match the query we're about to
            # summarize, self._last_bm25_cont_res doesn't either, and
            # _online_chat_worker would end up building context from
            # whatever few (or zero, or wrong-query) results happen to be
            # cached instead of the full set -- explaining why only the
            # first file or two ever get referenced. Wait (bounded, same
            # 6s cap as the mail/notes merge above) for the cache to
            # actually catch up to this query before proceeding.
            if getattr(self, "_last_query", None) != query:
                if _deadline is None:
                    _deadline = time.time() + 6.0
                if time.time() < _deadline:
                    self.root.after(150, lambda: self._auto_ai_chat_after_search(query, _deadline))
                    return
                print(f"[Online Chat] _last_query never caught up to {query!r} within 6s -- proceeding anyway with whatever context is cached")

            prev = getattr(self, "_chat_auto_sent_for", None)
            # v-fix (Vấn đề 2 -- keyword ĐẦU TIÊN sau khi mở app không dùng
            # History dù đã search nhiều lần trong quá khứ): điều kiện cũ
            # `prev not in (None, query)` cố tình LOẠI TRỪ trường hợp
            # prev is None -- nhưng prev luôn là None ở lượt auto-trigger
            # ĐẦU TIÊN của mỗi phiên làm việc (self._chat_auto_sent_for chỉ
            # được khởi tạo = None lúc mở app, xem __init__), bất kể keyword
            # đó đã từng được search/lưu History bao nhiêu lần ở NHỮNG LẦN
            # MỞ APP TRƯỚC. Kết quả: lượt search đầu tiên luôn rơi thẳng
            # xuống nhánh gọi API bên dưới (bỏ qua bước tra cứu cache), chỉ
            # từ lượt search THỨ HAI trở đi (khi prev đã là 1 query thật)
            # thì nhánh tra cứu này mới có cơ hội chạy. Dùng `prev != query`
            # thay vì `prev not in (None, query)` để prev is None cũng đi
            # qua bước tra cứu cache như mọi lần đổi keyword khác -- gọi
            # _new_online_chat() vẫn an toàn khi prev is None (chỉ reset
            # các list vốn đã rỗng + tăng session_id, xem docstring của nó).
            if prev != query:
                # v-new (theo yêu cầu: tiết kiệm API token -- keyword đã có
                # sẵn AI Chat History thì không gọi API lại): trước khi bắt
                # đầu 1 phiên MỚI (và do đó sắp gọi API) cho keyword này,
                # kiểm tra xem nó đã từng có 1 phiên trả lời THÀNH CÔNG nào
                # được lưu trong HISTORY_DB_FILE chưa (bất kể lưu từ lượt
                # search nào trước đó, kể cả phiên trước của app). Nếu có,
                # hiển thị lại nguyên phiên đó thay vì tốn 1 lượt gọi
                # Gemini/GPT-OSS. Chỉ can thiệp đúng NHÁNH "khác keyword"
                # này -- nhánh "cùng keyword, đã có sẵn self._chat_history"
                # ngay dưới (elif) vẫn giữ nguyên hành vi cũ (không gọi lại
                # trong live session), và bấm "New Chat" thủ công vẫn luôn
                # cho ra câu trả lời MỚI như trước (xem
                # _load_cached_chat_session_for_query/_reuse_cached_chat_session
                # bên dưới để biết lý do New Chat không bị cache chặn).
                _cached = self._load_cached_chat_session_for_query(query)
                if _cached:
                    self._reuse_cached_chat_session(query, _cached)
                    return
                self._new_online_chat()   # different query -- start fresh
            elif self._chat_history:
                return  # already summarized (or being chatted about) this exact query

            # v-fix (theo yêu cầu: quay lại hiển thị kết quả phân tích
            # NGAY trong khung AI Chat như bản gốc, không cần chờ user tự
            # bấm Enter nữa -- người dùng chỉ search khoảng 10-20 keyword/
            # ngày nên không còn lo tốn quota grounding như trước; xem
            # yêu cầu trước đó về prefill-only đã được HUỶ). Bỏ luôn giới
            # hạn AUTO_AI_CHAT_DAILY_LIMIT/ngày -- không còn ý nghĩa khi
            # người dùng đã chủ động chấp nhận việc tốn quota.
            self._chat_auto_sent_for = query
            # v-fix (Vấn đề: AI Chat History không theo từ khóa mới): this
            # used to only update the LIVE AI Chat panel (self._chat_txt)
            # for a new query -- ChatHistoryPanel (the saved-history pane)
            # is a separate widget that only redraws when the user manually
            # clicks a row in Search History (see
            # HistoryPanel._notify_chat_history_panel), so it kept showing
            # whatever OLD query was last clicked even though a brand new
            # keyword's chat was actively being summarized -- the new
            # content WAS being saved correctly (Search History itself
            # updated fine), just not reflected in this pane until the user
            # remembered to click on it again. Push it to follow the new
            # query immediately instead of waiting for a manual click.
            self._notify_chat_history_panel_for(query)
            self._refresh_chat_preview_nav(query, force_live=True)
            _lang = CHAT_LANG_OPTIONS.get(getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_LANG_OPTIONS[CHAT_DEFAULT_LANG])
            auto_msg = _lang["auto_msg"](query)
            self._append_chat_line("user", auto_msg)
            self._set_chat_busy(True)
            # v-fix (Vấn đề 2): snapshot session_id NOW (query was just set
            # above, on this same thread) -- see _online_chat_worker docstring.
            session_id = getattr(self, "_chat_session_id", 0)
            threading.Thread(target=self._online_chat_worker,
                              args=(auto_msg, session_id, query), daemon=True).start()
        except Exception as _e:
            print(f"[Online Chat] auto trigger error: {_e}")

    def _online_chat_worker(self, user_message, session_id=None, query=None, suppress_web_search=False, force_web_search=False, use_ai_search_context=False):
        """Background worker: call the online AI (default Gemini Flash,
        auto-switching to GPT-OSS 120B on rate-limit) with context built
        from REAL file content excerpts pulled from self._last_bm25_cont_res
        (latest local BM25/FTS5 results) -- never letting the AI make
        things up, same idea as how app_astro-weather forces Gemini to call
        a tool instead of guessing.

        v-fix (Vấn đề 2 -- Chat history bị lẫn nội dung): session_id/query
        are now captured by the CALLER at the moment this worker thread is
        spawned (see _send_online_chat_msg / _auto_ai_chat_after_search),
        and threaded through as explicit arguments -- instead of this
        function reading self._chat_session_id / self._chat_auto_sent_for
        itself, which only happens after the AI has replied (can be several
        seconds later). If the user switches to a different search keyword
        or clicks "New Chat" while this request is still in flight, those
        instance attributes will have already moved on to the NEW
        session/query by the time the reply comes back -- reading them at
        save-time silently filed this OLD answer under the NEW session
        (e.g. an answer about "Simpack ADAS" ending up saved under, and
        displayed inside, the "Simpack flexible" chat). We now always save
        this turn to chat_history under the session it actually belongs
        to, and only splice the reply into the live chat panel / running
        context (self._chat_history) / busy-state if that session is still
        the one currently on screen -- a stale/superseded reply is saved
        for the record but never shown."""
        if session_id is None:
            session_id = getattr(self, "_chat_session_id", 0)
        if query is None:
            query = getattr(self, "_chat_auto_sent_for", "") or ""

        def _is_current():
            # True only if no newer session (New Chat / different search)
            # has started since this request was spawned.
            return session_id == getattr(self, "_chat_session_id", 0)

        def _reenable():
            # Don't clear busy-state on behalf of a NEWER, still-running
            # request that may have started for the current session in the
            # meantime -- only the request that actually owns the current
            # session may re-enable sending.
            if _is_current():
                self._set_chat_busy(False)

        model_choice = getattr(self, "_chat_active_model", DEFAULT_ONLINE_MODEL)

        # v-new: skip a model already known rate-limited this session too,
        # not just a missing key -- see _model_usable/_mark_model_rate_limited.
        # This is what makes the switch "sticky": once GPT-OSS (or Gemini)
        # has been marked rate-limited, EVERY subsequent message starts
        # straight on the other model instead of wasting a call retrying
        # the one already known to be exhausted.
        if not self._model_usable(model_choice):
            alt = self._next_alt_model(model_choice)
            if alt:
                model_choice = alt
                self._chat_active_model = alt
                self._sync_chat_model_picker()
            else:
                self._mark_api_needs_attention(True)  # v-new: light up 🔑 Update API
                warn_text = CHAT_NO_API_KEY_MSG.get(
                    getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_NO_API_KEY_MSG["VI"])
                # v-fix (Vấn đề 1 -- VI/EN/JP không dịch được cảnh báo này):
                # this used to ONLY call _append_chat_line (draws it into
                # the Text widget) without ever adding it to
                # self._chat_history. _translate_current_chat() bails out
                # immediately whenever self._chat_history is empty (see
                # "if not self._chat_history: return"), so clicking
                # VI/EN/JP silently did nothing -- there was nothing in
                # chat_history to translate. Now recorded in chat_history
                # too, so the CHAT_NO_API_KEY_MSG static-swap branch added
                # to _translate_chat_worker can actually find and localize
                # it.
                self.root.after(0, lambda: (
                    (self._append_chat_line("assistant", warn_text, True)
                     if _is_current() else None),
                    (self._chat_history.append({"role": "assistant", "text": warn_text})
                     if _is_current() else None),
                    _reenable()))
                return

        # v-new (nút Online): user-controlled, default OFF -- see v-fix:
        # Online giờ luôn True mặc định (xem __init__ và CHAT_ONLINE_SUFFIX).
        # v-fix (Vấn đề cũ -- đã sửa: comment trước đây nói sai là "Groq/
        # GPT-OSS has no built-in web-search tool" nên luôn ép đổi sang
        # Gemini bất cứ khi nào Online bật): Groq's browser_search tool
        # THỰC SỰ hoạt động cho đúng model "gptoss" (xem
        # _call_online_ai_chat) -- giờ chỉ ép đổi model khi model đang
        # active KHÔNG nằm trong WEB_SEARCH_CAPABLE_MODELS (tức chỉ còn
        # đúng 1 trường hợp: "qwen"), ưu tiên Gemini trước rồi mới tới
        # GPT-OSS, theo đúng MODEL_FALLBACK_ORDER.
        #
        # v-fix (Vấn đề: 14 sources dù model default là GPT-OSS, và/hoặc
        # cảnh báo "model đang dùng GPT-OSS" không khớp với model thực sự
        # vừa trả lời): khối CHỐT model_choice này trước đây nằm SAU đoạn
        # build context (_max_ctx_files/system_prompt bên dưới) -- nên cap
        # số file gửi đi (theo GPTOSS_MAX_CONTEXT_FILES hay
        # ONLINE_AI_MAX_CONTEXT_FILES) luôn được tính theo model_choice CŨ,
        # từ TRƯỚC khi Online có cơ hội ép đổi model -- ngữ cảnh gửi đi
        # không khớp với model THỰC SỰ được gọi lượt này, và dòng cảnh báo
        # cuối bài (CHAT_NO_WEB_SEARCH_NOTE) cũng dựa trên model_choice ở
        # thời điểm sai. Chuyển việc chốt model_choice lên NGAY ĐÂY, chạy
        # trước mọi bước build context, để cap file / snippet size /
        # system_prompt / cảnh báo phía sau đều nhất quán với đúng model sẽ
        # thực sự trả lời.
        # v-fix (theo yêu cầu mới: TẮT hẳn grounding tự động -- lượt phân
        # tích đầu tiên giờ luôn chạy CỤC BỘ trước cho nhanh, tránh đúng
        # độ trễ "grounding thật sự chạy" 1-1.5 phút gặp phải khi để tự
        # động bật ở lượt đầu. Online/grounding giờ CHỈ bật khi user chủ
        # động bấm nút 🌐 riêng (force_web_search=True -- xem
        # _send_web_search_chat_msg), không còn tự động theo lượt chat
        # nào nữa.
        web_search = bool(force_web_search) and not suppress_web_search
        wanted_web_search_but_unavailable = False
        if web_search and model_choice not in WEB_SEARCH_CAPABLE_MODELS:
            # v-new: don't force onto a model just marked rate-limited this
            # session (_online_ai_available only checks the key exists, not
            # remaining quota) -- otherwise, once that model's daily quota
            # is exhausted, EVERY turn would still retry it first
            # (guaranteed 429), only falling back afterwards -- doubling
            # API calls/latency per message, and leaving the raw error
            # unhandled if the fallback is ALSO exhausted at that exact
            # moment. _model_usable() keeps the chat sticky on a working
            # model until the preferred one actually succeeds again.
            _switched = next((m for m in ("gemini", "gptoss") if self._model_usable(m)), None)
            if _switched:
                model_choice = _switched
                self._chat_active_model = _switched
                self._sync_chat_model_picker()
            if model_choice not in WEB_SEARCH_CAPABLE_MODELS:
                wanted_web_search_but_unavailable = True

        # v-new (yêu cầu: label "🤖 AI Search offline" -- dùng kết quả
        # Jina/BGE (self._ai_cont_res, đã merge sẵn BM25+semantic, xem
        # _ai_search_and_update) thay vì chỉ self._last_bm25_cont_res
        # thuần BM25, cho câu trả lời PHONG PHÚ hơn khi user chủ động bấm
        # label này. Rơi về BM25 như cũ nếu vì lý do gì đó chưa có
        # _ai_cont_res (phòng hờ, không nên xảy ra vì _send_ai_search_chat_msg
        # đã tự check trước khi gọi worker với use_ai_search_context=True).
        if use_ai_search_context and getattr(self, "_ai_cont_res", None):
            bm25_cont = self._ai_cont_res or []
        else:
            bm25_cont = self._last_bm25_cont_res or []
        # v-fix: this used to ALWAYS take only the first ONLINE_AI_TOP_N_FILES
        # results (ranked for the original search query) as context, no
        # matter what the user actually asked in this chat turn. So a
        # follow-up question about something further down the result list
        # (visible right there in the File Content tab) never made it into
        # the AI's context, and the AI incorrectly said "not found". Now we
        # also search the FULL result list for filenames matching keywords
        # from THIS specific message, and put those matches first -- the
        # original top-N is still included as a fallback/base context.
        all_paths = [r[0] for r in bm25_cont]
        keywords = [w for w in re.findall(r"[^\W_]+", (user_message or "").lower(), re.UNICODE) if len(w) >= 3]
        matched_paths = []
        if keywords:
            for p in all_paths:
                base = os.path.basename(p).lower()
                if any(kw in base for kw in keywords):
                    matched_paths.append(p)

        # v-fix (Outlook/OneNote gần như không bao giờ vào context của AI
        # Chat): _merge_mail_notes_async always appends mail/notes AFTER the
        # normal BM25 file list, then _sort_priority bands by POSITION in
        # that combined list -- so once there are >= ONLINE_AI_TOP_N_FILES
        # normal files ahead of them, mail/notes fall outside the plain
        # `all_paths[:ONLINE_AI_TOP_N_FILES]` slice every time, no matter
        # how relevant their own BM25 score actually was. Split into 3
        # pools (each pool keeps its own relative relevance order) and
        # reserve a small guaranteed slice for mail/notes, instead of
        # letting normal files silently crowd them out.
        normal_paths, outlook_paths, onenote_paths = [], [], []
        for p in all_paths:
            if _is_outlook_pseudo_path(p):
                outlook_paths.append(p)
            elif _is_onenote_pseudo_path(p):
                onenote_paths.append(p)
            else:
                normal_paths.append(p)

        context_paths = []
        # v-new (yêu cầu: nút "🤖 AI Search offline" phải phân tích SÂU HƠN,
        # dùng nhiều nguồn hơn -- chấp nhận chậm hơn): dùng budget lớn hơn
        # (*_DETAILED) CHỈ khi user chủ động bấm nút đó (use_ai_search_context
        # =True); chat thường/nút 🌐 vẫn giữ nguyên budget cũ như trước.
        if use_ai_search_context:
            _max_ctx_files = GPTOSS_MAX_CONTEXT_FILES_DETAILED if model_choice != "gemini" else ONLINE_AI_MAX_CONTEXT_FILES_DETAILED
        else:
            _max_ctx_files = GPTOSS_MAX_CONTEXT_FILES if model_choice != "gemini" else ONLINE_AI_MAX_CONTEXT_FILES
        # v-new (theo yêu cầu: không tham khảo 2 file trùng TÊN, kể cả
        # khác đuôi PDF/PPT/DOC hay khác số version _v1.1/_v1.2/_v1.3):
        # trước đây chỉ chặn trùng PATH y hệt (`p not in context_paths`),
        # rồi sau đó chỉ chặn trùng basename+extension Y HỆT -- 2 bản PDF
        # cùng tên ở 2 thư mục khác nhau bị gộp lại đúng, nhưng
        # "...v1.3.pdf" vs "...v1.2.pdf" (khác version) và "...pdf" vs
        # "...pptx" (khác đuôi, cùng tên gốc) vẫn bị tính là 2 nguồn khác
        # nhau -- ngốn 2 slot cho nội dung gần như giống hệt nhau, lẽ ra 1
        # trong 2 slot đó nên nhường cho 1 nguồn KHÁC HẲN NỘI DUNG để câu
        # trả lời có nhiều thông tin đa dạng hơn. Giờ gộp theo
        # _ctx_dedup_key() (tên gốc, bỏ đuôi + bỏ hậu tố version), và
        # trong mỗi nhóm trùng tên, luôn dùng bản có version CAO NHẤT
        # (_ctx_version_tuple) làm đại diện -- vd nhóm v1.1/v1.2/v1.3 chỉ
        # giữ lại v1.3. Không áp dụng cho pseudo-path Outlook/OneNote
        # (basename của chúng chỉ là token máy đọc, không có ý nghĩa để
        # so trùng).
        _best_ctx_for_key = {}
        for _p in all_paths:
            if _is_outlook_pseudo_path(_p) or _is_onenote_pseudo_path(_p):
                continue
            _k = _ctx_dedup_key(_p)
            if _k not in _best_ctx_for_key or _ctx_version_tuple(_p) > _ctx_version_tuple(_best_ctx_for_key[_k]):
                _best_ctx_for_key[_k] = _p
        _seen_ctx_basenames = set()
        def _add_ctx(p):
            if len(context_paths) >= _max_ctx_files:
                return
            if not (_is_outlook_pseudo_path(p) or _is_onenote_pseudo_path(p)):
                key = _ctx_dedup_key(p)
                p = _best_ctx_for_key.get(key, p)  # always the latest version among same-title files
                if key in _seen_ctx_basenames or p in context_paths:
                    return  # cùng tên file (mọi đuôi/version) đã có nguồn khác rồi -- nhường slot cho nguồn đa dạng hơn
                _seen_ctx_basenames.add(key)
            elif p in context_paths:
                return
            context_paths.append(p)
        # v-fix (OneNote/Outlook biến mất khỏi context khi dùng GPT-OSS):
        # the mail/notes reserve below (ONLINE_AI_MAIL_RESERVE_FILES=3 each)
        # was sized assuming the OLD fixed cap of ONLINE_AI_MAX_CONTEXT_FILES
        # =14 (= 8 normal + 3 outlook + 3 onenote, exactly). Once GPT-OSS got
        # its own much smaller cap (GPTOSS_MAX_CONTEXT_FILES=5, added to fix
        # the Groq 413 token-limit error), the FIRST loop below
        # (normal_paths[:ONLINE_AI_TOP_N_FILES], i.e. up to 8 files) already
        # filled the entire 5-slot budget on its own -- by the time the
        # outlook/onenote loops ran, len(context_paths) was already at cap,
        # so _add_ctx() silently did nothing for either of them. That's
        # exactly why "Is there any information in OneNote?" got "no
        # excerpt found" even though OneNote hits clearly showed up in the
        # search results. Scale each pool's reserved slice to the ACTUAL
        # budget instead of a fixed number, so mail/notes always keep at
        # least 1 guaranteed slot no matter how small the cap is.
        # v-fix (theo yêu cầu: 14 slot = 8 file thường (pdf/ppt/excel...) +
        # 3 Outlook + 3 OneNote CỐ ĐỊNH; nếu Outlook/OneNote có ÍT HƠN 3
        # kết quả thì phần dư phải dồn lại cho file thường, KHÔNG được bỏ
        # phí): bản CŨ tính _outlook_reserve/_onenote_reserve = 3 luôn
        # (miễn có ít nhất 1 kết quả), rồi trừ cứng vào _normal_take = 14-
        # 3-3 = 8 -- đúng số 8/3/3 khi cả 2 nguồn đủ ≥3 kết quả, nhưng nếu
        # ví dụ Outlook chỉ có 1 kết quả thật, 2 slot "dự trữ" còn lại của
        # nó bị BỎ TRỐNG (chỉ vòng lặp "5) fill remaining" ở cuối mới có
        # cơ hội lấp, và nó lấp theo thứ tự all_paths gốc -- có thể lại là
        # Outlook/OneNote tiếp theo chứ không chắc là file thường). Sửa:
        # tính SỐ THỰC SẼ LẤY của Outlook/OneNote bằng min(3, số thực có),
        # rồi _normal_take = 14 - (số thực lấy đó) -- phần dư tự động
        # chảy sang file thường ngay từ bước tính, không cần chờ tới
        # bước "fill remaining".
        _mail_reserve_table = ONLINE_AI_MAIL_RESERVE_BY_MODEL_DETAILED if use_ai_search_context else ONLINE_AI_MAIL_RESERVE_BY_MODEL
        _mail_reserve = _mail_reserve_table.get(
            model_choice, min(ONLINE_AI_MAIL_RESERVE_FILES, max(1, _max_ctx_files // 4)))
        # v-new (theo yêu cầu: pdf/ppt/doc ưu tiên hơn Outlook/OneNote):
        # bản CŨ cho Outlook/OneNote giữ chỗ TRƯỚC (mỗi loại tối đa
        # _mail_reserve=3, tức 3+3=6) rồi mới tới pdf/ppt/doc -- với
        # GPT-OSS (GPTOSS_MAX_CONTEXT_FILES=5), 6 > 5 nghĩa là Outlook +
        # OneNote có thể chiếm HẾT SẠCH cả 5 slot, không còn chỗ nào cho
        # pdf/ppt/doc dù nội dung các file đó thường đầy đủ & chất lượng
        # hơn hẳn 1 email/note ngắn. Giờ đảo ngược ưu tiên: pdf/ppt/doc
        # (normal_paths + matched_paths) được LẤY BUDGET TRƯỚC; Outlook/
        # OneNote chỉ được đảm bảo TỐI THIỂU 1 slot mỗi loại (nếu có kết
        # quả, để không biến mất hoàn toàn như lỗi cũ "OneNote luôn không
        # tìm được"), phần còn lại của budget mới dành cho pdf/ppt/doc;
        # nếu sau khi pdf/ppt/doc lấy hết vẫn còn dư chỗ, Outlook/OneNote
        # mới được nới rộng thêm lên tối đa _mail_reserve (3) mỗi loại.
        # Thứ tự add cũng chính là thứ tự hiển thị trong "Nguồn:", nên
        # pdf/ppt/doc giờ luôn hiện TRƯỚC msg/OneNote.
        _outlook_min_reserve = 1 if outlook_paths else 0
        _onenote_min_reserve = 1 if onenote_paths else 0
        _normal_budget = max(0, _max_ctx_files - _outlook_min_reserve - _onenote_min_reserve)
        for p in matched_paths:                                     # 1) filename matches for THIS message, normal files only
            if p not in outlook_paths and p not in onenote_paths and len(context_paths) < _normal_budget:
                _add_ctx(p)
        for p in normal_paths:                                      # 2) top-ranked normal files (pdf/ppt/doc/xlsx...) -- PRIORITY
            if len(context_paths) >= _normal_budget:
                break
            _add_ctx(p)
        # v-fix (Vấn đề 4: msg/OneNote KHÔNG BAO GIỜ xuất hiện trong
        # "Nguồn:" dù có kết quả khớp thật -- trong khi pdf/ppt luôn có):
        # 2 vòng lặp Outlook rồi OneNote chạy TUẦN TỰ trước đây, và
        # _add_ctx() chỉ kiểm tra cap TOÀN CỤC (_max_ctx_files), không có
        # giới hạn RIÊNG cho từng loại. Nghĩa là vòng Outlook chạy TRƯỚC
        # được phép "nới" tới tối đa _mail_reserve (vd 2 cho Gemini) --
        # nếu ngân sách còn dư đủ chỗ, nó ăn HẾT sạch phần dư đó (kể cả
        # phần lẽ ra dành cho OneNote), khiến khi tới lượt OneNote chạy,
        # context_paths đã chạm cap và _add_ctx() return ngay -- OneNote
        # luôn bị "trắng tay" bất kể có bao nhiêu kết quả khớp thật. Sửa:
        # xen kẽ (round-robin) từng phần tử Outlook/OneNote thay vì lấy
        # hết loại này rồi mới tới loại kia, để cả 2 cùng cạnh tranh công
        # bằng cho từng slot còn trống -- đảm bảo OneNote (và Outlook)
        # luôn giữ được ít nhất slot tối thiểu của mình
        # (_outlook_min_reserve/_onenote_min_reserve đã trừ sẵn ở
        # _normal_budget) trước khi loại kia được phép lấy thêm slot thứ
        # 2/3.
        _ol_reserve_list = outlook_paths[:_mail_reserve]
        _on_reserve_list = onenote_paths[:_mail_reserve]
        for _i in range(max(len(_ol_reserve_list), len(_on_reserve_list))):
            if _i < len(_ol_reserve_list):
                _add_ctx(_ol_reserve_list[_i])                     # 3) mail -- xen kẽ, không còn ăn hết trước
            if _i < len(_on_reserve_list):
                _add_ctx(_on_reserve_list[_i])                     # 4) notes -- xen kẽ, không còn bị "trắng tay"
        for p in all_paths:                                         # 5) fill any remaining budget (rare -- only if 1-4 undershot)
            _add_ctx(p)

        # v-fix (Vấn đề 2 -- AI phân tích NHẦM file khi user hỏi lại theo số
        # [N] cũ, vd "phân tích kỹ hơn về Nguồn [5]"): context_paths ở trên
        # được xếp lại MỖI LƯỢT theo matched_paths (keyword của đúng tin
        # nhắn NÀY) -- nên [5] của lượt này gần như chắc chắn KHÔNG phải
        # cùng 1 file với [5] mà user nhìn thấy ở lượt trước. Nếu tin nhắn
        # hiện tại có nhắc số [N] nào đó đã từng được gán (self.
        # _chat_citation_paths, xem bên dưới), ép file thật của đúng số đó
        # vào context lượt này (giống cách "+" Add file luôn được ưu tiên)
        # để model thực sự đọc đúng file user đang hỏi, thay vì đoán/lấy
        # nhầm file khác đang ngồi ở vị trí đó lượt này.
        _referenced_paths = [
            self._chat_citation_paths[n - 1] for n in sorted(_extract_citation_ids(user_message or ""))
            if 1 <= n <= len(self._chat_citation_paths)
        ]
        for _rp in reversed(_referenced_paths):
            if _rp in context_paths:
                context_paths.remove(_rp)
            if len(context_paths) >= _max_ctx_files and context_paths:
                context_paths.pop()
            context_paths.insert(0, _rp)

        # v-new (nút "+" Add file): file người dùng tự đính kèm LUÔN được
        # đưa vào context, ở vị trí [1] -- không phụ thuộc thứ hạng BM25
        # hay có nằm trong budget top-N hay không. Nếu context đã đầy,
        # nhường chỗ bằng cách bỏ bớt slot cuối (ít liên quan nhất) thay
        # vì vượt _max_ctx_files (tránh lại đúng lỗi 413 Groq đã từng gặp).
        _attached_path = None
        _af = getattr(self, "_chat_attached_file", None)
        if _af and _af.get("path"):
            _attached_path = _af["path"]
            if _attached_path in context_paths:
                context_paths.remove(_attached_path)
            if len(context_paths) >= _max_ctx_files and context_paths:
                context_paths.pop()
            context_paths.insert(0, _attached_path)

        # v-fix (yêu cầu 4 -- nút 🌐 không cần phân tích lại nguồn cục bộ):
        # bỏ hẳn context cục bộ vừa chọn ở trên cho ĐÚNG lượt bấm 🌐 -- lượt
        # này không cấp số [N] cho file nào cả (để bảng số self.
        # _chat_citation_paths của các lượt chat cục bộ khác trong cùng
        # phiên không bị xáo trộn), và cũng không đọc/gửi kèm nội dung file
        # nào cho model (real_ctx_paths/_fetch_content_snippets phía dưới
        # tự động rỗng theo). Model chỉ cần tự tra cứu internet -- xem
        # CHAT_WEB_ONLY_SYSTEM_PROMPT, thay hẳn cho system_prompt bên dưới.
        if force_web_search:
            context_paths = []

        # v-fix (Vấn đề 2): mở rộng bảng số [N] CỐ ĐỊNH của cả phiên chat --
        # file NÀO đã có số rồi thì GIỮ NGUYÊN số đó (không đổi theo lượt),
        # chỉ file mới toanh trong lượt này mới được cấp số tiếp theo. Nhờ
        # vậy [5] luôn là cùng 1 file trong toàn bộ phiên hỏi-đáp, kể cả khi
        # thứ tự ưu tiên nội bộ (matched_paths) đảo lộn khác nhau mỗi lượt.
        for _p in context_paths:
            if _p not in self._chat_citation_paths:
                self._chat_citation_paths.append(_p)

        _lang = CHAT_LANG_OPTIONS.get(getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_LANG_OPTIONS[CHAT_DEFAULT_LANG])

        # v-fix (source Outlook/OneNote luôn "không trích xuất được nội
        # dung"): _fetch_content_snippets only reads content_store, which is
        # only ever populated for REAL files indexed off disk -- Outlook/
        # OneNote pseudo-paths are never written there. Their real content
        # already came back from search_outlook()/search_onenote() at
        # search time though (the "snippet" field, cached in
        # self._outlook_meta / self._onenote_meta by _merge_mail_notes_async)
        # -- use that instead of hitting content_store for these two.
        if use_ai_search_context:
            _snip_chars = GPTOSS_SNIPPET_CHARS_DETAILED if model_choice != "gemini" else ONLINE_AI_SNIPPET_CHARS_DETAILED
        else:
            _snip_chars = GPTOSS_SNIPPET_CHARS if model_choice != "gemini" else ONLINE_AI_SNIPPET_CHARS
        real_ctx_paths = [p for p in context_paths
                           if not (_is_outlook_pseudo_path(p) or _is_onenote_pseudo_path(p))]
        real_snip_map = {s["path"]: s for s in _fetch_content_snippets(real_ctx_paths, keywords, _snip_chars)} \
                         if real_ctx_paths else {}

        # v-new (nút "+" Add file): file đính kèm thường KHÔNG nằm trong
        # content_store (chỉ được index nếu nó tình cờ cũng nằm trong 1
        # thư mục đã quét) -- _fetch_content_snippets ở trên sẽ không tìm
        # ra nội dung thật của nó. Dùng thẳng text đã extract sẵn lúc bấm
        # "+" (self._chat_attached_file["text"]) thay vì đợi/để trống.
        if _attached_path and _af:
            real_snip_map[_attached_path] = {"path": _attached_path, "snippet": _af["text"][:_snip_chars]}

        def _pseudo_snippet(p):
            if _is_outlook_pseudo_path(p):
                meta = self._outlook_meta.get(p, {})
                body = (meta.get("snippet") or "").strip()
                text = (f"[Outlook mail] Subject: {meta.get('subject','')}\n"
                        f"From: {meta.get('sender','')}  Received: {meta.get('received','')}\n"
                        f"Folder: {meta.get('folder_path','')}\n{body}")
                return {"path": p, "snippet": text.strip() or _lang["no_extract"]}
            if _is_onenote_pseudo_path(p):
                meta = self._onenote_meta.get(p, {})
                body = (meta.get("snippet") or "").strip()
                text = (f"[OneNote page] {meta.get('section_path','')} > {meta.get('title','')}\n{body}")
                return {"path": p, "snippet": text.strip() or _lang["no_extract"]}
            return {"path": p, "snippet": _lang["no_extract"]}

        def _cite_num(p):
            # v-fix (Vấn đề 2): số [N] gửi cho model (và sau này in ra
            # "Nguồn:") giờ LUÔN tra theo self._chat_citation_paths (bảng
            # số cố định của cả phiên, vừa cập nhật ở trên) -- KHÔNG dùng
            # vị trí cục bộ trong context_paths của riêng lượt này nữa.
            try:
                return self._chat_citation_paths.index(p) + 1
            except ValueError:
                return 0  # không nên xảy ra -- mọi context_paths đã được add ở trên rồi

        def _ctx_header(s):
            p = s["path"]
            n = _cite_num(p)
            if _is_outlook_pseudo_path(p) or _is_onenote_pseudo_path(p):
                # No real basename/path worth showing the model -- the raw
                # token is machine-only (entry_id/store_id/page_id) and the
                # model would otherwise just quote it verbatim back at the
                # user. Give it the same nice label the UI shows instead.
                return f"[{n}] {self._display_label_for_source(p)}"
            # v-new: đánh dấu rõ đây là file user tự đính kèm (nút "+"),
            # không phải 1 kết quả BM25 -- giúp model hiểu tại sao [1] có
            # vẻ không liên quan tới truy vấn tìm kiếm ban đầu.
            prefix = "📎 (User-attached file) " if p == _attached_path else ""
            return f"[{n}] {prefix}{os.path.basename(p)}  ({p})"

        snippets = [real_snip_map.get(p) or _pseudo_snippet(p) for p in context_paths]
        if snippets:
            context_block = "\n\n".join(
                f"{_ctx_header(s)}\n"
                f"{s['snippet'] or _lang['no_extract']}"
                for s in snippets
            )
        else:
            context_block = _lang["no_context"]

        system_prompt = _lang["system_prompt"] + f"\n\nCác đoạn trích từ file tìm được / File excerpts found / 見つかったファイルの抜粋:\n\n{context_block}"

        # v-new (theo yêu cầu: "14 sources dù model đang trả lời là GPT-OSS
        # -- 14 là cap của Gemini"): khi mid-turn rate-limit fallback (bên
        # dưới) đổi model_choice SANG "gptoss", context/system_prompt ở
        # trên vẫn được build với cap 14 file của Gemini (đúng model TẠI
        # THỜI ĐIỂM build) -- gửi nguyên context đó cho GPT-OSS thì số
        # "nguồn" hiển thị (14) không khớp với model thực sự vừa trả lời
        # (GPT-OSS chỉ nên có tối đa 5). Hàm này build lại system_prompt
        # với ĐÚNG cap của model mới, dùng lại snippets đã fetch sẵn (theo
        # đúng thứ tự ưu tiên cũ) -- không cần fetch lại nội dung file.
        def _rebuild_system_prompt_for(mc):
            cap = GPTOSS_MAX_CONTEXT_FILES if mc != "gemini" else ONLINE_AI_MAX_CONTEXT_FILES
            if len(snippets) <= cap:
                return system_prompt  # already within the new model's budget, nothing to trim
            trimmed = snippets[:cap]
            blk = "\n\n".join(
                f"{_ctx_header(s)}\n{s['snippet'] or _lang['no_extract']}"
                for s in trimmed
            ) if trimmed else _lang["no_context"]
            return _lang["system_prompt"] + f"\n\nCác đoạn trích từ file tìm được / File excerpts found / 見つかったファイルの抜粋:\n\n{blk}"

        # v-fix (Vấn đề 2): only append the "you may use internet
        # knowledge" instruction when the active model can actually back
        # that up with a real web-search tool -- Gemini's google_search
        # grounding, or Groq's browser_search (verified: works for
        # "gptoss", see _call_online_ai_chat). "qwen" has neither -- it
        # stays on the strict local-files-only prompt, and the reply gets
        # an explicit CHAT_NO_WEB_SEARCH_NOTE appended below so the user
        # isn't left wondering why "Online" had no effect. (model_choice
        # and wanted_web_search_but_unavailable were already decided
        # further up, BEFORE context was built -- see the v-fix note
        # there.)
        if force_web_search:
            # v-fix (yêu cầu 4): thay HẲN system_prompt bằng bản riêng chỉ
            # cho lượt tra cứu internet -- không nối thêm vào system_prompt
            # cục bộ/CHAT_ONLINE_SUFFIX (vốn ép model viết thêm 1 mục
            # "Thông tin bổ sung (internet)" bọc ngoài phần phân tích cục bộ
            # -- không cần thiết nữa vì lượt này vốn dĩ không có phần phân
            # tích cục bộ nào để "bổ sung" thêm). context_block ở trên chắc
            # chắn rỗng (context_paths đã bị xoá ngay phía trên) nên bỏ qua
            # luôn, không lãng phí trong prompt gửi đi.
            system_prompt = CHAT_WEB_ONLY_SYSTEM_PROMPT.get(
                getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_WEB_ONLY_SYSTEM_PROMPT["EN"])
        elif web_search and model_choice in WEB_SEARCH_CAPABLE_MODELS:
            system_prompt += CHAT_ONLINE_SUFFIX.get(
                getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_ONLINE_SUFFIX["EN"])

        # v-new (yêu cầu 4 -- dán ảnh vào ô chat): nếu file đính kèm hiện
        # tại là 1 ảnh (self._chat_attached_file["is_image"], xem
        # _attach_image_bytes), chỉ Gemini thực sự "nhìn" được nó (đọc
        # inline_data trong _call_online_ai_chat) -- Groq (GPT-OSS/Qwen)
        # không hỗ trợ vision cho 2 model text-only này, nên thêm 1 dòng
        # ghi chú vào system prompt để model không bịa nội dung ảnh, và
        # user cũng không thắc mắc vì sao câu trả lời "không thấy ảnh".
        _img_payload = None
        _af_for_img = getattr(self, "_chat_attached_file", None)
        if _af_for_img and _af_for_img.get("is_image") and _af_for_img.get("image_bytes"):
            if model_choice == "gemini":
                _img_payload = {"bytes": _af_for_img["image_bytes"], "mime": _af_for_img.get("image_mime", "image/png")}
            else:
                system_prompt += CHAT_IMAGE_NOT_SUPPORTED_NOTE.get(
                    getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_IMAGE_NOT_SUPPORTED_NOTE["EN"])

        # v-fix (Vấn đề 2 -- lỗi 413 "Request too large" trên Groq): the
        # full running self._chat_history keeps growing every turn and, on
        # Gemini, easily fits -- but Groq's 8000-TOKENS-PER-MINUTE cap for
        # gpt-oss-120b/qwen3.6-27b is far smaller and this (system prompt +
        # entire history + new message) regularly blew past it after just a
        # few turns. Cap how much history is actually sent to Groq to the
        # most recent turns only when a Groq model is being called; Gemini
        # calls keep sending the full history as before.
        history = list(self._chat_history)
        if model_choice != "gemini" and len(history) > GPTOSS_MAX_HISTORY_MSGS:
            history = history[-GPTOSS_MAX_HISTORY_MSGS:]

        # v-new (streaming -- hiệu ứng gõ chữ dần): 2 closure nhỏ hop text
        # từ thread nền (đang chạy _call_online_ai_chat) sang main thread
        # qua root.after(0, ...) -- Tk không thread-safe, mọi thao tác lên
        # Text widget phải chạy trên main thread. _is_current() đảm bảo
        # không đổ chữ vào 1 phiên chat đã cũ (user đã New Chat / đổi từ
        # khoá) trong lúc câu trả lời vẫn đang chạy dở.
        def _stream_cb(_t):
            if _is_current():
                self.root.after(0, lambda _t=_t: self._append_chat_stream_chunk(_t))

        def _reset_cb():
            if _is_current():
                self.root.after(0, self._reset_chat_stream)

        if _is_current():
            self.root.after(0, self._start_chat_stream)

        answer, err, web_sources, used_web_search = _call_online_ai_chat(
            model_choice, system_prompt, history, user_message, web_search=web_search, image=_img_payload,
            on_chunk=_stream_cb, on_reset=_reset_cb)
        # v-fix (yêu cầu: 429 grounding -> tự retry KHÔNG grounding thay vì
        # nhảy sang GPT-OSS): _call_online_ai_chat (nhánh "gemini") giờ tự
        # âm thầm thử lại KHÔNG có tool google_search ngay bên trong nó khi
        # gặp 429 grounding, và chỉ thật sự trả lỗi (err) nếu CẢ 2 lượt
        # (có + không tool) đều fail. Khi lượt "không tool" thành công,
        # used_web_search=False dù web_search=True -- người dùng vẫn nhận
        # được câu trả lời (không tự chuyển hẳn sang GPT-OSS), chỉ là lượt
        # này không có nguồn internet thật, nên coi như "muốn web-search
        # nhưng không dùng được" để hiển thị đúng ghi chú bên dưới.
        if web_search and not used_web_search and not err:
            wanted_web_search_but_unavailable = True

        # Automatically switch models on rate-limit, retrying RIGHT AWAY
        # within this same chat turn (no need for the user to press send
        # again). v-note: if Online mode forced Gemini/GPT-OSS above,
        # falling back to a non-web-search model here loses web grounding
        # for this one retry -- acceptable trade-off vs. failing the turn
        # outright.
        #
        # v-fix (Vấn đề: khi GPT-OSS đang là model active và bị 429 --Groq
        # TPD hết-- lỗi trả thẳng về cho user "⚠️ Lỗi GPT-OSS 120B (Groq):
        # 429..." thay vì tự chuyển sang model khác): điều kiện cũ chỉ
        # check `model_choice == "gemini"`, tức là fallback chỉ chạy MỘT
        # CHIỀU Gemini->GPT-OSS. Giờ dùng _next_alt_model() để tự xoay
        # vòng qua CẢ 3 model theo MODEL_FALLBACK_ORDER (Gemini -> GPT-OSS
        # -> Qwen), thay vì chỉ đổi 1 lần giữa đúng 2 model cố định --
        # vòng lặp dừng khi thành công, hết model để thử, hoặc lỗi không
        # còn là rate-limit nữa.
        print(f"[Online Chat][Debug] first attempt model={model_choice!r} err={err!r}")
        _tried_models = {model_choice}
        while err and _is_online_ai_rate_limit_err(err):
            # v-new: remember model_choice is exhausted RIGHT NOW so every
            # later message this session starts on a working model instead
            # of re-hitting this one and failing again -- see
            # _mark_model_rate_limited / _model_usable.
            self._mark_model_rate_limited(model_choice, True)
            alt_model = self._next_alt_model(model_choice)
            if not alt_model or alt_model in _tried_models:
                break  # no other usable/untried model left in the chain
            _log_online_ai_fallback(
                f"model={model_choice!r} -> switching to {alt_model!r}\nraw error: {err}")
            print(f"[Online Chat][Debug] rate-limit detected -> trying alt_model={alt_model!r}")
            model_choice = alt_model
            _tried_models.add(model_choice)
            self._chat_active_model = model_choice
            self._sync_chat_model_picker()
            if web_search and model_choice not in WEB_SEARCH_CAPABLE_MODELS:
                wanted_web_search_but_unavailable = True
            system_prompt = _rebuild_system_prompt_for(model_choice)  # v-new: re-cap context to match the model that will actually answer
            if model_choice != "gemini" and len(history) > GPTOSS_MAX_HISTORY_MSGS:
                history = history[-GPTOSS_MAX_HISTORY_MSGS:]  # v-new: also re-cap history on fallback, same reason
            # v-new (yêu cầu 4): model vừa đổi (fallback) -- tính lại xem
            # ảnh có còn gửi được không (chỉ Gemini đọc được), và nếu
            # KHÔNG, thêm ghi chú "không đọc được ảnh" vào system_prompt
            # VỪA rebuild ở trên cho model mới.
            if _af_for_img and _af_for_img.get("is_image") and _af_for_img.get("image_bytes"):
                if model_choice == "gemini":
                    _img_payload = {"bytes": _af_for_img["image_bytes"], "mime": _af_for_img.get("image_mime", "image/png")}
                else:
                    _img_payload = None
                    system_prompt += CHAT_IMAGE_NOT_SUPPORTED_NOTE.get(
                        getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_IMAGE_NOT_SUPPORTED_NOTE["EN"])
            answer, err, web_sources, used_web_search = _call_online_ai_chat(
                model_choice, system_prompt, history, user_message,
                web_search=web_search if model_choice in WEB_SEARCH_CAPABLE_MODELS else False,
                image=_img_payload, on_chunk=_stream_cb, on_reset=_reset_cb)
            if web_search and model_choice in WEB_SEARCH_CAPABLE_MODELS and not used_web_search and not err:
                wanted_web_search_but_unavailable = True
            print(f"[Online Chat][Debug] fallback attempt model={model_choice!r} err={err!r}")

        # v-fix: cả 3 model đều vừa báo rate-limit (hoặc không còn model
        # nào khác để thử) -- đây chính là lúc mà trước giờ user phải tự
        # vào Settings bấm 'Update API' -> Save để hết lỗi (xem
        # _reset_online_ai_clients). Tự làm việc đó ngay tại đây: xoá
        # client SDK cache rồi thử lại đúng model vừa fail LẦN CUỐI trước
        # khi bó tay báo lỗi cho user.
        if err and _is_online_ai_rate_limit_err(err):
            # v-new: mark this one rate-limited too (it may be model_choice
            # unchanged from above if no alt was usable, or the last alt
            # that itself just failed) -- every model is now known
            # exhausted until a call to one of them succeeds again.
            self._mark_model_rate_limited(model_choice, True)
            _log_online_ai_fallback(
                f"ALL models now rate-limited (model={model_choice!r} just failed too)\nraw error: {err}")
            print(f"[Online Chat][Debug] all models rate-limited -> "
                  f"resetting cached clients, retrying model={model_choice!r} once more")
            _reset_online_ai_clients()
            answer, err, web_sources, used_web_search = _call_online_ai_chat(
                model_choice, system_prompt, history, user_message,
                web_search=web_search if model_choice in WEB_SEARCH_CAPABLE_MODELS else False,
                image=_img_payload, on_chunk=_stream_cb, on_reset=_reset_cb)
            if web_search and model_choice in WEB_SEARCH_CAPABLE_MODELS and not used_web_search and not err:
                wanted_web_search_but_unavailable = True
            print(f"[Online Chat][Debug] post-reset retry model={model_choice!r} err={err!r}")
            if not err:
                # the reset+retry actually worked -- this model has quota
                # again after all (e.g. a transient client-cache glitch,
                # not a real quota exhaustion), so un-mark it.
                self._mark_model_rate_limited(model_choice, False)

        if err:
            # v-new: light up 🔑 Update API for anything the user needs to
            # act on -- missing key, quota/rate-limit (429), or "request too
            # large" (413/TPM, e.g. the Groq 8000-token error) -- but NOT
            # for a one-off transient network error, which isn't fixed by
            # updating a key.
            err_l = str(err).lower()
            if (_is_online_ai_rate_limit_err(err) or "413" in err_l
                    or "too large" in err_l or "thiếu gemini_api_key" in err_l
                    or "thiếu groq_api_key" in err_l):
                self._mark_api_needs_attention(True)
            self.root.after(0, lambda: (
                (self._append_chat_line("assistant", f"⚠️ {err}", True) if _is_current() else None),
                _reenable()))
            return
        self._mark_api_needs_attention(False)  # v-new: this call succeeded -- clear any earlier warning
        # v-new: model_choice is whichever model actually produced this
        # answer -- a real success means its quota is no longer exhausted,
        # so un-mark it (harmless no-op if it was never marked).
        self._mark_model_rate_limited(model_choice, False)

        # v1.9: strip any citation/source list the AI wrote on its own
        # (best-effort, in case it ignores the system-prompt instruction
        # above) so only ONE, consistently-formatted source list (with
        # full paths) ever shows up -- previously the AI's own free-text
        # list (names only, no path) and Python's own list (names + path)
        # could both appear, in 2 different formats, on different turns.
        answer = self._strip_ai_own_citation_list(answer)
        # v-new (theo yêu cầu: mục "Thông tin bổ sung (internet)" nên nằm
        # DƯỚI mục "Nguồn:" nội bộ, thay vì ở trên): CHAT_ONLINE_SUFFIX
        # giờ ép model luôn viết mục này với tiêu đề CHÍNH XÁC ở CUỐI câu
        # trả lời -- tách nó ra khỏi answer TRƯỚC KHI nối "Nguồn:" nội bộ
        # (_append_citation_filenames, bên dưới), rồi nối lại nó SAU
        # "Nguồn:" -- xem _split_supplementary_section().
        answer, supplementary = self._split_supplementary_section(
            answer, getattr(self, "_chat_lang", CHAT_DEFAULT_LANG))
        # v1.6: deterministically append a source list mapping every [N]
        # citation actually used in the answer to its real filename, built
        # straight from the file list Python already sent as context -- the
        # AI doesn't always remember to spell this out itself, so previously
        # the file names for [1]~[8] showed up inconsistently depending on
        # the answer. See _append_citation_filenames(). v-new: label is now
        # language-aware ("Nguồn"/"Source"/"出典").
        # v-fix (Vấn đề 2): dùng self._chat_citation_paths (bảng [N] cố
        # định của cả phiên) thay vì `snippets` (chỉ chứa file của RIÊNG
        # lượt này) -- để nếu answer lỡ trích 1 số [N] từ lượt TRƯỚC (dù
        # hiếm, model đôi khi tự nhắc lại), dòng "Nguồn:" vẫn map đúng
        # file, không lỗi/mất dòng.
        answer = self._append_citation_filenames(answer, self._chat_citation_paths, _lang["source_label"])
        # v-new: re-attach the "Thông tin bổ sung (internet)" section
        # (analysis + its own sources, written by the AI) NOW, after the
        # internal "Nguồn:" list, instead of wherever the AI originally
        # placed it in its raw output.
        if supplementary:
            answer = answer.rstrip() + "\n\n" + supplementary
        # v-new: if Gemini's web-search grounding actually kicked in for
        # this turn (e.g. the "phần mềm đối thủ khác" case -- info NOT in
        # the local files), list the real internet page(s) it drew from
        # right under the "Thông tin bổ sung (internet)" section above (or,
        # if the AI didn't write that section itself this turn, under its
        # own small heading) so it's visually distinct from local file
        # citations, marked 🌐.
        answer = self._append_web_sources(answer, web_sources, _lang["source_label"],
                                           has_own_heading=bool(supplementary))
        # v-fix (Vấn đề 2): tell the user plainly when Online mode couldn't
        # actually run a real internet search this turn (GPT-OSS/Groq has
        # no grounding tool) instead of silently answering local-only while
        # implying otherwise.
        if wanted_web_search_but_unavailable:
            answer += CHAT_NO_WEB_SEARCH_NOTE.get(
                getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_NO_WEB_SEARCH_NOTE["EN"])

        # v-new (yêu cầu: label "🤖 Bấm vào đây để xem thêm thông tin từ AI
        # Search offline" -- NẰM PHÍA TRÊN link "🌐 ..." bên dưới): chỉ
        # thêm dòng này cho đúng LƯỢT PHÂN TÍCH CỤC BỘ (force_web_search
        # False) và khi label này CHƯA được dùng cho câu hỏi hiện tại
        # (self._chat_ai_search_link_used). Trạng thái enable/grey của
        # label được quyết định NGAY TẠI ĐÂY, lúc sinh câu trả lời: '1'
        # (bấm được, xanh) nếu 🤖 AI Search (nút ngoài) đã chạy XONG cho
        # ĐÚNG keyword hiện tại (self._ai_mode_active + self._ai_active_query
        # khớp `query` + đã có self._ai_cont_res) -- ngược lại '0' (xám,
        # không bấm được) nếu user chưa từng bấm 🤖 AI Search lần nào.
        if not force_web_search and not getattr(self, "_chat_ai_search_link_used", False):
            _ai_search_ready = bool(
                getattr(self, "_ai_mode_active", False)
                and getattr(self, "_ai_active_query", None) == query
                and getattr(self, "_ai_cont_res", None))
            _ai_link_text = CHAT_AI_SEARCH_LINK_TEXT.get(
                getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_AI_SEARCH_LINK_TEXT["EN"])
            _state_ch = "1" if _ai_search_ready else "0"
            answer = answer.rstrip() + f"\n\n{_CHAT_AISEARCH_LINK_MARKER}{_state_ch}{_ai_link_text}"

        # v-new (yêu cầu: thêm 1 dòng text (có link) ở CUỐI phần phân tích
        # cục bộ, để user bấm thẳng vào đó thay vì phải bấm nút 🌐 riêng):
        # chỉ thêm dòng này cho đúng LƯỢT PHÂN TÍCH CỤC BỘ (force_web_search
        # False -- xem docstring _online_chat_worker: mọi lượt chat thường
        # giờ luôn chạy cục bộ trước), model đang active THỰC SỰ có khả
        # năng tìm internet thật (WEB_SEARCH_CAPABLE_MODELS -- Qwen thì
        # không, thêm link cũng vô nghĩa), và nút 🌐 CHƯA được dùng cho
        # câu hỏi này (self._chat_web_search_used) -- giống hệt điều kiện
        # nút 🌐 tự greyout, để không mời bấm 1 việc đã làm rồi. Dòng này
        # được bọc bởi _CHAT_INTERNET_LINK_MARKER (ký tự vô hình, không ai
        # nhìn thấy) để _insert_chat_text_with_links nhận diện và biến
        # thành link click-được, gọi lại đúng hàm _send_web_search_chat_msg
        # -- xem chỗ render.
        if (not force_web_search and model_choice in WEB_SEARCH_CAPABLE_MODELS
                and not getattr(self, "_chat_web_search_used", False)):
            _link_text = CHAT_INTERNET_LINK_TEXT.get(
                getattr(self, "_chat_lang", CHAT_DEFAULT_LANG), CHAT_INTERNET_LINK_TEXT["EN"])
            answer = answer.rstrip() + f"\n\n{_CHAT_INTERNET_LINK_MARKER}{_link_text}"

        # Always persist under the session/query this request actually
        # belongs to (captured at spawn time above) -- never under
        # whatever self._chat_session_id/_chat_auto_sent_for happen to be
        # NOW, several seconds later.
        self._save_chat_hist("user", user_message, session_id, query)
        self._save_chat_hist("assistant", answer or "", session_id, query)
        # v-fix: push the just-saved reply into ChatHistoryPanel right away
        # instead of waiting for its own 5s auto-refresh tick (see
        # _notify_chat_history_panel_for).
        self._notify_chat_history_panel_for(query)
        # v-new: refresh ◀/▶ nav so the just-saved session becomes
        # reachable -- force_live only if this reply belongs to the
        # keyword still active right now (query == entry text), so it
        # doesn't yank someone currently browsing an older keyword's
        # preview back to a reply for a keyword they've since left.
        try:
            if self.entry_var.get().strip() == query:
                self._refresh_chat_preview_nav(query, force_live=(self._chat_preview_idx == -1))
        except Exception:
            pass

        if _is_current():
            self._chat_history.append({"role": "user", "text": user_message})
            self._chat_history.append({"role": "assistant", "text": answer or ""})

        def _done():
            if _is_current():
                self._finish_chat_stream(answer or _lang["no_answer"])
            _reenable()
        self.root.after(0, _done)

    def _strip_ai_own_citation_list(self, answer):
        """Best-effort: remove any source/citation list the AI generated on
        its own at the end of the answer (e.g. a "Nguồn tài liệu trích
        dẫn:" heading followed by "* [1] `file.pptx`" bullets, names only,
        no path) so only Python's own deterministic "Nguồn:" list (always
        with full paths -- see _append_citation_filenames) shows up,
        instead of 2 differently-formatted source sections stacking up."""
        if not answer:
            return answer
        lines = answer.split("\n")
        bullet_re = re.compile(r"^\s*[\*\-•]?\s*[`\"]?\[\d+[\d,;\-–~\s]*\]")
        heading_re = re.compile(r"(nguồn|trích dẫn|reference|citation|source|出典|参考文献|参考)", re.IGNORECASE)
        cut = None
        i = len(lines) - 1
        while i >= 0:
            line = lines[i]
            if bullet_re.match(line):
                cut = i
                i -= 1
                continue
            if not line.strip():
                i -= 1
                continue
            if heading_re.search(line) and cut is not None:
                cut = i
            break
        if cut is not None:
            lines = lines[:cut]
        return "\n".join(lines).rstrip()

    def _append_web_sources(self, answer, web_sources, source_label="Nguồn", has_own_heading=False):
        """v-new: append real internet URLs Gemini's web-search grounding
        actually used for this turn (see _extract_web_sources), marked 🌐
        so they read as distinct from the local-file [N] citations added
        by _append_citation_filenames just above. No-op if grounding
        wasn't used / found nothing (most local-file-only answers).

        has_own_heading (v-new): True when the AI's own "Thông tin bổ sung
        (internet)" section (see _split_supplementary_section) was already
        re-attached right above this -- in that case these extra grounded
        links are just additional bullets under THAT heading, no second
        "Nguồn (internet):" heading needed."""
        if not answer or not web_sources:
            return answer
        lines = [f"🌐 {s['title']}\n    {s['url']}" for s in web_sources]
        if has_own_heading or answer.rstrip().endswith(":") or f"\n{source_label}:\n" in answer:
            return answer.rstrip() + "\n" + "\n".join(lines)
        return answer.rstrip() + f"\n\n{source_label} (internet):\n" + "\n".join(lines)

    def _display_label_for_source(self, path):
        """v-fix (Vấn đề: "Nguồn:" trong AI Chat vẫn hiện pseudo-path xấu
        dù đã tra self._outlook_meta/_onenote_meta): that first fix was
        still wrong -- _merge_mail_notes_async (which populates those meta
        dicts) is capped at only ~6s before AI Chat gives up waiting and
        moves on (see _auto_ai_chat_after_search), but the merge itself
        was regularly taking MINUTES, so the meta dict essentially never
        had the entry ready in time and every citation kept silently
        falling back to the ugly raw string -- same symptom, same root
        cause, just one layer deeper than the first fix reached.

        Real fix: OneNote's section path AND title are already embedded
        DIRECTLY inside the pseudo path string itself (see
        _make_onenote_pseudo_path: "ONENOTE::{page_id}::{section}::
        {title}.one") -- parsing them back out is pure string splitting,
        instant, and needs NO cache/merge/COM call at all, ever. Outlook's
        pseudo path didn't carry its folder path the same way (see
        _make_outlook_pseudo_path), so that's fixed too, below, to embed
        it the exact same way OneNote already does -- both are now fast
        AND correctly formatted unconditionally, with the old meta-dict
        lookup kept only as a defensive fallback for stale pseudo paths
        saved before this change (e.g. from AI Chat History)."""
        if _is_outlook_pseudo_path(path):
            try:
                _, _entry_id, _store_id, folder_and_subject = path.split("::", 3)
            except Exception:
                folder_and_subject = None
            if folder_and_subject and "::" in folder_and_subject:
                folder_path, subject_msg = folder_and_subject.split("::", 1)
                subject = subject_msg[:-4] if subject_msg.lower().endswith(".msg") else subject_msg
                return f"Outlook > {folder_path} > {subject}.msg" if folder_path else f"Outlook > {subject}.msg"
            meta = self._outlook_meta.get(path)  # fallback: old-format pseudo path (pre-fix) or meta cache
            if meta:
                subject = (meta.get("subject") or "(no subject)").strip() or "(no subject)"
                folder_path = meta.get("folder_path") or ""
                return f"Outlook > {folder_path} > {subject}.msg" if folder_path else f"Outlook > {subject}.msg"
            return "Outlook > (email)"  # last-resort fallback -- never show the raw OUTLOOK::... token
        if _is_onenote_pseudo_path(path):
            try:
                _, _page_id, section_path, title_one = path.split("::", 3)
            except Exception:
                section_path, title_one = None, None
            if title_one is not None:
                title = title_one[:-4] if title_one.lower().endswith(".one") else title_one
                return f"OneNote > {section_path} > {title}.one" if section_path else f"OneNote > {title}.one"
            meta = self._onenote_meta.get(path)  # fallback: old-format pseudo path / meta cache
            if meta:
                title = (meta.get("title") or "(untitled)").strip() or "(untitled)"
                section_path = meta.get("section_path") or ""
                return f"OneNote > {section_path} > {title}.one" if section_path else f"OneNote > {title}.one"
            return "OneNote > (page)"  # last-resort fallback -- never show the raw ONENOTE::... token
        return os.path.basename(path)

    def _split_supplementary_section(self, answer, lang_key):
        """v-new (theo yêu cầu: mục "Thông tin bổ sung (internet)" nên nằm
        DƯỚI "Nguồn:" nội bộ): CHAT_ONLINE_SUFFIX ép model luôn đặt mục
        này ở CUỐI câu trả lời, với tiêu đề CHÍNH XÁC theo ngôn ngữ đang
        dùng. Tìm dòng chứa tiêu đề đó, cắt answer làm 2 phần tại ĐẦU
        dòng đó: (phần trước, phần từ tiêu đề trở đi). Trả về
        (answer_không_có_mục_đó, mục_đó_hoặc_None) -- best-effort, không
        raise nếu không tìm thấy (model không phải lúc nào cũng viết mục
        này, ví dụ khi Online mode tắt hoặc câu hỏi không cần thông tin
        ngoài file cục bộ)."""
        if not answer:
            return answer, None
        phrase_map = {
            "VI": "Thông tin bổ sung (internet)",
            "EN": "Additional information (internet)",
            "JP": "追加情報",
        }
        phrase = phrase_map.get(lang_key, phrase_map["EN"])
        idx = answer.find(phrase)
        if idx == -1:
            return answer, None
        line_start = answer.rfind("\n", 0, idx)
        line_start = line_start + 1 if line_start != -1 else 0
        before = answer[:line_start].rstrip()
        supplementary = answer[line_start:].strip()
        if not supplementary:
            return answer, None
        return before, supplementary

    def _append_citation_filenames(self, answer, citation_paths, source_label="Nguồn"):
        """Append a "Nguồn:" list mapping every [N] citation marker actually
        used in the answer to its real filename + full path.

        v-fix (Vấn đề 2): citation_paths is now self._chat_citation_paths --
        the STABLE [N]->path table built up over the WHOLE chat session
        (see _online_chat_worker) -- instead of the old `snippets` list,
        which only held THIS turn's re-ranked context and made [N] mean a
        different file almost every follow-up message.

        v1.8 fix: the AI doesn't always cite one number per bracket like
        "[4] [5]" -- it sometimes groups them as "[4, 5]" or even a range
        like "[4-6]", which the old single-digit regex didn't match at
        all, so those citations silently got no "Nguồn:" entry. Now every
        bracket group is parsed for however many numbers (and ranges) it
        contains."""
        if not answer or not citation_paths:
            return answer
        # v-fix (Vấn đề: mục "Nguồn:"/"Source:" đôi khi biến mất hoàn
        # toàn): dùng _extract_citation_ids() dùng chung thay vì tự parse
        # ASCII-only ở đây -- xem docstring của hàm đó. Trước đây khi
        # Gemini trả lời chủ yếu bằng tiếng Nhật và tự chuyển sang trích
        # dẫn kiểu 【N】 thay vì [N] (thường do bị ảnh hưởng bởi cách trích
        # các thuật ngữ tiếng Nhật khác ngay trong cùng câu trả lời), regex
        # cũ không khớp -> cited rỗng -> return answer nguyên văn, KHÔNG
        # thấy path/link nguồn đâu cả dù answer rõ ràng có trích dẫn.
        cited = sorted(n for n in _extract_citation_ids(answer) if 1 <= n <= len(citation_paths))
        if not cited:
            return answer
        def _cite_line(n):
            p = citation_paths[n - 1]
            nice = self._display_label_for_source(p)
            # v-fix (Vấn đề: Nguồn Outlook/OneNote không click mở file
            # được, trong khi PDF/PPT thì ok): trước đây pseudo-path
            # (OUTLOOK::.../ONENOTE::...) bị CỐ TÌNH bỏ khỏi dòng này vì
            # hiện nó ra rất xấu -- nhưng làm vậy cũng xoá luôn phần dữ
            # liệu mà _insert_chat_text_with_links/_insert_text_with_
            # source_links cần để tách "label — path" và biến path thành
            # link click-được (_open_path_for_chat). Không có "— path",
            # dòng Outlook/OneNote không match regex citation nữa -> mất
            # link, y hệt bug đang gặp. Giờ LUÔN nối "— {p}" cho mọi loại
            # nguồn như file thật -- pseudo-path xấu vẫn được giữ ở đây để
            # có data click, nhưng sẽ bị ẨN khỏi hiển thị ở bước render
            # (xem _insert_chat_text_with_links / _insert_text_with_
            # source_links), chỉ hiện label thân thiện "nice" có gạch
            # dưới, y hệt cách PDF/PPT vẫn luôn hoạt động.
            return f"[{n}] {nice}  —  {p}"
        lines = [_cite_line(n) for n in cited]
        return answer.rstrip() + f"\n\n{source_label}:\n" + "\n".join(lines)

    def _load_next_chat_session_id(self):
        """Return 1 + the highest session_id already saved in
        HISTORY_DB_FILE (across ALL previous app runs), or 1 if the table
        doesn't exist/is empty yet. Called once at startup so this run's
        session_id counter can never collide with a session_id used in a
        past run -- see the comment where this is assigned in __init__."""
        try:
            if not os.path.exists(HISTORY_DB_FILE):
                return 1
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='chat_history'")
            if c.fetchone()[0] == 0:
                conn.close()
                return 1
            c.execute("SELECT COALESCE(MAX(session_id), 0) FROM chat_history")
            max_id = c.fetchone()[0] or 0
            conn.close()
            return int(max_id) + 1
        except Exception as _e:
            print(f"[Chat History] _load_next_chat_session_id FAILED, defaulting to 1: {_e}")
            return 1

    def _find_chat_sessions_for_query(self, query_text):
        """Same lookup as ChatHistoryPanel._find_sessions_for_query, kept
        as its own copy here so the live AI Chat panel's ◀/▶ preview nav
        (see _refresh_chat_preview_nav) doesn't need a reference into that
        separate widget. Returns session_ids for query_text, oldest first."""
        q_norm = (query_text or "").strip().lower()
        if not q_norm or not os.path.exists(HISTORY_DB_FILE):
            return []
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='chat_history'")
            if c.fetchone()[0] == 0:
                conn.close()
                return []
            c.execute("SELECT session_id, MIN(id) AS first_id FROM chat_history "
                      "WHERE lower(trim(query))=? GROUP BY session_id ORDER BY first_id ASC", (q_norm,))
            rows = [r[0] for r in c.fetchall()]
            conn.close()
            return rows
        except Exception:
            return []

    def _load_chat_session_rows(self, session_id, query_text):
        """Read-only fetch of one saved session's (role, text) turns, for
        ◀/▶ preview rendering -- never touches self._chat_history."""
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("SELECT role, text FROM chat_history WHERE session_id=? AND lower(trim(query))=? ORDER BY id ASC",
                      (session_id, (query_text or "").strip().lower()))
            rows = c.fetchall()
            conn.close()
            return rows
        except Exception:
            return []

    def _refresh_chat_preview_nav(self, query, force_live=True):
        """Recompute which OLDER sessions exist for `query` (excluding the
        current live self._chat_session_id) and update the ◀/▶ button
        states/label. Called whenever the active keyword changes, whenever
        a new reply gets saved, and after "New Chat". If force_live, also
        snap the nav position back to -1 (live) -- used when the keyword
        itself just changed or a fresh New Chat was started, so the user
        isn't left "stuck" browsing an old page that no longer matches
        what's on screen; left False for a same-query reply landing, so it
        doesn't yank someone currently browsing an older page back to live."""
        try:
            sessions = self._find_chat_sessions_for_query(query)
            live_id = getattr(self, "_chat_session_id", None)
            older = [sid for sid in sessions if sid != live_id]
            self._chat_preview_query = query
            self._chat_preview_sessions = older
            if force_live or self._chat_preview_idx >= len(older):
                self._chat_preview_idx = -1
                self._set_chat_preview_readonly(False)  # in case a previous preview page left input disabled
            self._update_chat_preview_nav_ui()
        except Exception as _e:
            print(f"[Chat Preview] refresh error: {_e}")

    def _update_chat_preview_nav_ui(self):
        try:
            older = self._chat_preview_sessions
            idx = self._chat_preview_idx
            total = len(older) + 1
            page = total if idx == -1 else idx + 1
            btn_prev = getattr(self, "_chat_nav_prev_btn", None)
            btn_next = getattr(self, "_chat_nav_next_btn", None)
            lbl = getattr(self, "_chat_nav_lbl", None)
            can_go_older = bool(older) and (idx == -1 or idx > 0)
            can_go_newer = idx != -1
            if btn_prev is not None and btn_prev.winfo_exists():
                btn_prev.config(state="normal" if can_go_older else "disabled",
                                 fg="#1a5fb4" if can_go_older else "#c3c6cb")
            if btn_next is not None and btn_next.winfo_exists():
                btn_next.config(state="normal" if can_go_newer else "disabled",
                                 fg="#1a5fb4" if can_go_newer else "#c3c6cb")
            if lbl is not None and lbl.winfo_exists():
                if len(older) == 0:
                    lbl.config(text="")
                else:
                    lbl.config(text=f"{page}/{total}")
        except Exception:
            pass

    def _chat_preview_prev(self):
        """◀ -- step to an OLDER saved session for the current keyword."""
        older = self._chat_preview_sessions
        if not older:
            return
        idx = self._chat_preview_idx
        if idx == -1:
            idx = len(older) - 1
        elif idx > 0:
            idx -= 1
        else:
            return  # already at the oldest
        self._chat_preview_idx = idx
        self._render_chat_preview()

    def _chat_preview_next(self):
        """▶ -- step to a NEWER session, ending back at the live chat."""
        idx = self._chat_preview_idx
        if idx == -1:
            return  # already live
        older = self._chat_preview_sessions
        idx = idx + 1 if idx < len(older) - 1 else -1
        self._chat_preview_idx = idx
        self._render_chat_preview()

    def _render_chat_preview(self):
        """Redraw self._chat_txt for whatever _chat_preview_idx currently
        points at -- either the live, still-updating conversation (-1) or
        a read-only snapshot of an older saved session for the same
        keyword, loaded straight from HISTORY_DB_FILE via
        _load_chat_session_rows (never touches self._chat_history, so
        browsing old pages can't corrupt or lose the live conversation).

        v-fix (theo yêu cầu): input box/nút gửi KHÔNG còn bị khóa khi đang
        xem lại 1 session cũ nữa -- vẫn là cùng 1 từ khóa, nên gửi tiếp
        tin nhắn trong lúc xem lại là hợp lý. Xem _send_online_chat_msg:
        gửi tin nhắn trong lúc đang xem trang cũ sẽ tự quay lại "live"
        trước, rồi mới gửi tiếp vào đó."""
        try:
            idx = self._chat_preview_idx
            if idx == -1:
                self._redraw_chat_panel_from_history()
            else:
                session_id = self._chat_preview_sessions[idx]
                rows = self._load_chat_session_rows(session_id, self._chat_preview_query)
                txt = self._chat_txt
                txt.config(state="normal")
                txt.delete("1.0", "end")
                txt.config(state="disabled")
                for role, text in rows:
                    self._append_chat_line(role, text)
            self._update_chat_preview_nav_ui()
        except Exception as _e:
            print(f"[Chat Preview] render error: {_e}")

    def _set_chat_preview_readonly(self, readonly):
        """v-fix (theo yêu cầu): không còn khóa input/nút gửi khi xem lại
        session cũ (readonly=True is now a no-op, kept only so callers that
        pass readonly=False to re-enable things after an older build's
        lock still work harmlessly)."""
        if readonly:
            return
        try:
            if getattr(self, "_chat_entry", None) is not None and self._chat_entry.winfo_exists():
                self._chat_entry.config(state="normal")
            if getattr(self, "_chat_send_btn", None) is not None and self._chat_send_btn.winfo_exists():
                self._set_send_btn_enabled(True)
            self._set_new_chat_btn_enabled(True)
            if getattr(self, "_chat_placeholder", None) is not None and self._chat_placeholder.winfo_exists():
                self._chat_placeholder.config(text=self._current_chat_placeholder_text())
        except Exception:
            pass

    def _new_online_chat(self):
        """'New Chat' button -- clears the conversation history, resets to
        the default Gemini Flash model for the next chat turn. v1.9: also
        starts a new "session" id, so the messages before/after this point
        are grouped as separate entries in AI Chat History (ChatHistoryPanel).

        v1.11 FIX: this used to also reset self._chat_auto_sent_for to
        None -- meaning that after the user manually clicked "New Chat" (as
        opposed to this being called internally because a genuinely
        different search query came in, see _auto_ai_chat_after_search),
        the chat lost track of which search keyword it belonged to. Any
        further messages then got saved to chat_history with an empty
        query, so they never showed up under that keyword's "AI Chat
        History" row -- they were saved but effectively orphaned. Per
        request: "1 search keyword = 1 AI Chat History entry, unless the
        user hits New Chat while on that keyword, in which case that also
        counts as 1 more entry (still under the same keyword)". So we now
        deliberately KEEP self._chat_auto_sent_for pointing at the current
        keyword across a manual New Chat -- only the NEW session_id changes,
        which is what actually splits it into a separate saved entry.

        v-fix (Vấn đề 2): also force busy-state back to idle here. Now that
        a stale in-flight reply from the OLD session is no longer allowed
        to touch busy-state for the NEW session (see _online_chat_worker),
        clicking "New Chat" while a request was still running would
        otherwise leave the Send button disabled forever -- nothing else
        would ever clear it for this new session."""
        self._chat_history = []
        self._chat_citation_paths = []   # v-fix (Vấn đề 2): numbering [N] phải reset theo phiên mới, không kéo từ phiên cũ sang
        # v-new (yêu cầu: link 🌐/🤖 chỉ cần bấm 1 lần/câu hỏi): phiên chat
        # mới = câu hỏi mới, nên reset lại cả 2 cờ để 2 link lại dùng được
        # bình thường (grey/enable của chúng nằm ngay trong nội dung do
        # _online_chat_worker sinh ra, không còn 1 widget riêng để gọi lại).
        self._chat_web_search_used = False
        self._chat_ai_search_link_used = False
        self._chat_ai_search_grey_tags = []  # v-new: phiên mới -- bỏ theo dõi các tag xám của phiên cũ
        # v-fix (New Chat luôn reset về Gemini): giờ giữ đúng model người
        # dùng đã chọn ở dropdown header (self._chat_preferred_model) thay
        # vì luôn quay lại Gemini Flash mặc định.
        self._chat_active_model = getattr(self, "_chat_preferred_model", DEFAULT_ONLINE_MODEL)
        self._sync_chat_model_picker()
        # v-fix (yêu cầu 1, lượt 2 -- Gemini vẫn bị bỏ qua, chuyển thẳng
        # sang GPT-OSS cho MỌI keyword mới, dù KHÔNG hề có lỗi 429 nào
        # trong log lần này): nguyên nhân thật sự nằm ở đây, không phải ở
        # API key nữa. self._chat_model_rate_limited (xem
        # _mark_model_rate_limited/_model_usable) là 1 cờ "dính" (sticky)
        # cho CẢ PHIÊN LÀM VIỆC CỦA APP, không phải cho riêng 1 keyword --
        # nếu Gemini từng 429 dù chỉ 1 lần trên BẤT KỲ keyword nào trước
        # đó, cờ này set True và KHÔNG hề được xoá ở đây trước giờ --
        # khiến _model_usable("gemini") trả False cho MỌI keyword sau đó
        # trong suốt phiên app đang chạy, kể cả khi _chat_active_model vừa
        # được reset về "gemini" ở dòng ngay trên: _online_chat_worker vẫn
        # âm thầm đổi sang GPT-OSS ở bước _model_usable() TRƯỚC KHI kịp
        # thử gọi Gemini -- không hề có request Gemini nào được gửi đi,
        # nên không có dòng 429 nào trong log để thấy.
        # v4.0 KHÔNG có cơ chế "dính" theo cả phiên app này -- nó chỉ nhớ
        # model đang dùng của ĐÚNG 1 CUỘC HỘI THOẠI hiện tại (tự reset về
        # Gemini mỗi khi bắt đầu 1 keyword mới, y hệt dòng trên), nên mỗi
        # keyword mới luôn được thử lại Gemini từ đầu -- đúng theo đề xuất
        # của user. Xoá cờ rate-limited của cả 3 model ở đây để mô phỏng
        # lại đúng hành vi đó: BẮT ĐẦU 1 KEYWORD MỚI = quên hết lịch sử
        # rate-limit của các keyword TRƯỚC, luôn cho Gemini 1 cơ hội mới.
        # (Nếu Gemini vẫn thật sự hết quota, nó sẽ tự 429 lại ngay trong
        # LƯỢT NÀY và cơ chế fallback-trong-1-lượt ở _online_chat_worker
        # vẫn xử lý bình thường như trước -- chỉ khác là giờ nó THỰC SỰ
        # THỬ chứ không bỏ qua từ đầu.)
        for _mk in ("gemini", "gptoss", "qwen"):
            self._mark_model_rate_limited(_mk, False)
        self._chat_session_id = getattr(self, "_chat_session_id", 0) + 1
        self._set_chat_busy(False)  # re-enable Send even if an old request was still running
        # v-new: New Chat always lands back on "live" -- the session_id
        # just changed, so whatever the ◀/▶ nav was pointing at before is
        # no longer meaningful.
        try:
            self._refresh_chat_preview_nav(getattr(self, "_chat_auto_sent_for", "") or "", force_live=True)
        except Exception:
            pass
        try:
            self._chat_txt.config(state="normal")
            self._chat_txt.delete("1.0", "end")
            self._chat_txt.config(state="disabled")
            self._chat_entry.delete("1.0", "end")
            self._toggle_chat_placeholder()
            self._autosize_chat_entry()
            self._chat_placeholder.config(text=self._current_chat_placeholder_text())
            self._toggle_chat_placeholder()
            self._safe_focus_chat_entry()
        except Exception:
            pass
        # v-new (nút "+" Add file): "New Chat" cũng gỡ luôn file đính kèm
        # (nếu có) -- cuộc trò chuyện mới không nên vô tình vẫn còn nhét
        # nội dung 1 file từ cuộc trò chuyện đã bị xoá vào context.
        self._chat_attached_file = None
        self._update_attach_chip_ui()

    def _smart_search_realtime(self, q, version, box_x, box_y, box_h, adv_mode=False):
        try:
            if len(q) < 2: return
            cleaned_q, op, size_val = parse_size_filter(q)
            kw = [k.lower() for k in _split_query_tokens(cleaned_q) if k]
            # v5.8: same stopword strip as the live MFT scan -- this is a
            # SEPARATE search path (DB-backed) that feeds _last_bm25_file_res,
            # which the "Default" pane switches to as soon as AI Search /
            # Advanced mode is toggled on. Without this, "to" (from "relevant
            # to HILS") still slipped into the OR-fallback LIKE clause below
            # and matched every ".toc" file the moment AI Search was clicked,
            # even though the live-scan path (used before AI Search) was
            # already fixed.
            kw = _strip_stopwords(kw)

            # Use persistent connection — no open/close overhead per keystroke
            conn = self.db_conn
            if conn is None:
                return   # DB not ready yet (indexing in progress)
            try:
                conn.execute("SELECT 1")  # quick liveness check
            except Exception:
                self.db_conn = None
                return   # connection closed (indexing just started)
            c = conn.cursor()
            
            if not kw and op is not None:
                file_sql = f"SELECT type, name, path, size FROM files WHERE size {op} ? LIMIT 1000"
                c.execute(file_sql, (size_val,))
                file_res = c.fetchall()
                cont_res = []
            elif not kw:
                conn.close(); return
            else:
                # name-only AND: all keywords must appear in the file/folder name
                # Only use keywords >= 2 chars for name matching to avoid false positives like 'a'
                size_clause = f" AND size {op} ?" if op is not None else ""
                size_extra  = [size_val] if op is not None else []
                # Name/folder search: require at least one keyword >= 3 chars (anchor).
                # Short keywords (1-2 chars) are included in LIKE only when anchored by a long keyword.
                # e.g. "simpack a" → ok (simpack anchors), "a b" → skip (no anchor).
                has_anchor = any(_is_anchor_kw(k) for k in kw)
                if has_anchor:
                    name_conds  = ["name LIKE ?" for k in kw]
                    name_params = [f"%{k}%" for k in kw]
                    file_name_sql = ("SELECT type, name, path, size FROM files WHERE type != 'Folder' AND "
                                + " AND ".join(name_conds) + size_clause + " LIMIT 1000")
                    c.execute(file_name_sql, name_params + size_extra)
                    file_name_res = c.fetchall()

                    # ── OR fallback: if AND returns nothing, try OR so partial matches surface ──
                    if not file_name_res and len(kw) > 1:
                        or_conds  = ["name LIKE ?" for k in kw]
                        or_params = [f"%{k}%" for k in kw]
                        file_name_sql_or = ("SELECT type, name, path, size FROM files WHERE type != 'Folder' AND ("
                                    + " OR ".join(or_conds) + ")" + size_clause + " LIMIT 1000")
                        c.execute(file_name_sql_or, or_params + size_extra)
                        file_name_res = c.fetchall()
                    # v-fix (Vấn đề: "鉄道不整" [query CHỈ 1 từ khoá, ghép
                    # liền "鉄道"+"不整" không dấu cách] ra 0 kết quả ở SQL
                    # DB-indexed search dù "鉄道"/"不整" riêng lẻ có kết
                    # quả): khi kw chỉ có ĐÚNG 1 token, fallback OR ở trên
                    # (chỉ chạy khi len(kw)>1) không hề chạy -- và bản thân
                    # SQL "name LIKE '%鉄道不整%'" đòi hỏi đúng 4 ký tự
                    # dính liền đó trong tên, không tìm thấy nếu tên chỉ
                    # chứa 2 từ ở 2 chỗ khác nhau. Thử mọi điểm cắt hợp lệ
                    # tách token thành 2 nửa, OR các cặp "cả 2 nửa cùng
                    # LIKE" lại -- cùng ý tưởng _cjk_glued_or_query (vốn
                    # dùng cho FTS5/File Content) nhưng viết lại bằng SQL
                    # LIKE cho tìm theo tên.
                    if not file_name_res and len(kw) == 1 and len(kw[0]) >= 4 and not any(ch.isascii() for ch in kw[0]):
                        _tok = kw[0]
                        _n = len(_tok)
                        _cuts = list(range(2, _n - 1))
                        if len(_cuts) > 12:
                            _step = len(_cuts) / 12
                            _cuts = sorted({_cuts[int(i * _step)] for i in range(12)} | {_n // 2})
                        if _cuts:
                            _split_conds = ["(name LIKE ? AND name LIKE ?)" for _ in _cuts]
                            _split_params = []
                            for _c in _cuts:
                                _split_params.extend([f"%{_tok[:_c]}%", f"%{_tok[_c:]}%"])
                            file_name_sql_split = ("SELECT type, name, path, size FROM files WHERE type != 'Folder' AND ("
                                        + " OR ".join(_split_conds) + ")" + size_clause + " LIMIT 1000")
                            c.execute(file_name_sql_split, _split_params + size_extra)
                            file_name_res = c.fetchall()
                    # ─────────────────────────────────────────────────────────────────────────────

                    folder_name_sql = ("SELECT type, name, path, size FROM files WHERE type = 'Folder' AND "
                                + " AND ".join(name_conds) + size_clause + " LIMIT 1000")
                    c.execute(folder_name_sql, name_params + size_extra)
                    folder_name_res = c.fetchall()
                    # v-fix: cùng fallback split-token như file_name_sql ở
                    # trên, áp dụng cho folder (bản gốc chưa từng có fallback
                    # nào cho folder, kể cả OR-fallback nhiều từ khoá).
                    if not folder_name_res and len(kw) == 1 and len(kw[0]) >= 4 and not any(ch.isascii() for ch in kw[0]):
                        _tok = kw[0]
                        _n = len(_tok)
                        _cuts = list(range(2, _n - 1))
                        if len(_cuts) > 12:
                            _step = len(_cuts) / 12
                            _cuts = sorted({_cuts[int(i * _step)] for i in range(12)} | {_n // 2})
                        if _cuts:
                            _split_conds = ["(name LIKE ? AND name LIKE ?)" for _ in _cuts]
                            _split_params = []
                            for _c in _cuts:
                                _split_params.extend([f"%{_tok[:_c]}%", f"%{_tok[_c:]}%"])
                            folder_name_sql_split = ("SELECT type, name, path, size FROM files WHERE type = 'Folder' AND ("
                                        + " OR ".join(_split_conds) + ")" + size_clause + " LIMIT 1000")
                            c.execute(folder_name_sql_split, _split_params + size_extra)
                            folder_name_res = c.fetchall()

                    # v7.10: "Whole word" toggle -- SQL LIKE has no notion of
                    # word boundaries, so fetch the (broader) substring hits
                    # above as usual, then post-filter in Python when the
                    # toggle is on. This drops buried-substring hits like
                    # "readasync.xml" for "adas" while keeping "ADAS_systems.pdf".
                    if self.whole_word_var.get():
                        file_name_res   = [r for r in file_name_res
                                            if all(_kw_matches_with_glued_fallback(k, r[1].lower(), True) for k in kw)]
                        folder_name_res = [r for r in folder_name_res
                                            if all(_kw_matches_with_glued_fallback(k, r[1].lower(), True) for k in kw)]

                    file_res = file_name_res + folder_name_res
                else:
                    file_res = []
                
                # ── Content search (v10.63) ──────────────────────────────────
                # SPEED STRATEGY for large DB (44GB):
                #
                # LIKE '%phrase%' on content_store = full table scan = slow on 44GB
                # FTS5 trigram MATCH = uses index = fast even on large DB
                #
                # New approach:
                #   1. Extract alphanum tokens from phrase (FTS-safe parts)
                #   2. FTS MATCH → get small candidate path set (fast, indexed)
                #   3. LIKE verify full phrase on candidate set only (fast, small set)
                #
                # This avoids scanning all 44GB for every keystroke.
                # LIKE is only run on the ~few hundred paths FTS returns, not millions.

                import re as _re

                _STOP_WORDS = {
                    'a','an','the','is','in','on','at','to','of','or','and','as',
                    'be','by','do','for','has','had','he','her','him','his',
                    'how','i','if','it','its','me','my','no','not','off',
                    'our','out','own','so','than','that','them','then',
                    'they','this','us','was','we','who','why','will','with',
                    'you','your','also','been','but','can','did','does','from',
                    'get','got','have','into','just','may','new','now','one',
                    'see','set','she','time','what','when','which','would',
                }

                def _is_fts_safe(k):
                    # NOTE: content_index uses SQLite's fts5 'trigram' tokenizer, which
                    # can only match terms >= 3 characters — SQLite silently returns ZERO
                    # rows (not an error) for shorter terms, even a valid 2-char CJK word
                    # like 解析. So this length check must stay at 3 regardless of script;
                    # short CJK anchors are instead routed to the LIKE fallback below via
                    # _is_anchor_kw(), which has no such length limit.
                    FTS5_OPS = set('+-*:^"()：、。・<>@[]{}|\\/?!#$%&=~`\'')
                    if len(k) < 3: return False
                    if any(ch in FTS5_OPS for ch in k): return False
                    if '-' in k: return False   # hyphen is NOT operator in FTS5
                    if '.' in k and not k.replace('.','').isdigit(): return False  # dots cause issues
                    if _re.search(r'[^\w\u3000-\u9fff\uff00-\uffef\u4e00-\u9fff]', k): return False
                    return True

                # ── Tunable display limits ────────────────────────────────────
                # v9.15: these used to branch on adv_mode ("Advanced" vs
                # "Realtime"), but adv_mode is ALWAYS False here — the
                # Advanced button doesn't actually call this function with
                # True; it runs its own separate one-shot full-search inline
                # in _on_advanced() and shows it in a split pane below the
                # normal results. So this branch was dead code, and the
                # "Realtime" numbers (300/100/0.05) were the only ones ever
                # in effect for every keystroke.
                # Per user testing, using the looser limits doesn't cost
                # noticeable speed, so they're now the default for everyone,
                # not just an opt-in via a second click. The Advanced button
                # is currently hidden (kept in the code — see
                # self._adv_search_btn — to re-show later if needed).
                FTS_LIMIT      = 5000  # full result set
                DISPLAY_LIMIT  = 2000  # show all
                BM25_THRESHOLD = 0.0   # no threshold, show everything
                # ─────────────────────────────────────────────────────────────

                def _fts_candidate_paths(cursor, tokens):
                    """Use FTS index to get candidate path set — fast even on 44GB."""
                    fts_tokens = [k for k in tokens
                                  if len(k) >= 3 and k.lower() not in _STOP_WORDS and _is_fts_safe(k)]
                    if not fts_tokens:
                        return None   # no FTS tokens available
                    fts_terms = " AND ".join(f'"{k}"' for k in fts_tokens)
                    try:
                        cursor.execute(
                            "SELECT path FROM content_index WHERE content MATCH ? LIMIT ?",
                            (fts_terms, FTS_LIMIT))
                        return set(r[0] for r in cursor.fetchall())
                    except Exception as _e:
                        print(f"FTS error [{fts_terms}]: {_e}")
                        return None

                def _cjk_glued_fallback(cursor, token):
                    """v-new: see module-level _cjk_glued_or_query() for the
                    full rationale (also used by the Outlook/OneNote merge
                    for the exact same reason). Builds one combined MATCH
                    query (OR of AND-pairs across every split point) instead
                    of looping N separate FTS queries."""
                    widened_q = _cjk_glued_or_query(token)
                    if not widened_q:
                        return set()
                    try:
                        cursor.execute(
                            "SELECT path FROM content_index WHERE content MATCH ? LIMIT ?",
                            (widened_q, FTS_LIMIT))
                        return set(r[0] for r in cursor.fetchall())
                    except Exception as _e:
                        print(f"FTS glued-fallback error [{widened_q}]: {_e}")
                        return set()

                def _cjk_short_glued_like_fallback(cursor, token):
                    """v-new (Vấn đề: "鉄道不整"/"鉄道量" — cụm CJK dính liền
                    DÀI 3-5 KÝ TỰ, 0 kết quả ở cả FTS chính lẫn
                    _cjk_glued_fallback): đã kiểm chứng bằng thực nghiệm rằng
                    FTS5 tokenizer='trigram' KHÔNG THỂ match bất kỳ chuỗi nào
                    ngắn hơn 3 ký tự trong câu MATCH — quan trọng là giới hạn
                    này áp dụng cho CHÍNH CÁI TERM ĐANG TRUY VẤN, không phải
                    nội dung được lập chỉ mục, nên dù nội dung có chứa đúng
                    chuỗi đó, quote 1 term <3 ký tự (vd "鉄道", "道") trong
                    MATCH vẫn LUÔN trả về 0 hàng. Do đó, cách "tách 2 nửa rồi
                    AND lại" của _cjk_glued_or_query CHỈ có cơ hội hoạt động
                    khi MỖI NỬA ĐỀU >=3 ký tự — tức token gốc phải dài >=6
                    ký tự (khớp đúng ngưỡng len(phrase_tokens[0]) >= 6 ở nơi
                    gọi _cjk_glued_fallback bên trên). Với token 3-5 ký tự,
                    MỌI cách tách đều để lại ít nhất 1 nửa <3 ký tự — không
                    thể sửa bằng cách đổi ngưỡng, đây là giới hạn cứng của
                    chính tokenizer 'trigram', không phải bug logic.
                    Fallback này né hẳn FTS MATCH cho các nửa ngắn, dùng LIKE
                    quét thẳng content_store (LIKE không có giới hạn độ dài
                    tối thiểu như trigram) — đồng thời NỚI LỎNG yêu cầu: chỉ
                    cần cả 2 nửa cùng xuất hiện Ở ĐÂU ĐÓ trong tài liệu (không
                    cần liền kề đúng điểm cắt như bản FTS), vì rất có thể lý
                    do gốc khiến "鉄道不整" không match được chính là vì 2 từ
                    đó không nằm liền nhau trong nội dung đã lưu (cách trích
                    xuất text từ file gốc chèn khoảng trắng/xuống dòng ở giữa
                    — không liên quan gì đến OCR, xem thảo luận trước).
                    Đánh đổi: đây là full-table-scan trên content_store,
                    CHẬM hơn hẳn so với path FTS có index — nhưng chỉ kích
                    hoạt khi: (1) FTS chính đã thất bại, VÀ (2) token đúng 1
                    từ CJK dài 3-5 ký tự -- một trường hợp hẹp, không phải
                    đường đi mặc định cho mọi query."""
                    n = len(token)
                    if n < 3 or n > 5:
                        return None
                    if any(ch in '"\\%_' for ch in token):
                        return None   # unsafe for a LIKE pattern
                    conds, params = [], []
                    for cpos in range(1, n):
                        left, right = token[:cpos], token[cpos:]
                        conds.append("(content LIKE ? AND content LIKE ?)")
                        params.extend([f"%{left}%", f"%{right}%"])
                    sql = "SELECT DISTINCT path FROM content_store WHERE " + " OR ".join(conds)
                    try:
                        cursor.execute(sql, params)
                        return set(r[0] for r in cursor.fetchall())
                    except Exception as _e:
                        print(f"[BM25] short-CJK LIKE fallback error [{token}]: {_e}")
                        return set()

                def _like_on_paths(cursor, phrase, paths_set):
                    """Run LIKE verify only on known candidate paths — avoids full scan."""
                    if not paths_set:
                        return []
                    ph = ",".join("?" * len(paths_set))
                    cursor.execute(
                        f"SELECT path FROM content_store WHERE path IN ({ph}) AND content LIKE ?",
                        list(paths_set) + [f"%{phrase}%"])
                    return set(r[0] for r in cursor.fetchall())

                def _paths_to_rows(cursor, paths_set):
                    if not paths_set: return []
                    ph = ",".join("?" * len(paths_set))
                    cursor.execute(
                        f"SELECT f.path, f.size FROM files f WHERE f.path IN ({ph})",
                        list(paths_set))
                    return cursor.fetchall()

                phrase = cleaned_q.strip()
                cont_res = []

                if _is_anchor_kw(phrase):
                    # Extract alphanum tokens from phrase for FTS narrowing
                    # v-fix: was missing CJK punctuation (、。！？「」『』（）) that
                    # _QUERY_SPLIT_RE already handles for File Name search. Without
                    # these, an ASCII word glued to a Japanese sentence across a
                    # 、 (e.g. "...サーバーは、FlexLM") became ONE token containing
                    # 、 — which is itself banned in _is_fts_safe()'s FTS5_OPS set,
                    # so the WHOLE token (including "FlexLM") got silently dropped
                    # from the search instead of being split into two usable pieces.
                    phrase_tokens = [k for k in _re.split(r'[\s、。！？「」『』（）:+\-<>@"()#.\[\]{}|\\/?!&=~`]+', phrase) if k]

                    # Step 1: FTS → candidate paths (indexed, fast)
                    candidates = _fts_candidate_paths(c, phrase_tokens)

                    if candidates is not None and len(candidates) > 0:
                        if len(phrase_tokens) == 1 and _is_fts_safe(phrase_tokens[0]):
                            # Single clean token → FTS result is already exact, no LIKE needed
                            cont_res = _paths_to_rows(c, candidates)
                        else:
                            # Multi-token or special chars → LIKE verify on candidate subset
                            verified = _like_on_paths(c, phrase, candidates)
                            if verified:
                                cont_res = _paths_to_rows(c, verified)
                            else:
                                # LIKE verify missed → FTS results are good enough
                                cont_res = _paths_to_rows(c, candidates)
                    elif candidates is not None and len(candidates) == 0:
                        # FTS returned nothing → for a single glued CJK
                        # token (no natural delimiter to split on), try the
                        # split-fallback above before giving up entirely —
                        # see _cjk_glued_fallback docstring.
                        if len(phrase_tokens) == 1 and len(phrase_tokens[0]) >= 6:
                            widened = _cjk_glued_fallback(c, phrase_tokens[0])
                            cont_res = _paths_to_rows(c, widened) if widened else []
                        elif (len(phrase_tokens) == 1 and 3 <= len(phrase_tokens[0]) < 6
                              and not any(ch.isascii() for ch in phrase_tokens[0])):
                            # v-new: cụm CJK dính liền 3-5 ký tự (vd "鉄道不整",
                            # "鉄道量") — quá ngắn để _cjk_glued_fallback (cần
                            # mỗi nửa >=3 ký tự tức token gốc >=6) có cơ hội
                            # match qua FTS trigram. Xem docstring
                            # _cjk_short_glued_like_fallback ở trên để biết lý
                            # do đây phải là 1 lượt LIKE full-scan riêng, không
                            # thể chỉ đơn giản hạ ngưỡng của lượt FTS phía trên.
                            _t_like0 = time.time()
                            widened = _cjk_short_glued_like_fallback(c, phrase_tokens[0])
                            print(f"[BM25] short-CJK LIKE fallback for "
                                  f"'{phrase_tokens[0]}' took "
                                  f"{time.time()-_t_like0:.2f}s, "
                                  f"{len(widened) if widened else 0} hits")
                            cont_res = _paths_to_rows(c, widened) if widened else []
                        else:
                            cont_res = []
                    else:
                        # No FTS tokens available (all special chars) → LIKE only on full table
                        # This is the slow path, only hits for queries like "+81" alone
                        useful_like = [k for k in phrase_tokens
                                       if _is_anchor_kw(k) and k.lower() not in _STOP_WORDS]
                        if useful_like:
                            best = max(useful_like, key=len)  # use longest token only
                            try:
                                c.execute(
                                    "SELECT path FROM content_store WHERE content LIKE ? LIMIT 3000",
                                    (f"%{best}%",))
                                fallback_paths = set(r[0] for r in c.fetchall())
                                if len(phrase_tokens) > 1:
                                    fallback_paths = _like_on_paths(c, phrase, fallback_paths)
                                cont_res = _paths_to_rows(c, fallback_paths)
                            except: pass
                
            # Do NOT close conn — it's the persistent db_conn, reused every search

            # ── BM25 ranking for File Name + Folder Name ────────────────────
            if file_res and len(cleaned_q.split()) >= 1:
                try:
                    from rank_bm25 import BM25Okapi
                    import re as _re_bm25fn
                    import time as _time_fn
                    def _tok_fn(text):
                        return [t.lower() for t in _re_bm25fn.split(r'[\s\W]+', str(text)) if len(t) >= 2]
                    q_tok_fn = _tok_fn(cleaned_q)
                    if q_tok_fn:
                        # v1.5: filename weighted 3x (more important than full path)
                        def _weighted_fn(path):
                            fname = os.path.basename(path)
                            return _tok_fn(f"{fname} {fname} {fname} {path}")
                        corpus_fn = [_weighted_fn(r[2]) for r in file_res]
                        bm25_fn = BM25Okapi(corpus_fn)
                        raw_fn = bm25_fn.get_scores(q_tok_fn)
                        max_fn = max(raw_fn) if max(raw_fn) > 0 else 1.0

                        # v1.5: extension boost for File Name tab
                        _EXT_BOOST_FN = {
                            '.pdf': 1.30, '.docx': 1.30, '.doc': 1.25,
                            '.xlsx': 1.20, '.xls': 1.15, '.csv': 1.10,
                            '.pptx': 1.20, '.ppt': 1.15, '.msg': 1.15,
                            '.one': 1.10, '.txt': 1.00, '.md': 1.00,
                            '.py': 0.90,  '.log': 0.75, '.ini': 0.70,
                        }
                        # v1.5: recency boost — newer files rank higher
                        _now_fn = _time_fn.time()
                        def _recency_fn(path):
                            try:
                                age_days = (_now_fn - os.path.getmtime(path)) / 86400
                                return 1.0 / (1.0 + age_days / 365)
                            except Exception:
                                return 0.5

                        def _final_fn(i, path):
                            bm25_norm = raw_fn[i] / max_fn
                            ext       = os.path.splitext(path)[1].lower()
                            ext_b     = _EXT_BOOST_FN.get(ext, 1.0)
                            rec_b     = _recency_fn(path)
                            # v9.7 fix: relevance must dominate ranking — file-type
                            # preference is now a small additive nudge (±0.03 max),
                            # not a multiplier directly on the relevance term. The
                            # old formula (bm25_norm * ext_b * 0.75) let a 30%
                            # extension boost outrank a genuinely more relevant
                            # result in a different format (e.g. a highly relevant
                            # .zip/.spck file losing to a weakly-relevant .pdf just
                            # because .pdf carries ext_b=1.30 vs .zip's default 1.0).
                            # ext_b's 0.70–1.30 range maps to a ±0.03 nudge here —
                            # relevant only as a tie-breaker between near-equal
                            # matches, never enough to flip a real relevance gap.
                            return bm25_norm * 0.90 + (ext_b - 1.0) * 0.10 + rec_b * 0.05

                        file_res = [file_res[i] for i in sorted(
                            range(len(file_res)),
                            key=lambda i: _final_fn(i, file_res[i][2]),
                            reverse=True)]
                except Exception:
                    pass  # keep original order if BM25 fails
            # v9.15: was `2000 if adv_mode else 100` — adv_mode is always False
            # here (see note above), so 100 was the only cap ever applied.
            # Bumped to the looser cap for everyone; see note above for why.
            _file_display_limit = 2000
            file_res = file_res[:_file_display_limit]
            # ────────────────────────────────────────────────────────────────

            # ── BM25 ranking for File Content ────────────────────────────────
            # Default: BM25 only (fast). AI Search button triggers hybrid later.
            sem_res = []
            bm25_scores = {}
            size_map = {}

            # ── Display limits ──────────────────────────────────────────────────────
            # v9.15: was branched on adv_mode (always False here — see note above);
            # now always the looser "show everything" values.
            DISPLAY_LIMIT  = 2000  # all results
            BM25_THRESHOLD = 0.0   # no threshold filtering
            # ─────────────────────────────────────────────────────────────────────────────

            if cont_res:
                try:
                    from rank_bm25 import BM25Okapi
                    import re as _re_bm25

                    def _tokenize(text):
                        return [t.lower() for t in _re_bm25.split(r"[\s\W]+", str(text)) if len(t) >= 2]

                    q_tokens = _tokenize(cleaned_q)
                    if q_tokens:
                        # ⚡ SPEED: rank by filename + 2 parent dirs — skip content DB fetch.
                        # FTS already guarantees file contains keyword → BM25 only re-ranks.
                        # ~20x faster than content-fetch, results not truncated.
                        paths_in_cont = [r[0] for r in cont_res]
                        size_map = {r[0]: r[1] for r in cont_res}

                        path_order = []
                        corpus = []
                        for path in paths_in_cont:
                            fname  = os.path.basename(path)
                            parent = os.path.basename(os.path.dirname(path))
                            gp     = os.path.basename(os.path.dirname(os.path.dirname(path)))
                            # v1.5: filename weighted 3x — most relevant signal for local search
                            text = f"{fname} {fname} {fname} {parent} {parent} {gp}"
                            path_order.append(path)
                            corpus.append(_tokenize(text))

                        bm25 = BM25Okapi(corpus)
                        raw_scores = bm25.get_scores(q_tokens)
                        max_s = max(raw_scores) if max(raw_scores) > 0 else 1.0

                        # v1.5: extension boost + recency boost for File Content tab
                        _EXT_BOOST_C = {
                            '.pdf': 1.30, '.docx': 1.30, '.doc': 1.25,
                            '.xlsx': 1.20, '.xls': 1.15, '.csv': 1.10,
                            '.pptx': 1.20, '.ppt': 1.15, '.msg': 1.15,
                            '.one': 1.10, '.txt': 1.00, '.md': 1.00,
                            '.py': 0.90,  '.log': 0.75, '.ini': 0.70,
                        }
                        import time as _time_c
                        _now_c = _time_c.time()
                        def _recency_c(path):
                            try:
                                age_days = (_now_c - os.path.getmtime(path)) / 86400
                                return 1.0 / (1.0 + age_days / 365)
                            except Exception:
                                return 0.5

                        bm25_scores = {}
                        for i, path in enumerate(path_order):
                            bm25_norm = raw_scores[i] / max_s
                            ext       = os.path.splitext(path)[1].lower()
                            ext_b     = _EXT_BOOST_C.get(ext, 1.0)
                            rec_b     = _recency_c(path)
                            # v9.7 fix: same as the File Name tab above — relevance
                            # must dominate, extension type is now a small additive
                            # nudge (±0.03 max) instead of a multiplier that could
                            # outrank a genuinely more relevant result in a
                            # different file format.
                            bm25_scores[path] = bm25_norm * 0.90 + (ext_b - 1.0) * 0.10 + rec_b * 0.05
                    else:
                        size_map = {r[0]: r[1] for r in cont_res}
                except Exception as _bm25_e:
                    print(f"[BM25] Error: {_bm25_e}")
                    size_map = {r[0]: r[1] for r in cont_res}

            # Sort by BM25 score → threshold → display cap
            all_paths = list(bm25_scores.keys()) if bm25_scores else [r[0] for r in cont_res]
            hybrid = []
            for path in all_paths:
                score = bm25_scores.get(path, 0.0)
                hybrid.append((path, size_map.get(path, 0), score))
            hybrid.sort(key=lambda x: x[2], reverse=True)
            max_h = hybrid[0][2] if hybrid else 1.0
            # Noise filter: drop results scoring below threshold relative to top result
            if max_h > 0 and bm25_scores:
                cutoff = max_h * BM25_THRESHOLD
                hybrid = [x for x in hybrid if x[2] >= cutoff]
            # Limit displayed rows to keep UI responsive
            hybrid = hybrid[:DISPLAY_LIMIT]
            cont_res = [(p, sz, int((sc / max_h) * 99) if max_h > 0 else 0)
                        for p, sz, sc in hybrid]
            # ────────────────────────────────────────────────────────────────

            # ── Outlook/OneNote mail+notes merge (v1.5) ──────────────────────
            # No longer run INLINE here -- COM automation calls to Outlook/
            # OneNote can be slow (sometimes much more than a second -- see
            # the known OneNote COM bitness issue), and used to block the
            # entire File Content tab render until both finished, while File
            # Name/Folder Name stayed fast because they're painted early by
            # the separate MFT scan thread. Now spawned as a background
            # thread AFTER the fast BM25-only render just below, so mail/
            # notes hits get appended in a second, later refresh once
            # they're ready instead of holding up the whole tab. See
            # _merge_mail_notes_async().
            # ─────────────────────────────────────────────────────────────

            if version == self.current_search_id:
                # Priority sort at source: office/pdf/msg first, txt/md middle,
                # log/html/code/... last. All places reusing this cache (restore,
                # Advanced reset, AI merge) receive pre-sorted data.
                cont_res  = self._sort_priority(cont_res, 0)
                file_res  = self._sort_priority(file_res, 1)
                # Cache BM25 results + query for later AI Search merge
                self._last_query = q
                self._last_bm25_cont_res = cont_res
                self._last_bm25_file_res = file_res
                # v7.7 FIX: remember that the DB-backed (comprehensive) result
                # set has been painted for this search id, and snapshot it, so
                # the live MFT scan (narrower scope, may still be running/
                # finish later) merges with it in _render_mft_file_tree /
                # _render_mft_folder_tree instead of overwriting it outright.
                self._db_rendered_sid = version
                self._db_rendered_file_res = list(file_res)
                # v5.4: only clear AI mode if this refresh is for a genuinely
                # DIFFERENT query than the one AI results are currently shown
                # for. Some UI interactions (e.g. selecting a model in the AI
                # dropdown) can end up re-triggering this realtime BM25 refresh
                # for the SAME query in the background -- that used to always
                # silently kick the view back to BM25 and reset the AI Search
                # button, even though the user never asked to leave AI mode.
                # v9.12 fix: once AI Search has been turned on, keep it
                # "sticky" across new keywords instead of silently falling
                # back to BM25 and making the user re-click "AI Search"
                # every time — the semantic model is already loaded/cached
                # in memory (_load_semantic_model() is a no-op if the right
                # model is already loaded), so re-running AI search for a
                # NEW query here is cheap, not a full reload.
                # _ai_search_and_update() TOGGLES: calling it while
                # _ai_mode_active is already True restores BM25 instead of
                # refreshing, so clear the flag first here to make it
                # re-engage for the new query instead of turning itself off.
                if self._ai_mode_active and q != getattr(self, "_ai_active_query", None):
                    self._ai_mode_active = False
                    self._ai_pending_query = q  # v-fix: tell update_or_show_results a sticky AI re-run for q is in flight
                    threading.Thread(target=self._ai_search_and_update, args=(q,), daemon=True).start()
                elif not (self._ai_mode_active and q == getattr(self, "_ai_active_query", None)):
                    self._ai_mode_active = False
                self._adv_mode_active = adv_mode
                # Update Advanced button label to reflect current mode
                def _update_adv_btn():
                    try:
                        if self._adv_search_btn and self._adv_search_btn.winfo_exists():
                            if adv_mode:
                                self._adv_search_btn.config(text="Simple ↩", fg="#90ee90", bg="#1a3a1a", state="normal")
                            else:
                                self._adv_search_btn.config(text="Advanced", fg="#33363c", bg="#e6e8ec", state="normal")
                    except Exception: pass
                self.root.after(0, _update_adv_btn)
                self.root.after(0, lambda: self.update_or_show_results(file_res, cont_res, q, box_x, box_y, box_h, version, sem_res))
                # v-new (theo yêu cầu: AI Chat luôn hoạt động mặc định,
                # không cần bấm "AI Search" trước): trước đây chỉ
                # _on_ai_search()/_open_ai_split_win() (nút AI Search) mới
                # tự show panel + gọi _auto_ai_chat_after_search -- nên nếu
                # user chưa từng bấm AI Search, panel chat luôn ẩn dù BM25
                # (File Content) đã có kết quả. AI Chat chỉ đọc
                # self._last_bm25_cont_res (vừa cache ở trên, cùng q) --
                # không phụ thuộc AI Search/offline embedding model chút
                # nào -- nên show + trigger summary ngay tại đây, cho MỌI
                # lượt search, không riêng gì lượt có AI Search bật.
                # _auto_ai_chat_after_search tự bỏ qua nếu query đã lỗi
                # thời (đã gõ thêm) hoặc đang busy, nên gọi ở mỗi debounce
                # là an toàn, không spam.
                self.root.after(0, self._show_online_chat_panel)
                self.root.after(150, lambda: self._auto_ai_chat_after_search(q))
                # Kick off the (potentially slow) Outlook/OneNote merge only
                # AFTER the fast render above has already been scheduled --
                # it appends mail/notes hits in a follow-up refresh once ready.
                if (OUTLOOK_SEARCH_AVAILABLE or ONENOTE_SEARCH_AVAILABLE) and kw:
                    # v1.10: flagged so _auto_ai_chat_after_search can wait
                    # for this to finish before reading self._last_bm25_cont_res
                    # -- otherwise the AI Chat auto-summary could fire against
                    # the fast-but-incomplete BM25-only result set and never
                    # see the mail/notes hits that show up moments later.
                    self._mail_notes_merge_pending = True
                    # v-fix: đẩy vào queue depth=1 thay vì spawn thread mới --
                    # nếu có job đang chờ (chưa kịp chạy) thì bị THAY THẾ bởi
                    # job mới nhất; nếu 1 job đang CHẠY thì job mới xếp hàng
                    # ngay sau, không chạy song song. Xem _mail_notes_worker_loop.
                    job = (kw, q, version, file_res, cont_res, box_x, box_y, box_h, sem_res)
                    try:
                        self._mail_notes_queue.get_nowait()  # drop stale pending job, if any
                    except queue.Empty:
                        pass
                    self._mail_notes_queue.put_nowait(job)
        except Exception as e:
            if "closed database" in str(e).lower() or "cannot operate" in str(e).lower():
                self.db_conn = None  # mark as closed, will reopen after indexing
            else:
                print(f"Search Error: {e}")

    def _mail_notes_worker_loop(self):
        """v-fix: SINGLE long-lived worker thread that drains
        self._mail_notes_queue and runs _merge_mail_notes_async jobs ONE AT
        A TIME, forever. This is what actually stops the Outlook/OneNote
        pile-up (see the comment where the queue is created in __init__):
        as long as this loop is busy on an old job, any newer jobs just sit
        in/replace the queue slot instead of spawning their own competing
        thread + sqlite connection against the 350MB search_outlook.db."""
        while True:
            job = self._mail_notes_queue.get()  # blocks until something is queued
            # Drain to the newest job in case more piled up while we were
            # busy with the previous one (queue is maxsize=1 so normally at
            # most one extra, but loop defensively anyway).
            while True:
                try:
                    job = self._mail_notes_queue.get_nowait()
                except queue.Empty:
                    break
            try:
                self._merge_mail_notes_async(*job)
            except Exception as e:
                print(f"[MailNotesWorker] job failed: {e}")

    def _merge_mail_notes_async(self, kw, q, version, file_res, cont_res_base,
                                 box_x, box_y, box_h, sem_res):
        """Runs on _mail_notes_worker_loop (never more than one at a time --
        see that method). Executes the Outlook + OneNote COM/sqlite searches
        and, if they turn up anything and the user hasn't already moved on
        to a different query, append them to the already-rendered File
        Content tab via a follow-up update_or_show_results call. Split out
        of _smart_search_realtime so slow COM automation never blocks the
        initial (BM25-only) render -- see the comment above its call site."""
        # v-fix: bail out immediately if the user has already moved on to a
        # newer query -- BEFORE paying for the (potentially 60s+) sqlite
        # calls below, not just after (the old code only checked `version`
        # once the query had already finished, so a stale job still fully
        # ran and blocked the worker for nothing).
        if version != self.current_search_id:
            return
        # v-new: timing instrumentation to find out WHERE the "vài phút cho
        # 30~50 file" delay actually is -- search_outlook()/search_onenote()
        # are pure sqlite FTS5 queries with no COM calls (see those modules'
        # own docstrings), so they SHOULD be sub-second even for large
        # mailboxes; these prints will show if that's actually true on your
        # machine, or if the time is going somewhere else (tree
        # insert/render, DB lock contention, etc). Safe/cheap to leave in.
        _t0 = time.time()
        try:
            mail_res = []
            if OUTLOOK_SEARCH_AVAILABLE and kw:
                try:
                    _t_o0 = time.time()
                    outlook_fts_q = " AND ".join(f'"{k}"' for k in kw if k)
                    o_results = outlook_search.search_outlook(outlook_fts_q, limit=200) if outlook_fts_q else []
                    # v-new: exact glued-token match found nothing — try the
                    # split-fallback (see _cjk_glued_or_query) before giving
                    # up, same reasoning/fix as the File Content search.
                    if not o_results and len(kw) == 1 and len(kw[0]) >= 6:
                        _widened_o_q = _cjk_glued_or_query(kw[0])
                        if _widened_o_q:
                            o_results = outlook_search.search_outlook(_widened_o_q, limit=200)
                    # v-fix (yêu cầu 1: query nhiều từ, vd "Structural
                    # scenario", luôn ra 0 kết quả Outlook dù có nhiều thông
                    # tin liên quan): outlook_fts_q ở trên dùng AND cứng --
                    # bắt buộc TỪNG từ khoá phải cùng xuất hiện chung 1 email
                    # thì mới tính là khớp. File Content (BM25 cục bộ) không
                    # khắt khe như vậy -- 1 file chỉ cần chứa MỘT PHẦN các từ
                    # khoá là đã được xếp hạng/hiện ra, nên với query nhiều
                    # từ, pdf/ppt/doc luôn có hàng chục kết quả trong khi
                    # Outlook/OneNote đứng yên ở 0, dù dữ liệu liên quan vẫn
                    # nằm rải rác trong nhiều email khác nhau (mỗi email chỉ
                    # chứa 1 trong 2 từ, không có email nào chứa ĐỦ CẢ 2).
                    # Fallback nới lỏng phía trên chỉ áp dụng cho ĐÚNG 1 từ
                    # khoá dính liền kiểu CJK -- không giúp gì cho trường hợp
                    # nhiều từ (tiếng Anh/Việt) này. Giờ nếu AND ra 0 kết quả
                    # VÀ có từ 2 từ khoá trở lên, thử lại với OR (khớp email
                    # nào chứa ít nhất 1 trong các từ khoá, xếp hạng theo
                    # bm25() nên email chứa NHIỀU từ khoá hơn vẫn lên đầu) --
                    # rộng hơn AND nhưng vẫn còn hơn hẳn im lặng bỏ qua toàn
                    # bộ Outlook chỉ vì không có email nào khớp tuyệt đối.
                    if not o_results and len(kw) > 1:
                        _or_o_q = " OR ".join(f'"{k}"' for k in kw if k)
                        o_results = outlook_search.search_outlook(_or_o_q, limit=200) if _or_o_q else []
                    print(f"[Timing][Outlook] search_outlook() took {time.time()-_t_o0:.2f}s, {len(o_results)} hits")
                    for i, r in enumerate(o_results):
                        pseudo_path = _make_outlook_pseudo_path(
                            r["entry_id"], r["store_id"], r["subject"], r.get("folder_path", ""))
                        self._outlook_meta[pseudo_path] = r
                        # Simple rank-based score (already best-first from
                        # search_outlook's ORDER BY bm25(...) ASC) rather than
                        # trying to numerically compare sqlite's bm25() scale
                        # against the file-content hybrid score above — the two
                        # corpora aren't comparable, and _sort_priority right
                        # below already keeps mail from outranking a clearly-
                        # more-relevant file match via its band logic.
                        mail_res.append((pseudo_path, r.get("size", 0), max(5, 95 - i)))
                except Exception as _oe:
                    print(f"[Outlook] search merge failed: {_oe}")
            if ONENOTE_SEARCH_AVAILABLE and kw:
                try:
                    _t_n0 = time.time()
                    onenote_fts_q = " AND ".join(f'"{k}"' for k in kw if k)
                    n_results = onenote_search.search_onenote(onenote_fts_q, limit=200) if onenote_fts_q else []
                    # v-new: same glued-token split-fallback as Outlook above.
                    if not n_results and len(kw) == 1 and len(kw[0]) >= 6:
                        _widened_n_q = _cjk_glued_or_query(kw[0])
                        if _widened_n_q:
                            n_results = onenote_search.search_onenote(_widened_n_q, limit=200)
                    # v-fix (yêu cầu 1): same AND-too-strict problem/fix as
                    # Outlook right above -- see the comment there.
                    if not n_results and len(kw) > 1:
                        _or_n_q = " OR ".join(f'"{k}"' for k in kw if k)
                        n_results = onenote_search.search_onenote(_or_n_q, limit=200) if _or_n_q else []
                    print(f"[Timing][OneNote] search_onenote() took {time.time()-_t_n0:.2f}s, {len(n_results)} hits")
                    for i, r in enumerate(n_results):
                        pseudo_path = _make_onenote_pseudo_path(r["page_id"], r["section_path"], r["title"])
                        self._onenote_meta[pseudo_path] = r
                        mail_res.append((pseudo_path, 0, max(5, 95 - i)))
                except Exception as _ne:
                    print(f"[OneNote] search merge failed: {_ne}")

            if not mail_res:
                return   # nothing to add -- the fast BM25-only render already stands
            if version != self.current_search_id:
                return   # user moved on to a different query -- discard stale results

            _t_sort0 = time.time()
            cont_res = self._sort_priority(list(cont_res_base) + mail_res, 0)
            self._last_bm25_cont_res = cont_res
            print(f"[Timing] _sort_priority merge took {time.time()-_t_sort0:.2f}s -- "
                  f"{len(mail_res)} mail/notes rows, total elapsed so far {time.time()-_t0:.2f}s")
            _t_ui0 = time.time()
            self.root.after(0, lambda: (
                self.update_or_show_results(
                    file_res, cont_res, q, box_x, box_y, box_h, version, sem_res),
                print(f"[Timing] update_or_show_results (tree rebuild) took "
                      f"{time.time()-_t_ui0:.2f}s, TOTAL merge->on-screen {time.time()-_t0:.2f}s")))
        finally:
            # v1.10: always clear the "pending" flag, even on error/early-
            # return/exception, so _auto_ai_chat_after_search never waits
            # forever for a merge that already finished (or never had
            # anything to add).
            self._mail_notes_merge_pending = False


    def update_or_show_results(self, file_res, cont_res, query, box_x, box_y, box_h, version, sem_res=None):
        if sem_res is None: sem_res = []
        if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win):
            self.show_results(file_res, cont_res, query, self.has_args, box_x, box_y, box_h, version, sem_res)
            return
            
        if version != self.current_search_id: return
        
        self.active_result_win.title(f"Results: {query}")
        self._update_chat_placeholder(query)
        
        self.tree_c.search_id = version
        self.tree_f.search_id = version
        self.tree_fol.search_id = version
        
        for item in self.tree_c.get_children(): self.tree_c.delete(item)
        for item in self.tree_f.get_children(): self.tree_f.delete(item)
        for item in self.tree_fol.get_children(): self.tree_fol.delete(item)
        
        # cont_res / file_res are already priority-sorted at source (search thread),
        # filter preserves relative order so only file/folder split is needed here.
        only_files = [r for r in file_res if str(r[0]).lower() != "folder"]
        only_folders = [r for r in file_res if str(r[0]).lower() == "folder"]
        # v7.7 FIX: the DB-backed search is comprehensive but only knows about
        # already-indexed files. Fold in any live MFT-scan matches (same
        # search id) for very recent files/folders the index doesn't have
        # yet, instead of silently dropping them when this DB render lands.
        if version == self.current_search_id:
            only_files   = self._merge_mft_with_db(self._mft_file_res or [], is_folder=False) \
                           if only_files or self._mft_file_res else only_files
            only_folders = self._merge_mft_with_db(self._mft_folder_res or [], is_folder=True) \
                           if only_folders or self._mft_folder_res else only_folders
        self._all_files_data = only_files
        self._all_content_data = cont_res
        self.size_op_var.set(">"); self.size_num_var.set(""); self.size_unit_var.set("MB")
        self.ext_filter_var.set("")
        self.c_size_op_var.set(">"); self.c_size_num_var.set(""); self.c_size_unit_var.set("MB")
        self.c_ext_filter_var.set("")
        self.name_filter_var.set("")
        self.c_name_filter_var.set("")
        self.nb.tab(0, text=" File Name ")
        self.nb.tab(1, text=" Folder Name ")

        # v4.4: "Search again" from the History tab (Help) used to leave the
        # notebook sitting on whatever tab it was triggered from, so results
        # were invisible until the user manually clicked "File Name". Jump
        # there automatically -- but only for that flow (flag set in
        # use_query_from_hist), not on every live keystroke re-search, which
        # would otherwise yank the user off a tab they're actively reading.
        if getattr(self, "_force_file_tab", False):
            self._force_file_tab = False
            try:
                # v7.7: default results tab is now File Content (index 2),
                # not File Name (index 0) -- see matching change in
                # show_results.
                self.nb.select(2)
            except Exception:
                pass

        # Reset AI Search + Advanced button for new query — but not if this
        # is just a background refresh for the SAME query AI mode is already
        # showing results for (see matching guard in _smart_search_realtime).
        _same_ai_query = self._ai_mode_active and query == getattr(self, "_ai_active_query", None)
        # v-fix: a sticky AI re-run for this exact query may already be in
        # flight (kicked off from _smart_search_realtime, which clears
        # _ai_mode_active first as a re-arm trick -- see the comment where
        # _ai_pending_query is defined in __init__). Treat that the same as
        # "_same_ai_query" so we don't tear down the AI split window / AI
        # chat pane and flash back to Default right before that background
        # thread finishes and turns AI mode back on for this query.
        _same_ai_query = _same_ai_query or (query == getattr(self, "_ai_pending_query", None))
        if not _same_ai_query:
            self._ai_mode_active = False
            self._adv_mode_active = False
            self._adv_page = 0
            self._adv_all_cont  = []
            self._adv_all_files = []
            self._ai_cont_res   = []
            self._ai_file_res   = []
            # Close any open split windows
            self._close_adv_split()
            self._close_ai_split()
            try:
                if self._ai_search_btn and self._ai_search_btn.winfo_exists():
                    self._ai_search_btn.config(text="🤖 AI Search", fg="#7ec8e3",
                                               bg="#1e3a5f", state="normal")
                if self._adv_search_btn and self._adv_search_btn.winfo_exists():
                    self._adv_search_btn.config(text="Advanced", fg="#33363c",
                                                bg="#e6e8ec", state="normal")
            except Exception: pass
            try:
                if hasattr(self, '_hybrid_status_lbl') and self._hybrid_status_lbl.winfo_exists():
                    self._hybrid_status_lbl.config(text="📊 BM25", fg="#aaaaaa")
            except Exception: pass
        
        def bg_load_ui(tree, data, mode, current_vid):
            if not data or getattr(tree, 'search_id', 0) != current_vid: return
            CHUNK = 500
            first_chunk, remaining_chunk = data[:CHUNK], data[CHUNK:]
            offset = [0]
            for item in first_chunk:
                if getattr(tree, 'search_id', 0) != current_vid: return
                try:
                    offset[0] += 1
                    self._insert_row(tree, item, offset[0], mode)
                except: pass

            def _update_count2():
                if mode == 1 and self.filter_count_label:
                    self.filter_count_label.config(text=f"{len(self._last_bm25_file_res or [])} files")
                elif mode == 0 and self.content_filter_count_label:
                    self.content_filter_count_label.config(text=f"{len(self._last_bm25_cont_res or [])} files")

            def load_rest(step=0):
                if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win): return
                if getattr(tree, 'search_id', 0) != current_vid: return
                sub_chunk = remaining_chunk[step: step + CHUNK]
                if sub_chunk:
                    for item in sub_chunk:
                        if getattr(tree, 'search_id', 0) != current_vid: return
                        try:
                            offset[0] += 1
                            self._insert_row(tree, item, offset[0], mode)
                        except: pass
                    self.root.after(5, lambda: load_rest(step + CHUNK))
                else:
                    _update_count2()
            if remaining_chunk:
                self.root.after(10, lambda: load_rest(0))
            else:
                self.root.after(20, _update_count2)

        bg_load_ui(self.tree_c, cont_res, 0, version)
        bg_load_ui(self.tree_f, only_files, 1, version)
        bg_load_ui(self.tree_fol, only_folders, 2, version)

        # Reset pane-tree dicts to single-pane state for new search
        self._c_pane_trees   = {"main": self.tree_c}
        self._f_pane_trees   = {"main": self.tree_f}
        self._fol_pane_trees = {"main": self.tree_fol}

    def _parse_tier_filter(self, raw_text):
        """Parse a *standalone* tier expression (no '--update data' prefix) —
        used by the Update DB button so the user can just type 'tier 1,2' or
        even '1,2' into the Searchbox and click the button, instead of typing
        the full '--update data tier 1,2' command and pressing Enter.
        Accepts: '', 'tier 1', 'tier 1,2', '1,2', '1 2 3', 'tier 1 tier 2', ...
        Returns (selected_tiers, is_tier_expr):
          - ('', True)         -> box empty -> (None, True)  = no filter, all tiers
          - valid tier text    -> ({0-indexed tiers}, True)
          - anything else      -> (None, False) = not a tier expression at all
                                   (e.g. a normal search query like 'SR01403940'),
                                   caller should fall back to a full scan.
        """
        text = (raw_text or "").strip()
        if not text:
            return None, True
        rest = re.sub(r'(?i)\btier\b', ' ', text)
        if not re.fullmatch(r'[\d,\s]+', rest.strip()):
            return None, False
        nums = re.findall(r'\d+', rest)
        parsed = {int(n) - 1 for n in nums if 1 <= int(n) <= 4}
        if not parsed:
            return None, False
        return parsed, True

    def _on_search_entry_return(self, event=None):
        """v-fix (Vấn đề 1: gõ 軌道不整 (kanji qua IME tiếng Nhật) vào ô
        Search không ra kết quả gì -- cả 3 tab -- nhưng PASTE cùng chuỗi
        đó lại ra ngay): phím Enter dùng để XÁC NHẬN (commit)候補 kanji
        đang chọn trong IME thường TRÙNG đúng phím Enter mà widget này
        cũng bind để chạy tìm kiếm (handle_action). Trên Windows, khi
        IME xử lý phím Enter đó, Tk nhận được KeyPress "Return" và gọi
        callback bind() gần NHƯ ĐỒNG THỜI với lúc IME mới ghi chuỗi kanji
        đã chốt vào bộ đệm text của Entry -- có 1 khoảng race nhỏ nơi
        callback bind() có thể chạy TRƯỚC khi entry_var đã thực sự được
        cập nhật xong, nên handle_action() đọc phải chuỗi CŨ (rỗng hoặc
        dở dang romaji) thay vì "軌道不整" -- paste không qua IME nên
        không có race này. Dời việc gọi handle_action() ra sau 1 tick
        (root.after(1, ...)) để vòng lặp sự kiện Tcl/Tk kịp xử lý xong
        phần ghi text của IME trước khi ta đọc entry_var -- cùng kiểu
        fix/lý do đã áp dụng cho <<Paste>> ở trên."""
        self.root.after(1, self.handle_action)

    def handle_action(self):
        raw_query = self.entry_var.get().strip()
        if not raw_query: return
        
        if getattr(self, '_hist_idle_timer', None):
            try: self.root.after_cancel(self._hist_idle_timer)
            except Exception: pass
            self._hist_idle_timer = None
        self._save_hist(raw_query)
        q_norm = raw_query
        if q_norm.lower() in ["--exit", "--quit"]: (self.root.destroy() or None); return
            
        q_stripped = q_norm.strip()
        q_stripped_lower = q_stripped.lower()
        if q_stripped_lower == "--update data" or q_stripped_lower.startswith("--update data "):
            # v2.5: optional tier filter — lets you index just the tiers you
            # care about instead of the whole drive. Accepts any of:
            #   --update data                    -> all tiers (1-4)
            #   --update data tier 1              -> tier 1 only
            #   --update data tier 1,2            -> tiers 1 and 2
            #   --update data tier 1 tier 2       -> tiers 1 and 2
            #   --update data tier 1 2 3          -> tiers 1, 2, 3
            _tier_arg = q_stripped[len("--update data"):].strip()
            _selected_tiers = None  # None = all tiers (default, unchanged behavior)
            if _tier_arg:
                _rest = _tier_arg.lower().replace("tier", " ")
                _nums = re.findall(r'\d+', _rest)
                _parsed = set()
                for n in _nums:
                    tn = int(n)
                    if 1 <= tn <= 4:
                        _parsed.add(tn - 1)  # store 0-indexed to match _ext_tier()
                if _parsed:
                    _selected_tiers = _parsed
                else:
                    self.status_label.config(text="Bad tier filter — use e.g. 'tier 1,2'", fg="#ff5555")
                    self.entry_var.set(""); return
            _tier_desc = ("ALL (1-4)" if _selected_tiers is None
                          else ",".join(str(t + 1) for t in sorted(_selected_tiers)))
            self._update_db_running = True  # lock AI Search/model combo immediately, don't wait for the thread
            self._set_index_status(f"Updating (tier {_tier_desc})...", "#ffcc00"); self._ramp_blink_start(); self.root.update_idletasks()
            threading.Thread(target=self.indexing_worker, args=(_selected_tiers,), daemon=True).start()
            self.entry_var.set(""); return

        if q_stripped_lower == "--update outlook":
            # v-outlook: standalone command — same safe, ramp-light-independent
            # logic as the Update DB dialog's "Outlook mail" checkbox.
            self._start_outlook_only_update()
            self.entry_var.set(""); return

        if q_stripped_lower == "--update onenote":
            # v-onenote: standalone command — same idea, OneNote only.
            self._start_onenote_only_update()
            self.entry_var.set(""); return

        q_norm = unicodedata.normalize('NFKC', q_norm)

        # v3.4 FIX: this used to just .lift() the window and do nothing
        # else whenever active_result_win already existed — harmless when
        # Enter is pressed after typing (on_key_release's debounce had
        # already fired the real search on keystrokes), but silently
        # broken for any caller that sets entry_var programmatically and
        # calls handle_action() directly without going through
        # <KeyRelease> first — e.g. "Search again" from History, which
        # left the keyword sitting in the box with no results ever
        # fetched. Now this branch always launches a fresh search itself
        # (same calls on_key_release makes), and additionally lifts the
        # window to the front if one was already open.
        if self.active_result_win and tk.Toplevel.winfo_exists(self.active_result_win):
            self.active_result_win.lift()
        self.current_search_id += 1; self.root.update_idletasks()
        box_x, box_y, box_h = self.root.winfo_x(), self.root.winfo_y(), self._search_bar_box_h()
        threading.Thread(target=self._mft_scan_search, args=(
            q_norm, self.current_search_id, box_x, box_y, box_h), daemon=True).start()
        if self.db_conn is not None:
            threading.Thread(target=self._smart_search_realtime, args=(
                q_norm, self.current_search_id, box_x, box_y, box_h, False), daemon=True).start()

    # ── Priority sort helpers ──────────────────────────────────────────────────
    # v2.5: user-defined 4-tier system — used for BOTH (a) the order files are
    # extracted in during --update data (tier 1 finishes across the whole
    # filesystem before tier 2 starts, etc. — so an interrupted/partial run
    # always has the highest-value content indexed first) and (b) the "#"
    # sort order shown in File Content search results. Single source of
    # truth so the two never drift apart.
    #
    # Tier 1: primary office documents — the main things people search for
    _PRIORITY_EXTS_TIER0 = {
        '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf', '.csv',
    }
    # Tier 2: everyday readable content — email, notes, plain text, logs
    _PRIORITY_EXTS_TIER1 = {
        '.msg', '.eml', '.txt', '.one', '.log',
    }
    # Tier 3: scripts / config / domain-specific engineering files
    #   .spck = Simpack model files — domain-specific to this environment
    _PRIORITY_EXTS_TIER2 = {
        '.spck', '.py', '.env', '.ini', '.cfg', '.conf',
        '.js', '.ts', '.bat', '.ps1', '.json', '.yaml', '.yml', '.toml',
    }
    # Tier 4: markup / query / misc structured text
    _PRIORITY_EXTS_TIER3 = {
        '.htm', '.html', '.sql', '.xml', '.css', '.md', '.rst', '.tsv',
    }
    _ALL_TIERS = [_PRIORITY_EXTS_TIER0, _PRIORITY_EXTS_TIER1, _PRIORITY_EXTS_TIER2, _PRIORITY_EXTS_TIER3]

    @classmethod
    def _ext_tier(cls, ext):
        for i, tier_exts in enumerate(cls._ALL_TIERS):
            if ext in tier_exts:
                return i
        return len(cls._ALL_TIERS)  # unlisted extension -> put at the very end

    # v9.10: tier only breaks ties WITHIN a band of this many
    # originally-adjacent (by relevance) results, instead of being an
    # absolute override of relevance order across the whole result set.
    # Tune this if results still feel too format-biased (raise it) or too
    # relevance-only / not enough office-doc preference (lower it).
    _TIER_SORT_BAND_SIZE = 8

    def _sort_priority(self, rows, mode):
        """Sort by tier (see _ext_tier above) — used as a gentle
           tie-breaker among near-equally-relevant results, NOT an absolute
           override of the relevance order `rows` already arrived in.

           v9.10 fix: the old version sorted PURELY by (grp, tier,
           drive_score), throwing away the incoming relevance order
           entirely except as a same-tier tie-breaker. That meant ANY
           .pdf/.docx (tier 0) always sorted above EVERY other file format,
           regardless of how relevant it actually was — a barely-relevant
           PDF could rank above a highly relevant .zip/.spck simply because
           archives/simulation files aren't in the tier list (fall through
           to the lowest tier). Now the incoming rank is banded into groups
           of _TIER_SORT_BAND_SIZE and used as the PRIMARY key -- tier can
           only reorder results that were already close in relevance, never
           flip a real relevance gap.

           Outermost key unchanged: match group — rows tagged as an
           OR-fallback match (row[5] == 1, used by MFT multi-keyword
           search) always sort BELOW rows that matched the full AND query
           (row[5] == 0 or untagged).
        """
        def _key(indexed_row):
            i, r = indexed_row
            path = r[0] if mode == 0 else r[2]
            drive = path[0].upper() if path and len(path) >= 2 and path[1] == ':' else ''
            drive_score = 99 if drive == 'C' else 0
            ext = os.path.splitext(path)[1].lower()
            tier = self._ext_tier(ext)
            grp = r[5] if len(r) > 5 else 0
            band = i // self._TIER_SORT_BAND_SIZE
            # v-new (yêu cầu: cửa sổ AI Search result phải xếp giống hệt
            # BM25 -- pdf/ppt/doc/excel... luôn ở TRÊN, msg/OneNote luôn ở
            # DƯỚI CÙNG): trước đây msg/.one chỉ bị đẩy xuống trong PHẠM VI
            # 1 band (_TIER_SORT_BAND_SIZE=8) -- nếu 1 email/note có điểm
            # liên quan rất cao và rơi vào band đầu, nó vẫn chen lẫn giữa
            # các pdf/ppt của band đó, và band đó luôn hiện TRƯỚC mọi band
            # sau dù band sau toàn pdf/ppt liên quan cao hơn. Giờ tách
            # riêng: mail_note=1 (msg/eml/one) luôn xuống dưới TẤT CẢ các
            # file khác (band nào cũng vậy) -- áp dụng chung cho cả cửa sổ
            # BM25 lẫn AI Search vì 2 nơi dùng chung hàm này, nên luôn nhất
            # quán với nhau. Thứ tự tương đối GIỮA các tier khác (0/1/2/3)
            # vẫn giữ nguyên kiểu "band" như cũ -- không đụng tới, tránh
            # lặp lại lỗi v9.10 (1 pdf ít liên quan đè lên nội dung liên
            # quan hơn) đã fix trước đây.
            mail_note = 1 if ext in ('.msg', '.eml', '.one') else 0
            return (grp, mail_note, band, tier, drive_score, i)
        return [r for _, r in sorted(enumerate(rows), key=_key)]

    # ── Sortable columns (class methods — shared across ALL trees in ALL panes) ──
    def _sort_tree(self, tree, col, reverse):
        """Sort treeview by column; toggle direction on repeated click.
        Works with any Treeview (main/adv/ai, across all 3 tabs)."""
        try:
            items = [(tree.set(k, col), k) for k in tree.get_children("")]
            # Try numeric sort for size (contains digits + unit)
            def _sort_key(val):
                v = val[0]
                # File Size: convert "1.2 MB" → bytes for proper numeric sort
                _units = {'B': 1, 'KB': 1024, 'MB': 1024**2, 'GB': 1024**3, 'TB': 1024**4}
                parts = v.strip().split()
                if len(parts) == 2 and parts[1] in _units:
                    try: return (0, float(parts[0]) * _units[parts[1]])
                    except: pass
                # Date: already ISO-ish "YYYY-MM-DD HH:MM" → sorts correctly as string
                # Fallback: case-insensitive string
                return (1, v.lower())
            items.sort(key=lambda x: _sort_key(x), reverse=reverse)
            for idx, (_, k) in enumerate(items):
                tree.move(k, "", idx)
            # v5.8: #0 is icon-only now (no row numbers to renumber after sort)
            # Update heading with arrow indicator
            for c in tree["columns"]:
                tree.heading(c, text=tree.heading(c)["text"].replace(" ▲","").replace(" ▼",""))
            arrow = " ▼" if reverse else " ▲"
            tree.heading(col, text=tree.heading(col)["text"] + arrow,
                         command=lambda: self._sort_tree(tree, col, not reverse))
        except Exception as _se:
            print(f"Sort error: {_se}")

    def _make_sortable(self, tree, sortable_cols):
        """Bind click-to-sort for the specified columns. Call this immediately when a tree
        is created (including Advanced/AI panes born from a split) so Size/
        Modified/Type columns are always sortable, regardless of which pane the tree is in."""
        for col in sortable_cols:
            tree.heading(col, command=lambda c=col: self._sort_tree(tree, c, False))

    # ══════════════════════════════════════════════════════════════════════════
    # SPLIT LAYOUT ENGINE
    # Each Tab uses the same 3-phase layout strategy:
    #
    #   Default  (1 pane):  [  Main  ]
    #   Advanced (2 panes): [  Main  ]   ← top 50%
    #                       [ Adv.   ]   ← bottom 50%
    #   +AI      (3 panes): [Main|AI ]   ← top-left 60% | right 40%
    #                       [Adv.|AI ]   ← bottom-left  | right shared
    #
    # Layout as required:
    #   Advanced only  → top/bottom 50%/50%  (fixed, no sash)
    #   AI only        → left/right  50%/50% (fixed, no sash)
    #   Advanced + AI  → outer H-split 60/40; left side continues with V-split 50/50
    # ══════════════════════════════════════════════════════════════════════════

    # ── Treeview factory helpers ──────────────────────────────────────────────
    def _build_tree_content(self, parent, label_text, label_bg, label_fg, show_label=True):
        """Build a content-style tree (path, size, mtime, type) in parent."""
        lbl_widget = None
        if show_label:
            lbl_widget = tk.Label(parent, text=label_text, bg=label_bg, fg=label_fg,
                     font=("Segoe UI", 9, "bold"))
            lbl_widget.pack(fill="x")
        tree = ttk.Treeview(parent,
                            columns=("path","size","mtime","ftype","fp"),
                            show="tree headings")
        tree.heading("#0", text="")
        tree.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        for col, w, txt in [("path",750,"File Location"),
                             ("size",90,"Size"), ("mtime",130,"Modified"), ("ftype",55,"Type")]:
            tree.heading(col, text=txt)
            tree.column(col, width=w, stretch=False, minwidth=30)
        tree["displaycolumns"] = ("path","size","mtime","ftype")
        sb  = ttk.Scrollbar(parent, orient="vertical",   command=tree.yview)
        sbx = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb.set, xscrollcommand=sbx.set)
        sb.pack(side="right", fill="y"); sbx.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree._header_label = lbl_widget  # store ref for dynamic update
        tree._base_label   = label_text  # base text (e.g. '📄 Default (33)')
        self._make_sortable(tree, ["size", "mtime", "ftype"])  # wire sort for all panes
        return tree

    def _build_tree_file(self, parent, label_text, label_bg, label_fg, show_label=True):
        """Build a file-name-style tree (name, size, mtime, type, loc) in parent."""
        lbl_widget = None
        if show_label:
            lbl_widget = tk.Label(parent, text=label_text, bg=label_bg, fg=label_fg,
                     font=("Segoe UI", 9, "bold"))
            lbl_widget.pack(fill="x")
        tree = ttk.Treeview(parent,
                            columns=("name","size","mtime","ftype","loc","fp"),
                            show="tree headings")
        tree.heading("#0", text="")
        tree.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        for col, w, txt in [("name",360,"File Name"),
                             ("size",90,"Size"), ("mtime",130,"Modified"),
                             ("ftype",55,"Type"), ("loc",520,"Location")]:
            tree.heading(col, text=txt)
            tree.column(col, width=w, stretch=False, minwidth=30)
        tree["displaycolumns"] = ("name","size","mtime","ftype","loc")
        sb  = ttk.Scrollbar(parent, orient="vertical",   command=tree.yview)
        sbx = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb.set, xscrollcommand=sbx.set)
        sb.pack(side="right", fill="y"); sbx.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree._header_label = lbl_widget
        tree._base_label   = label_text
        self._make_sortable(tree, ["size", "mtime", "ftype"])  # wire sort for all panes
        return tree

    def _build_tree_folder(self, parent, label_text, label_bg, label_fg, show_label=True):
        """Build a folder-name-style tree (name, mtime, loc) in parent."""
        lbl_widget = None
        if show_label:
            lbl_widget = tk.Label(parent, text=label_text, bg=label_bg, fg=label_fg,
                     font=("Segoe UI", 9, "bold"))
            lbl_widget.pack(fill="x")
        tree = ttk.Treeview(parent,
                            columns=("name","mtime","loc","fp"),
                            show="tree headings")
        tree.heading("#0", text="")
        tree.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        for col, w, txt in [("name",420,"Folder Name"),
                             ("mtime",130,"Modified"), ("loc",650,"Location")]:
            tree.heading(col, text=txt)
            tree.column(col, width=w, stretch=False, minwidth=30)
        tree["displaycolumns"] = ("name","mtime","loc")
        sb  = ttk.Scrollbar(parent, orient="vertical",   command=tree.yview)
        sbx = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb.set, xscrollcommand=sbx.set)
        sb.pack(side="right", fill="y"); sbx.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree._header_label = lbl_widget
        tree._base_label   = label_text
        self._make_sortable(tree, ["mtime"])  # wire sort for all panes
        return tree

    # ── Data populators ───────────────────────────────────────────────────────
    def _fill_content(self, tree, data):
        """Populate a content tree from _all_content_data items."""
        for row in tree.get_children(): tree.delete(row)
        for i, item in enumerate(self._sort_priority(data, 0), 1):
            try:
                fpath = item[0]
                pf = self._pseudo_path_row_fields(fpath)
                if pf:
                    name_txt, size_str, mtime_str, ftype_str, _loc_str, img = pf
                    tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(
                        name_txt, size_str, mtime_str, ftype_str, fpath))
                    continue
                img = get_tree_icon_image(fpath, False)
                name_txt = fpath if img else get_file_icon(fpath, False) + fpath
                tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(
                    name_txt,
                    format_size(item[1]) if item[1] else "",
                    get_live_mtime(fpath), get_file_type(fpath), fpath))
            except: pass

    def _fill_files(self, tree, data):
        """Populate a file tree from _all_files_data items."""
        for row in tree.get_children(): tree.delete(row)
        for i, item in enumerate(self._sort_priority(data, 1), 1):
            try:
                fpath = item[2]
                pf = self._pseudo_path_row_fields(fpath)
                if pf:
                    name_txt, size_str, mtime_str, ftype_str, loc_str, img = pf
                    tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(
                        name_txt, size_str, mtime_str, ftype_str, loc_str, fpath))
                    continue
                img = get_tree_icon_image(fpath, False)
                name_txt = item[1] if img else get_file_icon(fpath, False) + item[1]
                tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(
                    name_txt,
                    format_size(item[3]) if item[3] else "",
                    get_live_mtime(fpath), get_file_type(fpath),
                    os.path.dirname(fpath), fpath))
            except: pass

    def _fill_folders(self, tree, data):
        """Populate a folder tree from folder items."""
        for row in tree.get_children(): tree.delete(row)
        for i, item in enumerate(data, 1):
            try:
                img = get_tree_icon_image(item[2], True)
                name_txt = item[1] if img else get_file_icon(item[2], True) + item[1]
                tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(name_txt, get_live_mtime(item[2]), item[2], item[2]))
            except: pass

    # ── Layout builder — called for EACH tab_frame independently ─────────────
    def _rebuild_tab_layout(self, tab_frame, tab_mode,
                            main_data, adv_data, ai_data,
                            main_tree_attr):
        """
        Destroy and rebuild the layout inside tab_frame.

        tab_mode : "c"   → File Content tab  (main = content tree)
                   "f"   → File Name tab      (main = file tree)
                   "fol" → Folder Name tab    (main = folder tree)

        Layout states (same for every tab):
          adv=False, ai=False  →  1 pane  : main only
          adv=True,  ai=False  →  2 panes : top(main) / bottom(adv)  50/50
          adv=False, ai=True   →  2 panes : left(main) / right(ai)   60/40
          adv=True,  ai=True   →  3 panes : outer H 60/40;
                                            left V-split top(main)/bot(adv) 50/50;
                                            right = ai
        adv_mode / ai_mode are stored in self._adv_mode_active / self._ai_mode_active.
        """
        adv = self._adv_mode_active
        ai  = self._ai_mode_active

        # Clear everything in the container
        for w in tab_frame.winfo_children(): w.destroy()

        # ── helpers per tab_mode ─────────────────────────────────────────────
        # Show ">5000" only when result count hits the 5000 hard cap
        _HARD_CAP = 5000
        def _fmt(n):
            return f">{_HARD_CAP}" if n >= _HARD_CAP else str(n)
        n_main = len(main_data)
        n_adv  = len(adv_data)
        n_ai   = len(ai_data)
        s_main = _fmt(n_main)
        s_adv  = _fmt(n_adv)
        s_ai   = _fmt(n_ai)

        def make_main(parent, show_label=True):
            lbl = f"📄 Default ({s_main})" if tab_mode == "c" else f"📁 Default ({s_main})"
            if tab_mode == "c":
                return self._build_tree_content(parent, lbl, BG_COLOR, TEXT_COLOR, show_label=show_label)
            elif tab_mode == "f":
                return self._build_tree_file(parent, lbl, BG_COLOR, TEXT_COLOR, show_label=show_label)
            else:
                return self._build_tree_folder(parent, lbl, BG_COLOR, TEXT_COLOR, show_label=show_label)

        def make_adv(parent):
            lbl = f"📄 Advanced ({s_adv})" if tab_mode == "c" else f"📁 Advanced ({s_adv})"
            if tab_mode == "c":
                return self._build_tree_content(parent, lbl, "#1a3a2a", "#aaffaa")
            elif tab_mode == "f":
                return self._build_tree_file(parent, lbl, "#1a2a3a", "#7ec8e3")
            else:
                return self._build_tree_folder(parent, lbl, "#1a1a2a", "#aaaaff")

        def make_ai(parent):
            _model_tags = {"jina_v3": "Jina", "bge_gemma2": "Gemma2"}
            _size_tag = _model_tags.get(_sem_model_key, _sem_model_key)
            lbl = f"🤖 AI Search [{_size_tag}] ({s_ai})"
            if tab_mode == "c":
                return self._build_tree_content(parent, lbl, "#1a1a3a", "#a0a0ff")
            elif tab_mode == "f":
                return self._build_tree_file(parent, lbl, "#2a1a2a", "#d090f0")
            else:
                return self._build_tree_folder(parent, lbl, "#2a1a1a", "#ffaa88")

        def fill_main(tree):
            if tab_mode == "c":   self._fill_content(tree, main_data)
            elif tab_mode == "f": self._fill_files(tree, main_data)
            else:                 self._fill_folders(tree, main_data)

        def fill_adv(tree):
            if tab_mode == "c":   self._fill_content(tree, adv_data)
            elif tab_mode == "f": self._fill_files(tree, adv_data)
            else:                 self._fill_folders(tree, adv_data)

        def fill_ai(tree):
            if tab_mode == "c":   self._fill_content(tree, ai_data)
            elif tab_mode == "f": self._fill_files(tree, ai_data)
            else:                 self._fill_folders(tree, ai_data)

        # ── Determine which pane-tree dict to update ─────────────────────────
        if tab_mode == "c":   pane_dict = self._c_pane_trees
        elif tab_mode == "f": pane_dict = self._f_pane_trees
        else:                 pane_dict = self._fol_pane_trees
        pane_dict.clear()

        # ── Build layout — strict 50/50 using place geometry ─────────────────
        SEP = 3  # separator thickness in px

        if not adv and not ai:
            # ── 1 pane ───────────────────────────────────────────────────────
            pane = tk.Frame(tab_frame, bg=BG_COLOR)
            pane.place(relx=0, rely=0, relwidth=1.0, relheight=1.0)
            main_tree = make_main(pane, show_label=False)
            fill_main(main_tree)
            setattr(self, main_tree_attr, main_tree)
            pane_dict["main"] = main_tree
            self._wire_tree_clicks(main_tree, tab_mode)

        elif adv and not ai:
            # ── 2 panes top/bottom strict 50/50 ──────────────────────────────
            sep_h = tk.Frame(tab_frame, bg="#555", height=SEP)
            sep_h.place(relx=0, rely=0.5, relwidth=1.0, height=SEP, anchor="w")

            top = tk.Frame(tab_frame, bg=BG_COLOR)
            top.place(relx=0, rely=0, relwidth=1.0, relheight=0.5)
            bot = tk.Frame(tab_frame, bg=BG_COLOR)
            bot.place(relx=0, rely=0.5, relwidth=1.0, relheight=0.5)

            main_tree = make_main(top); fill_main(main_tree)
            adv_tree  = make_adv(bot);  fill_adv(adv_tree)
            setattr(self, main_tree_attr, main_tree)
            pane_dict["main"] = main_tree
            pane_dict["adv"]  = adv_tree
            self._wire_tree_clicks(main_tree, tab_mode)
            self._wire_tree_clicks(adv_tree,  tab_mode)

        elif not adv and ai:
            # ── 2 panes left/right strict 50/50 ──────────────────────────────
            sep_v = tk.Frame(tab_frame, bg="#555", width=SEP)
            sep_v.place(relx=0.5, rely=0, width=SEP, relheight=1.0, anchor="n")

            left  = tk.Frame(tab_frame, bg=BG_COLOR)
            left.place(relx=0, rely=0, relwidth=0.5, relheight=1.0)
            right = tk.Frame(tab_frame, bg=BG_COLOR)
            right.place(relx=0.5, rely=0, relwidth=0.5, relheight=1.0)

            main_tree = make_main(left);  fill_main(main_tree)
            ai_tree   = make_ai(right);   fill_ai(ai_tree)
            setattr(self, main_tree_attr, main_tree)
            pane_dict["main"] = main_tree
            pane_dict["ai"]   = ai_tree
            self._wire_tree_clicks(main_tree, tab_mode)
            self._wire_tree_clicks(ai_tree,   tab_mode)

        else:
            # ── 3 panes: left 50% | right 50%; left split top/bottom 50/50 ──
            sep_v = tk.Frame(tab_frame, bg="#555", width=SEP)
            sep_v.place(relx=0.5, rely=0, width=SEP, relheight=1.0, anchor="n")

            left_frame = tk.Frame(tab_frame, bg=BG_COLOR)
            left_frame.place(relx=0, rely=0, relwidth=0.5, relheight=1.0)
            right = tk.Frame(tab_frame, bg=BG_COLOR)
            right.place(relx=0.5, rely=0, relwidth=0.5, relheight=1.0)

            sep_h = tk.Frame(left_frame, bg="#555", height=SEP)
            sep_h.place(relx=0, rely=0.5, relwidth=1.0, height=SEP, anchor="w")

            top = tk.Frame(left_frame, bg=BG_COLOR)
            top.place(relx=0, rely=0, relwidth=1.0, relheight=0.5)
            bot = tk.Frame(left_frame, bg=BG_COLOR)
            bot.place(relx=0, rely=0.5, relwidth=1.0, relheight=0.5)

            main_tree = make_main(top); fill_main(main_tree)
            adv_tree  = make_adv(bot);  fill_adv(adv_tree)
            ai_tree   = make_ai(right); fill_ai(ai_tree)
            setattr(self, main_tree_attr, main_tree)
            pane_dict["main"] = main_tree
            pane_dict["adv"]  = adv_tree
            pane_dict["ai"]   = ai_tree
            self._wire_tree_clicks(main_tree, tab_mode)
            self._wire_tree_clicks(adv_tree,  tab_mode)
            self._wire_tree_clicks(ai_tree,   tab_mode)

    # ── Public split triggers (called by Advanced / AI Search buttons) ────────

    def _wire_tree_clicks(self, tree, tab_mode):
        """Wire double-click, single-click (cell-edit), and right-click context menu
        to any Treeview — including panes created after the initial layout by split."""
        open_file     = getattr(self, "_fn_open_file",     None)
        open_explorer = getattr(self, "_fn_open_explorer", None)
        show_ctx      = getattr(self, "_fn_show_ctx_menu", None)
        on_click      = getattr(self, "_fn_on_tree_click", None)
        dbl_select    = getattr(self, "_fn_dbl_select",    None)

        if not all([open_file, open_explorer, show_ctx, on_click, dbl_select]):
            return  # handlers not ready yet

        is_folder  = (tab_mode == "fol")
        is_content = not is_folder  # both "c" and "f" tabs can open files

        # Use default-argument capture to avoid late-binding closure trap
        if is_folder:
            tree.bind("<Double-1>", lambda e, t=tree: (dbl_select(t), open_explorer(t)))
        else:
            tree.bind("<Double-1>", lambda e, t=tree: (dbl_select(t), open_file(t)))

        tree.bind("<Button-1>", lambda e, t=tree: on_click(e, t))
        tree.bind("<Button-3>", lambda e, t=tree, f=is_folder, c=is_content:
                  show_ctx(e, t, is_folder=f, is_content=c))

    def _open_adv_split_win(self):
        """Advanced activated → rebuild ALL tabs with adv=True layout."""
        if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win):
            return
        # FIX 3: main=BM25 realtime, adv=full Advanced results, ai=AI results
        bm25_cont    = self._last_bm25_cont_res or []
        bm25_files   = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() != "folder"]
        bm25_folders = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() == "folder"]
        adv_cont     = self._adv_all_cont  or []
        adv_files    = [r for r in (self._adv_all_files or []) if str(r[0]).lower() != "folder"]
        adv_folders  = [r for r in (self._adv_all_files or []) if str(r[0]).lower() == "folder"]
        ai_cont      = self._ai_cont_res  or []
        ai_files     = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() != "folder"]
        ai_folders   = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() == "folder"]

        self._rebuild_tab_layout(self.c_tree_frame,   "c",
            bm25_cont,    adv_cont,    ai_cont,    "tree_c")
        self._rebuild_tab_layout(self.f_tree_frame,   "f",
            bm25_files,   adv_files,   ai_files,   "tree_f")
        self._rebuild_tab_layout(self.fol_tree_frame, "fol",
            bm25_folders, adv_folders, ai_folders, "tree_fol")

    def _close_adv_split(self):
        """Advanced deactivated → rebuild ALL tabs without adv layout."""
        if not (hasattr(self, 'c_tree_frame') and self.c_tree_frame.winfo_exists()):
            return
        bm25_cont    = self._last_bm25_cont_res or []
        bm25_files   = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() != "folder"]
        bm25_folders = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() == "folder"]
        ai_cont      = self._ai_cont_res  or []
        ai_files     = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() != "folder"]
        ai_folders   = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() == "folder"]
        self._rebuild_tab_layout(self.c_tree_frame,   "c",
            bm25_cont,    [], ai_cont,    "tree_c")
        self._rebuild_tab_layout(self.f_tree_frame,   "f",
            bm25_files,   [], ai_files,   "tree_f")
        self._rebuild_tab_layout(self.fol_tree_frame, "fol",
            bm25_folders, [], ai_folders, "tree_fol")

    def _open_ai_split_win(self):
        """AI Search activated → rebuild ALL tabs with ai=True layout."""
        if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win):
            return
        # FIX 3: main=BM25 realtime, adv=Advanced full (if active), ai=AI results
        bm25_cont    = self._last_bm25_cont_res or []
        bm25_files   = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() != "folder"]
        bm25_folders = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() == "folder"]
        adv_cont     = self._adv_all_cont  or []
        adv_files    = [r for r in (self._adv_all_files or []) if str(r[0]).lower() != "folder"]
        adv_folders  = [r for r in (self._adv_all_files or []) if str(r[0]).lower() == "folder"]
        ai_cont      = self._ai_cont_res  or []
        ai_files     = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() != "folder"]
        ai_folders   = [r for r in (self._ai_file_res  or []) if str(r[0]).lower() == "folder"]

        self._rebuild_tab_layout(self.c_tree_frame,   "c",
            bm25_cont,    adv_cont,    ai_cont,    "tree_c")
        self._rebuild_tab_layout(self.f_tree_frame,   "f",
            bm25_files,   adv_files,   ai_files,   "tree_f")
        self._rebuild_tab_layout(self.fol_tree_frame, "fol",
            bm25_folders, adv_folders, ai_folders, "tree_fol")
        # v1.2: offline AI Search active -> tab layout now includes an AI
        # results pane. v-fix (theo yêu cầu: AI Chat panel không còn liên
        # quan gì đến nút "AI Search" nữa -- nút đó chỉ điều khiển offline/
        # embedding model kết quả, còn AI Chat panel (Gemini/GPT-OSS) đã tự
        # hiện độc lập ngay từ lúc có kết quả tìm kiếm đầu tiên, xem
        # _smart_search_realtime): KHÔNG gọi _show_online_chat_panel() ở
        # đây nữa -- nó đã hiện sẵn rồi, gọi lại là thừa và (quan trọng
        # hơn) việc gắn show/hide panel vào đúng lúc bật/tắt AI Search là
        # nguồn gốc của bug "chat panel chớp 0.1s rồi biến mất" (xem
        # _close_ai_split bên dưới).
        # v1.3: kick off an automatic first answer so the chat panel isn't
        # left empty right after clicking AI Search.
        query = getattr(self, "_ai_active_query", "") or self.entry_var.get().strip()
        self.root.after(150, lambda: self._auto_ai_chat_after_search(query))

    def _close_ai_split(self):
        """AI Search deactivated -> rebuild ALL tabs without ai layout."""
        if not (hasattr(self, 'c_tree_frame') and self.c_tree_frame.winfo_exists()):
            return
        bm25_cont    = self._last_bm25_cont_res or []
        bm25_files   = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() != "folder"]
        bm25_folders = [r for r in (self._last_bm25_file_res or []) if str(r[0]).lower() == "folder"]
        adv_cont     = self._adv_all_cont  or []
        adv_files    = [r for r in (self._adv_all_files or []) if str(r[0]).lower() != "folder"]
        adv_folders  = [r for r in (self._adv_all_files or []) if str(r[0]).lower() == "folder"]
        self._rebuild_tab_layout(self.c_tree_frame,   "c",
            bm25_cont,    adv_cont,    [], "tree_c")
        self._rebuild_tab_layout(self.f_tree_frame,   "f",
            bm25_files,   adv_files,   [], "tree_f")
        self._rebuild_tab_layout(self.fol_tree_frame, "fol",
            bm25_folders, adv_folders, [], "tree_fol")
        # v-fix (bug: AI Chat panel chớp ~0.1s rồi biến mất mỗi lần gõ từ
        # khoá mới): trước đây _close_ai_split() (chạy MỖI lần có kết quả
        # tìm kiếm mới, kể cả những lần merge Outlook/OneNote chạy async
        # SAU lần render đầu tiên -- xem update_or_show_results) luôn gọi
        # _hide_online_chat_panel() ở cuối, coi panel chat online như một
        # phần của offline "AI Search" (embedding/hybrid) -- 2 thứ này
        # KHÔNG liên quan gì đến nhau: AI Search (nút 🤖) chỉ điều khiển
        # kết quả offline/embedding model, còn AI Chat (Gemini/GPT-OSS) đã
        # tự hiện độc lập ngay khi có kết quả BM25 đầu tiên. Vì
        # update_or_show_results gọi _close_ai_split() cho MỌI lần render
        # (không chỉ khi user chủ động tắt AI Search), panel chat bị ẩn
        # lại gần như ngay sau khi vừa hiện lên bởi lần render kế tiếp
        # (vd. merge mail/notes chạy async ~1s sau đó) -- đúng cảnh
        # "hiện ra rồi biến mất" người dùng thấy. Không còn ẩn panel chat
        # ở đây nữa; nó chỉ nên biến mất khi cửa sổ kết quả đóng hẳn.
        pass

    def _open_triple_split_win(self):
        """Both Advanced + AI active → same as calling open_adv/ai (flags already set)."""
        self._open_adv_split_win()

    def _pseudo_path_row_fields(self, path):
        """v-fix (Vấn đề: sau khi vá mode==1 của _insert_row, "Bad" raw
        OUTLOOK::/ONENOTE:: token vẫn có thể lọt qua vào tab File Name --
        vì _render_mft_file_tree / _apply_size_filter_to_tree /
        _apply_folder_filter_to_tree đều tự tree.insert() riêng, KHÔNG đi
        qua _insert_row): helper dùng chung, format y hệt logic mode==0/1
        của _insert_row. Trả về (name_txt, size_str, mtime_str, ftype_str,
        location_str) nếu path là Outlook/OneNote pseudo-path, else None
        (nghĩa là: xử lý bình thường như file thật)."""
        if _is_outlook_pseudo_path(path):
            meta = self._outlook_meta.get(path, {})
            try:
                _, _entry_id, _store_id, folder_and_subject = path.split("::", 3)
                folder_path, subject_msg = folder_and_subject.split("::", 1) \
                    if "::" in folder_and_subject else ("", folder_and_subject)
                subject = subject_msg[:-4] if subject_msg.lower().endswith(".msg") else subject_msg
            except Exception:
                folder_path, subject = "", ""
            if not subject:
                subject = (meta.get("subject") or "(no subject)").strip() or "(no subject)"
                folder_path = meta.get("folder_path") or ""
            display = f"Outlook > {folder_path} > {subject}.msg" if folder_path else f"Outlook > {subject}.msg"
            img = get_tree_icon_image(path, is_folder=False)
            name_txt = display if img else "📧 " + display
            return (name_txt, format_size(meta.get("size")), format_meta_datetime(meta.get("received")),
                    "MSG", folder_path or "Outlook", img)
        if _is_onenote_pseudo_path(path):
            meta = self._onenote_meta.get(path, {})
            try:
                _, _page_id, section_path, title_one = path.split("::", 3)
                title = title_one[:-4] if title_one.lower().endswith(".one") else title_one
            except Exception:
                section_path, title = "", ""
            if not title:
                title = (meta.get("title") or "(untitled)").strip() or "(untitled)"
                section_path = meta.get("section_path") or ""
            display = f"OneNote > {section_path} > {title}.one" if section_path else f"OneNote > {title}.one"
            img = get_tree_icon_image(path, is_folder=False)
            name_txt = display if img else "📓 " + display
            return (name_txt, "", format_meta_datetime(meta.get("last_modified") or meta.get("created")),
                    "ONE", section_path or "OneNote", img)
        return None

    def _insert_row(self, tree, item, rn, mode):
        """Insert one row into a result tree — uses DB size, no disk I/O.
        v5.8: column #0 is icon-ONLY now (no row number). ttk.Treeview only
        supports image= on the tree column (#0), which is always the
        leftmost column and can't be reordered relative to the data
        columns, so the layout is [icon] → File Name → Size → ... .
        Emoji fallback if the icon backend isn't available or extraction
        fails for that row (emoji gets prefixed onto the Name text itself,
        since it can't be a real image=)."""
        if mode == 0:
            # item = (path, size, score) from BM25 content search
            path = item[0]
            size = item[1] if len(item) > 1 else None
            if _is_outlook_pseudo_path(path):
                # v-outlook: no real file on disk. Folder/subject are parsed
                # straight out of the pseudo-path token itself (embedded by
                # _make_outlook_pseudo_path) -- instant, no dependency on
                # self._outlook_meta having been populated yet by the
                # (sometimes slow) mail/notes merge. Same fix already
                # applied to _display_label_for_source; this brings the
                # Treeview row in sync with it so BM25/AI Search results
                # never show the raw OUTLOOK::... token while waiting on
                # that cache. Meta cache is only used below for the
                # secondary fields (sender/size/received), which are fine
                # to be briefly blank.
                meta = self._outlook_meta.get(path, {})
                try:
                    _, _entry_id, _store_id, folder_and_subject = path.split("::", 3)
                    folder_path, subject_msg = folder_and_subject.split("::", 1) \
                        if "::" in folder_and_subject else ("", folder_and_subject)
                    subject = subject_msg[:-4] if subject_msg.lower().endswith(".msg") else subject_msg
                except Exception:
                    folder_path, subject = "", ""
                if not subject:  # old-format pseudo path (pre-fix) -- fall back to meta cache
                    print(f"[Insert Row][Debug] Outlook path failed direct parse, path={path!r}")
                    subject = (meta.get("subject") or "(no subject)").strip() or "(no subject)"
                    folder_path = meta.get("folder_path") or ""
                sender = meta.get("sender") or ""
                display = f"Outlook > {folder_path} > {subject}.msg" if folder_path else f"Outlook > {subject}.msg"
                img = get_tree_icon_image(path, is_folder=False)  # extension-based (.msg) lookup — works without a real file
                name_txt = display if img else "📧 " + display
                readable_sz = format_size(meta.get("size"))
                mtime = format_meta_datetime(meta.get("received"))
                ftype = "MSG"
                kwargs = dict(text="", values=(name_txt, readable_sz, mtime, ftype, path))
                if img:
                    kwargs["image"] = img
                tree.insert("", tk.END, **kwargs)
                return
            if _is_onenote_pseudo_path(path):
                # v-onenote: same fix as Outlook above -- parse section/title
                # straight out of the pseudo-path token, don't depend on
                # self._onenote_meta being populated yet for the display name.
                meta = self._onenote_meta.get(path, {})
                try:
                    _, _page_id, section_path, title_one = path.split("::", 3)
                    title = title_one[:-4] if title_one.lower().endswith(".one") else title_one
                except Exception:
                    section_path, title = "", ""
                if not title:  # old-format pseudo path (pre-fix) -- fall back to meta cache
                    title = (meta.get("title") or "(untitled)").strip() or "(untitled)"
                    section_path = meta.get("section_path") or ""
                display = f"OneNote > {section_path} > {title}.one" if section_path else f"OneNote > {title}.one"
                img = get_tree_icon_image(path, is_folder=False)  # extension-based (.one) lookup — works without a real file
                name_txt = display if img else "📓 " + display
                readable_sz = ""  # OneNote pages have no meaningful byte size in this context
                mtime = format_meta_datetime(meta.get("last_modified") or meta.get("created"))
                ftype = "ONE"
                kwargs = dict(text="", values=(name_txt, readable_sz, mtime, ftype, path))
                if img:
                    kwargs["image"] = img
                tree.insert("", tk.END, **kwargs)
                return
            img = get_tree_icon_image(path, is_folder=False)
            name_txt = path if img else get_file_icon(path, is_folder=False) + path
            readable_sz = format_size(size) if size else ""
            mtime = get_live_mtime(path)
            ftype = get_file_type(path)
            kwargs = dict(text="", values=(name_txt, readable_sz, mtime, ftype, path))
            if img:
                kwargs["image"] = img
            tree.insert("", tk.END, **kwargs)
        elif mode == 1:
            # File Name tab: (type, name, path, size)
            path = item[2]
            # v-fix (Vấn đề: bấm "Search again" từ Search History luôn tự
            # nhảy sang tab File Name -- see use_query_from_hist -- nhưng
            # nhánh này chưa từng check _is_outlook_pseudo_path/
            # _is_onenote_pseudo_path như mode==0 (tab Content) đã có, nên
            # nếu 1 item Outlook/OneNote lọt vào đây, cột Name lẫn cột
            # Location đều hiện thẳng token thô "OUTLOOK::entry::store::
            # folder::subject.msg" thay vì "Outlook > folder > subject.msg"
            # -- đúng cái "Bad" user thấy). Format giống hệt logic ở
            # mode==0 để nhất quán giữa 2 tab.
            if _is_outlook_pseudo_path(path):
                meta = self._outlook_meta.get(path, {})
                try:
                    _, _entry_id, _store_id, folder_and_subject = path.split("::", 3)
                    folder_path, subject_msg = folder_and_subject.split("::", 1) \
                        if "::" in folder_and_subject else ("", folder_and_subject)
                    subject = subject_msg[:-4] if subject_msg.lower().endswith(".msg") else subject_msg
                except Exception:
                    folder_path, subject = "", ""
                if not subject:
                    subject = (meta.get("subject") or "(no subject)").strip() or "(no subject)"
                    folder_path = meta.get("folder_path") or ""
                display = f"Outlook > {folder_path} > {subject}.msg" if folder_path else f"Outlook > {subject}.msg"
                img = get_tree_icon_image(path, is_folder=False)
                name_txt = display if img else "📧 " + display
                kwargs = dict(text="", values=(name_txt, format_size(meta.get("size")),
                                                format_meta_datetime(meta.get("received")), "MSG",
                                                folder_path or "Outlook", path))
                if img:
                    kwargs["image"] = img
                tree.insert("", tk.END, **kwargs)
                return
            if _is_onenote_pseudo_path(path):
                meta = self._onenote_meta.get(path, {})
                try:
                    _, _page_id, section_path, title_one = path.split("::", 3)
                    title = title_one[:-4] if title_one.lower().endswith(".one") else title_one
                except Exception:
                    section_path, title = "", ""
                if not title:
                    title = (meta.get("title") or "(untitled)").strip() or "(untitled)"
                    section_path = meta.get("section_path") or ""
                display = f"OneNote > {section_path} > {title}.one" if section_path else f"OneNote > {title}.one"
                img = get_tree_icon_image(path, is_folder=False)
                name_txt = display if img else "📓 " + display
                kwargs = dict(text="", values=(name_txt, "",
                                                format_meta_datetime(meta.get("last_modified") or meta.get("created")), "ONE",
                                                section_path or "OneNote", path))
                if img:
                    kwargs["image"] = img
                tree.insert("", tk.END, **kwargs)
                return
            img = get_tree_icon_image(path, is_folder=False)
            name_txt = item[1] if img else get_file_icon(path, is_folder=False) + item[1]
            readable_sz = format_size(item[3]) if item[3] else ""
            mtime = get_live_mtime(path)
            ftype = get_file_type(path)
            kwargs = dict(text="", values=(name_txt, readable_sz, mtime, ftype, os.path.dirname(path), path))
            if img:
                kwargs["image"] = img
            tree.insert("", tk.END, **kwargs)
        else:
            # Folder Name tab: (type, name, path, size)
            img = get_tree_icon_image(item[2], is_folder=True)
            name_txt = item[1] if img else get_file_icon(item[2], is_folder=True) + item[1]
            mtime = get_live_mtime(item[2])
            kwargs = dict(text="", values=(name_txt, mtime, item[2], item[2]))
            if img:
                kwargs["image"] = img
            tree.insert("", tk.END, **kwargs)

    def _draw_search_icon(self, parent, size=18, color=None):
        """v4.9: flat vector magnifying-glass icon drawn on a Canvas — looks
        consistent across OSes/fonts and matches the theme, unlike the
        🔍 emoji glyph which renders as a colorful, font-dependent picture
        that clashes with a flat UI. Thicker strokes + longer handle + a
        blue tint by default, closer to the Windows 11 / Everything look."""
        color = color or "#0078d7"
        cv = tk.Canvas(parent, width=size, height=size, bg=BG_COLOR, highlightthickness=0)
        r = size * 0.30
        cx, cy = size * 0.40, size * 0.40
        lw = max(2, round(size * 0.13))  # stroke width scales with icon size
        cv.create_oval(cx - r, cy - r, cx + r, cy + r, outline=color, width=lw)
        ang = math.radians(45)
        x1 = cx + (r + lw * 0.3) * math.cos(ang)
        y1 = cy + (r + lw * 0.3) * math.sin(ang)
        x2 = x1 + size * 0.36 * math.cos(ang)
        y2 = y1 + size * 0.36 * math.sin(ang)
        cv.create_line(x1, y1, x2, y2, fill=color, width=lw, capstyle=tk.ROUND)
        return cv

    def _toggle_collapse(self):
        """Collapse/expand button ("▼"/"▲" between "✕" and the AI-model
        dropdown). Collapsing does NOT close the results -- it only
        pack_forget()s results_frame (so the Notebook/AI Search panes/AI
        Chat stay alive underneath, nothing is destroyed) and shrinks the
        window down to just the top bar (Searchbox + Update DB / AI Search
        / AI-model row), remembering the exact current geometry -- including
        any manual resize done via the "◢" grip -- so expanding again
        restores it exactly."""
        if not self.results_frame or not self.results_frame.winfo_exists():
            return
        win = self.root
        if not self._collapsed:
            # -- Collapse --
            self._pre_collapse_geometry = win.geometry()
            self.results_frame.pack_forget()
            m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", self._pre_collapse_geometry)
            if m:
                w, _h, x, y = m.groups()
            else:
                w, x, y = "900", str(self.x_pos), "5"
            # bg_f is a fixed height=35 frame (see __init__) -- collapsing
            # just shrinks the window to exactly that bar height.
            win.geometry(f"{w}x35+{x}+{y}")
            self.collapse_btn.config(text="▼")  # collapsed -- click to open back up
            self._collapsed = True
        else:
            # -- Expand --
            self.results_frame.pack(fill="both", expand=True, side="top")
            if self._pre_collapse_geometry:
                win.geometry(self._pre_collapse_geometry)
            self.collapse_btn.config(text="▲")  # expanded -- click to close it back down
            self._collapsed = False
        try:
            self.root.after(60, self._force_repaint)
        except Exception:
            pass

    def _close_results_window(self):
        """v2.8: results now render inside self.root instead of a separate
        Toplevel, so "closing" them means tearing down the results-only widgets
        (notebook, filter bars, AI/Advanced/Update DB buttons, resize grip) and
        shrinking the same window back down to the small floating search box --
        there is no second window to destroy/deiconify anymore."""
        self._in_results_mode = False
        try:
            if self.results_frame and self.results_frame.winfo_exists():
                self.results_frame.destroy()
        except Exception:
            pass
        self.results_frame = None
        self._online_chat_panel = None  # destroyed along with results_frame above
        try:
            if self._results_extra_bar and self._results_extra_bar.winfo_exists():
                self._results_extra_bar.destroy()
        except Exception:
            pass
        self._results_extra_bar = None
        try:
            self.close_btn.pack_forget()
        except Exception:
            pass
        try:
            self.collapse_btn.pack_forget()
        except Exception:
            pass
        self._collapsed = False
        self._pre_collapse_geometry = None
        # v4.4 fix: while results were showing, self._r_p (the ramp-light
        # frame) was repacked with after=self._results_extra_bar so it would
        # sit just left of the button group (see show_results()). Now that
        # _results_extra_bar has just been destroyed above, that "after="
        # anchor no longer exists -- simply forgetting/destroying the anchor
        # widget left r_p with a dangling packing reference and it stopped
        # being drawn at all (the ramp light vanishing after closing results).
        # Explicitly re-pack it back to its normal idle-mode spot (far right
        # edge of the search bar) so it's guaranteed visible again.
        try:
            self._r_p.pack_forget()
            self._r_p.pack(side="right", fill="y", padx=(1, 5), pady=5)
        except Exception:
            pass
        self.entry.bind("<Escape>", lambda e: self.root.destroy())
        try:
            self.root.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
        except Exception:
            pass
        self.active_result_win = None
        self._result_win = None
        self.entry_var.set("")
        try:
            self.root.geometry(f"{SMALL_SIZE}+{self.x_pos}+5")
            self.root.after(50, self.entry.focus_set)
            # v5.9: same DWM repaint nudge as _real_shrink() -- closing results
            # shrinks the window the same way, and could leave the ramp light
            # (correct color, just unpainted) looking like it had vanished.
            self.root.after(60, self._force_repaint)
        except Exception:
            pass

    def show_results(self, file_res, cont_res, query, should_exit, box_x, box_y, box_h, version, sem_res=None):
        if sem_res is None: sem_res = []
        # v2.8: results now render inside THIS SAME window (self.root) instead of
        # a separate Toplevel that popped up while root was withdrawn. That old
        # approach felt like a stutter/flash the instant the first character was
        # typed -- one whole OS window vanishing and a different one appearing.
        # Now the top search bar (self.bg_f, built once in __init__) never moves
        # or gets recreated; we just grow the window and add a results frame
        # below it. `win` is kept as a local alias so the rest of this (long)
        # function -- originally written against a standalone Toplevel -- needs
        # no further changes below this point.
        win = self.root
        self._in_results_mode = True
        self.res_edit = None
        # New Notebook instance below always defaults to its first-added tab
        # (File Name) anyway, so just clear the flag here for consistency.
        self._force_file_tab = False

        if should_exit:
            win.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
        else:
            win.protocol("WM_DELETE_WINDOW", self._close_results_window)
        win.geometry(f"{RESULT_SIZE}+{box_x}+{box_y}")

        self.active_result_win = win; self._result_win = win

        # Reveal the "✕" close button on the persistent search bar (hidden while
        # idle) and point it at the right close behavior for this session.
        _close_cmd = (lambda: self.root.destroy()) if should_exit else self._close_results_window
        self.close_btn.configure(command=_close_cmd)
        self.close_btn.pack(side="right", padx=(0, 12), pady=5, before=self._r_p)

        # Reveal the "▼/▲" collapse button, packed right after "✕" (before=
        # self._r_p, same anchor close_btn uses) so it lands between "✕" and
        # the AI-model cluster (_results_extra_bar, packed further below).
        # A brand-new results session always starts expanded.
        self._collapsed = False
        self._pre_collapse_geometry = None
        self.collapse_btn.config(text="▲")
        self.collapse_btn.pack(side="right", padx=(0, 4), pady=5, before=self._r_p)

        self.entry.bind("<Escape>", (lambda e: self.root.destroy()) if should_exit
                                      else (lambda e: self._close_results_window()))

        # v2.5 fix: focus_set() alone only sets *Tk-internal* focus — if the
        # window itself doesn't have OS-level window focus yet (e.g. this runs
        # from a background thread's after(0, ...) callback), keystrokes kept
        # landing nowhere until the user clicked the box manually. focus_force()
        # grabs real OS input focus for the window + widget, so typing continues
        # without a click.
        win.lift()
        win.focus_force()
        self.entry.focus_force()
        self.entry.icursor("end")

        # The results-only content (notebook, AI/Advanced/Update DB buttons,
        # resize grip) lives in its own frame below the persistent search bar,
        # so closing results is just "destroy this one frame" -- see
        # _close_results_window().
        self.results_frame = tk.Frame(win, bg=BG_COLOR)
        self.results_frame.pack(fill="both", expand=True, side="top")

        # AI Search / Advanced / Update DB buttons still live on the search bar
        # row (same as before), but grouped in their own sub-frame so they can
        # all be torn down together when results close instead of accumulating
        # on the persistent bar across searches.
        self._results_extra_bar = tk.Frame(self.bg_f, bg=BG_COLOR)
        self._results_extra_bar.pack(side="right")
        search_bar = self._results_extra_bar  # local alias used further below

        # v3.4: while results are showing, move the ramp light so it sits
        # between the Searchbox and the button group (Advanced/AI Search/
        # Update DB) instead of staying pinned at the far-right edge past all
        # the buttons. `after=` tells pack() to treat r_p as if it had been
        # packed right after search_bar, so it lands just to search_bar's
        # left (i.e. between Entry and the buttons) instead of at the very
        # end. When results close and search_bar is destroyed, r_p simply
        # settles back to the far-right edge on its own (idle-mode look,
        # unchanged from before).
        try:
            self._r_p.pack(side="right", padx=(1, 5), pady=5, after=self._results_extra_bar)
        except Exception:
            pass

        # v2.6: overrideredirect(True) removes the native titlebar/border, which
        # also removes the OS resize grip — the window used to be permanently
        # stuck at RESULT_SIZE (1350x700). Draw a small draggable "◢" handle in
        # the bottom-right corner that lets the user resize by hand; default
        # size/position stay exactly as before if the user never touches it.
        # Parented to results_frame so it's cleaned up automatically on close.
        _MIN_W, _MIN_H = 900, 500
        resize_grip = tk.Label(self.results_frame, text="◢", bg=BG_COLOR, fg="#666666",
                                font=("Segoe UI", 11, "bold"), cursor="size_nw_se")
        resize_grip.place(relx=1.0, rely=1.0, anchor="se", x=-2, y=-2)

        def _start_resize(e):
            self._resize_startx, self._resize_starty = e.x_root, e.y_root
            self._resize_startw = win.winfo_width()
            self._resize_starth = win.winfo_height()

        def _do_resize(e):
            if not (self.active_result_win and tk.Toplevel.winfo_exists(self.active_result_win)):
                return
            new_w = max(_MIN_W, self._resize_startw + (e.x_root - self._resize_startx))
            new_h = max(_MIN_H, self._resize_starth + (e.y_root - self._resize_starty))
            self.active_result_win.geometry(f"{new_w}x{new_h}")

        resize_grip.bind("<Button-1>", _start_resize)
        resize_grip.bind("<B1-Motion>", _do_resize)
        resize_grip.lift()  # keep the handle clickable above the Notebook
        self._resize_grip = resize_grip  # v1.2: re-lifted in _show_online_chat_panel too

        # v1.2: results_frame now uses grid (instead of just packing the
        # Notebook) to make room for the fixed online AI chat panel below
        # it -- row 0 (Notebook) expands to fill all remaining height, row 1
        # (chat panel) keeps a fixed height and is only shown when offline
        # AI Search is active (see _show_online_chat_panel/
        # _hide_online_chat_panel). resize_grip still uses place(), so it
        # doesn't conflict with grid here.
        self.results_frame.grid_rowconfigure(0, weight=1)
        self.results_frame.grid_rowconfigure(1, weight=0)
        self.results_frame.grid_columnconfigure(0, weight=1)
        self.nb = ttk.Notebook(self.results_frame)
        self.nb.grid(row=0, column=0, sticky="nsew")
        # v1.8: the online chat panel sits below the WHOLE Notebook (shared
        # across File Name/Folder Name/File Content), so without this it
        # kept showing even when the user switched to the Help tab, which
        # has its own "AI Chat History" pane instead (see ChatHistoryPanel)
        # and doesn't need the live chat floating below it. Hide/re-show it
        # based on which tab is actually selected.
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._sync_online_chat_panel_visibility())
        # The online chat panel is (re)created every time results open, but
        # only shown (grid) when offline AI Search is active -- hidden by
        # default.
        self._online_chat_panel = None
        # v-new (theo yêu cầu: AI Chat luôn hiển thị mặc định): trước đây
        # panel chat chỉ tự hiện khi self._ai_mode_active (tức offline "AI
        # Search" -- Jina-v3/BGE-Gemma2 -- đã được bật) dù bản thân AI Chat
        # KHÔNG hề phụ thuộc vào kết quả embedding đó, nó chỉ đọc
        # self._last_bm25_cont_res (BM25 -- xem _online_chat_worker). Luôn
        # hiện panel ngay khi cửa sổ kết quả mở, bất kể AI Search có bật
        # hay không.
        self.root.after(50, lambda: (self._show_online_chat_panel(),
                                      self._update_chat_placeholder(query)))
        # v-fix (AI Chat vẫn cần bấm AI Search ở LẦN SEARCH ĐẦU TIÊN dù đã
        # sửa để panel luôn hiện): panel hiện rồi không có nghĩa là nó đã
        # TỰ TÓM TẮT -- việc tự tóm tắt trước đây chỉ được kích từ
        # _smart_search_realtime (chạy SAU khi show_results ở đây đã tạo
        # cửa sổ), hoặc từ nút AI Search. show_results() là điểm vào DUY
        # NHẤT & chắc chắn có mọi lần 1 cửa sổ kết quả MỚI được tạo, bất kể
        # nó được gọi từ đường nào -- gọi thẳng _auto_ai_chat_after_search
        # ở đây luôn, không phụ thuộc caller nào khác có nhớ gọi nó hay
        # không.
        self.root.after(200, lambda: self._auto_ai_chat_after_search(query))
        only_files = [r for r in file_res if str(r[0]).lower() != "folder"]
        only_folders = [r for r in file_res if str(r[0]).lower() == "folder"]
        self._all_files_data = only_files
        self._all_content_data = cont_res
        self._ai_cont_res  = []
        self._ai_file_res  = []
        self.size_op_var.set(">"); self.size_num_var.set(""); self.size_unit_var.set("MB")
        self.ext_filter_var.set("")
        self.c_size_op_var.set(">"); self.c_size_num_var.set(""); self.c_size_unit_var.set("MB")
        self.c_ext_filter_var.set("")
        self.name_filter_var.set("")
        self.c_name_filter_var.set("")

        # v2.3: Tab order — File Name (1st), Folder Name (2nd), File Content (last/right)
        # Tabs are added to notebook in correct order below via nb.add
        # File Name + Folder Name: MFT realtime (no DB needed)
        # File Content: BM25/AI results, populated after --update data

        # 1. TAB CONTENT (frame created here, added to nb LAST after f/fol frames below)
        c_frame = tk.Frame(self.nb)

        # ── Content Filter bar ─────────────────────────────────────────────
        c_filter_bar = tk.Frame(c_frame, bg=BG_COLOR, pady=2)
        c_filter_bar.pack(fill="x", padx=6, pady=(4, 2))
        c_filter_bar.pack_propagate(False)
        c_filter_bar.config(height=32)

        tk.Label(c_filter_bar, text="Size:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(4, 2))

        c_op_menu = ttk.OptionMenu(c_filter_bar, self.c_size_op_var, self.c_size_op_var.get(),
                                    "Any", ">", ">=", "<", "<=", "=",
                                    command=lambda *_: self._apply_content_filter_to_tree())
        c_op_menu.config(width=4); c_op_menu.pack(side="left", padx=2)

        c_size_entry = tk.Entry(c_filter_bar, textvariable=self.c_size_num_var, width=7,
                                 font=("Segoe UI", 9), bg=ENTRY_BG, fg=TEXT_COLOR,
                                 insertbackground=TEXT_COLOR, bd=1, relief="flat")
        c_size_entry.pack(side="left", padx=2)
        self.c_size_num_var.trace_add("write", self._apply_content_filter_to_tree)

        c_unit_menu = ttk.OptionMenu(c_filter_bar, self.c_size_unit_var, self.c_size_unit_var.get(),
                                      "B", "KB", "MB", "GB",
                                      command=lambda *_: self._apply_content_filter_to_tree())
        c_unit_menu.config(width=4); c_unit_menu.pack(side="left", padx=2)

        # ── Content Ext filter ─────────────────────────────────────────────
        tk.Label(c_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 11)).pack(side="left", padx=4)
        tk.Label(c_filter_bar, text="Ext:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 3))

        _C_EXT_PRESETS = [
            ("All",   ""),
            ("PDF",   "pdf"),
            ("Word",  "doc,docx"),
            ("Excel", "xls,xlsx,csv"),
            ("PPT",   "ppt,pptx"),
            ("Img",   "png,jpg,jpeg,bmp,gif"),
            ("Txt",   "txt,log"),
            ("Msg",   "msg"),
            ("One",   "one"),
        ]
        def _set_c_ext(val):
            self.c_ext_filter_var.set(val)
        for _lbl, _val in _C_EXT_PRESETS:
            tk.Button(c_filter_bar, text=_lbl, font=("Segoe UI", 8),
                      bg="#e6e8ec", fg="#33363c", bd=0, padx=5, pady=1,
                      activebackground="#d3d6dc", cursor="hand2",
                      command=lambda v=_val: _set_c_ext(v)).pack(side="left", padx=1)

        tk.Label(c_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)

        c_ext_entry = tk.Entry(c_filter_bar, textvariable=self.c_ext_filter_var, width=14,
                                font=("Segoe UI", 9), bg=ENTRY_BG, fg="#888",
                                insertbackground=TEXT_COLOR, bd=1, relief="flat")
        c_ext_entry.pack(side="left", padx=(0, 2))
        def _c_ext_focus_in(e):
            if not self.c_ext_filter_var.get(): c_ext_entry.config(fg=TEXT_COLOR)
        def _c_ext_focus_out(e):
            if not self.c_ext_filter_var.get(): c_ext_entry.config(fg="#888")
        c_ext_entry.bind("<FocusIn>",  _c_ext_focus_in)
        c_ext_entry.bind("<FocusOut>", _c_ext_focus_out)
        self.c_ext_filter_var.trace_add("write", self._apply_content_filter_to_tree)

        tk.Button(c_filter_bar, text="✕", command=lambda: self.c_ext_filter_var.set(""),
                  bg="#c9ccd2", fg="#222222", font=("Segoe UI", 8), bd=0, padx=4,
                  activebackground="#b0b3ba", cursor="hand2").pack(side="left", padx=(0, 4))

        # ── Whole word toggle (greyed out — layout parity with File Name tab) ─
        # Content search matches on tokenized/BM25 terms rather than raw
        # substrings, so "Whole word" has no meaning here. Kept in the UI
        # (disabled) purely so Size/Ext/Whole word/Name all line up at the
        # same horizontal position across File Name / Folder Name / File
        # Content tabs.
        tk.Label(c_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        _c_ww_cb = tk.Checkbutton(c_filter_bar, text="Whole word", variable=self.c_whole_word_dummy_var,
                                    bg=BG_COLOR, fg="#666", selectcolor=ENTRY_BG,
                                    font=("Segoe UI", 9), state="disabled")
        _c_ww_cb.pack(side="left", padx=(0, 4))
        add_tooltip(_c_ww_cb, "Whole word doesn't apply to File Content search — content matching uses its own tokenization.")
        # ──────────────────────────────────────────────────────────────────

        # ── Content Name filter ────────────────────────────────────────────
        tk.Label(c_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        tk.Label(c_filter_bar, text="Name:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 2))
        c_name_entry = tk.Entry(c_filter_bar, textvariable=self.c_name_filter_var, width=18,
                                 font=("Segoe UI", 9), bg=ENTRY_BG, fg="#888",
                                 insertbackground=TEXT_COLOR, bd=1, relief="flat")
        c_name_entry.pack(side="left", padx=(0, 2))
        def _c_name_focus_in(e):
            if not self.c_name_filter_var.get(): c_name_entry.config(fg=TEXT_COLOR)
        def _c_name_focus_out(e):
            if not self.c_name_filter_var.get(): c_name_entry.config(fg="#888")
        c_name_entry.bind("<FocusIn>",  _c_name_focus_in)
        c_name_entry.bind("<FocusOut>", _c_name_focus_out)
        self.c_name_filter_var.trace_add("write", self._apply_content_filter_to_tree)
        tk.Button(c_filter_bar, text="✕", command=lambda: self.c_name_filter_var.set(""),
                  bg="#c9ccd2", fg="#222222", font=("Segoe UI", 8), bd=0, padx=4,
                  activebackground="#b0b3ba", cursor="hand2").pack(side="left", padx=(0, 4))
        # ──────────────────────────────────────────────────────────────────

        # v9.13.7 / v9.16: "Search files" (was "Search image") -- extract
        # text from one picked file (OCR for images, direct read for
        # text/Office/PDF/etc. via get_file_content()) and search with it.
        # Placed here (not the main search bar) because File Content is the
        # only tab this is actually useful for -- filename/folder search
        # wouldn't do anything useful with extracted body text.
        tk.Button(c_filter_bar, text="+", font=("Segoe UI", 11, "bold"),
                  bg=BG_COLOR, fg="#7ec8e3", bd=0, activebackground=BG_COLOR,
                  cursor="hand2", command=lambda: self._search_by_image()
                  ).pack(side="right", padx=(0, 4))
        tk.Label(c_filter_bar, text="Search files:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="right", padx=(6, 2))

        self.content_filter_count_label = None  # removed

        # ── AI Search button ───────────────────────────────────────────────────
        _ai_btn_state = "normal"
        self._hybrid_status_lbl = None  # removed from UI

        def _on_ai_search():
            q_now = self._last_query or query
            if not q_now: return
            # Disable button immediately to prevent double-click while running
            try:
                self._ai_search_btn.config(state="disabled", text="⏳ Running...",
                                            fg="#ffcc00", bg="#2a2a1a")
            except Exception: pass
            # v1.2: AI Search always uses the offline model (Jina-v3/
            # BGE-Gemma2) -- chatting with the online AI (Gemini Flash /
            # GPT-OSS) now lives in its own chat panel below
            # (self._online_chat_panel), no longer tied to the AI Search
            # button / Online checkbox (see _build_online_chat_panel /
            # _send_online_chat_msg).
            # v1.6: show the chat panel and fire its own background summary
            # request RIGHT NOW, in parallel with the offline embedding
            # search below -- it only needs the already-available BM25
            # content excerpts (self._last_bm25_cont_res), not the AI
            # Search embedding results, so there's no reason to wait for
            # _ai_search_and_update() to finish first (that used to make
            # the chat look sequential/slower than AI Search).
            self._show_online_chat_panel()
            self._auto_ai_chat_after_search(q_now)
            threading.Thread(target=self._ai_search_and_update, args=(q_now,), daemon=True).start()

        # ── AI Model selector: dropdown to pick one of 3 models ────────────
        # jina_v3 / bge_gemma2 — each model has different embedding dimension/
        # vector space so semantic data is stored in separate tables in DB
        # (semantic_index / semantic_index_jina_v3 / semantic_index_bge_gemma2).
        # Switching does not delete other model data — just run --update data
        # ONCE per model, then switch freely without rebuilding.
        _model_keys_order = ["jina_v3", "bge_gemma2"]
        _model_display = [SEMANTIC_MODELS[k]["label"] for k in _model_keys_order]
        _display_to_key = {SEMANTIC_MODELS[k]["label"]: k for k in _model_keys_order}

        def _check_table_has_data(model_key):
            try:
                table = SEMANTIC_MODELS.get(model_key, {}).get("table", "semantic_index")
                conn = sqlite3.connect(DB_FILE, timeout=5)
                c = conn.cursor()
                c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table,))
                has_table = c.fetchone()[0] > 0
                count = 0
                if has_table:
                    c.execute(f"SELECT count(*) FROM {table}")
                    count = c.fetchone()[0]
                conn.close()
                return count
            except Exception:
                return 0

        def _refresh_ai_status_lbl(model_key, ok=None):
            label = SEMANTIC_MODELS.get(model_key, {}).get("label", model_key)
            data_count = _check_table_has_data(model_key)
            try:
                if ok is False:
                    self._ai_model_status_lbl.config(text=f"✗ {label}: load error", fg="#ff6666")
                else:
                    # v5.8: dropped "✓ <model>: ready" — the ramp light going
                    # Blue already communicates "ready" for the current
                    # model, so this label stayed silent otherwise (and for
                    # a genuinely-not-ready model, the dropdown/AI-Search/
                    # Advanced buttons already grey out via _sync_ai_adv_lock).
                    self._ai_model_status_lbl.config(text="", fg="#888888")
            except Exception:
                pass
            return data_count

        def _on_ai_model_change(event=None):
            # v1.2: this dropdown now only picks between the 2 offline
            # models (Jina-v3/BGE-Gemma2) -- choosing the online AI (Gemini/
            # GPT-OSS) has fully moved to the separate chat panel below, no
            # longer sharing a dropdown/checkbox with AI Search.
            global _sem_model_key
            new_key = _display_to_key.get(self._ai_model_combo.get(), DEFAULT_SEMANTIC_MODEL)
            self.ai_model_var.set(new_key)
            _sem_model_key = new_key
            # v-new (yêu cầu: nhớ AI Search model đã chọn qua configure.ini):
            _config_set("Models", "ai_search_model", new_key)

            data_count = _check_table_has_data(new_key)
            label = SEMANTIC_MODELS[new_key]["label"]
            if data_count == 0:
                # No embedding for this model yet — inform user to run --update data once
                messagebox.showinfo(
                    "AI Model Switch",
                    f"Switched to model: {label}\n\n"
                    f"This model has NO semantic data in search_data.db yet.\n"
                    f"Run \"--update data\" ONCE to build the AI index for this model.\n\n"
                    f"After that, you can switch freely between all 3 models "
                    f"without rebuilding (unless you re-scan all drives)."
                )
            # Load new model in background immediately so the next AI search has no delay.
            # jina-v3/bge-gemma2 on CPU can take a long time (especially bge-gemma2, ~9B params)
            # so show loading status clearly so the user knows it's loading, not frozen.
            def _bg_load():
                ok = _load_semantic_model(new_key)
                self.root.after(0, lambda: _refresh_ai_status_lbl(new_key, ok))
                # v5.8: if AI Search results are currently on screen, re-run
                # the semantic search with the newly selected model instead
                # of silently leaving the OLD model's results displayed.
                # Without this, switching Jina v3 -> BGE Gemma2 kept showing
                # Jina's hits/count until the user manually clicked BM25
                # then AI Search again -- calling _ai_search_and_update
                # directly here would just TOGGLE IT OFF (that method treats
                # a second call while _ai_mode_active as "restore BM25"), so
                # _ai_mode_active is cleared first to force a fresh run.
                if ok and data_count and self._ai_mode_active:
                    q_now = self._last_query or query
                    if q_now:
                        def _rerun_btn_state():
                            try:
                                self._ai_search_btn.config(state="disabled", text="⏳ Running...",
                                                            fg="#ffcc00", bg="#2a2a1a")
                            except Exception: pass
                        self.root.after(0, _rerun_btn_state)
                        self._ai_mode_active = False
                        self._ai_search_and_update(q_now)
            self._ai_model_status_lbl.config(text=f"⏳ loading {label}...", fg="#ffcc00")
            threading.Thread(target=_bg_load, daemon=True).start()

        # v2.6: moved onto the top search_bar (same row as the Searchbox) instead of
        # the per-tab c_filter_bar, per user request — stays visible on every tab and
        # no longer eats vertical space above the Content tab's results tree.
        ai_model_frame = tk.Frame(search_bar, bg=BG_COLOR)
        ai_model_frame.pack(side="right", padx=(2, 4))

        # v1.2: the "🌐 Online" checkbox is gone, along with using this same
        # dropdown to pick Gemini/GPT-OSS -- AI Search now ALWAYS runs
        # offline (Jina-v3/BGE-Gemma2). Chatting with the online AI moved to
        # its own chat panel below the results (see
        # _build_online_chat_panel), defaulting to Gemini Flash and
        # auto-switching to GPT-OSS on rate-limit, no manual picking needed.
        self._ai_model_combo = ttk.Combobox(
            ai_model_frame, values=_model_display,
            state="disabled", width=16, font=("Segoe UI", 8))
        _cur_key = self.ai_model_var.get() if self.ai_model_var.get() in SEMANTIC_MODELS else DEFAULT_SEMANTIC_MODEL
        self._ai_model_combo.set(SEMANTIC_MODELS[_cur_key]["label"])
        self._ai_model_combo.pack(side="left")
        self._ai_model_combo.bind("<<ComboboxSelected>>", _on_ai_model_change)

        self._ai_model_status_lbl = tk.Label(
            ai_model_frame, text="", font=("Segoe UI", 7, "italic"),
            bg=BG_COLOR, fg="#888888")
        self._ai_model_status_lbl.pack(side="left", padx=(4, 0))
        # Show data availability status for the current model when the result window opens
        _refresh_ai_status_lbl(_cur_key)
        # ──────────────────────────────────────────────────────────────────

        self._ai_search_btn = tk.Button(
            search_bar, text="🤖 AI Search",
            font=("Segoe UI", 8, "bold"),
            bg="#1e3a5f", fg="#7ec8e3",
            relief="raised", bd=2,
            padx=8, pady=2,
            activebackground="#2a5080", activeforeground="#a8d8f0",
            cursor="hand2",
            # v1.2: AI Search is always offline -- waits for
            # _sync_ai_adv_lock() to unlock once the ramp light turns Blue
            # (BM25 + offline embedding ready).
            state="disabled",
            command=_on_ai_search)
        self._ai_search_btn.pack(side="right", padx=(4, 2), pady=4)
        def _ai_search_btn_tooltip_text():
            if not ENABLE_AI_SEARCH_FEATURE:
                return "AI Search is disabled in this build."
            return "AI-powered semantic search (slower, smarter). Click again anytime to re-run."
        add_tooltip(self._ai_search_btn, _ai_search_btn_tooltip_text)

        # ── Advanced button (pagination) ───────────────────────────────────
        # Page 0 = realtime top 100. Each click appends next 500. Last page → reset to page 0.
        PAGE_SIZE = 500

        def _adv_label():
            page = self._adv_page
            total = len(self._adv_all_cont) if hasattr(self, '_adv_all_cont') else 0
            total_f = len(self._adv_all_files) if hasattr(self, '_adv_all_files') else 0
            if page == 0:
                return "Advanced"
            return "Simple ↩"

        def _on_advanced():
            q_now = self._last_query or query
            if not q_now: return

            # First click on page 0: run full search to get all results
            if self._adv_page == 0:
                try:
                    self._adv_search_btn.config(state="disabled", text="⏳ Loading...", fg="#ffcc00", bg="#2a2a1a")
                except Exception: pass

                def _run_full_search():
                    # Run search with adv_mode=True to get full result set
                    import sqlite3 as _sq3
                    try:
                        conn = self.db_conn
                        if conn is None: return
                        c2 = conn.cursor()
                        from rank_bm25 import BM25Okapi
                        import re as _re2

                        cleaned_q2, op2, size_val2 = parse_size_filter(q_now)
                        kw2 = [k.lower() for k in cleaned_q2.split() if k]
                        kw2 = _strip_stopwords(kw2)  # v5.8: see _smart_search_realtime comment
                        size_clause2 = f" AND size {op2} ?" if op2 else ""
                        size_extra2  = [size_val2] if op2 else []
                        has_anchor2  = any(_is_anchor_kw(k) for k in kw2)

                        # File Name full search
                        file_res2 = []
                        if has_anchor2:
                            nc = ["name LIKE ?" for k in kw2]
                            np2 = [f"%{k}%" for k in kw2]
                            c2.execute("SELECT type, name, path, size FROM files WHERE type != 'Folder' AND " + " AND ".join(nc) + size_clause2 + " LIMIT 5000", np2 + size_extra2)
                            fn2 = c2.fetchall()
                            if not fn2 and len(kw2) > 1:
                                c2.execute("SELECT type, name, path, size FROM files WHERE type != 'Folder' AND (" + " OR ".join(nc) + ")" + size_clause2 + " LIMIT 5000", np2 + size_extra2)
                                fn2 = c2.fetchall()
                            c2.execute("SELECT type, name, path, size FROM files WHERE type = 'Folder' AND " + " AND ".join(nc) + size_clause2 + " LIMIT 5000", np2 + size_extra2)
                            file_res2 = fn2 + c2.fetchall()
                            # v7.10: same "Whole word" post-filter as the main
                            # realtime DB path -- SQL LIKE can't express word
                            # boundaries, so filter the substring hits in Python.
                            if self.whole_word_var.get():
                                file_res2 = [r for r in file_res2
                                              if all(_kw_matches_with_glued_fallback(k, r[1].lower(), True) for k in kw2)]

                        # Content full search via FTS
                        import re as _re3
                        _STOP2 = {'a','an','the','is','in','on','at','to','of','or','and','as','be','by','do','for','has','had','he','her','him','his','how','i','if','it','its','me','my','no','not','off','our','out','own','so','than','that','them','then','they','this','us','was','we','who','why','will','with','you','your','also','been','but','can','did','does','from','get','got','have','into','just','may','new','now','one','see','set','she','time','what','when','which','would'}
                        FTS5_OPS = set('+-*:^"()：、。・<>@[]{}|\\/?!#$%&=~`\'')
                        # NOTE: same trigram-tokenizer limitation as the realtime search —
                        # FTS terms under 3 chars silently match zero rows, so this stays
                        # at len>=3; short CJK anchors fall through to the LIKE path below.
                        def _safe(k): return len(k)>=3 and not any(ch in FTS5_OPS for ch in k) and '-' not in k and not _re3.search(r'[^\w　-鿿＀-￯一-鿿]', k)
                        phrase2 = cleaned_q2.strip()
                        cont_res2 = []
                        if _is_anchor_kw(phrase2):
                            # v-fix: same CJK-punctuation gap as the realtime path above.
                            pt2 = [k for k in _re3.split(r'[\s、。！？「」『』（）:+\-<>@"()#.\[\]{}|\/?!&=~`]+', phrase2) if k]
                            fts_tok2 = [k for k in pt2 if len(k)>=3 and k.lower() not in _STOP2 and _safe(k)]
                            if fts_tok2:
                                fts_q2 = " AND ".join(f'"{k}"' for k in fts_tok2)
                                try:
                                    c2.execute("SELECT path FROM content_index WHERE content MATCH ? LIMIT 5000", (fts_q2,))
                                    cands2 = set(r[0] for r in c2.fetchall())
                                    if cands2:
                                        ph2 = ",".join("?"*len(cands2))
                                        c2.execute(f"SELECT f.path, f.size FROM files f WHERE f.path IN ({ph2})", list(cands2))
                                        cont_res2 = c2.fetchall()
                                except Exception: pass
                            else:
                                # No token was usable by the FTS trigram index (e.g. a
                                # short CJK anchor like 解析 — trigram needs >= 3 chars
                                # and would otherwise silently return zero results).
                                # Fall back to a direct LIKE scan on content_store —
                                # slower (full table scan) but has no length floor.
                                try:
                                    c2.execute("SELECT path FROM content_store WHERE content LIKE ? LIMIT 5000", (f"%{phrase2}%",))
                                    cands2 = set(r[0] for r in c2.fetchall())
                                    if cands2:
                                        ph2 = ",".join("?"*len(cands2))
                                        c2.execute(f"SELECT f.path, f.size FROM files f WHERE f.path IN ({ph2})", list(cands2))
                                        cont_res2 = c2.fetchall()
                                except Exception: pass

                        # BM25 rank file results
                        def _tok2(t): return [x.lower() for x in _re2.split(r'[\s\W]+', str(t)) if len(x)>=2]
                        qt2 = _tok2(cleaned_q2)
                        if file_res2 and qt2:
                            try:
                                corp_f = [_tok2(os.path.basename(r[2])+" "+r[2]) for r in file_res2]
                                b2f = BM25Okapi(corp_f)
                                sc_f = b2f.get_scores(qt2)
                                file_res2 = [file_res2[i] for i in sorted(range(len(file_res2)), key=lambda i: sc_f[i], reverse=True)]
                            except Exception: pass

                        # BM25 rank content results by filename
                        bm25_c2 = {}
                        sz_map2 = {r[0]: r[1] for r in cont_res2}
                        if cont_res2 and qt2:
                            try:
                                po2, corp_c2 = [], []
                                for p in [r[0] for r in cont_res2]:
                                    fn = os.path.basename(p); par = os.path.basename(os.path.dirname(p)); gp = os.path.basename(os.path.dirname(os.path.dirname(p)))
                                    po2.append(p); corp_c2.append(_tok2(f"{fn} {par} {gp}"))
                                b2c = BM25Okapi(corp_c2)
                                sc_c = b2c.get_scores(qt2)
                                mx2 = max(sc_c) if max(sc_c) > 0 else 1.0
                                bm25_c2 = {po2[i]: sc_c[i]/mx2 for i in range(len(po2))}
                            except Exception: pass

                        hyb2 = sorted([(p, sz_map2.get(p,0), bm25_c2.get(p,0.0)) for p in sz_map2], key=lambda x: x[2], reverse=True)
                        mx_h2 = hyb2[0][2] if hyb2 else 1.0
                        cont_res2_scored = [(p, sz, int((sc/mx_h2)*99) if mx_h2>0 else 0) for p,sz,sc in hyb2]

                        # Priority sort: office/pdf/msg first, txt/md middle, log/html/code/... last
                        cont_res2_scored = self._sort_priority(cont_res2_scored, 0)
                        file_res2        = self._sort_priority(file_res2, 1)

                        # Store full results
                        self._adv_all_cont  = cont_res2_scored
                        self._adv_all_files = file_res2

                        def _apply_page1():
                            self._adv_page = 1
                            # Display ALL results immediately (no pagination)
                            all_cont  = self._adv_all_cont
                            all_files = self._adv_all_files
                            self._all_content_data = all_cont
                            self._all_files_data   = [r for r in all_files if str(r[0]).lower() != "folder"]
                            only_all_f   = [r for r in all_files if str(r[0]).lower() != "folder"]
                            only_all_fol = [r for r in all_files if str(r[0]).lower() == "folder"]
                            # Clear current tree and reload everything
                            for item in self.tree_c.get_children():   self.tree_c.delete(item)
                            for item in self.tree_f.get_children():   self.tree_f.delete(item)
                            for item in self.tree_fol.get_children(): self.tree_fol.delete(item)
                            for i, item in enumerate(all_cont):
                                try: self._insert_row(self.tree_c, item, i+1, 0)
                                except: pass
                            for i, item in enumerate(only_all_f):
                                try: self._insert_row(self.tree_f, item, i+1, 1)
                                except: pass
                            for i, item in enumerate(only_all_fol):
                                try: self._insert_row(self.tree_fol, item, i+1, 2)
                                except: pass
                            try:
                                if self._adv_search_btn and self._adv_search_btn.winfo_exists():
                                    self._adv_search_btn.config(state="normal", text=_adv_label(), fg="#90ee90", bg="#1a3a1a")
                                if self.content_filter_count_label:
                                    self.content_filter_count_label.config(text=f"{len(all_cont)} files")
                                if self.filter_count_label:
                                    self.filter_count_label.config(text=f"{len(only_all_f)} files")
                            except Exception: pass
                            # Open Advanced split INSIDE tab
                            self._adv_mode_active = True
                            self.root.after(0, self._open_adv_split_win)
                        self.root.after(0, _apply_page1)
                    except Exception as _e2:
                        print(f"[Advanced] Error: {_e2}")
                        try:
                            self.root.after(0, lambda: self._adv_search_btn.config(state="normal", text="Advanced", fg="#33363c", bg="#e6e8ec"))
                        except: pass

                threading.Thread(target=_run_full_search, daemon=True).start()
                return

            # Second click: reset to realtime (page 0) — show "Simple" mode
            self._adv_page = 0
            self._adv_mode_active = False
            # Restore single tree in File Content tab
            self._close_adv_split()
            # Clear trees and reload realtime results only
            rt_c  = self._last_bm25_cont_res or []
            rt_f  = self._last_bm25_file_res or []
            self._all_content_data = rt_c
            self._all_files_data   = [r for r in rt_f if str(r[0]).lower() != "folder"]
            only_rt_f   = [r for r in rt_f if str(r[0]).lower() != "folder"]
            only_rt_fol = [r for r in rt_f if str(r[0]).lower() == "folder"]
            for item in self.tree_c.get_children():   self.tree_c.delete(item)
            for item in self.tree_f.get_children():   self.tree_f.delete(item)
            for item in self.tree_fol.get_children(): self.tree_fol.delete(item)
            for i, item in enumerate(rt_c):
                try: self._insert_row(self.tree_c, item, i+1, 0)
                except: pass
            for i, item in enumerate(only_rt_f):
                try: self._insert_row(self.tree_f, item, i+1, 1)
                except: pass
            for i, item in enumerate(only_rt_fol):
                try: self._insert_row(self.tree_fol, item, i+1, 2)
                except: pass
            try:
                self._adv_search_btn.config(text="Advanced", fg="#33363c", bg="#e6e8ec", state="normal")
                if self.content_filter_count_label:
                    self.content_filter_count_label.config(text=f"{len(rt_c)} files")
                if self.filter_count_label:
                    self.filter_count_label.config(text=f"{len(only_rt_f)} files")
            except Exception: pass
            return

        self._adv_search_btn = tk.Button(
            search_bar, text="Advanced",
            font=("Segoe UI", 8, "bold"),
            bg="#e6e8ec", fg="#33363c",
            relief="raised", bd=2,
            padx=8, pady=2,
            activebackground="#d3d6dc", activeforeground="#111111",
            cursor="hand2",
            command=_on_advanced)
        # v9.15: Advanced button HIDDEN per user request — the default
        # (Simple) search now always shows the fuller/looser result set that
        # Advanced used to require a second click for (see FTS_LIMIT/
        # DISPLAY_LIMIT/BM25_THRESHOLD notes in _smart_search_realtime), so
        # the button's split-pane "show more results" behavior is redundant
        # for now. The widget above is still fully created and wired to
        # _on_advanced — to bring it back, just uncomment the .pack() line
        # below (and its tooltip) and nothing else needs to change.
        # self._adv_search_btn.pack(side="right", padx=(4, 2), pady=4)
        # add_tooltip(self._adv_search_btn,
        #             lambda: ("Back to simple results (top matches only)"
        #                      if getattr(self, "_adv_mode_active", False)
        #                      else "Show more results (paginate through all matches)"))
        # Apply current ramp-light status now that Advanced / AI Search / model
        # combobox all exist (they're rebuilt fresh every time show_results runs).
        self._sync_ai_adv_lock()
        # ──────────────────────────────────────────────────────────────────

        # ── Update DB button (same action as typing "--update data") ────────
        # Packed AFTER Advanced so it lands to Advanced's left — i.e. right next
        # to the Searchbox, between the Searchbox and the Advanced button.
        # v3.4: also reads the Searchbox for a bare tier expression — type
        # "tier 1,2" (or just "1,2") then click this button instead of typing
        # the full "--update data tier 1,2" command and pressing Enter.
        # If the box is empty or holds an ordinary search query, behaves
        # exactly as before (full Tier 1-4 scan).
        def _run_update_db(selected_tiers, selected_models=None, ocr_enabled=False, force_reindex=False):
            """Actually kick off the indexing_worker thread. Extracted out of
            the old _on_update_db so both the new dialog's 'Start Update'
            button and (if ever needed again) any other caller can trigger a
            scan the same way."""
            if self._update_db_running or self._mailnotes_update_running:
                return
            tier_desc = ("ALL (1-4)" if selected_tiers is None
                         else ",".join(str(t + 1) for t in sorted(selected_tiers)))
            try:
                self._update_db_btn.config(state="disabled", text="Updating...",
                                            fg="#ffcc00", bg="#2a2a1a")
            except Exception:
                pass
            self._update_db_running = True  # lock AI Search/model combo immediately, don't wait for the thread
            self._set_index_status(f"Updating (tier {tier_desc})...", "#ffcc00")
            self._ramp_blink_start()
            self.entry_var.set("")
            threading.Thread(target=self.indexing_worker,
                              args=(selected_tiers, selected_models, ocr_enabled, force_reindex), daemon=True).start()

        if not hasattr(self, "_model_abort_events"):
            self._model_abort_events = {}  # model_key -> threading.Event(), set on Abort click

        def _install_model_online(model_key, status_lbl, install_btn, dialog, cb=None, mv=None):
            """Download a model's weights from HuggingFace Hub into
            models/<dir_name>/ next to the app. Requires internet + the
            huggingface_hub package. Runs in a background thread so the
            dialog stays responsive; UI updates are marshalled back via
            self.root.after(0, ...) since Tk isn't thread-safe.
            v5.9: while downloading, the button becomes "Abort" (rather than
            just greyed out) — huggingface_hub's snapshot_download can't be
            killed mid-transfer from here, but clicking Abort immediately
            resets the button/status and flags the in-flight download as
            abandoned, so its eventual result (success or error) is silently
            ignored instead of overwriting the UI once it finally returns."""
            info = SEMANTIC_MODELS.get(model_key, {})
            label = info.get("label", model_key)
            repo = info.get("hf_repo")
            if not repo:
                messagebox.showerror("Install model", f"No download source configured for {label}.", parent=dialog)
                return
            if not messagebox.askyesno(
                "Install model — internet required",
                f"Download {label} now from HuggingFace ({repo})?\n\n"
                f"This requires internet access and may take a while /\n"
                f"use significant disk space (several GB for larger models).",
                parent=dialog):
                dialog.lift()
                return
            dest_dir = os.path.join(_MODEL_ROOT_CANDIDATES[0], info["dir_names"][0])
            abort_event = threading.Event()
            self._model_abort_events[model_key] = abort_event

            def _do_abort():
                abort_event.set()
                try:
                    status_lbl.config(text="cancelled", fg="#888888")
                    install_btn.config(state="normal", text="Install online...",
                                        command=lambda: _install_model_online(model_key, status_lbl, install_btn, dialog, cb, mv))
                except Exception:
                    pass

            try:
                install_btn.config(state="normal", text="Abort", command=_do_abort)
                status_lbl.config(text="downloading...", fg="#ffcc00")
            except Exception:
                pass

            def _bg():
                try:
                    from huggingface_hub import snapshot_download
                except ImportError:
                    def _fail():
                        if abort_event.is_set():
                            return
                        status_lbl.config(text="huggingface_hub not installed", fg="#ff6666")
                        install_btn.config(state="normal", text="Install online...",
                                            command=lambda: _install_model_online(model_key, status_lbl, install_btn, dialog, cb, mv))
                        messagebox.showerror(
                            "Install model",
                            "The 'huggingface_hub' package isn't installed.\n"
                            "Run: pip install huggingface_hub --break-system-packages",
                            parent=dialog)
                        dialog.lift()
                    self.root.after(0, _fail)
                    return
                try:
                    os.makedirs(dest_dir, exist_ok=True)
                    snapshot_download(repo_id=repo, local_dir=dest_dir)
                    if abort_event.is_set():
                        return  # user aborted meanwhile — UI already reset, don't touch it
                    def _ok():
                        SEMANTIC_MODEL_DIRS[model_key] = dest_dir
                        status_lbl.config(text="✓ installed", fg="#4caf50")
                        install_btn.config(state="normal", text="Reinstall...",
                                           command=lambda: _install_model_online(model_key, status_lbl, install_btn, dialog, cb, mv))
                        # v7.10: model just became installed -- un-grey the
                        # checkbox and tick it by default so the user doesn't
                        # have to close/reopen this dialog to use it this run.
                        if cb is not None:
                            try: cb.config(state="normal")
                            except Exception: pass
                        if mv is not None:
                            try: mv.set(True)
                            except Exception: pass
                        messagebox.showinfo("Install model", f"{label} installed successfully.", parent=dialog)
                        dialog.lift()
                    self.root.after(0, _ok)
                except Exception as e:
                    if abort_event.is_set():
                        return  # aborted — the exception is just the interrupted transfer, ignore it
                    err = str(e)
                    def _err():
                        status_lbl.config(text="download failed", fg="#ff6666")
                        install_btn.config(state="normal", text="Install online...",
                                            command=lambda: _install_model_online(model_key, status_lbl, install_btn, dialog, cb, mv))
                        messagebox.showerror("Install model", f"Download failed for {label}:\n{err}", parent=dialog)
                        dialog.lift()
                    self.root.after(0, _err)

            threading.Thread(target=_bg, daemon=True).start()

        def _install_vi_diacritics_online(status_lbl, install_btn, dialog):
            """Download the Vietnamese diacritics-restoration model (base
            model + LoRA adapter, two separate HF repos) into
            models/vi-diacritics/{base,adapter}/. Same safe
            snapshot_download-into-a-folder mechanism as
            _install_model_online, just for 2 repos instead of 1, and with
            ignore_patterns to skip redundant TensorFlow/Flax/ONNX weight
            copies that vinai/bartpho-syllable ships alongside the PyTorch
            ones we actually use (this is what made an earlier bundled-into-
            the-exe attempt balloon to several GB for no benefit)."""
            if not messagebox.askyesno(
                "Install Vietnamese diacritics restoration — internet required",
                "Download the Vietnamese diacritics-restoration model now?\n\n"
                "Improves AI Search when queries are typed without dấu "
                "(e.g. \"he thong kiem soat\" instead of \"hệ thống kiểm soát\").\n\n"
                "Requires internet access, ~500MB-1GB disk space.",
                parent=dialog):
                dialog.lift()
                return
            local_root = _vi_dia_local_dir()
            abort_event = threading.Event()

            def _do_abort():
                abort_event.set()
                try:
                    status_lbl.config(text="cancelled", fg="#888888")
                    install_btn.config(state="normal", text="Install online...",
                                        command=lambda: _install_vi_diacritics_online(status_lbl, install_btn, dialog))
                except Exception:
                    pass

            try:
                install_btn.config(state="normal", text="Abort", command=_do_abort)
                status_lbl.config(text="downloading...", fg="#ffcc00")
            except Exception:
                pass

            def _bg():
                try:
                    from huggingface_hub import snapshot_download
                except ImportError:
                    def _fail():
                        if abort_event.is_set():
                            return
                        status_lbl.config(text="huggingface_hub not installed", fg="#ff6666")
                        install_btn.config(state="normal", text="Install online...",
                                            command=lambda: _install_vi_diacritics_online(status_lbl, install_btn, dialog))
                        messagebox.showerror(
                            "Install model",
                            "The 'huggingface_hub' package isn't installed.\n"
                            "Run: pip install huggingface_hub --break-system-packages",
                            parent=dialog)
                        dialog.lift()
                    self.root.after(0, _fail)
                    return
                try:
                    os.makedirs(local_root, exist_ok=True)
                    # v9.13.2 fix: force real file copies, not symlinks.
                    # huggingface_hub's default download layout stores the
                    # actual bytes as a hash-named blob and creates a
                    # symlink (e.g. config.json -> ../../blobs/<hash>) so
                    # local_dir "looks like" a normal repo. Windows without
                    # Developer Mode/admin rights can't create symlinks --
                    # when that fails, some huggingface_hub versions leave
                    # local_dir with NO file actually named "config.json"
                    # at all (just the hash-named blob), which is exactly
                    # why _vi_dia_installed() kept reporting "not
                    # installed" even after a successful download.
                    # local_dir_use_symlinks=False forces genuine file
                    # copies with the correct names instead. Older/newer
                    # huggingface_hub versions differ on whether this kwarg
                    # still exists, so fall back to calling without it.
                    try:
                        snapshot_download(repo_id=VI_DIA_REPO, local_dir=local_root,
                                           ignore_patterns=VI_DIA_IGNORE_PATTERNS,
                                           local_dir_use_symlinks=False)
                    except TypeError:
                        snapshot_download(repo_id=VI_DIA_REPO, local_dir=local_root,
                                           ignore_patterns=VI_DIA_IGNORE_PATTERNS)
                    if abort_event.is_set():
                        return
                    def _ok():
                        status_lbl.config(text="✓ installed", fg="#4caf50")
                        install_btn.config(state="normal", text="Reinstall...",
                                           command=lambda: _install_vi_diacritics_online(status_lbl, install_btn, dialog))
                        dialog.lift()
                    self.root.after(0, _ok)
                except Exception as e:
                    if abort_event.is_set():
                        return
                    err = str(e)
                    def _err():
                        status_lbl.config(text="download failed", fg="#ff6666")
                        install_btn.config(state="normal", text="Install online...",
                                            command=lambda: _install_vi_diacritics_online(status_lbl, install_btn, dialog))
                        messagebox.showerror("Install model", f"Download failed:\n{err}", parent=dialog)
                        dialog.lift()
                    self.root.after(0, _err)

            threading.Thread(target=_bg, daemon=True).start()

        def _on_update_db():
            """v5.9: clicking Update DB now opens an options dialog instead
            of scanning immediately — pick which file-type tiers to index,
            check/install AI Search models (Jina-v3 / BGE-Gemma2), and pick
            which model(s) actually get embeddings built this run."""
            if self._update_db_running or self._mailnotes_update_running:
                messagebox.showinfo("Update DB", "An update is already running.")
                return

            # Pre-fill tier checkboxes from a tier expression already typed
            # in the Searchbox (e.g. "tier 1,2") — preserves the old quick-
            # type shortcut as a convenience default; otherwise only Tier 1
            # starts checked (fastest default: primary office docs/PDF).
            raw_box_text = self.entry_var.get().strip()
            pre_tiers, is_tier_expr = self._parse_tier_filter(raw_box_text)
            default_on = set(pre_tiers) if (raw_box_text and is_tier_expr) else {0}

            dlg = tk.Toplevel(self.root)
            dlg.title("Update DB — Options")
            dlg.configure(bg=BG_COLOR)
            dlg.transient(self.root)
            # v5.9b: NOT topmost — this was actively unwanted (it was staying
            # pinned above unrelated apps like Chrome). A normal Toplevel
            # behaves like any other window: on top of self.root initially,
            # but can be covered by whatever the user clicks next, same as
            # any other dialog in the app.
            dlg.resizable(False, False)
            # placeholder position -- actually centered on self.root once every
            # widget below is packed (see the centering block near the bottom,
            # right after the Start/Cancel buttons)
            dlg.geometry(f"+{self.root.winfo_x()}+{self.root.winfo_y() + 40}")

            # ── Outlook mail (v-outlook) ─────────────────────────────────────
            # A separate, mutually-exclusive mode from the file-DB update
            # below: checking this greys out every other option in the
            # dialog (Tiers/OCR/AI models/Vi diacritics — none of that
            # applies to indexing mail) and "Start Update" then ONLY runs
            # outlook_search.index_outlook_mail(), completely independent of
            # DB_FILE/db_conn/the ramp light — same guarantee as the
            # standalone "--update outlook" command.
            outlook_f = tk.LabelFrame(dlg, text="Outlook Mail",
                                       bg=BG_COLOR, fg="#33363c",
                                       font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            outlook_f.pack(fill="x", padx=12, pady=(12, 6))
            outlook_only_var = tk.BooleanVar(value=False)
            _outlook_cb_state = "normal" if OUTLOOK_SEARCH_AVAILABLE else "disabled"
            if self._mailnotes_update_running or self._update_db_running:
                _outlook_cb_state = "disabled"
            outlook_cb = tk.Checkbutton(
                outlook_f, text="Update Outlook mail (skip file DB entirely this run)",
                variable=outlook_only_var, bg=BG_COLOR, anchor="w",
                font=("Segoe UI", 9), justify="left", wraplength=460,
                state=_outlook_cb_state)
            outlook_cb.pack(fill="x", anchor="w", pady=2)
            if not OUTLOOK_SEARCH_AVAILABLE:
                tk.Label(outlook_f, text="outlook_search.py not found next to this script.",
                         bg=BG_COLOR, fg="#888888", font=("Segoe UI", 8)).pack(anchor="w")
            elif self._mailnotes_update_running:
                tk.Label(outlook_f, text="A mail/notes index is already running.",
                         bg=BG_COLOR, fg="#ffcc00", font=("Segoe UI", 8)).pack(anchor="w")
            # ──────────────────────────────────────────────────────────────

            # ── OneNote (v-onenote) ─────────────────────────────────────────
            # Same idea and same mutual-exclusivity-with-Tiers as Outlook
            # above. If BOTH Outlook and OneNote are checked, Start Update
            # runs Outlook first, then OneNote — see _start_mail_notes_update.
            onenote_f = tk.LabelFrame(dlg, text="OneNote",
                                       bg=BG_COLOR, fg="#33363c",
                                       font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            onenote_f.pack(fill="x", padx=12, pady=(0, 6))
            onenote_only_var = tk.BooleanVar(value=False)
            _onenote_cb_state = "normal" if ONENOTE_SEARCH_AVAILABLE else "disabled"
            if self._mailnotes_update_running or self._update_db_running:
                _onenote_cb_state = "disabled"
            onenote_cb = tk.Checkbutton(
                onenote_f, text="Update OneNote (skip file DB entirely this run)",
                variable=onenote_only_var, bg=BG_COLOR, anchor="w",
                font=("Segoe UI", 9), justify="left", wraplength=460,
                state=_onenote_cb_state)
            onenote_cb.pack(fill="x", anchor="w", pady=2)
            if not ONENOTE_SEARCH_AVAILABLE:
                tk.Label(onenote_f, text="onenote_search.py not found next to this script.",
                         bg=BG_COLOR, fg="#888888", font=("Segoe UI", 8)).pack(anchor="w")
            elif self._mailnotes_update_running:
                tk.Label(onenote_f, text="A mail/notes index is already running.",
                         bg=BG_COLOR, fg="#ffcc00", font=("Segoe UI", 8)).pack(anchor="w")
            # ──────────────────────────────────────────────────────────────

            # ──────────────────────────────────────────────────────────────

            # ── "Also Update Database" checkbox — standalone, no frame ─────
            # v-simplify2: only makes sense once Outlook and/or OneNote is
            # checked (it means "don't skip the File DB stage after those
            # finish"), so it starts disabled/unchecked and only becomes
            # clickable once at least one of them is checked -- instead of
            # sitting there enabled-but-meaningless the rest of the time.
            chain_all_var = tk.BooleanVar(value=False)
            chain_cb = tk.Checkbutton(
                dlg, text="Also Update Database (File DB) after",
                variable=chain_all_var, bg=BG_COLOR, fg="#33363c", anchor="w",
                font=("Segoe UI", 9), state="disabled")
            chain_cb.pack(fill="x", padx=12, pady=(0, 6))

            def _refresh_chain_cb_state(*_a):
                enabled = (outlook_only_var.get() or onenote_only_var.get()) \
                          and not (self._mailnotes_update_running or self._update_db_running)
                try:
                    chain_cb.config(state=("normal" if enabled else "disabled"))
                except Exception:
                    pass
                if not enabled:
                    # Outlook/OneNote both got unchecked (or a run started
                    # elsewhere) -- a checked-but-disabled box would look
                    # like it's still going to do something, so clear it.
                    chain_all_var.set(False)
            outlook_only_var.trace_add("write", _refresh_chain_cb_state)
            onenote_only_var.trace_add("write", _refresh_chain_cb_state)
            _refresh_chain_cb_state()

            # ── Tier 1-4 file-type checkboxes ───────────────────────────────
            tiers_f = tk.LabelFrame(dlg, text="File types to index (Tiers)",
                                     bg=BG_COLOR, fg="#33363c",
                                     font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            tiers_f.pack(fill="x", padx=12, pady=(0, 6))

            # v-outlook: (widget, state_when_unlocked) pairs — toggled by the
            # "Update Outlook mail" / "Update OneNote" checkboxes above. Restoring uses each
            # widget's own correct normal state rather than blindly "normal"
            # (e.g. an uninstalled AI model's checkbox must stay disabled).
            _greyable = []

            tier_labels = ["Tier 1  (Office / PDF)", "Tier 2  (Email / Text / Notes)",
                           "Tier 3  (Scripts / Config)", "Tier 4  (Markup / Query / Misc)"]
            tier_var_list = []
            for i, exts in enumerate(self._ALL_TIERS):
                v = tk.BooleanVar(value=(i in default_on))
                tier_var_list.append(v)
                ext_str = ", ".join(sorted(exts))
                cb = tk.Checkbutton(tiers_f, text=f"{tier_labels[i]}: {ext_str}",
                                     variable=v, bg=BG_COLOR, anchor="w",
                                     font=("Segoe UI", 9), justify="left",
                                     wraplength=460)
                cb.pack(fill="x", anchor="w", pady=2)
                _greyable.append((cb, "normal"))

            # v9.13: OCR images checkbox — off by default. Separate from the
            # Tier checkboxes above (images don't belong to any of the 4
            # Tiers), but placed in the same frame since it's conceptually
            # the same kind of choice: "what content gets extracted".
            ocr_var = tk.BooleanVar(value=False)
            ocr_cb = tk.Checkbutton(
                tiers_f, text=("OCR images (.jpg/.png/...) — slower, downloads OCR model on first use"
                               if ENABLE_AI_SEARCH_FEATURE else
                               "OCR images (disabled in this build — requires torch, not bundled)"),
                variable=ocr_var, bg=BG_COLOR, anchor="w",
                font=("Segoe UI", 9), justify="left", wraplength=460,
                state=("normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))
            ocr_cb.pack(fill="x", anchor="w", pady=(6, 2))
            _greyable.append((ocr_cb, "normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))

            # v-new: "Force re-index" -- var defined here alongside the
            # other extraction-related checkboxes, but the actual widget is
            # placed in btn_row below (between Cancel/Start Update, per
            # request) since it's more of an action-modifier than a
            # tier/content choice.
            force_var = tk.BooleanVar(value=False)

            # ── AI Search models section (separate frame, same dialog) ─────
            # v5.9b: each model now has its own "build embeddings this run"
            # checkbox (checked by default = same as before/all models) --
            # unchecking a model skips its (slow) embedding pass entirely,
            # e.g. if the user only ever uses Jina-v3 and doesn't care about
            # BGE-Gemma2, unchecking it noticeably speeds up Update DB.
            ai_f = tk.LabelFrame(dlg, text=("Search AI models" if ENABLE_AI_SEARCH_FEATURE
                                             else "Search AI models (disabled in this build)"),
                                  bg=BG_COLOR, fg="#33363c",
                                  font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            ai_f.pack(fill="x", padx=12, pady=(0, 6))

            model_var_list = []  # (model_key, BooleanVar) — which models to embed this run
            for model_key, info in SEMANTIC_MODELS.items():
                row = tk.Frame(ai_f, bg=BG_COLOR)
                row.pack(fill="x", pady=3)
                mdir = _find_model_dir_for(model_key)
                installed = bool(mdir) and os.path.isfile(os.path.join(mdir, "config.json"))
                label = info.get("label", model_key)

                # v7.10: a model that isn't installed locally can't build
                # embeddings this run regardless of the checkbox state, so
                # start it unchecked + disabled (greyed out, checkbox AND
                # its "Build embeddings for <model>" label together) instead
                # of letting the user tick a box that does nothing. It's
                # re-enabled automatically the moment install succeeds (see
                # _install_model_online's _ok()) without needing to reopen
                # this dialog.
                #
                # v10.17 (ENABLE_AI_SEARCH_FEATURE=False build): force this
                # whole section unchecked + disabled regardless of what's
                # installed on disk — see the flag's definition near the top
                # of the file for what this build variant is for.
                mv = tk.BooleanVar(value=(installed and ENABLE_AI_SEARCH_FEATURE))
                model_var_list.append((model_key, mv))
                cb = tk.Checkbutton(row, text=f"Build embeddings for {label}",
                                     variable=mv, bg=BG_COLOR, font=("Segoe UI", 9),
                                     state=("normal" if (installed and ENABLE_AI_SEARCH_FEATURE) else "disabled"),
                                     disabledforeground="#a0a0a0")
                cb.pack(side="left")
                _greyable.append((cb, "normal" if (installed and ENABLE_AI_SEARCH_FEATURE) else "disabled"))

                status_txt = ("✓ installed" if installed else "not installed locally")
                status_fg = "#4caf50" if installed else "#888888"
                status_lbl = tk.Label(row, text=status_txt, bg=BG_COLOR, fg=status_fg,
                                       font=("Segoe UI", 9))
                status_lbl.pack(side="left", padx=(8, 0))

                # v5.9b: install button is NEVER greyed out / disabled, even
                # once "installed" — if a previous download was Aborted, the
                # model folder can exist but be incomplete, and a disabled
                # button would make it impossible to try installing again.
                # Always clickable so the user can (re)install / overwrite
                # at any time.
                #
                # v10.17: EXCEPT in the ENABLE_AI_SEARCH_FEATURE=False build,
                # where there's no torch/sentence_transformers baked into the
                # exe to ever load an installed model with anyway — so the
                # install button is disabled here too instead of offering a
                # download that would just sit unused.
                install_btn = tk.Button(row, text=("Reinstall..." if installed else "Install online..."),
                                         font=("Segoe UI", 8), cursor="hand2",
                                         state=("normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))
                install_btn.config(command=lambda mk=model_key, sl=status_lbl, ib=install_btn, ckb=cb, mvv=mv:
                                    _install_model_online(mk, sl, ib, dlg, ckb, mvv))
                install_btn.pack(side="right")
                _greyable.append((install_btn, "normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))



            # ── Vietnamese diacritics restoration (optional, separate from
            # the embedding models above — this is a query-preprocessing
            # helper, not something you pick for search, so no "build
            # embeddings" checkbox here, just install status) ─────────────
            vidia_f = tk.LabelFrame(dlg, text=("Vietnamese diacritics restoration (optional)" if ENABLE_AI_SEARCH_FEATURE
                                                else "Vietnamese diacritics restoration (disabled in this build)"),
                                     bg=BG_COLOR, fg="#33363c",
                                     font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            vidia_f.pack(fill="x", padx=12, pady=(0, 6))
            vidia_row = tk.Frame(vidia_f, bg=BG_COLOR)
            vidia_row.pack(fill="x", pady=3)
            tk.Label(vidia_row, text="Improves AI Search for queries typed without dấu",
                     bg=BG_COLOR, font=("Segoe UI", 9)).pack(side="left")
            vidia_installed = _vi_dia_installed()
            vidia_status_lbl = tk.Label(
                vidia_row, text=("✓ installed" if vidia_installed else "not installed"),
                bg=BG_COLOR, fg=("#4caf50" if vidia_installed else "#888888"),
                font=("Segoe UI", 9))
            vidia_status_lbl.pack(side="left", padx=(8, 0))
            vidia_install_btn = tk.Button(
                vidia_row, text=("Reinstall..." if vidia_installed else "Install online..."),
                font=("Segoe UI", 8), cursor="hand2",
                state=("normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))
            vidia_install_btn.config(
                command=lambda: _install_vi_diacritics_online(vidia_status_lbl, vidia_install_btn, dlg))
            vidia_install_btn.pack(side="right")
            _greyable.append((vidia_install_btn, "normal" if ENABLE_AI_SEARCH_FEATURE else "disabled"))

            def _on_mail_notes_toggle(*_a):
                locked = (outlook_only_var.get() or onenote_only_var.get()) and not chain_all_var.get()
                for w, unlocked_state in _greyable:
                    try:
                        w.config(state=("disabled" if locked else unlocked_state))
                    except Exception:
                        pass
            outlook_only_var.trace_add("write", _on_mail_notes_toggle)
            onenote_only_var.trace_add("write", _on_mail_notes_toggle)
            chain_all_var.trace_add("write", _on_mail_notes_toggle)

            # ── Action buttons ───────────────────────────────────────────────
            btn_row = tk.Frame(dlg, bg=BG_COLOR)
            btn_row.pack(fill="x", padx=12, pady=(6, 12))

            def _start():
                if chain_all_var.get():
                    chosen = {i for i, v in enumerate(tier_var_list) if v.get()}
                    if not chosen:
                        messagebox.showwarning("Update DB", "Select at least one tier.", parent=dlg)
                        return
                    selected_tiers = None if chosen == {0, 1, 2, 3} else chosen
                    chosen_models = [mk for mk, v in model_var_list if v.get()]
                    if not chosen_models:
                        if not messagebox.askyesno(
                            "Update DB",
                            "No AI model selected — the File DB stage of this run will "
                            "update BM25/content search only. AI Search data will NOT "
                            "be built or refreshed.\n\n"
                            "Continue with Tier-only File DB after?", parent=dlg):
                            return
                        selected_models = []
                    else:
                        selected_models = None if len(chosen_models) == len(model_var_list) else chosen_models
                    dlg.destroy()
                    # Whichever of Outlook (~10h) / OneNote (~5-10min) the
                    # user actually checked runs first (sequentially, if
                    # both), then File DB (~1-2 days) -- run_update_db only
                    # fires once the mail/notes stage has fully finished,
                    # via on_complete.
                    self._start_mail_notes_update(
                        do_outlook=outlook_only_var.get(), do_onenote=onenote_only_var.get(),
                        on_complete=lambda: _run_update_db(selected_tiers, selected_models, ocr_var.get(), force_var.get()))
                    return
                if outlook_only_var.get() or onenote_only_var.get():
                    dlg.destroy()
                    self._start_mail_notes_update(
                        do_outlook=outlook_only_var.get(),
                        do_onenote=onenote_only_var.get())
                    return
                chosen = {i for i, v in enumerate(tier_var_list) if v.get()}
                if not chosen:
                    messagebox.showwarning("Update DB", "Select at least one tier.", parent=dlg)
                    return
                selected_tiers = None if chosen == {0, 1, 2, 3} else chosen
                chosen_models = [mk for mk, v in model_var_list if v.get()]
                # v8.4: no models ticked is now a valid, deliberate choice —
                # "Tier-only" update (BM25/content only, AI stage skipped
                # entirely this run). Confirm instead of blocking, since it's
                # easy to forget to tick a model and not realize AI won't be
                # touched.
                if not chosen_models:
                    if not messagebox.askyesno(
                        "Update DB",
                        "No AI model selected — this run will update BM25/content "
                        "search only. AI Search data will NOT be built or refreshed.\n\n"
                        "Continue with Tier-only update?", parent=dlg):
                        return
                    selected_models = []  # explicit empty = skip AI entirely (see indexing_worker)
                else:
                    selected_models = None if len(chosen_models) == len(model_var_list) else chosen_models
                dlg.destroy()
                _run_update_db(selected_tiers, selected_models, ocr_var.get(), force_var.get())

            tk.Button(btn_row, text="Start Update", font=("Segoe UI", 9, "bold"),
                      bg="#2196f3", fg="white", activebackground="#1976d2",
                      cursor="hand2", padx=10, command=_start).pack(side="right", padx=(6, 0))
            tk.Button(btn_row, text="Cancel", font=("Segoe UI", 9), cursor="hand2",
                      padx=10, command=dlg.destroy).pack(side="right")
            # v-fix (theo yêu cầu -- sang trái nút Cancel, thẳng hàng với
            # các checkbox bên trên): moved from between Cancel/Start
            # Update to the LEFT side of btn_row, no extra padx so its left
            # edge lines up with the Tier/OCR checkboxes above (tiers_f
            # also starts at padx=12 off the dialog). Off by default:
            # bypasses the mtime-based "file unchanged, skip re-extraction"
            # shortcut (see indexing_worker) so EVERY selected file's
            # content is re-extracted from scratch regardless of whether
            # it's changed on disk since the last scan -- needed after a
            # change to the extraction logic itself (e.g. the PDF
            # page-count fix), since already-indexed files would otherwise
            # never be touched again (their mtime on disk never changed).
            # Makes the run as slow as a first-time scan, hence off by
            # default / explicit opt-in.
            force_cb = tk.Checkbutton(
                btn_row, text="Force re-index", variable=force_var,
                bg=BG_COLOR, font=("Segoe UI", 9), cursor="hand2")
            force_cb.pack(side="left")
            # v-fix (theo phát hiện của user): "Force re-index" chỉ ảnh
            # hưởng tới luồng file DB (mtime-based skip logic trong
            # indexing_worker) -- khi Outlook-only/OneNote-only được chọn,
            # code ở _start (nhánh outlook_only_var.get()/onenote_only_var
            # .get() phía trên) return NGAY sau _start_mail_notes_update(),
            # chưa từng đọc force_var trong nhánh đó -- checkbox này hoàn
            # toàn vô tác dụng khi ở chế độ Outlook/OneNote-only, giống hệt
            # Tier/OCR/AI model ở trên, nhưng trước đây lại không được đưa
            # vào _greyable nên không tự mờ đi cùng chúng. Thêm vào để
            # _on_mail_notes_toggle() (đã định nghĩa ở trên, trigger bởi
            # outlook_only_var/onenote_only_var/chain_all_var) mờ luôn ô
            # này cho nhất quán.
            _greyable.append((force_cb, "normal"))
            add_tooltip(force_cb, "Re-extract content for every selected file, even ones "
                                   "already indexed and unchanged on disk (slower).")

            # v7.10: center the dialog on the Search GUI (self.root) instead
            # of just offsetting from its top-left corner. Done last, now
            # that every widget above has been packed, so winfo_reqwidth/
            # reqheight below reflect the dialog's real final size.
            dlg.update_idletasks()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            rx, ry = self.root.winfo_x(), self.root.winfo_y()
            dw, dh = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
            cx = rx + (rw - dw) // 2
            cy = ry + (rh - dh) // 2
            # keep it fully on-screen
            sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
            cx = max(0, min(cx, sw - dw))
            cy = max(0, min(cy, sh - dh))
            dlg.geometry(f"+{cx}+{cy}")

        _btn_text = "Update DB"
        _any_running = self._update_db_running or self._mailnotes_update_running
        if _any_running:
            # v9.11 fix: prefer the last real progress string (e.g. "AI 2/2:
            # 28%") over a bare "Updating..." — this button gets recreated
            # from scratch whenever the search box is minimized/reopened
            # mid-run, and previously always lost the percentage at that
            # point even though indexing was still happily running in the
            # background the whole time.
            _btn_text = self._last_index_status_text or "Updating..."
            if self._mailnotes_update_running and not self._update_db_running:
                _btn_text = (self._outlook_progress_text or self._onenote_progress_text
                             or "Updating...")
        self._update_db_btn = tk.Button(
            search_bar, text=_btn_text,
            font=("Segoe UI", 8, "bold"),
            bg="#e6e8ec", fg="#33363c",
            relief="raised", bd=2,
            padx=8, pady=2,
            activebackground="#d3d6dc", activeforeground="#111111",
            cursor="hand2",
            state=("disabled" if _any_running else "normal"),
            command=_on_update_db)
        self._update_db_btn.pack(side="right", padx=(4, 2), pady=4)
        add_tooltip(self._update_db_btn,
                    "Open Update DB options (choose file-type tiers,\n"
                    "check/install AI Search models) then start the scan.")
        # ──────────────────────────────────────────────────────────────────


        self.c_tree_frame = tk.Frame(c_frame)
        self.c_tree_frame.pack(fill="both", expand=True)

        self.tree_c = ttk.Treeview(self.c_tree_frame, columns=("icon_name", "size", "mtime", "ftype", "full_path"), show="tree headings")
        self.tree_c.heading("#0", text="")
        self.tree_c.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        self.tree_c.heading("icon_name", text="File Location")
        self.tree_c.heading("size",      text="File Size")
        self.tree_c.heading("mtime",     text="Date Modified")
        self.tree_c.heading("ftype",     text="Type")
        self.tree_c.column("icon_name",  width=850, stretch=False)
        self.tree_c.column("size",       width=90,  anchor="center", stretch=False)
        self.tree_c.column("mtime",      width=130, anchor="center", stretch=False)
        self.tree_c.column("ftype",      width=55,  anchor="center", stretch=False)
        self.tree_c["displaycolumns"] = ("icon_name", "size", "mtime", "ftype")
        sb_c = ttk.Scrollbar(self.c_tree_frame, orient="vertical", command=self.tree_c.yview)
        sb_c_x = ttk.Scrollbar(self.c_tree_frame, orient="horizontal", command=self.tree_c.xview)
        self.tree_c.configure(yscrollcommand=sb_c.set, xscrollcommand=sb_c_x.set)
        sb_c.pack(side="right", fill="y")
        sb_c_x.pack(side="bottom", fill="x")
        self.tree_c.pack(side="left", fill="both", expand=True)

        # 2. TAB FILE NAME — added FIRST to notebook (MFT default tab)
        f_frame = tk.Frame(self.nb)
        self.nb.add(f_frame, text=" File Name ")

        # ── Smart Size Filter bar ──────────────────────────────────────────
        filter_bar = tk.Frame(f_frame, bg=BG_COLOR, pady=2)
        filter_bar.pack(fill="x", padx=6, pady=(4, 2))
        filter_bar.pack_propagate(False)
        filter_bar.config(height=32)

        tk.Label(filter_bar, text="Size:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(4, 2))

        op_menu = ttk.OptionMenu(filter_bar, self.size_op_var, self.size_op_var.get(),
                                  "Any", ">", ">=", "<", "<=", "=",
                                  command=lambda *_: self._apply_size_filter_to_tree())
        op_menu.config(width=4); op_menu.pack(side="left", padx=2)

        size_num_entry = tk.Entry(filter_bar, textvariable=self.size_num_var, width=7,
                                   font=("Segoe UI", 9), bg=ENTRY_BG, fg=TEXT_COLOR,
                                   insertbackground=TEXT_COLOR, bd=1, relief="flat")
        size_num_entry.pack(side="left", padx=2)
        self.size_num_var.trace_add("write", self._apply_size_filter_to_tree)

        unit_menu = ttk.OptionMenu(filter_bar, self.size_unit_var, self.size_unit_var.get(),
                                    "B", "KB", "MB", "GB",
                                    command=lambda *_: self._apply_size_filter_to_tree())
        unit_menu.config(width=4); unit_menu.pack(side="left", padx=2)

        def _clear_size_filter():
            self.size_op_var.set(">"); self.size_num_var.set(""); self.size_unit_var.set("MB")
            self._apply_size_filter_to_tree()

        # ── Extension filter ───────────────────────────────────────────────
        tk.Label(filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 11)).pack(side="left", padx=4)
        tk.Label(filter_bar, text="Ext:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 3))

        # Quick preset buttons
        _EXT_PRESETS = [
            ("All",   ""),
            ("PDF",   "pdf"),
            ("Word",  "doc,docx"),
            ("Excel", "xls,xlsx,csv"),
            ("PPT",   "ppt,pptx"),
            ("Img",   "png,jpg,jpeg,bmp,gif"),
            ("Txt",   "txt,log"),
            ("Msg",   "msg"),
            ("One",   "one"),
        ]
        def _set_ext(val):
            self.ext_filter_var.set(val)   # trace fires _apply_size_filter_to_tree

        for _lbl, _val in _EXT_PRESETS:
            tk.Button(filter_bar, text=_lbl, font=("Segoe UI", 8),
                      bg="#e6e8ec", fg="#33363c", bd=0, padx=5, pady=1,
                      activebackground="#d3d6dc", cursor="hand2",
                      command=lambda v=_val: _set_ext(v)).pack(side="left", padx=1)

        tk.Label(filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)

        ext_entry = tk.Entry(filter_bar, textvariable=self.ext_filter_var, width=14,
                              font=("Segoe UI", 9), bg=ENTRY_BG, fg="#888",
                              insertbackground=TEXT_COLOR, bd=1, relief="flat")
        ext_entry.pack(side="left", padx=(0, 2))

        def _ext_focus_in(e):
            if not self.ext_filter_var.get(): ext_entry.config(fg=TEXT_COLOR)
        def _ext_focus_out(e):
            if not self.ext_filter_var.get(): ext_entry.config(fg="#888")
        ext_entry.bind("<FocusIn>",  _ext_focus_in)
        ext_entry.bind("<FocusOut>", _ext_focus_out)
        # trace already wired through ext_filter_var → _apply_size_filter_to_tree
        self.ext_filter_var.trace_add("write", self._apply_size_filter_to_tree)

        tk.Button(filter_bar, text="✕", command=lambda: self.ext_filter_var.set(""),
                  bg="#c9ccd2", fg="#222222", font=("Segoe UI", 8), bd=0, padx=4,
                  activebackground="#b0b3ba", cursor="hand2").pack(side="left", padx=(0, 4))
        # ──────────────────────────────────────────────────────────────────

        # ── Whole word toggle ───────────────────────────────────────────────
        # v7.10: OFF by default (raw substring, original behavior). When ON,
        # a keyword only counts as a match when it's bounded by non-letter
        # characters on both sides -- e.g. searching "adas" still matches
        # "ADAS_systems.pdf" / "VDIM_0ADAS_ESP..." but no longer matches
        # buried-substring false positives like "readasync.xml" or
        # "ReadAStringExample.mlx". Applies to both the live MFT scan (File
        # Name / Folder Name, no DB needed) and the DB-backed search.
        tk.Label(filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        _ww_cb = tk.Checkbutton(filter_bar, text="Whole word", variable=self.whole_word_var,
                                  bg=BG_COLOR, fg=TEXT_COLOR, selectcolor=ENTRY_BG,
                                  font=("Segoe UI", 9),
                                  command=self._rerun_current_search)
        _ww_cb.pack(side="left", padx=(0, 4))
        add_tooltip(_ww_cb,
                     "ON: \"adas\" matches ADAS_systems.pdf but not readasync.xml\n"
                     "OFF (default): \"adas\" matches anywhere, incl. inside other words")
        # ──────────────────────────────────────────────────────────────────

        # ── Name filter ────────────────────────────────────────────────────
        tk.Label(filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        tk.Label(filter_bar, text="Name:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 2))
        name_entry = tk.Entry(filter_bar, textvariable=self.name_filter_var, width=18,
                               font=("Segoe UI", 9), bg=ENTRY_BG, fg="#888",
                               insertbackground=TEXT_COLOR, bd=1, relief="flat")
        name_entry.pack(side="left", padx=(0, 2))
        def _name_focus_in(e):
            if not self.name_filter_var.get(): name_entry.config(fg=TEXT_COLOR)
        def _name_focus_out(e):
            if not self.name_filter_var.get(): name_entry.config(fg="#888")
        name_entry.bind("<FocusIn>",  _name_focus_in)
        name_entry.bind("<FocusOut>", _name_focus_out)
        def _apply_name_filter_everywhere(*_a):
            # v-new: name_filter_var is shared between File Name and Folder
            # Name tabs (Folder Name tab now has its own Name/Whole word UI
            # too — same variables, same position). Refresh both on any edit.
            self._apply_size_filter_to_tree()
            self._apply_folder_filter_to_tree()
        self.name_filter_var.trace_add("write", _apply_name_filter_everywhere)
        tk.Button(filter_bar, text="✕", command=lambda: self.name_filter_var.set(""),
                  bg="#c9ccd2", fg="#222222", font=("Segoe UI", 8), bd=0, padx=4,
                  activebackground="#b0b3ba", cursor="hand2").pack(side="left", padx=(0, 4))
        # ──────────────────────────────────────────────────────────────────

        self.filter_count_label = None  # removed - counts shown in pane headers
        # ──────────────────────────────────────────────────────────────────

        self.f_tree_frame = tk.Frame(f_frame)
        self.f_tree_frame.pack(fill="both", expand=True)

        self.tree_f = ttk.Treeview(self.f_tree_frame, columns=("icon_name", "size", "mtime", "ftype", "location", "full_path"), show="tree headings")
        self.tree_f.heading("#0", text="")
        self.tree_f.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        self.tree_f.heading("icon_name", text="File Name")
        self.tree_f.heading("size",      text="File Size")
        self.tree_f.heading("mtime",     text="Date Modified")
        self.tree_f.heading("ftype",     text="Type")
        self.tree_f.heading("location",  text="File Location")
        self.tree_f.column("icon_name",  width=380, minwidth=150, stretch=False)
        self.tree_f.column("size",       width=90,  anchor="center", stretch=False, minwidth=70)
        self.tree_f.column("mtime",      width=130, anchor="center", stretch=False, minwidth=100)
        self.tree_f.column("ftype",      width=55,  anchor="center", stretch=False, minwidth=40)
        self.tree_f.column("location",   width=820, minwidth=300, stretch=False)
        self.tree_f["displaycolumns"] = ("icon_name", "size", "mtime", "ftype", "location")
        sb_f  = ttk.Scrollbar(self.f_tree_frame, orient="vertical",   command=self.tree_f.yview)
        sb_fx = ttk.Scrollbar(self.f_tree_frame, orient="horizontal", command=self.tree_f.xview)
        self.tree_f.configure(yscrollcommand=sb_f.set, xscrollcommand=sb_fx.set)
        # Pack order matters: scrollbars first, then tree fills remaining space
        sb_f.pack(side="right",  fill="y")
        sb_fx.pack(side="bottom", fill="x")
        self.tree_f.pack(side="left", fill="both", expand=True)

        # 3. TAB FOLDER NAME — added SECOND to notebook
        fol_frame = tk.Frame(self.nb)
        self.nb.add(fol_frame, text=" Folder Name ")

        # ── Whole word + Name filter bar (same position/style as File Name
        #    tab). Reuses the SAME shared variables (self.whole_word_var,
        #    self.name_filter_var) as the File Name tab, so the filter stays
        #    in sync no matter which tab you edit it from.
        #
        #    v-fix: Size/Ext controls don't apply to folders, but leaving
        #    them out entirely shifted Whole word/Name to the far left,
        #    landing at a different horizontal position than in the File
        #    Name tab. Rebuilding the identical (but disabled/grayed-out,
        #    non-functional) Size/Ext widgets here reserves the exact same
        #    space, so Whole word/Name line up at the same spot in both
        #    tabs. ─────────────────────────────────────────────────────────
        fol_filter_bar = tk.Frame(fol_frame, bg=BG_COLOR, pady=2)
        fol_filter_bar.pack(fill="x", padx=6, pady=(4, 2))
        fol_filter_bar.pack_propagate(False)
        fol_filter_bar.config(height=32)

        tk.Label(fol_filter_bar, text="Size:", bg=BG_COLOR, fg="#666",
                 font=("Segoe UI", 9)).pack(side="left", padx=(4, 2))
        _fol_op_menu = ttk.OptionMenu(fol_filter_bar, self.size_op_var, self.size_op_var.get(),
                                       "Any", ">", ">=", "<", "<=", "=")
        _fol_op_menu.config(width=4, state="disabled"); _fol_op_menu.pack(side="left", padx=2)
        _fol_size_entry = tk.Entry(fol_filter_bar, textvariable=self.size_num_var, width=7,
                                    font=("Segoe UI", 9), bg=ENTRY_BG, fg="#666",
                                    bd=1, relief="flat", state="disabled")
        _fol_size_entry.pack(side="left", padx=2)
        _fol_unit_menu = ttk.OptionMenu(fol_filter_bar, self.size_unit_var, self.size_unit_var.get(),
                                         "B", "KB", "MB", "GB")
        _fol_unit_menu.config(width=4, state="disabled"); _fol_unit_menu.pack(side="left", padx=2)

        tk.Label(fol_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 11)).pack(side="left", padx=4)
        tk.Label(fol_filter_bar, text="Ext:", bg=BG_COLOR, fg="#666",
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 3))
        for _lbl in ["All", "PDF", "Word", "Excel", "PPT", "Img", "Txt", "Msg", "One"]:
            tk.Button(fol_filter_bar, text=_lbl, font=("Segoe UI", 8),
                      bg="#e6e8ec", fg="#999", bd=0, padx=5, pady=1,
                      state="disabled").pack(side="left", padx=1)
        tk.Label(fol_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        tk.Entry(fol_filter_bar, textvariable=self.ext_filter_var, width=14,
                  font=("Segoe UI", 9), bg=ENTRY_BG, fg="#666",
                  bd=1, relief="flat", state="disabled").pack(side="left", padx=(0, 2))
        tk.Button(fol_filter_bar, text="✕", bg="#c9ccd2", fg="#999", font=("Segoe UI", 8),
                  bd=0, padx=4, state="disabled").pack(side="left", padx=(0, 4))
        add_tooltip(_fol_size_entry, "Size/Ext filters don't apply to folders — shown here only to line up with the File Name tab.")
        # ──────────────────────────────────────────────────────────────────

        tk.Label(fol_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        _fol_ww_cb = tk.Checkbutton(fol_filter_bar, text="Whole word", variable=self.whole_word_var,
                                     bg=BG_COLOR, fg=TEXT_COLOR, selectcolor=ENTRY_BG,
                                     font=("Segoe UI", 9),
                                     command=self._rerun_current_search)
        _fol_ww_cb.pack(side="left", padx=(0, 4))
        add_tooltip(_fol_ww_cb,
                     "ON: \"adas\" matches ADAS_systems folder but not readasync\n"
                     "OFF (default): \"adas\" matches anywhere, incl. inside other words")

        tk.Label(fol_filter_bar, text="│", bg=BG_COLOR, fg="#555",
                 font=("Segoe UI", 9)).pack(side="left", padx=3)
        tk.Label(fol_filter_bar, text="Name:", bg=BG_COLOR, fg=TEXT_COLOR,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 2))
        fol_name_entry = tk.Entry(fol_filter_bar, textvariable=self.name_filter_var, width=18,
                                   font=("Segoe UI", 9), bg=ENTRY_BG, fg="#888",
                                   insertbackground=TEXT_COLOR, bd=1, relief="flat")
        fol_name_entry.pack(side="left", padx=(0, 2))
        def _fol_name_focus_in(e):
            if not self.name_filter_var.get(): fol_name_entry.config(fg=TEXT_COLOR)
        def _fol_name_focus_out(e):
            if not self.name_filter_var.get(): fol_name_entry.config(fg="#888")
        fol_name_entry.bind("<FocusIn>",  _fol_name_focus_in)
        fol_name_entry.bind("<FocusOut>", _fol_name_focus_out)
        # trace already wired once (shared var) in the File Name tab setup —
        # it refreshes both tabs, see _apply_name_filter_everywhere above.
        tk.Button(fol_filter_bar, text="✕", command=lambda: self.name_filter_var.set(""),
                  bg="#c9ccd2", fg="#222222", font=("Segoe UI", 8), bd=0, padx=4,
                  activebackground="#b0b3ba", cursor="hand2").pack(side="left", padx=(0, 4))
        # ──────────────────────────────────────────────────────────────────

        self.fol_tree_frame = tk.Frame(fol_frame)
        self.fol_tree_frame.pack(fill="both", expand=True)
        self.tree_fol = ttk.Treeview(self.fol_tree_frame, columns=("icon_name", "mtime", "location", "full_path"), show="tree headings")
        self.tree_fol.heading("#0", text="")
        self.tree_fol.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        self.tree_fol.heading("icon_name", text="Folder Name")
        self.tree_fol.heading("mtime",     text="Date Modified")
        self.tree_fol.heading("location",  text="Folder Location")
        self.tree_fol.column("icon_name",  width=440, minwidth=150, stretch=False)
        self.tree_fol.column("mtime",      width=130, anchor="center", stretch=False, minwidth=100)
        self.tree_fol.column("location",   width=750, minwidth=300, stretch=False)
        self.tree_fol["displaycolumns"] = ("icon_name", "mtime", "location")
        sb_fol  = ttk.Scrollbar(self.fol_tree_frame, orient="vertical",   command=self.tree_fol.yview)
        sb_folx = ttk.Scrollbar(self.fol_tree_frame, orient="horizontal", command=self.tree_fol.xview)
        self.tree_fol.configure(yscrollcommand=sb_fol.set, xscrollcommand=sb_folx.set)
        sb_fol.pack(side="right",  fill="y")
        sb_folx.pack(side="bottom", fill="x")
        self.tree_fol.pack(side="left", fill="both", expand=True)

        # v2.3: File Content tab added LAST (rightmost) — only populated after --update data
        self.nb.add(c_frame, text=" File Content ")

        # v7.7: File Content is now the DEFAULT tab shown when results first
        # open (was File Name). This is what makes the AI Chat panel appear
        # immediately without an extra click on "AI Search" -- the panel is
        # gated on the File Content tab being the active/visible one (see
        # _sync_online_chat_panel_visibility), so simply landing on this tab
        # by default is enough to reveal it right away. ttk.Notebook always
        # selects the FIRST added tab by default, and File Name is added
        # first (kept first for visual left-to-right tab order), so this
        # explicit select() is required to override that default.
        self.nb.select(c_frame)

        # v3.4: back to being a real Notebook tab (a floating toggle button
        # felt wrong/inconsistent with the other tabs). ttk.Notebook has no
        # way to pin a tab flush to the far-right edge -- tabs are always
        # left-aligned/sequential -- so as a fallback this uses a wide
        # disabled (unclickable) spacer tab to visually push Help further
        # away from File Content instead.
        _spacer_frame = tk.Frame(self.nb)
        self.nb.add(_spacer_frame, text=" " * 28)
        self.nb.tab(_spacer_frame, state="disabled")

        help_frame = tk.Frame(self.nb, bg=BG_COLOR)
        help_split = tk.PanedWindow(help_frame, orient="horizontal", bg=BG_COLOR,
                                     sashwidth=4, sashrelief="raised", bd=0)
        help_split.pack(fill="both", expand=True, padx=2, pady=2)
        # Bordered panels (grooved relief) so the split is easy to read even
        # though the surrounding window has no title bar / chrome of its own.
        help_left = tk.Frame(help_split, bg=BG_COLOR, relief="groove", bd=2)
        help_right = tk.Frame(help_split, bg=BG_COLOR, relief="groove", bd=2)
        help_split.add(help_left, minsize=250)
        help_split.add(help_right, minsize=250)
        self._build_help_content(help_left)
        # v1.8: help_right is now split top/bottom -- Search History on top
        # (unchanged), "AI Chat History" (read-only, Copy-only right-click,
        # persisted to HISTORY_DB_FILE) on the bottom -- instead of the live
        # AI Chat panel bleeding into the Help tab (it's now hidden here,
        # see _sync_online_chat_panel_visibility).
        help_right_split = tk.PanedWindow(help_right, orient="vertical", bg=BG_COLOR,
                                           sashwidth=4, sashrelief="raised", bd=0)
        help_right_split.pack(fill="both", expand=True)
        hr_top = tk.Frame(help_right_split, bg=BG_COLOR)
        hr_bottom = tk.Frame(help_right_split, bg=BG_COLOR)
        help_right_split.add(hr_top, minsize=100)
        help_right_split.add(hr_bottom, minsize=100)
        tk.Label(hr_top, text=" Search History ", font=("Segoe UI", 9, "bold"),
                 bg=BG_COLOR, fg="#7ec8e3", anchor="w").pack(fill="x", padx=4, pady=(4, 0))
        _history_panel = HistoryPanel(hr_top, self)
        tk.Label(hr_bottom, text=" AI Chat History ", font=("Segoe UI", 9, "bold"),
                 bg=BG_COLOR, fg="#7ec8e3", anchor="w").pack(fill="x", padx=4, pady=(4, 0))
        self.chat_history_panel = ChatHistoryPanel(hr_bottom, self)
        # v-new: HistoryPanel.__init__ already tried to auto-select the
        # latest row (see _apply_auto_follow), but self.chat_history_panel
        # didn't exist yet at that point (it's built right above this
        # line), so that first attempt silently did nothing. Re-apply now
        # that both panes exist, so AI Chat History shows the latest
        # keyword immediately instead of only after the first 5s tick.
        _history_panel._apply_auto_follow()
        self.nb.add(help_frame, text=" Help ")

        # Default the Help/History split to 50%-50% once the pane has a real
        # width (can't compute % of width before it's actually laid out).
        # Only fires once -- if the user drags the sash afterwards, further
        # window resizes won't snap it back to 50/50.
        _sash_done = {"v": False}
        def _center_help_sash(event=None):
            if _sash_done["v"]:
                return
            try:
                total_w = help_split.winfo_width()
                if total_w > 20:
                    help_split.sash_place(0, total_w // 2, 0)
                    _sash_done["v"] = True
            except Exception:
                pass
        help_split.bind("<Configure>", _center_help_sash, add="+")

        # Same 50%-50% snap for the new vertical split (Search History /
        # AI Chat History) inside help_right.
        # v-fix: this used to snap ONCE, on the very first <Configure> event
        # whose height happened to exceed 20px, then set a "done" flag and
        # never touch the sash again. In practice that first qualifying
        # event often fires while the window is still being laid out (not
        # yet at its final size), so the 50/50 split was computed against a
        # too-small height and then "locked in" -- once the window reached
        # its real size, Search History ended up taking most of the space
        # and AI Chat History was squeezed into a sliver, exactly as seen
        # in the screenshots. Now we keep re-centering to 50/50 on EVERY
        # resize, and only stop once the user manually drags the sash
        # themselves (tracked via a click on the sash region).
        _sash_user_moved = {"v": False}
        def _center_help_right_sash(event=None):
            if _sash_user_moved["v"]:
                return
            try:
                total_h = help_right_split.winfo_height()
                if total_h > 20:
                    help_right_split.sash_place(0, 0, total_h // 2)
            except Exception:
                pass
        def _mark_help_right_sash_user_moved(event=None):
            # A click on the sash itself means the user is dragging it --
            # stop auto-recentering so their manual placement sticks.
            try:
                if help_right_split.identify(event.x, event.y):
                    _sash_user_moved["v"] = True
            except Exception:
                pass
        help_right_split.bind("<Configure>", _center_help_right_sash, add="+")
        help_right_split.bind("<ButtonPress-1>", _mark_help_right_sash_user_moved, add="+")

        # v-fix: the <Configure> binds above only fire on a REAL resize
        # after this point -- but at the moment this runs, the Help tab is
        # very likely not the selected Notebook tab yet (File Name usually
        # is), so its frame has never been mapped and winfo_width()/
        # winfo_height() can report a bogus placeholder size (or 0), which
        # fails the "> 20" guard and silently skips centering -- nothing
        # ever fires again afterwards unless the user happens to manually
        # resize the window. Explicitly re-run both centering functions a
        # few times shortly after creation (covers the case where results
        # open straight onto File tabs) AND every time the Notebook tab
        # actually changes to Help (covers the case where the Help tab is
        # only visited/mapped for the first time well after results
        # opened) -- so the 50/50 split reliably lands correctly either way.
        for _delay in (60, 250, 600):
            self.root.after(_delay, _center_help_sash)
            self.root.after(_delay, _center_help_right_sash)

        def _on_tab_changed(event=None):
            try:
                if self.nb.select() == str(help_frame):
                    self.root.after(30, _center_help_sash)
                    self.root.after(30, _center_help_right_sash)
            except Exception:
                pass
        self.nb.bind("<<NotebookTabChanged>>", _on_tab_changed, add="+")

        self.tree_c.search_id = version
        self.tree_f.search_id = version
        self.tree_fol.search_id = version

        # Initialize pane-tree dicts so filters work from the very first pane
        self._c_pane_trees   = {"main": self.tree_c}
        self._f_pane_trees   = {"main": self.tree_f}
        self._fol_pane_trees = {"main": self.tree_fol}

        def bg_load_ui(tree, data, mode, current_vid):
            if not data or getattr(tree, 'search_id', 0) != current_vid: return
            CHUNK = 500
            first_chunk, remaining_chunk = data[:CHUNK], data[CHUNK:]
            offset = [0]
            for item in first_chunk:
                if getattr(tree, 'search_id', 0) != current_vid: return
                try:
                    offset[0] += 1
                    self._insert_row(tree, item, offset[0], mode)
                except: pass

            def _update_count():
                if mode == 1 and self.filter_count_label:
                    self.filter_count_label.config(text=f"{len(self._all_files_data)} files")
                elif mode == 0 and self.content_filter_count_label:
                    self.content_filter_count_label.config(text=f"{len(self._all_content_data)} files")

            def load_rest(step=0):
                if not self.active_result_win or not tk.Toplevel.winfo_exists(self.active_result_win): return
                if getattr(tree, 'search_id', 0) != current_vid: return
                sub_chunk = remaining_chunk[step: step + CHUNK]
                if sub_chunk:
                    for item in sub_chunk:
                        if getattr(tree, 'search_id', 0) != current_vid: return
                        try:
                            offset[0] += 1
                            self._insert_row(tree, item, offset[0], mode)
                        except: pass
                    self.root.after(5, lambda: load_rest(step + CHUNK))
                else:
                    _update_count()
            if remaining_chunk:
                self.root.after(10, lambda: load_rest(0))
            else:
                self.root.after(20, _update_count)

        bg_load_ui(self.tree_c, cont_res, 0, version)
        bg_load_ui(self.tree_f, only_files, 1, version)
        bg_load_ui(self.tree_fol, only_folders, 2, version)

        def on_tree_click(event, tree):
            row = tree.identify_row(event.y); col = tree.identify_column(event.x)
            if row and col:
                now = time.time()
                if getattr(tree, "_last_row", None) == row and getattr(tree, "_last_col", None) == col and (now - getattr(tree, "_last_time", 0)) > 0.4:
                    show_cell_edit(tree, row, col)
                tree._last_row, tree._last_col, tree._last_time = row, col, now

        def show_cell_edit(tree, row, col):
            if self.res_edit: self.res_edit.destroy()
            if col == "#0": return  # v5.6: icon+# combined column, nothing to edit
            idx = int(col[1:]) - 1
            val = str(tree.item(row, "values")[idx])
            if val.startswith(("📁 ", "📕 ", "📝 ", "📊 ", "📉 ", "🔮 ", "📄 ", "📎 ", "📧 ")): val = val[2:]
            bbox = tree.bbox(row, col)
            if not bbox: return  # cell not visible (row scrolled out or window resized)
            x, y, w, h = bbox
            self.res_edit = tk.Entry(tree, font=("Segoe UI", 9), bd=0); self.res_edit.insert(0, val)
            self.res_edit.place(x=x, y=y, width=w, height=h); self.res_edit.focus_set()
            add_only_copy_menu(self.res_edit); self.res_edit.bind("<FocusOut>", lambda e: self.res_edit.destroy())

        # Extensions that open directly (Office, PDF, text, logs, etc.)
        OPEN_DIRECT_EXTS = {
            '.docx', '.doc', '.docm', '.xlsx', '.xls', '.xlsm',
            '.pptx', '.ppt', '.pptm', '.pdf', '.msg', '.one',
            '.txt', '.csv', '.tsv', '.json', '.xml', '.yaml', '.yml',
            '.ini', '.cfg', '.conf', '.toml', '.md',
            '.log', '.html', '.htm', '.bat', '.ps1', '.sh',
            '.py', '.js', '.ts', '.cpp', '.c', '.h', '.java',
            '.fmu', '.scen', '.mdl', '.m', '.slx',
        }
        # Extensions to open with Notepad++ if available, else Notepad
        NOTEPAD_EXTS = {
            '.log', '.html', '.htm', '.bat', '.ps1', '.sh',
            '.py', '.js', '.ts', '.cpp', '.c', '.h', '.java',
            '.ini', '.cfg', '.conf', '.toml', '.yaml', '.yml',
            '.json', '.xml', '.md', '.txt', '.csv', '.tsv',
        }
        # Find Notepad++ once
        _npp_paths = [
            r"C:\Program Files\Notepad++\notepad++.exe",
            r"C:\Program Files (x86)\Notepad++\notepad++.exe",
        ]
        NOTEPAD_PP = next((p for p in _npp_paths if os.path.isfile(p)), None)

        def _normalize_path(raw):
            return os.path.abspath(os.path.normpath(raw.replace("¥", "\\")))

        def open_file(tree):
            """Double-click: open file directly if possible, else open folder."""
            sel = tree.selection()
            if not sel: return
            try:
                values = tree.item(sel[0])["values"]
                raw_path = str(values[-1])
                if _is_outlook_pseudo_path(raw_path):
                    _open_outlook_result(raw_path)
                    return
                if _is_onenote_pseudo_path(raw_path):
                    _open_onenote_result(raw_path)
                    return
                if raw_path.startswith(("📁 ", "📕 ", "📝 ", "📊 ", "📉 ", "🔮 ", "📄 ", "📎 ", "📧 ")):
                    raw_path = raw_path[2:]
                abs_p = _normalize_path(raw_path)
                ext = os.path.splitext(abs_p)[1].lower()
                has_ext = bool(ext)
                is_file = has_ext and not os.path.isdir(abs_p)
                if is_file and ext in OPEN_DIRECT_EXTS:
                    if ext in NOTEPAD_EXTS:
                        editor = NOTEPAD_PP or "notepad.exe"
                        subprocess.Popen([editor, abs_p])
                    else:
                        try:
                            os.startfile(abs_p)
                        except:
                            subprocess.Popen(f'cmd /c start "" "{abs_p}"', shell=True)
                    return
                # Not openable directly → fall through to open folder
                open_explorer(tree)
            except Exception as ex:
                print(f"Open File Error: {ex}")

        def _open_outlook_result(pseudo_path):
            """Open the mail in Outlook itself (Display()) — this must run
            off the UI thread since it's a COM call into another
            application and can briefly block while Outlook responds."""
            entry_id, store_id = _parse_outlook_pseudo_path(pseudo_path)
            if not entry_id:
                return
            def _do_open():
                ok, err = outlook_search.open_outlook_item(entry_id, store_id)
                if not ok:
                    print(f"[Outlook] Could not open item: {err}")
                    self.root.after(0, lambda: messagebox.showerror(
                        "Outlook", f"Không mở được email này trong Outlook:\n{err}"))
            threading.Thread(target=_do_open, daemon=True).start()

        def _open_onenote_result(pseudo_path):
            """Open the page in OneNote itself (NavigateTo) — same
            off-UI-thread reasoning as _open_outlook_result."""
            page_id = _parse_onenote_pseudo_path(pseudo_path)
            if not page_id:
                return
            def _do_open():
                ok, err = onenote_search.open_onenote_page(page_id)
                if not ok:
                    print(f"[OneNote] Could not open page: {err}")
                    self.root.after(0, lambda: messagebox.showerror(
                        "OneNote", f"Không mở được trang này trong OneNote:\n{err}"))
            threading.Thread(target=_do_open, daemon=True).start()

        def open_explorer(tree, is_content=False):
            sel = tree.selection()
            if not sel: return
            try:
                values = tree.item(sel[0])["values"]
                raw_path = str(values[-1])
                if _is_outlook_pseudo_path(raw_path):
                    _open_outlook_result(raw_path)  # "Open Folder" makes no sense for mail — just open the item itself
                    return
                if _is_onenote_pseudo_path(raw_path):
                    _open_onenote_result(raw_path)  # same reasoning — no folder concept for a OneNote page
                    return
                if raw_path.startswith(("📁 ", "📕 ", "📝 ", "📊 ", "📉 ", "🔮 ", "📄 ", "📎 ", "📧 ")):
                    raw_path = raw_path[2:]
                abs_p = _normalize_path(raw_path)
                has_ext = bool(os.path.splitext(abs_p)[1])
                is_file = has_ext and not os.path.isdir(abs_p)
                folder  = os.path.dirname(abs_p) if is_file else abs_p
                if is_file and len(abs_p) <= 250:
                    try:
                        subprocess.Popen(f'cmd /c explorer /select,"{abs_p}"', shell=True)
                        return
                    except:
                        pass
                try:
                    os.startfile(folder)
                except:
                    subprocess.Popen(f'cmd /c explorer "{folder}"', shell=True)
            except Exception as ex:
                print(f"Explorer Failure: {ex}")

        def show_context_menu(e, tree, is_folder=False, is_content=False):
            row = tree.identify_row(e.y)
            if not row:
                return
            tree.selection_set(row)
            val = list(tree.item(row)["values"])
            name_val = str(val[1])
            if name_val.startswith(("📁 ", "📕 ", "📝 ", "📊 ", "📉 ", "🔮 ", "📄 ", "📎 ", "📧 ")):
                name_val = name_val[2:]
            full_path = str(val[-1])
            ctx_win = self._result_win if hasattr(self, '_result_win') and self._result_win.winfo_exists() else win
            if _is_outlook_pseudo_path(full_path):
                meta = self._outlook_meta.get(full_path, {})
                m = tk.Menu(ctx_win, tearoff=0)
                m.add_command(label="Open in Outlook", command=lambda p=full_path: _open_outlook_result(p))
                m.add_separator()
                m.add_command(label="Copy Subject", command=lambda v=meta.get("subject", ""): (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
                m.add_command(label="Copy Sender",  command=lambda v=meta.get("sender", ""):  (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
                m.post(e.x_root, e.y_root)
                return
            if _is_onenote_pseudo_path(full_path):
                meta = self._onenote_meta.get(full_path, {})
                m = tk.Menu(ctx_win, tearoff=0)
                m.add_command(label="Open in OneNote", command=lambda p=full_path: _open_onenote_result(p))
                m.add_separator()
                m.add_command(label="Copy Title",   command=lambda v=meta.get("title", ""):        (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
                m.add_command(label="Copy Section",  command=lambda v=meta.get("section_path", ""): (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
                m.post(e.x_root, e.y_root)
                return
            ext = os.path.splitext(full_path)[1].lower()
            # Use self._result_win so this works from any pane (split or not)
            m = tk.Menu(ctx_win, tearoff=0)
            if is_content and ext in OPEN_DIRECT_EXTS:
                m.add_command(label="Open File",   command=lambda t=tree: open_file(t))
                m.add_command(label="Open Folder", command=lambda t=tree: open_explorer(t))
            else:
                m.add_command(label="Open Folder", command=lambda t=tree: open_explorer(t))
            m.add_separator()
            lbl_name = "Copy Folder Name" if is_folder else "Copy File Name"
            lbl_path = "Copy Folder Path" if is_folder else "Copy File Path"
            m.add_command(label=lbl_name, command=lambda v=name_val:  (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
            m.add_command(label=lbl_path, command=lambda v=full_path: (ctx_win.clipboard_clear(), ctx_win.clipboard_append(v)))
            m.post(e.x_root, e.y_root)

        def dbl_select(tree):
            row = tree.identify_row(tree.winfo_pointery() - tree.winfo_rooty())
            if row: tree.selection_set(row); tree.focus(row)

        # ── Store handlers as instance attrs so _wire_tree_clicks can use them
        self._fn_open_file     = open_file
        self._fn_open_explorer = open_explorer
        self._fn_show_ctx_menu = show_context_menu
        self._fn_on_tree_click = on_tree_click
        self._fn_dbl_select    = dbl_select

        # ── Sortable columns ──────────────────────────────────────────────
        # _sort_tree / _make_sortable have been promoted to class methods (see
        # above, near _build_tree_*) so all Advanced/AI panes created after split are also
        # automatically sortable. Only need to wire the 3 root trees (Default) here.
        self._make_sortable(self.tree_c,   ["size", "mtime", "ftype"])
        self._make_sortable(self.tree_f,   ["size", "mtime", "ftype"])
        self._make_sortable(self.tree_fol, ["mtime"])
        # ─────────────────────────────────────────────────────────────────

        # Wire click events on the initial (main) trees
        self._wire_tree_clicks(self.tree_f,   "f")
        self._wire_tree_clicks(self.tree_c,   "c")
        self._wire_tree_clicks(self.tree_fol, "fol")

        # v2.6: the Notebook (self.nb) was created/packed AFTER the resize_grip,
        # so it stacks visually on top and would swallow the grip's clicks.
        # Re-raise the grip now that every widget in this window has been built.
        try:
            resize_grip.lift()
        except Exception:
            pass
        # v-fix: also cover the very first search of a session (this
        # branch only, since update_or_show_results handles every search
        # after the window already exists) so the chat placeholder shows
        # "Ask anything about {query}" from the first search onward, not
        # only starting on the second one.
        self._update_chat_placeholder(query)

    # ── Smart Size + Extension Filter ────────────────────────────────────
    def _get_size_filter_bytes(self):
        """Return (op, bytes) from dropdown vars, or (None, None) when inactive."""
        op = self.size_op_var.get()
        if op == "Any":
            return None, None
        try:
            val = float(self.size_num_var.get())
        except ValueError:
            return None, None
        unit = self.size_unit_var.get()
        mult = {'B': 1, 'KB': 1024, 'MB': 1024**2, 'GB': 1024**3}.get(unit, 1)
        return op, int(val * mult)

    def _get_ext_filter(self):
        """Parse ext_filter_var → frozenset of .ext strings, e.g. {'.pdf','.docx'}, or None."""
        raw = self.ext_filter_var.get().strip()
        if not raw:
            return None
        exts = set()
        for e in raw.replace(',', ' ').split():
            e = e.strip().lower().lstrip('.')
            if e:
                exts.add('.' + e)
        return frozenset(exts) if exts else None

    def _update_pane_label(self, tree, shown, total):
        """Update the header label of a pane tree with filtered count."""
        try:
            lbl = getattr(tree, '_header_label', None)
            base = getattr(tree, '_base_label', '')
            if not lbl or not lbl.winfo_exists():
                return
            # Extract base text without count: e.g. '📄 Default (33)' → '📄 Default'
            import re as _re
            base_no_count = _re.sub(r'\s*\(.*\)\s*$', '', base).strip()
            if shown < total:
                lbl.config(text=f"{base_no_count} ({shown}/{total})")
            else:
                lbl.config(text=f"{base_no_count} ({total})")
        except Exception:
            pass

    def _apply_folder_filter_to_tree(self, *_):
        """Folder Name tab equivalent of _apply_size_filter_to_tree — only
        Name + Whole word apply here (no Size/Ext filters for folders).
        Reuses the exact same merge function as the live folder renderer
        (_merge_mft_with_db) so this path and _render_mft_folder_tree
        always produce the same order — avoiding the same MFT-first vs
        DB-first divergence that was reshuffling the File Name tab."""
        if not hasattr(self, 'tree_fol') or not self.active_result_win:
            return
        try:
            if not tk.Toplevel.winfo_exists(self.active_result_win):
                return
            if not self.tree_fol.winfo_exists():
                return
        except Exception:
            return

        mft_folders = [r for r in (self._mft_folder_res or []) if str(r[0]).lower() == "folder"]
        adv_folders = [r for r in (self._adv_all_files or [])  if str(r[0]).lower() == "folder"]
        ai_folders  = [r for r in (self._ai_file_res or [])    if str(r[0]).lower() == "folder"]

        main_folders = self._merge_mft_with_db(mft_folders, is_folder=True)
        main_folders = self._sort_priority(main_folders, 1)

        pane_sources = {"main": main_folders, "adv": adv_folders, "ai": ai_folders}
        name_needle = _norm_txt(self.name_filter_var.get().strip().lower())

        for pane_key, tree in self._fol_pane_trees.items():
            try:
                if not tree.winfo_exists(): continue
            except Exception:
                continue
            source = pane_sources.get(pane_key, main_folders)
            if name_needle:
                filtered = [r for r in source if name_needle in _norm_txt(os.path.basename(r[2]).lower())]
            else:
                filtered = source
            x0 = self._save_xview(tree)
            y0 = self._save_yview(tree)
            widths = self._save_col_widths(tree)
            for row in tree.get_children(): tree.delete(row)
            for rn, item in enumerate(filtered, start=1):
                try:
                    path = item[2]
                    img = get_tree_icon_image(path, is_folder=True)
                    name_txt = item[1] if img else "📁 " + item[1]
                    mt = item[4] if len(item) > 4 and item[4] is not None else get_live_mtime(path)
                    tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(name_txt, mt, path, path))
                except Exception:
                    pass
            self._restore_col_widths(tree, widths)
            self._restore_xview(tree, x0)
            self._restore_yview(tree, y0)
            self._update_pane_label(tree, len(filtered), len(source))

    def _apply_size_filter_to_tree(self, *_):
        """Client-side filter: size, extension, AND name substring.
        Applies to ALL visible panes in File Name tab, each with its own data source."""
        if not hasattr(self, 'tree_f') or not self.active_result_win:
            return
        try:
            if not tk.Toplevel.winfo_exists(self.active_result_win):
                return
            if not self.tree_f.winfo_exists():
                return
        except Exception:
            return

        # Data source per pane key — each pane must only show its own results
        # "main" pane = MFT (realtime scan) blended with BM25/DB, deduplicated
        # by path.
        #
        # v-fix: this used to do its OWN separate dedupe here (MFT rows
        # first, then BM25/DB rows) — the OPPOSITE priority order from
        # _merge_mft_with_db() (DB rows first, then MFT-only), which is what
        # the live-streaming renderer (_render_mft_file_tree) actually uses
        # to paint the screen. Since _sort_priority bands its OUTPUT order
        # around whatever order it's fed, feeding it "MFT-first" here vs.
        # "DB-first" there produced a different final order for the exact
        # same underlying data — so simply touching ANY filter (even
        # clicking "✕" on an already-empty box) would swap the visible
        # order away from what was on screen, with nothing about the
        # search itself having changed. Reusing the one shared merge
        # function guarantees this path and the live renderer always agree.
        mft_files  = [r for r in (self._mft_file_res or []) if str(r[0]).lower() != "folder"]
        adv_files  = [r for r in (self._adv_all_files or [])      if str(r[0]).lower() != "folder"]
        ai_files   = [r for r in (self._ai_file_res or [])        if str(r[0]).lower() != "folder"]

        main_files = [r for r in self._merge_mft_with_db(mft_files, is_folder=False)]

        # v-fix: this used to re-apply the category sort (office/pdf first,
        # log/code last) to ALL THREE sources, including adv_files/ai_files
        # — which silently threw away their BM25/AI relevance ranking every
        # time ANY filter changed (name, extension, or size), even just
        # clearing a filter back to empty. That's why a good AI/BM25 sort
        # would visibly reshuffle the moment you touched the Name filter's
        # "✕". Only "main" (the plain MFT/DB blended default view) is
        # supposed to follow the category convention — same as the File
        # Content tab's equivalent function, which never re-sorts its
        # adv_cont/ai_cont sources either.
        main_files = self._sort_priority(main_files, 1)

        pane_sources = {"main": main_files, "adv": adv_files, "ai": ai_files}

        total_shown = 0
        for pane_key, tree in self._f_pane_trees.items():
            try:
                if not tree.winfo_exists(): continue
            except Exception:
                continue
            source = pane_sources.get(pane_key, main_files)
            filtered = self._filter_file_rows(source)
            x0 = self._save_xview(tree)
            y0 = self._save_yview(tree)
            widths = self._save_col_widths(tree)
            for row in tree.get_children(): tree.delete(row)
            for rn, item in enumerate(filtered, start=1):
                try:
                    path = item[2]
                    pf = self._pseudo_path_row_fields(path)
                    if pf:
                        name_txt, size_str, mtime_str, ftype_str, loc_str, img = pf
                        tree.insert("", tk.END, text="", **({"image": img} if img else {}),
                                    values=(name_txt, size_str, mtime_str, ftype_str, loc_str, path))
                        continue
                    img = get_tree_icon_image(path, is_folder=False)
                    name_txt = item[1] if img else get_file_icon(path, is_folder=False) + item[1]
                    readable_sz = format_size(get_live_size(path, item[3]))
                    mtime = get_live_mtime(path)
                    ftype = get_file_type(path)
                    tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(name_txt, readable_sz, mtime, ftype, os.path.dirname(path), path))
                except Exception:
                    pass
            self._restore_col_widths(tree, widths)
            self._restore_xview(tree, x0)
            self._restore_yview(tree, y0)
            self._update_pane_label(tree, len(filtered), len(source))
            if pane_key == "main":
                total_shown = len(filtered)
    # ─────────────────────────────────────────────────────────────────────

    def _get_content_size_filter(self):
        op = self.c_size_op_var.get()
        if op == "Any":
            return None, None
        try:
            val = float(self.c_size_num_var.get())
        except ValueError:
            return None, None
        unit = self.c_size_unit_var.get()
        mult = {'B': 1, 'KB': 1024, 'MB': 1024**2, 'GB': 1024**3}.get(unit, 1)
        return op, int(val * mult)

    def _get_content_ext_filter(self):
        raw = self.c_ext_filter_var.get().strip()
        if not raw:
            return None
        exts = set()
        for e in raw.replace(',', ' ').split():
            e = e.strip().lower().lstrip('.')
            if e:
                exts.add('.' + e)
        return frozenset(exts) if exts else None

    def _apply_content_filter_to_tree(self, *_):
        """Client-side filter for File Content tab: size, extension, AND name substring.
        Applies to ALL visible panes, each with its own data source."""
        if not hasattr(self, 'tree_c') or not self.active_result_win:
            return
        try:
            if not tk.Toplevel.winfo_exists(self.active_result_win):
                return
            if not self.tree_c.winfo_exists():
                return
        except Exception:
            return

        op, size_bytes = self._get_content_size_filter()
        ext_filter = self._get_content_ext_filter()
        name_needle = _norm_txt(self.c_name_filter_var.get().strip().lower())

        # Data source per pane key — each pane must only show its own results
        bm25_cont = self._last_bm25_cont_res or []
        adv_cont  = self._adv_all_cont  or []
        ai_cont   = self._ai_cont_res   or []
        pane_sources = {"main": bm25_cont, "adv": adv_cont, "ai": ai_cont}

        def _filter_content(data):
            if op is None and ext_filter is None and not name_needle:
                return data
            cmp_fn = {'>':  lambda a, b: a >  b, '>=': lambda a, b: a >= b,
                      '<':  lambda a, b: a <  b, '<=': lambda a, b: a <= b,
                      '=':  lambda a, b: a == b}.get(op) if op else None
            result = []
            for item in data:
                fpath = item[0]
                if cmp_fn is not None:
                    if _is_outlook_pseudo_path(fpath):
                        live_size = (self._outlook_meta.get(fpath) or {}).get("size", 0) or 0
                    elif _is_onenote_pseudo_path(fpath):
                        live_size = 0  # OneNote pages have no meaningful byte size
                    else:
                        live_size = get_live_size(fpath, 0)
                    if not cmp_fn(live_size, size_bytes):
                        continue
                if ext_filter is not None:
                    _, ext = os.path.splitext(fpath)
                    if ext.lower() not in ext_filter:
                        continue
                if name_needle:
                    fname = _norm_txt(os.path.basename(fpath).lower())
                    if name_needle not in fname:
                        continue
                result.append(item)
            return result

        total_shown = 0
        total_src   = 0
        for pane_key, tree in self._c_pane_trees.items():
            try:
                if not tree.winfo_exists(): continue
            except Exception:
                continue
            source = pane_sources.get(pane_key, bm25_cont)
            filtered = _filter_content(source)
            for row in tree.get_children(): tree.delete(row)
            for rn, item in enumerate(filtered, start=1):
                try:
                    fpath = item[0]
                    pf = self._pseudo_path_row_fields(fpath)
                    if pf:
                        name_txt, size_str, mtime_str, ftype_str, _loc_str, img = pf
                        tree.insert("", tk.END, text="", **({"image": img} if img else {}),
                                    values=(name_txt, size_str, mtime_str, ftype_str, fpath))
                        continue
                    img = get_tree_icon_image(fpath, is_folder=False)
                    name_txt = fpath if img else get_file_icon(fpath, is_folder=False) + fpath
                    readable_sz = format_size(get_live_size(fpath, 0))
                    mtime = get_live_mtime(fpath)
                    ftype = get_file_type(fpath)
                    tree.insert("", tk.END, text="", **({"image": img} if img else {}), values=(name_txt, readable_sz, mtime, ftype, fpath))
                except Exception:
                    pass
            self._update_pane_label(tree, len(filtered), len(source))
            if pane_key == "main":
                total_shown = len(filtered)
                total_src   = len(source)
    # ─────────────────────────────────────────────────────────────────────

    def _save_hist(self, q):
        try:
            # v7.10 FIX: this used to connect to DB_FILE (search_data.db)
            # directly, which meant simply performing a search -- even
            # with --update data NEVER run -- created/grew search_data.db
            # on disk. On the next launch the ramp light saw that file
            # "exists" and jumped from Red to Yellow, even though nothing
            # was ever actually indexed. History now lives in its own
            # HISTORY_DB_FILE so searching can never touch search_data.db.
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            # v7.4 FIX: history table used to only get created inside
            # indexing_worker (--update data), so on a fresh/never-updated
            # DB (search_data.db = 0KB) every save silently failed with
            # "no such table: history". Ensure it exists here too.
            c.execute("CREATE TABLE IF NOT EXISTS history (id INTEGER PRIMARY KEY, query TEXT, date TEXT)")
            c.execute("INSERT INTO history (query, date) VALUES (?,?)", (q, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))); conn.commit(); conn.close()
            self._last_saved_hist_query = q
            print(f"[History] Saved: {q!r}")
        except Exception as _e:
            print(f"[History] Save FAILED for {q!r}: {_e}")

    def _save_chat_hist(self, role, text, session_id=None, query=None):
        """v1.8: persist one AI Chat turn (role: 'user'/'assistant') to the
        same HISTORY_DB_FILE used by Search History, in its own table --
        this is what feeds the "AI Chat History" pane in the Help tab (see
        ChatHistoryPanel), so past chats survive across searches and app
        restarts, same as Search History already does. v1.9: also tags the
        row with self._chat_session_id so ChatHistoryPanel can group turns
        into one collapsible entry per conversation instead of one row per
        message.

        v1.11 FIX: also store the exact search query this session belongs
        to (self._chat_auto_sent_for, set once in _auto_ai_chat_after_search
        when the session is auto-started for a search). ChatHistoryPanel
        used to have NO direct link between a session and its query -- it
        had to regex-reparse the auto-summary prompt out of the first saved
        message and string-compare that to the Search History row text,
        which silently failed to match on anything but an exact
        case/whitespace/punctuation match, making "AI Chat History" look
        like it was never updating. Storing the query directly removes that
        guesswork entirely.

        v-fix (per user request): the v1.12 behavior below has been REMOVED.
        It used to wipe out any older rows saved under the same query the
        FIRST time a new session's turn was saved -- meant as "1 keyword =
        1 saved entry", but in practice this deleted the original
        auto-summary the moment the user typed anything into a manual "New
        Chat" started on the SAME still-active keyword, so the AI Chat
        History pane ended up showing only whichever session happened to
        save last -- often looking like it was stuck on stale content with
        no sign the New Chat conversation ever happened.

        Now every session's rows are simply kept. ChatHistoryPanel already
        groups by query and concatenates ALL of that query's sessions in
        chronological order (see _find_sessions_for_query/show_for_query),
        so: searching "Simpack realtime" saves session A under that query;
        clicking "New Chat" and continuing (keyword unchanged) saves
        session B under the SAME query, and both A and B now show up
        together, in order, under the one "Simpack realtime" entry.
        Searching a different keyword ("Abaqus FAQ") starts its own
        session under ITS OWN query and never touches "Simpack realtime"'s
        rows at all.

        v-fix (Vấn đề 2): session_id/query are now accepted as explicit
        parameters -- the caller (_online_chat_worker) captures them at the
        moment the request was SENT and passes them straight through, so a
        reply that finishes after the user has already moved on to a
        different search/session still gets filed under the query it was
        actually asked about, not under whatever query happens to be
        active by the time the AI responds. Falls back to reading the
        instance attributes only if a caller doesn't pass them."""
        try:
            conn = sqlite3.connect(HISTORY_DB_FILE); c = conn.cursor()
            c.execute("CREATE TABLE IF NOT EXISTS chat_history "
                      "(id INTEGER PRIMARY KEY, role TEXT, text TEXT, date TEXT, session_id INTEGER DEFAULT 0)")
            # v1.9/v1.11: older DBs won't have these columns yet -- add them
            # on the fly (SQLite ALTER TABLE ADD COLUMN is safe/cheap).
            for _ddl in ("ALTER TABLE chat_history ADD COLUMN session_id INTEGER DEFAULT 0",
                         "ALTER TABLE chat_history ADD COLUMN query TEXT DEFAULT ''"):
                try:
                    c.execute(_ddl)
                except Exception:
                    pass  # already exists
            if session_id is None:
                session_id = getattr(self, "_chat_session_id", 0)
            if query is None:
                query = getattr(self, "_chat_auto_sent_for", "") or ""
            c.execute("INSERT INTO chat_history (role, text, date, session_id, query) VALUES (?,?,?,?,?)",
                      (role, text, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), session_id, query))
            conn.commit(); conn.close()
        except Exception as _e:
            print(f"[Chat History] Save FAILED: {_e}")

    def _maybe_save_hist_idle(self, snapshot):
        """Fired ~10s after the user stops typing without pressing Enter.
        The realtime preview means Enter is rarely needed to SEE results,
        but that also meant most searches never got saved to History at
        all. Save the query anyway once the user has genuinely paused on
        it -- but only if the box still holds that exact text (they didn't
        keep typing/edit since).

        v-fix: this used to also skip saving if `snapshot` matched
        `self._last_saved_hist_query` (the previous entry saved, ever --
        not just "just now"). That meant re-searching the SAME term later
        (minutes/hours/days apart) silently stopped being logged the
        moment it happened to equal whatever was last saved, which isn't
        what a "History" is supposed to do -- every genuine search pause
        is its own timestamped event and should be recorded, even if the
        query text repeats. The only thing actually worth guarding against
        (an idle-save firing right after the user already hit Enter for
        the exact same text) is already handled elsewhere: pressing Enter
        cancels the pending idle timer outright (see the after_cancel call
        before _save_hist in the Enter handler), so that timer can never
        fire a second time for the same instant."""
        self._hist_idle_timer = None
        current = self.entry_var.get().strip()
        # v-fix (Search History không lưu keyword): entry_var gets cleared to
        # "" as soon as the results window closes (_close_results_window) or
        # certain other UI resets run — which routinely happens WITHIN this
        # 10s window on the normal "type -> glance at results -> close" flow,
        # since realtime search never requires pressing Enter. The old check
        # compared against whatever the box holds RIGHT NOW, so a search the
        # user genuinely ran silently never got saved just because they'd
        # already closed the results before the timer fired. Only bail out
        # when the box now holds a DIFFERENT, NON-EMPTY query (i.e. the user
        # kept typing something else) — an empty box afterward is not
        # evidence the search never happened, just that the UI reset.
        if current and current != snapshot:
            print(f"[History] Idle-save skipped: box changed ({snapshot!r} -> {current!r})")
            return  # stale — user moved on to a different query since this timer was scheduled
        if not snapshot:
            print("[History] Idle-save skipped: empty query")
            return
        if snapshot.lower().startswith("-"):
            print(f"[History] Idle-save skipped: looks like a command ({snapshot!r})")
            return  # commands like --update data / --exit aren't search queries
        self._save_hist(snapshot)

    def _build_help_content(self, parent):
        """Left pane of the Help tab — static usage text. Used to be a
        separate popup opened via typing '--h'/'--help'; now it's just
        always-visible content inside the Help tab.

        v-new: VI / EN / JA buttons next to the "Help" label (same visual
        pattern as the AI Chat panel's language toggle -- see
        _build_online_chat_panel / _select_chat_lang) let the user switch
        the whole Help text between Vietnamese, English and Japanese.
        Unlike the AI Chat toggle, this does NOT call the AI -- the three
        versions are pre-written (see HELP_CONTENT) so switching is
        instant and works even with no internet / AI quota."""
        header = tk.Frame(parent, bg=BG_COLOR)
        header.pack(fill="x", padx=4, pady=(4, 0))
        tk.Label(header, text=" Help ", font=("Segoe UI", 9, "bold"),
                 bg=BG_COLOR, fg="#7ec8e3", anchor="w").pack(side="left")

        lang_row = tk.Frame(header, bg=BG_COLOR)
        lang_row.pack(side="left", padx=(8, 0))
        self._help_lang_btns = {}

        txt = tk.Text(parent, bg=BG_COLOR, fg=TEXT_COLOR, font=("Segoe UI", 10), padx=15, pady=15, bd=0)
        txt.pack(fill="both", expand=True)
        self._help_text_widget = txt

        def _select_help_lang(code):
            self._help_lang = code
            for c, b in self._help_lang_btns.items():
                if c == code:
                    b.config(bg="#1a5fb4", fg="#ffffff", relief="sunken")
                else:
                    b.config(bg="#e4e6ea", fg=TEXT_COLOR, relief="raised")
            txt.config(state="normal")
            txt.delete("1.0", "end")
            txt.insert("1.0", HELP_CONTENT.get(code, HELP_CONTENT[HELP_DEFAULT_LANG]).strip())
            txt.config(state="disabled")

        for code in ("VI", "EN", "JA"):
            btn = tk.Button(lang_row, text=code, font=("Segoe UI", 7, "bold"),
                             bd=1, padx=4, pady=0, cursor="hand2",
                             command=lambda c=code: _select_help_lang(c))
            btn.pack(side="left", padx=1)
            self._help_lang_btns[code] = btn

        _select_help_lang(getattr(self, "_help_lang", HELP_DEFAULT_LANG))

if __name__ == "__main__":
    _ensure_settings_table()
    _load_api_keys_from_config()  # v-new: đọc danh sách key Gemini/Groq từ configure.ini (tự migrate từ search_data.db nếu configure.ini chưa có) trước khi UI dựng lên
    root = tk.Tk()
    try:
        ttk.Style(root).configure("Treeview", rowheight=20)  # v4.9: room for 16px icons
    except Exception:
        pass
    app = RealtimeSmartSearchApp(root); root.mainloop()