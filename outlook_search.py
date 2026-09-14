"""
outlook_search.py — Index & search Outlook mail content locally over COM,
without ever exporting mail to .msg and without depending on whether
Windows Search has indexed Outlook on this machine (that's per-PC and
often disabled/broken on work laptops — this works the same on every
machine that just has Outlook installed and signed in).

Design notes (so this fits smartSearch_MFT_BM25_AI's existing architecture):
  - Own sqlite db file (search_outlook.db), separate from search_data.db —
    same reasoning as HISTORY_DB_FILE: searching should never create/grow
    search_data.db, since the ramp-light "is DB built?" check in the main
    app looks at that file's existence/size.
  - FTS5 with the 'porter unicode61' tokenizer (mail is normal prose, not
    filenames/paths, so the trigram tokenizer used for content_index isn't
    the right fit here — porter gives proper BM25 word-stem ranking).
  - Incremental: each mail's LastModificationTime is stored; a re-index
    only touches items that actually changed since the last run, so repeat
    "Update Outlook Index" runs after the first are fast.
  - All Outlook COM calls happen on whatever thread calls into this module.
    COM apartments are thread-local — if you call index_outlook_mail() or
    search_outlook()/open_outlook_item() from a background thread (which
    you should, for indexing at least — see integration notes), that
    thread must be the one that owns the CoInitialize() call. This module
    handles that internally per-call, so you don't need to think about it
    from the caller's side.

Requires: pywin32 (pip install pywin32) and a local Outlook desktop
install signed into the account you want to search (Classic or New
Outlook's underlying COM surface — New Outlook currently does not expose
this COM API; Classic Outlook does).
"""

import os
import sqlite3
import time
import traceback

print("[outlook_search] module loaded — build tag: storeid-fix-v5 (root cause fixed)")

_insert_error_count = 0  # throttles diagnostic error printing — see index_outlook_mail()

try:
    import pythoncom
    import pywintypes
    import win32com.client
    import win32com.server.util
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTLOOK_DB_FILE = os.path.join(_BASE_DIR, 'search_outlook.db')

# olMail = 43 in Outlook's OlObjectClass enum. Other item classes
# (appointments, contacts, tasks...) are skipped — this is a mail search.
OL_MAIL_CLASS = 43

# RPC/COM reject-call reason codes used by RetryRejectedCall below (from
# the Windows SDK's oleidl.h — pywin32 doesn't expose these as constants).
_SERVERCALL_ISHANDLED = 0
_SERVERCALL_REJECTED = 1
_SERVERCALL_RETRYLATER = 2
_PENDINGMSG_WAITDEFPROCESS = 2

# v-fix: automating Outlook from a background thread (this module always
# runs under its own pythoncom.CoInitialize(), on whatever thread the
# caller uses — see module docstring) can hit a well-known COM reentrancy
# stall: Outlook's own STA occasionally rejects an incoming call (busy
# showing a dialog, mid-sync, momentarily reentered, etc.) and asks the
# CALLER "try again later" via RetryRejectedCall. A process with no
# IMessageFilter registered gets COM's default handling for this, which in
# practice can end up waiting indefinitely on some Outlook/Windows builds
# -- no exception, no timeout, low CPU -- exactly matching "processed ==
# total but never finishes, .db size frozen for hours" while Outlook
# itself still looks perfectly responsive to the user. Registering this
# filter makes retries happen on OUR terms (short bounded backoff, then
# give up on that one call) instead of blocking forever.
#
# v-fix2: building the IID and defining/registering this class turned out
# to be pywin32-version-sensitive (pythoncom.IID / pythoncom.IID_Message-
# Filter don't exist on every install) and crashed the WHOLE APP at
# import time twice in a row -- completely disproportionate for what's
# meant to be an optional safety net. Everything below is now wrapped so
# ANY failure here just disables this one feature (prints a warning,
# _MSGFILTER_AVAILABLE stays False) instead of ever blocking the app or
# Outlook indexing from working at all, same as OUTLOOK_AVAILABLE itself.
_MSGFILTER_AVAILABLE = False
if OUTLOOK_AVAILABLE:
    try:
        # Fixed IMessageFilter GUID from the Windows SDK (oleidl.h /
        # objidl.h) -- same value on every Windows install. pywintypes.IID()
        # is the actual constructor that exists across pywin32 versions for
        # turning that string into a real IID object (pythoncom.IID(...)
        # and pythoncom.IID_IMessageFilter do NOT exist in every version).
        _IID_IMessageFilter = pywintypes.IID('{00000016-0000-0000-C000-000000000046}')

        class _OutlookMessageFilter:
            _com_interfaces_ = [_IID_IMessageFilter]
            _public_methods_ = ['HandleInComingCall', 'RetryRejectedCall', 'MessagePending']

            def HandleInComingCall(self, dwCallType, htaskCaller, dwTickCount, lpInterfaceInfo):
                return _SERVERCALL_ISHANDLED

            def RetryRejectedCall(self, htaskCallee, dwTickCount, dwRejectType):
                # dwTickCount is how long WE'VE already been waiting on this
                # one call, in ms. Keep retrying (short pauses) for up to
                # ~2 minutes; past that, give up (-1) so the call fails with
                # an exception instead of hanging forever -- our existing
                # try/except around every property access (see
                # index_outlook_mail) then just skips that one item/folder
                # and moves on, rather than the whole run sitting frozen.
                if dwRejectType == _SERVERCALL_RETRYLATER:
                    if dwTickCount > 120_000:
                        return -1  # cancel the call
                    return 250  # retry again in 250ms
                return -1  # SERVERCALL_REJECTED or anything else -- don't retry

            def MessagePending(self, htaskCallee, dwTickCount, dwPendingType):
                return _PENDINGMSG_WAITDEFPROCESS

        _MSGFILTER_AVAILABLE = True
    except Exception as e:
        print(f"[Outlook] IMessageFilter setup not available on this pywin32/Windows "
              f"install (non-fatal, indexing still runs without it): {e!r}")


def _register_message_filter():
    """Best-effort -- if this fails for any reason, indexing still runs
    exactly as before (just without the extra protection against a stuck
    reentrant call). Must be called AFTER pythoncom.CoInitialize() on the
    same thread that will make the Outlook COM calls."""
    if not _MSGFILTER_AVAILABLE:
        return
    try:
        pythoncom.CoRegisterMessageFilter(
            win32com.server.util.wrap(_OutlookMessageFilter(), _IID_IMessageFilter))
    except Exception as e:
        print(f"[Outlook] Could not register IMessageFilter (non-fatal, continuing without it): {e!r}")


def _connect():
    conn = sqlite3.connect(OUTLOOK_DB_FILE)
    c = conn.cursor()
    c.execute("""CREATE VIRTUAL TABLE IF NOT EXISTS outlook_content
                 USING fts5(subject, body, tokenize='porter unicode61')""")
    c.execute("""CREATE TABLE IF NOT EXISTS outlook_store (
                    entry_id      TEXT PRIMARY KEY,
                    store_id      TEXT,
                    subject       TEXT,
                    sender        TEXT,
                    received      TEXT,
                    folder_path   TEXT,
                    size          INTEGER,
                    last_modified TEXT,
                    rowid_fts     INTEGER
                 )""")
    # v-fix: older search_outlook.db files created before the size column
    # existed won't have it — add it in place so re-running doesn't require
    # deleting the db file.
    try:
        c.execute("ALTER TABLE outlook_store ADD COLUMN size INTEGER")
    except sqlite3.OperationalError:
        pass  # already has the column
    conn.commit()
    return conn, c


def _iter_folders(root_folder):
    """Yield root_folder itself, then every subfolder, recursively
    (Inbox, Sent Items, every custom folder/subfolder, etc.).

    v-fix: rewritten to use indexed `.Item(i)` access instead of
    `for sub in root_folder.Folders:`. The implicit for-loop uses COM's
    enumerator protocol under the hood (IEnumVARIANT.Next()) — this is
    the exact call that was observed hanging indefinitely (no error, no
    timeout, low CPU) when Outlook is busy/syncing. Indexed .Item(i)
    access goes through a different, more reliable COM code path and
    does not exhibit this hang."""
    yield root_folder
    try:
        subfolders = root_folder.Folders
        count = subfolders.Count
    except Exception:
        return
    for i in range(1, count + 1):
        try:
            sub = subfolders.Item(i)
        except Exception:
            continue
        yield from _iter_folders(sub)


def _iter_all_folders(ns):
    """Yield every folder across every open store (mailbox), one at a time
    — lazily, using indexed access (see _iter_folders docstring — same
    Next()-hang concern applies to ns.Stores, which is exactly where the
    reported hang happened)."""
    try:
        stores = ns.Stores
        store_count = stores.Count
    except Exception:
        return
    for i in range(1, store_count + 1):
        try:
            store = stores.Item(i)
            root = store.GetRootFolder()
        except Exception:
            continue
        yield from _iter_folders(root)


class _OutlookDisconnected(Exception):
    """Raised (internally, within this module only) when a COM call fails
    in a way that means Outlook's own process is gone (closed/crashed) --
    as opposed to a one-off quirk with a single item/folder. Distinguishing
    the two matters: a single bad item should just be skipped, but if
    Outlook itself is gone, EVERY remaining call in the current scan is
    about to fail the same way -- better to pause once and wait for
    Outlook to come back than to burn through thousands of items each
    individually failing (and, worse, having each one counted as a normal
    "skipped/unchanged" item, which used to make an interrupted run look
    falsely complete)."""
    pass


# Known HRESULTs for "the other process/server is gone" -- these are what
# actually get raised when Outlook.exe is closed while a COM call to it is
# in flight or about to be made. (0x800706BA RPC_S_SERVER_UNAVAILABLE,
# 0x800706BE RPC_S_CALL_FAILED, 0x80010108 RPC_E_DISCONNECTED,
# 0x800706BF RPC_S_CALL_FAILED_DNE) — signed 32-bit forms, as pywintypes
# reports them.
_RPC_DISCONNECT_HRESULTS = {-2147023170, -2147023169, -2147023166, -2147417848}


def _check_disconnect(e):
    """Call from inside an `except Exception as e:` block that would
    otherwise just skip-and-continue. Raises _OutlookDisconnected (letting
    it propagate past the normal skip handling) if `e` looks like Outlook
    itself went away; does nothing (falls through to the normal
    skip-this-one-item behavior) for any other kind of error."""
    hresult = getattr(e, 'hresult', None)
    if hresult is None and getattr(e, 'args', None):
        hresult = e.args[0] if isinstance(e.args[0], int) else None
    if hresult in _RPC_DISCONNECT_HRESULTS:
        raise _OutlookDisconnected(str(e))


def _wait_for_outlook_restart(stop_flag=None, progress_cb=None, poll_sec=15):
    """Poll (every poll_sec) until Outlook responds to a fresh, cheap COM
    call again. Returns True once it's back, False if stop_flag requested
    an abort while waiting. Deliberately has NO overall timeout -- the
    whole point is to support "close Outlook, come back to it hours later,
    reopen it" and have indexing pick back up on its own."""
    waited = 0
    while True:
        if stop_flag and stop_flag():
            return False
        try:
            test_app = win32com.client.Dispatch("Outlook.Application")
            test_app.GetNamespace("MAPI")
            return True
        except Exception:
            if progress_cb:
                progress_cb(0, 1, f"Outlook đã đóng — đang chờ mở lại... (đã chờ {waited // 60}p{waited % 60:02d}s)")
            time.sleep(poll_sec)
            waited += poll_sec


def index_outlook_mail(progress_cb=None, stop_flag=None):
    """Incrementally index all Outlook mail (every store/mailbox currently
    open in Outlook — shared mailboxes included) into OUTLOOK_DB_FILE.

    progress_cb(done, total, current_subject): optional, called periodically
    so the UI can drive a progress bar the same way indexing_worker() does
    for the main file DB.

    stop_flag: optional zero-arg callable; if it returns True, indexing
    stops at the next safe point (mirrors the cancellable Update DB flow).

    Returns (indexed_count, skipped_count, error) — error is None on success,
    or a short human-readable string on failure (e.g. Outlook not running/
    installed) so the UI can show it directly.
    """
    if not OUTLOOK_AVAILABLE:
        return 0, 0, "pywin32 chưa được cài (pip install pywin32)"

    pythoncom.CoInitialize()
    try:
        _register_message_filter()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
        except Exception as e:
            return 0, 0, f"Không kết nối được Outlook (đã cài/đăng nhập chưa?): {e}"

        conn, c = _connect()

        # entry_id -> last_modified already indexed, so unchanged mail can
        # be skipped entirely (no Body read = no slow COM round-trip). This
        # dict is also updated live as new items are indexed below (not just
        # reloaded from the DB), so that if Outlook disconnects and the scan
        # restarts (see the resume loop below), everything committed so far
        # in THIS run is skipped just as fast as older, already-committed mail.
        c.execute("SELECT entry_id, last_modified FROM outlook_store")
        known = dict(c.fetchall())

        # indexed/skipped ARE cumulative across the whole run (including any
        # restarts below) -- that's the real, correct total for the final
        # summary. processed/total are NOT: they're reset at the top of
        # each restart attempt (see v-fix below) since they represent "how
        # far along THIS walk of the folder tree is", and a restart is a
        # brand new walk from the top.
        indexed = 0
        skipped = 0
        disconnect_count = 0
        # Generous cap so a genuinely persistent, unrelated problem can't
        # spin this forever — but high enough that a user closing Outlook
        # several times over a long unattended run is never an issue.
        MAX_OUTLOOK_DISCONNECTS = 500

        # v-resume: outer retry loop. If Outlook's process itself goes away
        # mid-scan (closed by the user, crashed, etc.), EVERY COM object
        # held from before that point (folder, folder_items, item, even the
        # folder_gen generator's internal state) is permanently invalid —
        # COM/RPC handles don't survive the server process restarting.
        # There's no way to resume the OLD scan at the exact item it was on;
        # the only option is to wait for Outlook to come back, get a FRESH
        # Application/Namespace, and re-walk folders from the top. Thanks to
        # the `known` dict above, that re-walk is fast: every already-
        # indexed item (from this run or an earlier one) is skipped with a
        # single dict lookup, no Body read — so in practice this behaves
        # like "pause and resume", just implemented as "restart, but the
        # restart is cheap", not a true low-level resume.
        while True:
            # v-fix: processed/total used to be declared ONCE, outside this
            # loop, and kept climbing across every restart -- so 2-3
            # disconnects on even a modest mailbox could show "processed"
            # reaching 3x the real folder-tree size (each restart re-walks
            # the same folders and re-counts them, even though the actual
            # per-item work is skipped fast via `known`). Resetting them
            # here means the displayed count always reflects THIS walk only.
            processed = 0
            total = 0          # grows as folders are discovered — see note above
            folder_num = 0
            folder_gen = _iter_all_folders(ns)
            try:
                while True:
                    if stop_flag and stop_flag():
                        break
                    try:
                        folder = next(folder_gen)
                    except StopIteration:
                        break
                    except Exception as e:
                        _check_disconnect(e)  # re-raises _OutlookDisconnected if applicable
                        if progress_cb:
                            progress_cb(processed, total or 1, f"(Outlook COM lỗi, dừng index: {e})")
                        break

                    folder_num += 1
                    try:
                        folder_path = folder.FolderPath
                    except Exception:
                        folder_path = f"(folder #{folder_num})"

                    # v-fix: item.StoreID doesn't exist on this Outlook/win32com
                    # setup (see index_outlook_mail's insert-block comment) — use
                    # the folder's StoreID instead, fetched once per folder.
                    try:
                        folder_store_id = folder.StoreID
                    except Exception:
                        try:
                            folder_store_id = folder.Store.StoreID
                        except Exception:
                            folder_store_id = ""

                    try:
                        folder_items = folder.Items
                        folder_count = folder_items.Count
                    except Exception as e:
                        _check_disconnect(e)
                        # v-fix: this is exactly the call that can hang for a long
                        # time on some folders (large shared mailbox, IMAP folder
                        # not yet synced, public folder, etc). Report BEFORE the
                        # call too, so if it does hang, the last text on screen
                        # names the folder that's stuck instead of going silent.
                        if progress_cb:
                            progress_cb(processed, total or 1, f"(bỏ qua — lỗi mở) {folder_path}")
                        continue

                    total += folder_count
                    if progress_cb:
                        # Fires immediately on entering every folder, even ones
                        # with 0 mail — this is what makes the counter visibly
                        # move within the first second, long before any mail has
                        # actually been indexed.
                        progress_cb(processed, total, f"Đang quét: {folder_path} ({folder_count} mục)")

                    try:
                        for idx in range(1, folder_count + 1):
                            processed += 1
                            if stop_flag and stop_flag():
                                break
                            if progress_cb and idx % 3 == 0:
                                # v-fix: fires BEFORE touching this item's
                                # properties (Class/Body/etc — any of which can be
                                # the call that hangs, e.g. a network fetch in
                                # Outlook Online mode). If indexing freezes, the
                                # last text on screen now names the exact folder +
                                # item position it froze on, instead of a stale
                                # message from hundreds of items earlier.
                                progress_cb(processed, total, f"Đang mở: {folder_path} [{idx}/{folder_count}]")
                            try:
                                item = folder_items.Item(idx)
                            except Exception as e:
                                _check_disconnect(e)
                                skipped += 1
                                if processed % 500 == 0:
                                    conn.commit()
                                continue
                            try:
                                if item.Class != OL_MAIL_CLASS:
                                    if progress_cb and processed % 25 == 0:
                                        progress_cb(processed, total, folder_path)
                                    # v-fix: commit trigger below (after this if/elif
                                    # chain) used to only fire on *newly indexed*
                                    # mail. A folder full of non-mail items (or mail
                                    # already indexed/unchanged) could run for hours
                                    # without a single commit — DB looked frozen even
                                    # though "processed" kept climbing. Fall through
                                    # to the shared periodic-commit check instead of
                                    # continuing early.
                                    if processed % 500 == 0:
                                        conn.commit()
                                    continue
                            except Exception as e:
                                _check_disconnect(e)
                                continue

                            try:
                                entry_id = item.EntryID
                                lm_str = str(item.LastModificationTime)

                                if known.get(entry_id) == lm_str:
                                    skipped += 1
                                    if progress_cb and processed % 25 == 0:
                                        progress_cb(processed, total, "(bỏ qua — không đổi)")
                                    # Same v-fix as above: a run of already-indexed
                                    # "unchanged" mail must not starve out commits.
                                    if processed % 500 == 0:
                                        conn.commit()
                                    continue

                                subject = item.Subject or ""
                                body = item.Body or ""
                                sender = item.SenderName or ""
                                try:
                                    received = item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                                except Exception:
                                    received = ""
                                # v-fix: item.StoreID raised AttributeError on
                                # EVERY mail item in this environment (confirmed
                                # via the diagnostic build's error log) — that one
                                # line was silently killing 100% of inserts this
                                # whole time, which is why the DB stayed at 0 rows
                                # no matter how many items were processed. Folders
                                # expose StoreID reliably (it's a core MAPIFolder
                                # property), so get it from the folder instead —
                                # same value, just a more reliable source. Cache it
                                # per folder so this isn't a repeated COM call for
                                # every single mail in a large folder.
                                store_id = folder_store_id
                                try:
                                    size = int(item.Size)  # bytes — same unit as real file sizes elsewhere in the app
                                except Exception:
                                    size = 0

                                # FTS5 has no UPSERT — drop the old content row (if
                                # any) before inserting the fresh one on re-index.
                                c.execute("SELECT rowid_fts FROM outlook_store WHERE entry_id=?", (entry_id,))
                                row = c.fetchone()
                                if row and row[0] is not None:
                                    c.execute("DELETE FROM outlook_content WHERE rowid=?", (row[0],))

                                c.execute("INSERT INTO outlook_content (subject, body) VALUES (?,?)",
                                          (subject, body))
                                fts_rowid = c.lastrowid

                                c.execute("""INSERT INTO outlook_store
                                             (entry_id, store_id, subject, sender, received,
                                              folder_path, size, last_modified, rowid_fts)
                                             VALUES (?,?,?,?,?,?,?,?,?)
                                             ON CONFLICT(entry_id) DO UPDATE SET
                                               store_id=excluded.store_id, subject=excluded.subject,
                                               sender=excluded.sender, received=excluded.received,
                                               folder_path=excluded.folder_path, size=excluded.size,
                                               last_modified=excluded.last_modified,
                                               rowid_fts=excluded.rowid_fts""",
                                          (entry_id, store_id, subject, sender, received,
                                           folder_path, size, lm_str, fts_rowid))

                                known[entry_id] = lm_str  # v-resume: keep in-memory skip-list current
                                indexed += 1
                                if progress_cb and processed % 10 == 0:
                                    progress_cb(processed, total, subject[:60])
                                # Commit every 50 writes (not 200) so the .db file
                                # visibly grows early and often, not in one late lump.
                                if indexed % 50 == 0:
                                    conn.commit()
                            except Exception as e:
                                _check_disconnect(e)
                                skipped += 1
                                # v-fix: this except block was silently swallowing
                                # every insert/commit error with no trace at all —
                                # that's exactly why "18295 processed, 0 rows in
                                # DB" produced zero clues. Print the real error for
                                # the first several occurrences so the actual cause
                                # (locked DB, disk full, readonly file, bad value,
                                # etc.) becomes visible instead of invisible.
                                global _insert_error_count
                                if _insert_error_count < 8:
                                    _insert_error_count += 1
                                    print(f"[Outlook][INSERT ERROR #{_insert_error_count}] -> {e!r}")
                                    traceback.print_exc()
                                if processed % 500 == 0:
                                    conn.commit()
                                continue
                    except _OutlookDisconnected:
                        raise
                    except Exception:
                        pass

                    # Commit at the end of every folder too, so switching folders
                    # (Inbox -> Sent Items -> ...) is a visible step, not silence.
                    conn.commit()

                # Inner scan finished naturally (StopIteration or stop_flag) —
                # no disconnect happened this attempt, so we're done for real.
                break
            except _OutlookDisconnected as e:
                conn.commit()  # keep everything indexed so far, no matter what happens next
                disconnect_count += 1
                print(f"[Outlook] Disconnected (#{disconnect_count}) -- {e!r} -- waiting for Outlook to reopen...")
                if progress_cb:
                    progress_cb(processed, processed,
                                "Outlook đã đóng giữa chừng — đang chờ mở lại...")
                if disconnect_count > MAX_OUTLOOK_DISCONNECTS or (stop_flag and stop_flag()):
                    print("[Outlook] Giving up waiting (stop requested or too many disconnects).")
                    break
                if not _wait_for_outlook_restart(stop_flag, progress_cb):
                    print("[Outlook] Stop requested while waiting for Outlook.")
                    break
                print("[Outlook] Outlook responded again — resuming scan.")
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    ns = outlook.GetNamespace("MAPI")
                except Exception:
                    # Outlook.exe answered the lightweight test call inside
                    # _wait_for_outlook_restart a moment ago but is refusing
                    # this one (e.g. still finishing its own startup) — loop
                    # back to the top, which waits again before retrying.
                    continue
                # Loop back to the top: re-walk folders from scratch with the
                # fresh `ns`, using the now-larger `known` dict to skip fast.
                continue

        conn.commit()
        conn.close()
        if progress_cb:
            progress_cb(processed, total, "Done")
        return indexed, skipped, None
    finally:
        pythoncom.CoUninitialize()


def search_outlook(query, limit=200):
    """BM25 search over indexed mail (subject + body). Returns a list of
    dicts, best match first. Safe to call from the main thread — this is
    pure sqlite, no COM involved (COM is only needed for indexing and for
    opening an item afterwards)."""
    if not query or not query.strip():
        return []
    if not os.path.exists(OUTLOOK_DB_FILE):
        return []

    conn = sqlite3.connect(OUTLOOK_DB_FILE)
    c = conn.cursor()
    try:
        c.execute("""
            SELECT s.entry_id, s.store_id, s.subject, s.sender, s.received,
                   s.folder_path, s.size,
                   snippet(oc.outlook_content, 1, '[', ']', '...', 12) AS snip,
                   bm25(oc.outlook_content) AS score
            FROM outlook_content oc
            JOIN outlook_store s ON s.rowid_fts = oc.rowid
            WHERE oc.outlook_content MATCH ?
            ORDER BY score
            LIMIT ?
        """, (query, limit))
        rows = c.fetchall()
    except sqlite3.OperationalError:
        # Malformed FTS5 query (stray quotes/operators typed mid-search) —
        # fail soft with no results rather than raising into the UI thread.
        rows = []
    conn.close()

    return [
        {
            "entry_id": r[0], "store_id": r[1], "subject": r[2],
            "sender": r[3], "received": r[4], "folder_path": r[5],
            "size": r[6], "snippet": r[7], "score": r[8],
        }
        for r in rows
    ]


def open_outlook_item(entry_id, store_id):
    """Open the given mail item in Outlook's own reading window (brings
    Outlook to front, focused on that exact email) — same idea as
    double-clicking a file result opens Explorer/the file itself."""
    if not OUTLOOK_AVAILABLE:
        return False, "pywin32 chưa được cài"
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        item = ns.GetItemFromID(entry_id, store_id)
        item.Display()
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()


def get_indexed_mail_count():
    """Quick count for the UI (e.g. show '12,340 mail indexed' next to the
    Update Outlook Index button)."""
    if not os.path.exists(OUTLOOK_DB_FILE):
        return 0
    try:
        conn = sqlite3.connect(OUTLOOK_DB_FILE)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM outlook_store")
        n = c.fetchone()[0]
        conn.close()
        return n
    except Exception:
        return 0