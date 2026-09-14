"""
onenote_search.py — Index & search OneNote page content locally over COM.
Mirrors outlook_search.py's design (own sqlite db file, own try/except so
a missing dependency never breaks the rest of the app), adapted for how
differently OneNote's object model works compared to Outlook's.

Design notes (so this fits smartSearch_MFT_BM25_AI's existing architecture):
  - Own sqlite db file (search_onenote.db), separate from search_data.db
    AND search_outlook.db — same reasoning as HISTORY_DB_FILE/OUTLOOK_DB_FILE:
    searching/indexing OneNote should never touch either of those files.
  - FTS5 with 'porter unicode61' (OneNote page content is prose, same
    tokenizer choice as outlook_search.py's mail body/subject).
  - Incremental: each page's lastModifiedTime (from OneNote's own
    hierarchy XML) is stored; unchanged pages are skipped without
    re-reading their content.
  - Password-protected sections are skipped (their content literally
    can't be read via COM without the user unlocking them in the OneNote
    UI first) — counted and reported back, not silently dropped.

  - IMPORTANT — early binding required: this uses
    win32com.client.gencache.EnsureDispatch("OneNote.Application"), NOT
    plain win32com.client.Dispatch(). GetHierarchy() and GetPageContent()
    both return their real result through an [out] BSTR parameter in
    OneNote's COM interface. Late-bound Dispatch() has a long-documented
    history of not marshaling that particular kind of [out] string param
    back correctly (silently returns None/empty for these two calls
    specifically, even though other OneNote COM calls work fine with it).
    gencache.EnsureDispatch builds a real wrapper from OneNote's type
    library ahead of time, which handles this correctly. This is the
    single most common gotcha in OneNote COM automation, so it's called
    out explicitly here rather than left to be rediscovered the hard way.

  - Unlike Outlook (which has to be walked folder by folder, and where
    that walk itself turned out to be a source of hangs — see
    outlook_search.py's history), OneNote returns its ENTIRE
    notebook/section-group/section/page tree in a SINGLE GetHierarchy()
    call. There is no folder-by-folder discovery step here, and — because
    the full page count is known immediately from that one call — no
    "progress bar total keeps growing" ambiguity either.

Requires: pywin32 (pip install pywin32) and OneNote installed locally —
specifically the classic desktop version that ships with Office (same
family as classic Outlook). OneNote for Windows 10 (the Microsoft Store
app) has NO COM automation surface at all and cannot be used this way.
"""

import html
import os
import re
import sqlite3
import time
import traceback
import xml.etree.ElementTree as ET

try:
    import pythoncom
    import win32com
    import win32com.client

    # v10.20 FIX: win32com.client.gencache (required for OneNote -- see the
    # module docstring's "IMPORTANT — early binding required" note) caches
    # its auto-generated COM wrapper code under a folder it computes from
    # win32com's OWN package location on disk. That computation assumes a
    # normal pip-installed package sitting in site-packages -- inside a
    # PyInstaller --onedir/--onefile build there IS no such folder (or it's
    # read-only/temporary), so gencache either fails to find/write its
    # cache or points at a stale/empty one. The result is exactly the
    # bizarre, unreadable error this fix addresses: something like
    # "<built-in method GetFuncDesc of PyITypeInfo object at 0x...>"
    # instead of a real message, because gencache half-generated a broken
    # wrapper from OneNote's type library and EnsureDispatch blew up
    # partway through.
    #
    # The standard fix (well-documented in the pywin32/PyInstaller
    # community): force gencache to use a real, always-writable folder
    # under the current user's TEMP directory instead of the
    # package-relative path it computes by default. This makes
    # EnsureDispatch freshly (re)generate its wrapper from whatever
    # OneNote/Office version is ACTUALLY registered on THIS machine the
    # first time it runs here -- which also happens to be exactly right,
    # since the GitHub Actions build machine has no Office/OneNote
    # installed at all and could never have baked in a working cache
    # during the build anyway. Must run before the first
    # gencache.EnsureDispatch call anywhere in this module.
    import tempfile
    win32com.__gen_path__ = os.path.join(tempfile.gettempdir(), "gen_py")
    os.makedirs(win32com.__gen_path__, exist_ok=True)

    ONENOTE_AVAILABLE = True
except ImportError:
    ONENOTE_AVAILABLE = False

# Same "the other process is gone" HRESULTs as outlook_search.py — see that
# module's comment for the exact meanings. Duplicated rather than imported
# so this module keeps working standalone even if outlook_search.py is
# missing/broken (matches this file's own "never breaks the rest of the
# app" design note above).
_RPC_DISCONNECT_HRESULTS = {-2147023170, -2147023169, -2147023166, -2147417848}


def _is_disconnect(e):
    hresult = getattr(e, 'hresult', None)
    if hresult is None and getattr(e, 'args', None):
        hresult = e.args[0] if isinstance(e.args[0], int) else None
    return hresult in _RPC_DISCONNECT_HRESULTS


def _wait_for_onenote_restart(stop_flag=None, progress_cb=None, poll_sec=15):
    """Poll until OneNote responds to a fresh, cheap COM call again (or
    stop_flag says abort). No overall timeout — supports "close OneNote,
    reopen it whenever" and having indexing pick back up on its own."""
    waited = 0
    while True:
        if stop_flag and stop_flag():
            return False
        try:
            win32com.client.gencache.EnsureDispatch("OneNote.Application")
            return True
        except Exception:
            if progress_cb:
                progress_cb(0, 1, f"OneNote đã đóng — đang chờ mở lại... (đã chờ {waited // 60}p{waited % 60:02d}s)")
            time.sleep(poll_sec)
            waited += poll_sec


print("[onenote_search] module loaded — build tag: v1 (single-hierarchy-call design)")

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ONENOTE_DB_FILE = os.path.join(_BASE_DIR, 'search_onenote.db')

# OneNote's XML namespace. Microsoft's 2013 schema is what's documented
# and used almost universally today — it reads content written by any
# older OneNote version fine too (the schema is additive/compatible), so
# there's no need to also try 2010's namespace URI.
_NS = "{http://schemas.microsoft.com/office/onenote/2013/onenote}"

_insert_error_count = 0  # throttles diagnostic error printing (module-level, like outlook_search.py's)


def _connect():
    conn = sqlite3.connect(ONENOTE_DB_FILE)
    c = conn.cursor()
    c.execute("""CREATE VIRTUAL TABLE IF NOT EXISTS onenote_content
                 USING fts5(title, body, tokenize='porter unicode61')""")
    c.execute("""CREATE TABLE IF NOT EXISTS onenote_store (
                    page_id       TEXT PRIMARY KEY,
                    title         TEXT,
                    notebook      TEXT,
                    section_path  TEXT,
                    created       TEXT,
                    last_modified TEXT,
                    rowid_fts     INTEGER
                 )""")
    conn.commit()
    return conn, c


def _walk_hierarchy(elem, notebook_name, path_parts):
    """Recursively yield (page_element, notebook_name, section_path_list)
    for every <one:Page> under this element of the hierarchy XML tree.

    - Notebook: just remembers its name and recurses.
    - SectionGroup: extends the display path with its name and recurses.
      OneNote also represents its own internal Recycle Bin as a
      SectionGroup (isRecycleBin="true") — skipped so deleted pages don't
      show up in search results.
    - Section: extends the display path with its name and yields each
      child Page — UNLESS the section is password-locked (locked="true"),
      in which case it's skipped outright (GetPageContent would just
      throw on every page in it anyway)."""
    tag = elem.tag
    if tag == f"{_NS}Notebooks":
        for child in elem:
            yield from _walk_hierarchy(child, notebook_name, path_parts)
    elif tag == f"{_NS}Notebook":
        nb_name = elem.get("name", "") or notebook_name
        for child in elem:
            yield from _walk_hierarchy(child, nb_name, [nb_name] if nb_name else [])
    elif tag == f"{_NS}SectionGroup":
        if elem.get("isRecycleBin") == "true":
            return
        grp_name = elem.get("name", "")
        new_path = path_parts + ([grp_name] if grp_name else [])
        for child in elem:
            yield from _walk_hierarchy(child, notebook_name, new_path)
    elif tag == f"{_NS}Section":
        if elem.get("locked") == "true":
            return
        sec_name = elem.get("name", "")
        new_path = path_parts + ([sec_name] if sec_name else [])
        for child in elem:
            if child.tag == f"{_NS}Page":
                yield (child, notebook_name, new_path)
    # any other/unexpected tag is silently ignored — forward-compatible
    # with schema additions we don't know about.


_PSEUDO_TAG_RE = re.compile(r'<[^>]+>')


def _clean_run_text(raw):
    """A <one:T>'s CDATA content is sometimes NOT plain text -- OneNote
    wraps mixed-language/mixed-style runs in pseudo-HTML, e.g.:
        <span lang=ja>もし、</span><span lang=en-US>Abaqus</span><span lang=ja>だけで、</span>
    This lives INSIDE a CDATA section, so the XML parser does NOT treat
    it as real child elements -- it hands the whole thing back as one
    literal text string, <span> tags and all. Left as-is, these fake
    tags land in the middle of a sentence (often mid-word for Japanese,
    which has no spaces), corrupting the contiguous text and silently
    breaking any search for a phrase that spans a tag boundary -- e.g.
    "Abaqus CAEのライセンスサーバーは、FlexLM" was actually indexed as
    something like "Abaqus <span lang=ja>CAEの...", so searching the real
    sentence never matched.

    Fix: strip the pseudo-tags and HTML-unescape any entities (&quot; etc,
    also stored literally since CDATA disables entity expansion) so the
    result reads exactly like the sentence a human sees in OneNote.
    NOTE: this is a blunt regex, not a real HTML parser -- a rare note
    that genuinely types a literal "<...>" could get stripped too, but
    that's far rarer than the systematic breakage this fixes."""
    if not raw:
        return ""
    return html.unescape(_PSEUDO_TAG_RE.sub('', raw))


def _extract_text(content_xml):
    """Pull every bit of visible text out of a page's content XML, no
    matter how it's nested (plain paragraphs, tables, tagged items,
    OCR'd text from images if 'Make text in images searchable' is on,
    etc).

    v-fix: earlier versions joined EVERY text fragment from
    root.itertext() with an inserted " " between each one. OneNote very
    often splits ONE continuous sentence into several <one:T> runs
    within the SAME line purely for formatting reasons (spell-check
    flags, hyperlinks, autocorrect, bold/italic spans) -- with NO space
    between them in the real text. This is especially damaging for
    Japanese/Chinese, which have no spaces between words at all: a
    keyword that happened to straddle one of those internal run splits
    got a fabricated space shoved into the middle of it, silently
    breaking search for that keyword -- even though the exact same
    text, extracted intact elsewhere (e.g. an Outlook mail body, which
    has no such run-splitting), searched fine. This is why the same
    keyword could be found via the Msg/Outlook tab but not OneNote, even
    when both genuinely contained it.

    Fix (1): concatenate text runs WITHIN the same line/bullet (the direct
    <one:T> children of a given <one:OE>) with NO separator -- this
    preserves the exact original character sequence -- and only insert
    a separator BETWEEN different lines/bullets (different <one:OE>
    elements), which is a real, correct line break.

    Fix (2): each run's raw text is passed through _clean_run_text()
    first, since OneNote often stores mixed-language/mixed-style runs as
    pseudo-HTML <span>...</span> text (see _clean_run_text's docstring)
    that would otherwise land as literal garbage in the middle of the
    sentence, corrupting it just as badly as fix (1)'s fabricated spaces.

    Fix (3): text that lives OUTSIDE the <one:OE>/<one:T> structure --
    most notably OCR'd text from a pasted screenshot ("Make text in
    images searchable"), but also things like attachment captions --
    is NOT part of any <one:OE>'s direct <one:T> children, so fix (1)'s
    switch to only reading <one:OE>/<one:T> silently dropped it entirely
    (this is why re-indexed search_onenote.db shrank noticeably in size
    after that fix). This fallback branch walks every OTHER element too
    and picks up any leftover .text it finds, so nothing that used to be
    searchable (pre-fix-1) goes missing."""
    try:
        root = ET.fromstring(content_xml)
    except Exception:
        return ""

    lines = []

    def _has_descendant_oe(elem):
        for child in elem:
            if child.tag == f"{_NS}OE" or _has_descendant_oe(child):
                return True
        return False

    def _walk(elem):
        if elem.tag == f"{_NS}OE":
            # This element's OWN text runs only (direct <one:T> children).
            own_parts = [_clean_run_text("".join(t.itertext())) for t in elem.findall(f"{_NS}T")]
            own_line = "".join(own_parts).strip()
            if own_line:
                lines.append(own_line)
            # Still recurse below -- an OE can directly contain other
            # things (nested <one:OEChildren>, an embedded <one:Image>,
            # etc.) that need their own handling, not just T children.
            for child in elem:
                _walk(child)
            return

        if elem.tag == f"{_NS}T":
            return  # always consumed via its parent <one:OE> above

        if not _has_descendant_oe(elem):
            # A leaf-ish region with no <one:OE> lines anywhere below it
            # -- e.g. OCR'd text from a pasted screenshot/printout,
            # attachment captions, table metadata. Grab EVERYTHING under
            # it in one go via itertext() (which, unlike a shallow
            # .text read, also picks up any nested sub-elements' text
            # AND their .tail text -- important because large "Insert
            # Printout" OCR content is often structured with internal
            # sub-tags, and a shallow read was silently truncating it
            # right at the first such sub-tag, which is why big OCR/
            # printout pages lost tens of thousands of characters).
            txt = "".join(elem.itertext())
            cleaned = _clean_run_text(txt).strip()
            if cleaned:
                lines.append(cleaned)
            return  # already consumed everything below; don't recurse

        # Otherwise this is a structural container (Outline, OEChildren,
        # Table, Row, Cell, ...) with <one:OE> lines somewhere beneath it
        # -- just recurse and let those OEs handle themselves individually.
        for child in elem:
            _walk(child)

    _walk(root)
    return " ".join(lines)


def index_onenote(progress_cb=None, stop_flag=None):
    """Incrementally index every OneNote page currently visible to
    OneNote (every open notebook) into ONENOTE_DB_FILE.

    progress_cb(done, total, current_label): same shape as
    outlook_search.index_outlook_mail()'s callback.
    stop_flag: optional zero-arg callable; True = stop at next safe point.

    Returns (indexed_count, skipped_count, locked_count, error) — error is
    None on success, or a short human-readable string on failure.
    """
    global _insert_error_count

    if not ONENOTE_AVAILABLE:
        return 0, 0, 0, "pywin32 chưa được cài (pip install pywin32)"

    pythoncom.CoInitialize()
    try:
        try:
            # v-note: must be gencache.EnsureDispatch — see module docstring.
            onenote = win32com.client.gencache.EnsureDispatch("OneNote.Application")
        except Exception as e:
            return 0, 0, 0, f"Không kết nối được OneNote (đã cài chưa?): {e}"

        try:
            hierarchy_xml = onenote.GetHierarchy("", 4)  # hsPages=4 -> full tree, every notebook, down to pages
        except Exception as e:
            return 0, 0, 0, f"Không lấy được danh sách OneNote: {e}"

        try:
            root = ET.fromstring(hierarchy_xml)
        except Exception as e:
            return 0, 0, 0, f"Không đọc được cấu trúc OneNote (XML lỗi): {e}"

        pages = list(_walk_hierarchy(root, "", []))
        total = len(pages)

        conn, c = _connect()
        c.execute("SELECT page_id, last_modified FROM onenote_store")
        known = dict(c.fetchall())

        processed = 0
        indexed = 0
        skipped = 0
        locked = 0

        for page_elem, notebook_name, section_path in pages:
            if stop_flag and stop_flag():
                break
            processed += 1

            page_id = page_elem.get("ID", "")
            title = page_elem.get("name", "") or "(untitled)"
            created = page_elem.get("dateTime", "")
            lm = page_elem.get("lastModifiedTime", "") or created
            section_str = " > ".join([notebook_name] + section_path) if notebook_name else " > ".join(section_path)

            if progress_cb and processed % 5 == 0:
                progress_cb(processed, total, f"{section_str} > {title}")

            if not page_id:
                skipped += 1
                continue

            if lm and known.get(page_id) == lm:
                skipped += 1
                if processed % 500 == 0:
                    conn.commit()
                continue

            try:
                content_xml = onenote.GetPageContent(page_id)
                body = _extract_text(content_xml)
            except Exception as e:
                if _is_disconnect(e):
                    # OneNote's own process went away (closed by the user,
                    # crashed, etc). Unlike a locked/mid-sync page, this
                    # isn't specific to THIS page — every subsequent
                    # GetPageContent call would fail the same way. Pause,
                    # wait for OneNote to come back, get a fresh COM handle,
                    # and retry this SAME page (page_id is just a string,
                    # not a live COM object, so nothing about `pages` or our
                    # position in it needs to be redone).
                    conn.commit()
                    if progress_cb:
                        progress_cb(processed, total, "OneNote đã đóng giữa chừng — đang chờ mở lại...")
                    if not _wait_for_onenote_restart(stop_flag, progress_cb):
                        conn.commit()
                        conn.close()
                        if progress_cb:
                            progress_cb(processed, total, "Done (dừng giữa chừng)")
                        return indexed, skipped, locked, None
                    try:
                        onenote = win32com.client.gencache.EnsureDispatch("OneNote.Application")
                    except Exception:
                        pass  # still not ready -- fall through, retry below anyway
                    try:
                        content_xml = onenote.GetPageContent(page_id)
                        body = _extract_text(content_xml)
                    except Exception as e2:
                        # Retried once right after reconnecting and it still
                        # failed -- treat as a genuine per-page issue this
                        # time rather than looping forever on one page.
                        locked += 1
                        if _insert_error_count < 8:
                            _insert_error_count += 1
                            print(f"[OneNote][GetPageContent ERROR #{_insert_error_count}] {title!r} -> {e2!r}")
                        if processed % 500 == 0:
                            conn.commit()
                        continue
                else:
                    # Most commonly a locked/password-protected section that
                    # slipped through (e.g. locked mid-run), or a page mid-sync.
                    locked += 1
                    if _insert_error_count < 8:
                        _insert_error_count += 1
                        print(f"[OneNote][GetPageContent ERROR #{_insert_error_count}] {title!r} -> {e!r}")
                    if processed % 500 == 0:
                        conn.commit()
                    continue

            try:
                c.execute("SELECT rowid_fts FROM onenote_store WHERE page_id=?", (page_id,))
                row = c.fetchone()
                if row and row[0] is not None:
                    c.execute("DELETE FROM onenote_content WHERE rowid=?", (row[0],))

                c.execute("INSERT INTO onenote_content (title, body) VALUES (?,?)", (title, body))
                fts_rowid = c.lastrowid

                c.execute("""INSERT INTO onenote_store
                             (page_id, title, notebook, section_path, created, last_modified, rowid_fts)
                             VALUES (?,?,?,?,?,?,?)
                             ON CONFLICT(page_id) DO UPDATE SET
                               title=excluded.title, notebook=excluded.notebook,
                               section_path=excluded.section_path, created=excluded.created,
                               last_modified=excluded.last_modified, rowid_fts=excluded.rowid_fts""",
                          (page_id, title, notebook_name, section_str, created, lm, fts_rowid))

                indexed += 1
                if indexed % 50 == 0:
                    conn.commit()
            except Exception as e:
                skipped += 1
                if _insert_error_count < 8:
                    _insert_error_count += 1
                    print(f"[OneNote][INSERT ERROR #{_insert_error_count}] -> {e!r}")
                    traceback.print_exc()

            # v-note: same lesson learned from outlook_search.py — commit
            # on a processed-count cadence too (not just on successful
            # inserts), so a long run of skipped/locked pages still
            # flushes to disk periodically instead of going silent.
            if processed % 500 == 0:
                conn.commit()

        conn.commit()
        conn.close()
        if progress_cb:
            progress_cb(processed, total, "Done")
        return indexed, skipped, locked, None
    finally:
        pythoncom.CoUninitialize()


def search_onenote(query, limit=200):
    """BM25 search over indexed OneNote pages (title + body). Returns a
    list of dicts, best match first. Safe to call from the main thread —
    pure sqlite, no COM involved."""
    if not query or not query.strip():
        return []
    if not os.path.exists(ONENOTE_DB_FILE):
        return []

    conn = sqlite3.connect(ONENOTE_DB_FILE)
    c = conn.cursor()
    try:
        c.execute("""
            SELECT s.page_id, s.title, s.notebook, s.section_path,
                   s.created, s.last_modified,
                   snippet(oc.onenote_content, 1, '[', ']', '...', 12) AS snip,
                   bm25(oc.onenote_content) AS score
            FROM onenote_content oc
            JOIN onenote_store s ON s.rowid_fts = oc.rowid
            WHERE oc.onenote_content MATCH ?
            ORDER BY score
            LIMIT ?
        """, (query, limit))
        rows = c.fetchall()
    except sqlite3.OperationalError:
        # Malformed FTS5 query (stray quotes/operators typed mid-search) —
        # fail soft with no results, same as outlook_search.search_outlook.
        rows = []
    conn.close()

    return [
        {
            "page_id": r[0], "title": r[1], "notebook": r[2], "section_path": r[3],
            "created": r[4], "last_modified": r[5], "snippet": r[6], "score": r[7],
        }
        for r in rows
    ]


def open_onenote_page(page_id):
    """Bring OneNote to front, navigated to this exact page — same idea
    as outlook_search.open_outlook_item()."""
    if not ONENOTE_AVAILABLE:
        return False, "pywin32 chưa được cài"
    pythoncom.CoInitialize()
    try:
        onenote = win32com.client.gencache.EnsureDispatch("OneNote.Application")
        onenote.NavigateTo(page_id)
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()


def get_indexed_page_count():
    """Quick count for the UI (e.g. show 'N page indexed' next to the
    Update OneNote button)."""
    if not os.path.exists(ONENOTE_DB_FILE):
        return 0
    try:
        conn = sqlite3.connect(ONENOTE_DB_FILE)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM onenote_store")
        n = c.fetchone()[0]
        conn.close()
        return n
    except Exception:
        return 0
