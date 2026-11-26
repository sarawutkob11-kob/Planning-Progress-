"""SharePoint link -> local folder that the OneDrive client syncs (25-Nov-2026, approved).

No Graph / MSAL / login: OneDrive writes one registry key per synced library under
HKCU\\Software\\SyncEngines\\Providers\\OneDrive\\<id> (UrlNamespace, MountPoint, LibraryType). The library
whose UrlNamespace is the longest prefix of the link wins; the local folder is its MountPoint plus the rest
of the link's path. The app copies files there and OneDrive uploads them to SharePoint.

Link forms:
  (a) direct path    https://t.sharepoint.com/sites/Team/Shared%20Documents/UAT/Results
  (b) address bar    .../Forms/AllItems.aspx?id=%2Fsites%2FTeam%2F...  (id, else RootFolder)
                     .../_layouts/15/onedrive.aspx?id=%2Fpersonal%2F...
  (c) sharing link   /:f:/r/<path> -> <path>; other /:x:/... links are opaque -> error (open the folder in the
                     browser and copy the address bar instead)

Excel file links (26-Nov-2026): resolve_file_link() takes the same forms (sharing link /:x:/r/<path>) but must
end in .xlsx; Office Online pages (Doc.aspx / WopiFrame.aspx ?sourcedoc=) have no path and are refused. Errors
are ExcelLinkError (UI text proj.xl_<code>); the folder functions above are unchanged. canonical_file_url()
gives the file's plain URL (no query string) kept in project event details.
"""

from __future__ import annotations

import dataclasses
import logging
import re
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlsplit

from .projects import AttachmentDirError, ExcelLinkError

log = logging.getLogger(__name__)

PROVIDERS_KEY = r"Software\SyncEngines\Providers\OneDrive"
_SHARING = re.compile(r"^/:[a-z]+:/(?P<kind>[a-z]+)(?P<rest>/.*)?$", re.IGNORECASE)
_PATH_PARAMS = ("id", "rootfolder")
_OFFICE_PAGES = ("doc.aspx", "wopiframe.aspx")


@dataclasses.dataclass(frozen=True)
class SyncedLibrary:
    url_namespace: str
    mount_point: str
    library_type: str = ""


def synced_libraries() -> list[SyncedLibrary]:
    """Libraries the OneDrive client syncs on this PC (registry). [] when none / not Windows."""
    try:
        import winreg
    except ImportError:
        return []
    out: list[SyncedLibrary] = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, PROVIDERS_KEY)
    except OSError:
        return []
    with root:
        i = 0
        while True:
            try:
                name = winreg.EnumKey(root, i)
            except OSError:
                break
            i += 1
            lib = _read_provider(winreg, root, name)
            if lib is not None:
                out.append(lib)
    return out


def _read_provider(winreg, root, name: str) -> SyncedLibrary | None:
    def value(key, field: str) -> str:
        try:
            return str(winreg.QueryValueEx(key, field)[0] or "")
        except OSError:
            return ""
    try:
        with winreg.OpenKey(root, name) as key:
            url, mount = value(key, "UrlNamespace"), value(key, "MountPoint")
            kind = value(key, "LibraryType")
    except OSError:
        log.warning("OneDrive provider key %s could not be read", name, exc_info=True)
        return None
    return SyncedLibrary(url, mount, kind) if url and mount else None


def clean_link(link: str) -> str:
    return (link or "").strip().strip('"<>').strip()


def _segments(path: str) -> list[str]:
    return [s for s in unquote(path).replace("\\", "/").split("/") if s]


def _query_path(query: str) -> str | None:
    params = {k.casefold(): v for k, v in parse_qs(query).items()}
    for name in _PATH_PARAMS:
        if params.get(name) and params[name][0].strip():
            return params[name][0]
    return None


def link_location(link: str) -> tuple[str, list[str]]:
    """-> (host, path segments) the link points to. Raises AttachmentDirError (link_invalid / link_sharing)."""
    text = clean_link(link)
    parts = urlsplit(text)
    if parts.scheme.lower() not in ("http", "https") or not parts.hostname:
        raise AttachmentDirError("link_invalid", text, f"Not a SharePoint link: {text}")
    host = parts.hostname.casefold()
    in_query = _query_path(parts.query)
    if in_query is not None:                              # (b) address bar: ?id= / ?RootFolder=
        return host, _segments(in_query)
    path = parts.path
    sharing = _SHARING.match(path)
    if sharing:                                           # (c) sharing link
        if sharing.group("kind").casefold() != "r" or not sharing.group("rest"):
            raise AttachmentDirError("link_sharing", text, f"Sharing links can't be used: {text}")
        path = sharing.group("rest")
    segs = _segments(path)
    folded = [s.casefold() for s in segs]
    if "forms" in folded and segs[-1].casefold().endswith(".aspx"):  # library view without ?id=
        segs = segs[:folded.index("forms")]
    return host, segs


def _namespace(lib: SyncedLibrary) -> tuple[str, list[str]]:
    parts = urlsplit(lib.url_namespace.strip())
    return (parts.hostname or "").casefold(), [s.casefold() for s in _segments(parts.path)]


def resolve_link(link: str, libraries: list[SyncedLibrary] | None = None) -> Path:
    """Local folder for a SharePoint link (the folder is NOT checked here). Raises AttachmentDirError."""
    host, segs = link_location(link)
    if any(s.strip() in (".", "..") for s in segs):
        text = clean_link(link)
        raise AttachmentDirError("link_invalid", text, f"Not a SharePoint folder link: {text}")
    folded = [s.casefold() for s in segs]
    best: tuple[int, SyncedLibrary] | None = None
    for lib in synced_libraries() if libraries is None else libraries:
        ns_host, ns = _namespace(lib)
        if ns_host == host and ns and folded[:len(ns)] == ns and (best is None or len(ns) > best[0]):
            best = (len(ns), lib)
    if best is None:
        text = clean_link(link)
        raise AttachmentDirError("link_not_synced", text, f"No library synced by OneDrive on this PC matches {text}")
    n, lib = best
    return Path(lib.mount_point).joinpath(*segs[n:])


def _file_error(code: str, text: str) -> ExcelLinkError:
    messages = {
        "link_invalid": f"Not a SharePoint file link: {text}",
        "link_sharing": f"This kind of link can't be used: {text}",
        "link_not_xlsx": f"The link must point to an Excel .xlsx file: {text}",
        "link_not_synced": f"No library synced by OneDrive on this PC matches {text}",
    }
    return ExcelLinkError(code, text, messages.get(code, f"SharePoint link can't be used: {text}"))


def _is_office_page(text: str) -> bool:
    segs = _segments(urlsplit(text).path)
    return bool(segs) and segs[-1].strip().casefold() in _OFFICE_PAGES


def resolve_file_link(link: str, libraries: list[SyncedLibrary] | None = None) -> Path:
    """Local path of a SharePoint .xlsx link (existence NOT checked, no disk access).

    Raises ExcelLinkError (link_invalid / link_sharing / link_not_xlsx / link_not_synced), path = cleaned link.
    """
    text = clean_link(link)
    try:
        _host, segs = link_location(text)
    except AttachmentDirError as exc:
        raise _file_error(exc.code, text) from exc
    if _is_office_page(text):
        raise _file_error("link_sharing", text)
    if any(s.strip() in (".", "..") for s in segs):
        raise _file_error("link_invalid", text)
    if not segs or not segs[-1].strip().casefold().endswith(".xlsx"):
        raise _file_error("link_not_xlsx", text)
    try:
        return resolve_link(text, libraries)
    except AttachmentDirError as exc:
        raise _file_error(exc.code, text) from exc


def canonical_file_url(link: str) -> str:
    """Plain URL of the file a link points to, no query string: https://<host>/<encoded path segments>.

    Address-bar links give the path from ?id=, /:x:/r/ links the path after /r/. For event details only.
    Raises AttachmentDirError like link_location (call it after resolve_file_link succeeded).
    """
    host, segs = link_location(link)
    return f"https://{host}/" + "/".join(quote(s, safe="") for s in segs)
