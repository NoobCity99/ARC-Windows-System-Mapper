import os
import re
import time
import difflib
from typing import Dict, Iterable, List, Optional, Tuple

from models import AppEntry, RelatedFile


CONFIG_EXTENSIONS = {
    ".cfg",
    ".config",
    ".dat",
    ".db",
    ".ini",
    ".json",
    ".sqlite",
    ".xml",
    ".yaml",
    ".yml",
}


class RelatedFileScanner:
    def __init__(self, max_dirs: int = 25, max_files: int = 200, max_depth: int = 3) -> None:
        self.max_dirs = max_dirs
        self.max_files = max_files
        self.max_depth = max_depth
        self._config_cache: Dict[str, List[str]] = {}
        self._root_index: List[Tuple[str, str, str]] = []
        self._root_index_key: Optional[Tuple[bool, Tuple[str, ...]]] = None
        self._source_scores = {
            "install_location": 100,
            "appdata": 70,
            "localappdata": 70,
            "localappdata_low": 60,
            "programdata": 60,
            "documents": 50,
            "saved_games": 50,
            "drive": 40,
        }
        self._confidence_scores = {"High": 30, "Medium": 20, "Low": 10, "": 0}

    def reset_cache(self) -> None:
        self._config_cache.clear()
        self._root_index = []
        self._root_index_key = None

    def build_root_index(self, deep_scan: bool, extra_roots: Optional[List[str]] = None) -> List[Tuple[str, str, str]]:
        roots: List[str] = []
        for root in extra_roots or []:
            if not root:
                continue
            try:
                candidate = os.path.abspath(root)
            except OSError:
                continue
            if not os.path.isdir(candidate):
                continue
            roots.append(candidate)
        key = (deep_scan, tuple(roots))
        if self._root_index and self._root_index_key == key:
            return self._root_index
        index: List[Tuple[str, str, str]] = []
        if deep_scan:
            for root, source in _default_roots():
                try:
                    with os.scandir(root) as it:
                        for entry in it:
                            if not entry.is_dir(follow_symlinks=False):
                                continue
                            index.append((entry.path, source, entry.name.casefold()))
                except OSError:
                    continue
        for root in roots:
            try:
                with os.scandir(root) as it:
                    for entry in it:
                        if not entry.is_dir(follow_symlinks=False):
                            continue
                        index.append((entry.path, "drive", entry.name.casefold()))
            except OSError:
                continue
        self._root_index = index
        self._root_index_key = key
        return index

    def scan(
        self,
        apps: List[AppEntry],
        deep_scan: bool = False,
        include_files: bool = True,
        extra_roots: Optional[List[str]] = None,
    ) -> None:
        root_index = self.build_root_index(deep_scan, extra_roots=extra_roots)
        for app in apps:
            app.related_files = self.scan_for_app(
                app,
                deep_scan=deep_scan,
                include_files=include_files,
                root_index=root_index,
                extra_roots=extra_roots,
            )
        self._dedupe_related_files(apps)

    def scan_for_app(
        self,
        app: AppEntry,
        deep_scan: bool = False,
        include_files: bool = True,
        root_index: Optional[List[Tuple[str, str, str]]] = None,
        extra_roots: Optional[List[str]] = None,
    ) -> List[RelatedFile]:
        tokens_name = _tokens(app.name)
        tokens_pub = _tokens(app.publisher)
        cleaned_name = _cleaned(app.name)
        cleaned_pub = _cleaned(app.publisher)
        if not tokens_name and not tokens_pub:
            return []

        candidates: List[Tuple[str, str, str]] = []
        seen_dirs = set()

        install_location = _clean_path(app.install_location)
        if install_location and os.path.isfile(install_location):
            install_location = os.path.dirname(install_location)
        if install_location and os.path.isdir(install_location):
            candidates.append((install_location, "install_location", "High"))
            seen_dirs.add(_normalize_path(install_location))

        if deep_scan or extra_roots:
            if root_index is None:
                root_index = self.build_root_index(deep_scan, extra_roots=extra_roots)
            for path, source, name_cf in root_index:
                confidence = ""
                name_clean = _cleaned(name_cf)
                exact_name = bool(cleaned_name) and name_clean == cleaned_name
                exact_pub = bool(cleaned_pub) and name_clean == cleaned_pub
                if exact_name:
                    confidence = "High"
                elif any(token in name_cf for token in tokens_name):
                    confidence = "Medium"
                elif exact_pub or any(token in name_cf for token in tokens_pub):
                    confidence = "Low"
                else:
                    continue
                norm = _normalize_path(path)
                if norm in seen_dirs:
                    continue
                seen_dirs.add(norm)
                candidates.append((path, source, confidence))
                if len(candidates) >= self.max_dirs:
                    break

        related: List[RelatedFile] = []
        seen_files = set()
        for path, source, confidence in candidates:
            dir_score = self._score_related(path, source, confidence, is_file=False)
            related.append(RelatedFile(path=path, kind="dir", source=source, confidence=confidence, score=dir_score))
            if not include_files:
                continue
            for file_path in self._config_files_for_root(path):
                if len(seen_files) >= self.max_files:
                    break
                try:
                    norm = _normalize_path(file_path)
                except OSError:
                    continue
                if norm in seen_files:
                    continue
                seen_files.add(norm)
                file_confidence = confidence if confidence == "High" else "Medium"
                file_score = self._score_related(file_path, source, file_confidence, is_file=True)
                related.append(
                    RelatedFile(
                        path=file_path,
                        kind="file",
                        source="config_file",
                        confidence=file_confidence,
                        score=file_score,
                    )
                )
            if len(seen_files) >= self.max_files:
                break
        return related

    def _config_files_for_root(self, root: str) -> List[str]:
        try:
            norm_root = _normalize_path(root)
        except OSError:
            return []
        cached = self._config_cache.get(norm_root)
        if cached is not None:
            return cached
        files = _scan_config_files(root, self.max_depth, self.max_files)
        self._config_cache[norm_root] = files
        return files

    def deep_scan_for_app(
        self,
        app: AppEntry,
        roots: List[str],
        limits: "DeepScanLimits",
        ignored: Optional[Iterable[str]] = None,
    ) -> List[RelatedFile]:
        if not roots:
            return []
        tokens_name = _tokens(app.name)
        tokens_pub = _tokens(app.publisher)
        cleaned_name = _cleaned(app.name)
        cleaned_pub = _cleaned(app.publisher)
        if not tokens_name and not tokens_pub and not cleaned_name and not cleaned_pub:
            return []
        ignored_set = {_normalize_path(path) for path in (ignored or []) if path}
        results: List[RelatedFile] = []
        seen: set = set()
        deadline = time.monotonic() + limits.max_seconds if limits.max_seconds else None
        max_dirs = limits.max_dirs
        max_files = limits.max_files
        for root in roots:
            if not root or not os.path.isdir(root):
                continue
            try:
                root_norm = _normalize_path(root)
            except OSError:
                continue
            root_depth = root_norm.count(os.sep)
            for current, dirs, files in os.walk(root, topdown=True):
                if deadline and time.monotonic() >= deadline:
                    return results
                try:
                    current_norm = _normalize_path(current)
                except OSError:
                    continue
                if current_norm in ignored_set:
                    dirs[:] = []
                    continue
                depth = current_norm.count(os.sep) - root_depth
                if depth >= limits.max_depth:
                    dirs[:] = []
                name_cf = _cleaned(os.path.basename(current))
                score = _fuzzy_score(name_cf, cleaned_name, tokens_name, cleaned_pub, tokens_pub)
                if score >= limits.dir_threshold:
                    confidence = _confidence_from_score(score)
                    if current_norm not in seen:
                        seen.add(current_norm)
                        results.append(
                            RelatedFile(
                                path=current,
                                kind="dir",
                                source="deep_scan",
                                confidence=confidence,
                                score=score,
                            )
                        )
                        if len(results) >= max_dirs:
                            return results
                    for filename in files:
                        if len(results) >= max_files:
                            return results
                        ext = os.path.splitext(filename)[1].lower()
                        if ext not in CONFIG_EXTENSIONS and ext != ".exe":
                            continue
                        file_name = _cleaned(filename)
                        fscore = _fuzzy_score(file_name, cleaned_name, tokens_name, cleaned_pub, tokens_pub)
                        if fscore < limits.file_threshold:
                            continue
                        path = os.path.join(current, filename)
                        try:
                            norm = _normalize_path(path)
                        except OSError:
                            continue
                        if norm in ignored_set or norm in seen:
                            continue
                        seen.add(norm)
                        results.append(
                            RelatedFile(
                                path=path,
                                kind="file",
                                source="deep_scan",
                                confidence=_confidence_from_score(fscore),
                                score=fscore,
                            )
                        )
            if deadline and time.monotonic() >= deadline:
                break
        return results

    def _score_related(self, path: str, source: str, confidence: str, is_file: bool) -> int:
        score = self._source_scores.get(source, 0) + self._confidence_scores.get(confidence, 0)
        if is_file:
            score -= 5
        return score

    def _dedupe_related_files(self, apps: List[AppEntry]) -> None:
        best_by_path: Dict[str, Tuple[int, str]] = {}
        for app in apps:
            app_key = app.key()
            for related in app.related_files or []:
                path = related.path or ""
                if not path:
                    continue
                try:
                    norm = _normalize_path(path)
                except OSError:
                    norm = path.casefold()
                score = related.score or self._score_related(path, related.source, related.confidence, related.kind != "dir")
                best = best_by_path.get(norm)
                if best is None or score > best[0]:
                    best_by_path[norm] = (score, app_key)
        if not best_by_path:
            return
        for app in apps:
            app_key = app.key()
            kept: List[RelatedFile] = []
            seen: set = set()
            for related in app.related_files or []:
                path = related.path or ""
                if not path:
                    continue
                try:
                    norm = _normalize_path(path)
                except OSError:
                    norm = path.casefold()
                best = best_by_path.get(norm)
                if not best or best[1] != app_key:
                    continue
                if norm in seen:
                    continue
                seen.add(norm)
                kept.append(related)
            app.related_files = kept


def _clean_path(path: str) -> str:
    text = (path or "").strip()
    if not text:
        return ""
    text = text.strip('"')
    return os.path.expandvars(text)


def _normalize_path(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _tokens(text: str) -> List[str]:
    if not text:
        return []
    cleaned = re.sub(r"[^a-zA-Z0-9]+", " ", text).strip().casefold()
    if not cleaned:
        return []
    tokens = [cleaned]
    tokens.extend([chunk for chunk in cleaned.split() if len(chunk) >= 3])
    return list(dict.fromkeys(tokens))


def _cleaned(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"[^a-zA-Z0-9]+", " ", text).strip().casefold()


def _default_roots() -> Iterable[Tuple[str, str]]:
    roots: List[Tuple[str, str]] = []
    env = os.environ
    appdata = env.get("APPDATA", "")
    local_appdata = env.get("LOCALAPPDATA", "")
    program_data = env.get("PROGRAMDATA", "")
    userprofile = env.get("USERPROFILE", "")

    roots.extend(_maybe_root(appdata, "appdata"))
    roots.extend(_maybe_root(local_appdata, "localappdata"))
    if local_appdata:
        roots.extend(_maybe_root(os.path.join(local_appdata, "Low"), "localappdata_low"))
    roots.extend(_maybe_root(program_data, "programdata"))
    roots.extend(_maybe_root(os.path.join(userprofile, "Documents"), "documents"))
    roots.extend(_maybe_root(os.path.join(userprofile, "Saved Games"), "saved_games"))
    return roots


def _maybe_root(path: str, source: str) -> List[Tuple[str, str]]:
    if path and os.path.isdir(path):
        return [(path, source)]
    return []


def _scan_config_files(root: str, max_depth: int, max_files: int) -> List[str]:
    try:
        root_norm = _normalize_path(root)
    except OSError:
        return []
    root_depth = root_norm.count(os.sep)
    results: List[str] = []
    seen = set()
    for current, dirs, files in os.walk(root, topdown=True):
        try:
            current_norm = _normalize_path(current)
        except OSError:
            continue
        depth = current_norm.count(os.sep) - root_depth
        if depth >= max_depth:
            dirs[:] = []
        for filename in files:
            if len(results) >= max_files:
                return results
            ext = os.path.splitext(filename)[1].lower()
            if ext not in CONFIG_EXTENSIONS:
                continue
            path = os.path.join(current, filename)
            norm = _normalize_path(path)
            if norm in seen:
                continue
            seen.add(norm)
            results.append(path)
    return results


class DeepScanLimits:
    def __init__(
        self,
        max_dirs: int = 2000,
        max_files: int = 5000,
        max_depth: int = 5,
        max_seconds: float = 8.0,
        dir_threshold: int = 68,
        file_threshold: int = 72,
    ) -> None:
        self.max_dirs = max_dirs
        self.max_files = max_files
        self.max_depth = max_depth
        self.max_seconds = max_seconds
        self.dir_threshold = dir_threshold
        self.file_threshold = file_threshold


def _fuzzy_score(
    candidate: str,
    name_full: str,
    name_tokens: List[str],
    pub_full: str,
    pub_tokens: List[str],
) -> int:
    if not candidate:
        return 0
    best = 0.0
    for token in name_tokens:
        if not token:
            continue
        best = max(best, difflib.SequenceMatcher(None, candidate, token).ratio())
    if name_full:
        best = max(best, difflib.SequenceMatcher(None, candidate, name_full).ratio())
    pub_best = 0.0
    for token in pub_tokens:
        if not token:
            continue
        pub_best = max(pub_best, difflib.SequenceMatcher(None, candidate, token).ratio())
    if pub_full:
        pub_best = max(pub_best, difflib.SequenceMatcher(None, candidate, pub_full).ratio())
    if pub_best:
        best = max(best, pub_best * 0.9)
    return int(best * 100)


def _confidence_from_score(score: int) -> str:
    if score >= 85:
        return "High"
    if score >= 72:
        return "Medium"
    if score >= 60:
        return "Low"
    return ""
