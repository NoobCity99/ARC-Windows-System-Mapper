from dataclasses import dataclass, field
import datetime as _dt
import json
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class RelatedFile:
    path: str
    kind: str = "file"
    source: str = ""
    confidence: str = ""
    marked: str = ""
    score: int = 0
    _search_cache: str = field(default="", init=False, repr=False, compare=False)
    _search_cache_src: Tuple[str, str, str, str] = field(default=("", "", "", ""), init=False, repr=False, compare=False)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "path": self.path,
            "kind": self.kind,
            "source": self.source,
            "confidence": self.confidence,
            "marked": self.marked,
        }

    def search_blob(self) -> str:
        src = (self.path, self.source, self.confidence, self.marked)
        if src != self._search_cache_src:
            self._search_cache_src = src
            self._search_cache = " ".join(part for part in src if part).casefold()
        return self._search_cache


@dataclass
class AppEntry:
    name: str
    version: str = ""
    install_date: str = ""
    size_mb: Optional[int] = None
    publisher: str = ""
    install_location: str = ""
    website: str = ""
    group: str = ""
    related_files: List[RelatedFile] = field(default_factory=list)
    _key_cache: str = field(default="", init=False, repr=False, compare=False)
    _key_cache_src: Tuple[str, str] = field(default=("", ""), init=False, repr=False, compare=False)
    _name_key_cache: str = field(default="", init=False, repr=False, compare=False)
    _name_key_src: str = field(default="", init=False, repr=False, compare=False)
    _install_date_cache: Optional[_dt.date] = field(default=None, init=False, repr=False, compare=False)
    _install_date_src: str = field(default="", init=False, repr=False, compare=False)
    _search_cache: str = field(default="", init=False, repr=False, compare=False)
    _search_cache_src: Tuple[str, str] = field(default=("", ""), init=False, repr=False, compare=False)
    related_scanned: bool = field(default=False, init=False, repr=False, compare=False)
    related_scan_token: str = field(default="", init=False, repr=False, compare=False)

    def key(self) -> str:
        src = (self.name or "", self.version or "")
        if src != self._key_cache_src:
            self._key_cache_src = src
            name, version = src
            self._key_cache = f"{len(name)}|{name}|{len(version)}|{version}"
        return self._key_cache

    def legacy_key(self) -> str:
        return json.dumps([self.name, self.version], ensure_ascii=True)

    def name_key(self) -> str:
        name = self.name or ""
        if name != self._name_key_src:
            self._name_key_src = name
            self._name_key_cache = name.casefold()
        return self._name_key_cache

    def install_date_value(self) -> Optional[_dt.date]:
        raw = self.install_date or ""
        if raw != self._install_date_src:
            self._install_date_src = raw
            self._install_date_cache = None
            if raw:
                try:
                    self._install_date_cache = _dt.date.fromisoformat(raw)
                except ValueError:
                    self._install_date_cache = None
        return self._install_date_cache

    def search_blob(self) -> str:
        src = (self.name or "", self.publisher or "")
        if src != self._search_cache_src:
            self._search_cache_src = src
            self._search_cache = " ".join(part for part in src if part).casefold()
        return self._search_cache

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "group": self.group,
            "version": self.version,
            "install_date": self.install_date,
            "size_mb": self.size_mb,
            "publisher": self.publisher,
            "install_location": self.install_location,
            "website": self.website,
            "related_files": [item.to_dict() for item in self.related_files],
        }


@dataclass(frozen=True)
class AppGroup:
    name: str
