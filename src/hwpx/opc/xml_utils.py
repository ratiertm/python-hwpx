"""XML 파싱/직렬화를 위한 OPC 공통 유틸리티."""

from __future__ import annotations

from io import BytesIO
from typing import Mapping

from lxml import etree

# SPEC: e2e-namespace-2024-compat-001 -- 2024 네임스페이스 매핑 추가
# Mapping of non-2011 HWPML namespace URIs to their 2011 equivalents.
# Documents created with Hancom Office 2016+ use 2016 URIs; documents
# following the OWPML 2024 specification use owpml.org URIs.
# Normalising to 2011 at parse time lets the rest of the codebase use a
# single set of namespace constants without any lookup changes.
_HWPML_NS_TO_2011: tuple[tuple[bytes, bytes], ...] = (
    # 2016 → 2011 (Hancom Office 2016+)
    (b"http://www.hancom.co.kr/hwpml/2016/paragraph", b"http://www.hancom.co.kr/hwpml/2011/paragraph"),
    (b"http://www.hancom.co.kr/hwpml/2016/head", b"http://www.hancom.co.kr/hwpml/2011/head"),
    (b"http://www.hancom.co.kr/hwpml/2016/section", b"http://www.hancom.co.kr/hwpml/2011/section"),
    (b"http://www.hancom.co.kr/hwpml/2016/core", b"http://www.hancom.co.kr/hwpml/2011/core"),
    (b"http://www.hancom.co.kr/hwpml/2016/master-page", b"http://www.hancom.co.kr/hwpml/2011/master-page"),
    (b"http://www.hancom.co.kr/hwpml/2016/history", b"http://www.hancom.co.kr/hwpml/2011/history"),
    (b"http://www.hancom.co.kr/hwpml/2016/app", b"http://www.hancom.co.kr/hwpml/2011/app"),
    # 2024 → 2011 (OWPML 2024 specification, owpml.org domain)
    (b"http://www.owpml.org/owpml/2024/core", b"http://www.hancom.co.kr/hwpml/2011/core"),
    (b"http://www.owpml.org/owpml/2024/head", b"http://www.hancom.co.kr/hwpml/2011/head"),
    (b"http://www.owpml.org/owpml/2024/paragraph", b"http://www.hancom.co.kr/hwpml/2011/paragraph"),
    (b"http://www.owpml.org/owpml/2024/section", b"http://www.hancom.co.kr/hwpml/2011/section"),
    (b"http://www.owpml.org/owpml/2024/master-page", b"http://www.hancom.co.kr/hwpml/2011/master-page"),
    (b"http://www.owpml.org/owpml/2024/history", b"http://www.hancom.co.kr/hwpml/2011/history"),
    (b"http://www.owpml.org/owpml/2024/version", b"http://www.hancom.co.kr/hwpml/2011/app"),
)


# SPEC: e2e-namespace-2024-compat-002 -- normalize 함수 업데이트
def normalize_hwpml_namespaces(data: bytes) -> bytes:
    """Replace 2016 and 2024 HWPML namespace URIs with their 2011 equivalents.

    This is a byte-level transformation applied **before** XML parsing so that
    all downstream code can rely on a single, consistent set of namespace
    constants (the 2011 family).  The replacement is harmless for documents
    that already use 2011 namespaces because the newer byte sequences simply
    won't appear.

    Supported conversions:
    - 2016 (hancom.co.kr/hwpml/2016/*) → 2011
    - 2024 (owpml.org/owpml/2024/*) → 2011
    """
    for old, new in _HWPML_NS_TO_2011:
        if old in data:
            data = data.replace(old, new)
    return data


def parse_xml(data: bytes) -> etree._Element:
    """바이트 XML 문서를 파싱해 루트 요소를 반환한다.

    2016/2024 HWPML 네임스페이스는 파싱 전에 2011 버전으로 자동 정규화된다.
    """

    return etree.fromstring(normalize_hwpml_namespaces(data))


def parse_xml_with_namespaces(data: bytes) -> tuple[etree._Element, Mapping[str, str]]:
    """루트 요소와 네임스페이스 매핑(prefix -> uri)을 함께 반환한다."""

    root = parse_xml(data)
    namespaces = {"" if prefix is None else prefix: uri for prefix, uri in root.nsmap.items() if uri}
    return root, namespaces


def iter_declared_namespaces(data: bytes) -> Mapping[str, str]:
    """XML 선언 순서를 보존한 네임스페이스 매핑을 추출한다.

    2016/2024 HWPML 네임스페이스는 2011 버전으로 정규화된다.
    """

    normalized = normalize_hwpml_namespaces(data)
    namespaces: dict[str, str] = {}
    for _, elem in etree.iterparse(BytesIO(normalized), events=("start-ns",)):
        prefix, uri = elem
        namespaces[prefix or ""] = uri
    return namespaces


def extract_xml_declaration(data: bytes) -> bytes | None:
    """문서 상단의 XML 선언(`<?xml ... ?>`)을 추출한다."""

    stripped = data.lstrip()
    if not stripped.startswith(b"<?xml"):
        return None
    end = stripped.find(b"?>")
    if end == -1:
        return None
    return stripped[: end + 2]


def serialize_xml(element: etree._Element, *, xml_declaration: bool = False) -> bytes:
    """요소를 UTF-8 XML 바이트로 직렬화한다."""

    return etree.tostring(element, encoding="utf-8", xml_declaration=xml_declaration)
