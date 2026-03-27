# SPEC: e2e-namespace-2024-compat-007 -- 2024 전용 테스트
"""Tests for OWPML 2024 namespace normalization."""
from __future__ import annotations

from hwpx.opc.xml_utils import normalize_hwpml_namespaces, parse_xml
from hwpx.oxml.header import parse_char_property, CharFontRef


def test_2024_core_normalized():
    xml = b'<root xmlns:hc="http://www.owpml.org/owpml/2024/core"><hc:item/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/core" in result
    assert b"owpml.org" not in result


def test_2024_head_normalized():
    xml = b'<root xmlns:hh="http://www.owpml.org/owpml/2024/head"><hh:charPr id="0"/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/head" in result
    assert b"owpml.org" not in result


def test_2024_paragraph_normalized():
    xml = b'<root xmlns:hp="http://www.owpml.org/owpml/2024/paragraph"><hp:run/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/paragraph" in result


def test_2024_section_normalized():
    xml = b'<root xmlns:hs="http://www.owpml.org/owpml/2024/section"><hs:sec/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/section" in result


def test_2024_master_page_normalized():
    xml = b'<root xmlns:mp="http://www.owpml.org/owpml/2024/master-page"><mp:page/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/master-page" in result


def test_2024_history_normalized():
    xml = b'<root xmlns:hst="http://www.owpml.org/owpml/2024/history"><hst:log/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/history" in result


def test_2024_version_to_app():
    xml = b'<root xmlns:ver="http://www.owpml.org/owpml/2024/version"><ver:info/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/app" in result


def test_2024_charpr_parsing():
    """2024 namespace charPr parses into CharProperty correctly."""
    xml = b'''\
<hh:charPr xmlns:hh="http://www.owpml.org/owpml/2024/head"
    id="0" height="1000" textColor="#000000">
  <hh:fontRef hangul="0" latin="1"/>
  <hh:bold/>
</hh:charPr>'''
    normalized = normalize_hwpml_namespaces(xml)
    elem = parse_xml(b'<?xml version="1.0"?>' + normalized)
    prop = parse_char_property(elem)
    assert prop.id == 0
    assert prop.height == 1000
    assert prop.bold is True
    assert isinstance(prop.font_ref, CharFontRef)
    assert prop.font_ref.latin == 1


def test_mixed_2011_2024():
    """Mixed 2011 and 2024 namespaces — only 2024 is normalized."""
    xml = (
        b'<root xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
        b' xmlns:hp="http://www.owpml.org/owpml/2024/paragraph">'
        b'<hh:item/><hp:run/></root>'
    )
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/head" in result
    assert b"hancom.co.kr/hwpml/2011/paragraph" in result
    assert b"owpml.org" not in result


def test_2016_still_works():
    """2016 normalization still functions after 2024 additions."""
    xml = b'<root xmlns:hh="http://www.hancom.co.kr/hwpml/2016/head"><hh:charPr/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert b"hancom.co.kr/hwpml/2011/head" in result
    assert b"2016" not in result


def test_2011_unchanged():
    """2011 namespace passes through untouched."""
    xml = b'<root xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:charPr/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert result == xml


def test_unknown_namespace_passthrough():
    """Unknown namespaces are not modified."""
    xml = b'<root xmlns:foo="http://example.com/unknown"><foo:bar/></root>'
    result = normalize_hwpml_namespaces(xml)
    assert result == xml
