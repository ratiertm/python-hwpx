# SPEC: e2e-char-property-dataclass-003 -- 기존 HWPX 파일 파싱 검증
# SPEC: e2e-char-property-dataclass-005 -- 파싱 에러 처리
# SPEC: e2e-char-property-dataclass-008 -- 기능 동작 검증
# SPEC: e2e-char-property-dataclass-011 -- 라운드트립 테스트
"""Tests for CharProperty dataclass expansion (Phase A)."""
from __future__ import annotations

import xml.etree.ElementTree as stdlib_ET

from lxml import etree

from hwpx.oxml.header import (
    CharFontRef,
    CharOffset,
    CharOutline,
    CharProperty,
    CharRatio,
    CharRelSize,
    CharShadow,
    CharSpacing,
    CharStrikeout,
    CharUnderline,
    parse_char_property,
    serialize_char_property_into,
)

_HH = "{http://www.hancom.co.kr/hwpml/2011/head}"


def _make_charpr_xml(lxml: bool = True) -> etree._Element | stdlib_ET.Element:
    """Create a sample charPr XML element with representative children."""
    xml_str = f"""\
<hh:charPr xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
    id="0" height="1000" textColor="#000000" shadeColor="#FFFFFF"
    useFontSpace="0" useKerning="1" symMark="NONE">
  <hh:fontRef hangul="0" latin="1" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
  <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
  <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
  <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
  <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
  <hh:italic/>
  <hh:bold/>
  <hh:underline type="BOTTOM" shape="SOLID" color="#0000FF"/>
  <hh:strikeout shape="NONE" color="#000000"/>
  <hh:outline type="NONE"/>
  <hh:shadow type="DROP" color="#C0C0C0" offsetX="10" offsetY="10"/>
</hh:charPr>"""
    if lxml:
        return etree.fromstring(xml_str.encode())
    else:
        return stdlib_ET.fromstring(xml_str)


# --- Test parse into typed fields ---


def test_parse_basic_fields():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert prop.id == 0
    assert prop.height == 1000
    assert prop.text_color == "#000000"
    assert prop.shade_color == "#FFFFFF"
    assert prop.use_kerning is True
    assert prop.use_font_space is False
    assert prop.sym_mark == "NONE"


def test_parse_font_ref():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.font_ref, CharFontRef)
    assert prop.font_ref.hangul == 0
    assert prop.font_ref.latin == 1


def test_parse_lang_fields():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.ratio, CharRatio)
    assert prop.ratio.hangul == 100
    assert isinstance(prop.spacing, CharSpacing)
    assert prop.spacing.latin == 0
    assert isinstance(prop.rel_size, CharRelSize)
    assert prop.rel_size.japanese == 100
    assert isinstance(prop.offset, CharOffset)
    assert prop.offset.symbol == 0


def test_parse_bool_flags():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert prop.bold is True
    assert prop.italic is True
    assert prop.emboss is None
    assert prop.engrave is None
    assert prop.supscript is None
    assert prop.subscript is None


def test_parse_underline():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.underline, CharUnderline)
    assert prop.underline.type == "BOTTOM"
    assert prop.underline.shape == "SOLID"
    assert prop.underline.color == "#0000FF"


def test_parse_strikeout():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.strikeout, CharStrikeout)
    assert prop.strikeout.shape == "NONE"


def test_parse_outline():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.outline, CharOutline)
    assert prop.outline.type == "NONE"


def test_parse_shadow():
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    assert isinstance(prop.shadow, CharShadow)
    assert prop.shadow.type == "DROP"
    assert prop.shadow.color == "#C0C0C0"
    assert prop.shadow.offset_x == 10
    assert prop.shadow.offset_y == 10


# --- Test missing optional children ---


def test_parse_minimal_charpr():
    """charPr with only id and height — all optional children are None."""
    xml_str = f'<hh:charPr xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" id="5" height="2000"/>'
    elem = etree.fromstring(xml_str.encode())
    prop = parse_char_property(elem)

    assert prop.id == 5
    assert prop.height == 2000
    assert prop.font_ref is None
    assert prop.underline is None
    assert prop.bold is None
    assert prop.italic is None
    assert prop.other_children == {}


def test_parse_unknown_child():
    """Unknown child elements go to other_children."""
    xml_str = f"""\
<hh:charPr xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" id="3" height="1000">
  <hh:bold/>
  <hh:customTag foo="bar"/>
</hh:charPr>"""
    elem = etree.fromstring(xml_str.encode())
    prop = parse_char_property(elem)

    assert prop.bold is True
    assert "customTag" in prop.other_children


# --- Test stdlib ET compat ---


def test_parse_stdlib_et():
    """parse_char_property works with stdlib ET elements too."""
    elem = _make_charpr_xml(lxml=False)
    prop = parse_char_property(elem)

    assert prop.id == 0
    assert prop.height == 1000
    assert prop.bold is True
    assert prop.italic is True
    assert isinstance(prop.font_ref, CharFontRef)


# --- Test serialize round-trip ---


def test_serialize_roundtrip_lxml():
    """Parse -> serialize -> re-parse produces identical CharProperty."""
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    # Serialize into a fresh element
    new_elem = etree.Element(f"{_HH}charPr")
    new_elem.set("id", str(prop.id))
    serialize_char_property_into(prop, new_elem)

    # Re-parse
    prop2 = parse_char_property(new_elem)

    assert prop2.id == prop.id
    assert prop2.height == prop.height
    assert prop2.text_color == prop.text_color
    assert prop2.shade_color == prop.shade_color
    assert prop2.use_kerning == prop.use_kerning
    assert prop2.bold == prop.bold
    assert prop2.italic == prop.italic
    assert prop2.font_ref.hangul == prop.font_ref.hangul
    assert prop2.font_ref.latin == prop.font_ref.latin
    assert prop2.underline.type == prop.underline.type
    assert prop2.underline.color == prop.underline.color
    assert prop2.shadow.type == prop.shadow.type
    assert prop2.shadow.offset_x == prop.shadow.offset_x


def test_serialize_roundtrip_stdlib():
    """Parse -> serialize -> re-parse works with stdlib ET."""
    elem = _make_charpr_xml(lxml=False)
    prop = parse_char_property(elem)

    new_elem = stdlib_ET.Element(f"{_HH}charPr")
    new_elem.set("id", str(prop.id))
    serialize_char_property_into(prop, new_elem)

    prop2 = parse_char_property(new_elem)

    assert prop2.height == prop.height
    assert prop2.bold == prop.bold
    assert prop2.italic == prop.italic


def test_serialize_preserves_element_order():
    """OWPML schema order: fontRef, ratio, spacing, relSz, offset, italic, bold, underline, ..."""
    elem = _make_charpr_xml()
    prop = parse_char_property(elem)

    new_elem = etree.Element(f"{_HH}charPr")
    new_elem.set("id", "0")
    serialize_char_property_into(prop, new_elem)

    child_tags = [etree.QName(child).localname for child in new_elem]
    expected_order = [
        "fontRef", "ratio", "spacing", "relSz", "offset",
        "italic", "bold", "underline", "strikeout", "outline", "shadow",
    ]
    assert child_tags == expected_order


# --- Test IDE-friendly access pattern ---


def test_typed_field_access():
    """Verify dot-notation access works (IDE autocomplete friendly)."""
    prop = CharProperty(
        id=0,
        height=2000,
        text_color="#FF0000",
        font_ref=CharFontRef(hangul=1, latin=2),
        underline=CharUnderline(type="BOTTOM", shape="SOLID", color="#0000FF"),
        bold=True,
    )

    assert prop.font_ref.hangul == 1
    assert prop.font_ref.latin == 2
    assert prop.underline.type == "BOTTOM"
    assert prop.height == 2000
    assert prop.bold is True
