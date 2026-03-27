"""Shared namespace constants for the HWPML/OWPML XML schemas.

All modules that need HWPML namespace URIs should import from here
to avoid duplicating the definitions.

Three generations of namespace URIs exist:
- 2011 (hancom.co.kr/hwpml/2011/*) — canonical form used internally
- 2016 (hancom.co.kr/hwpml/2016/*) — Hancom Office 2016+
- 2024 (owpml.org/owpml/2024/*)    — OWPML 2024 specification

All are normalised to 2011 at parse time by opc.xml_utils.
"""

# HWPML 2011 (canonical / normalised form)
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HP = f"{{{HP_NS}}}"

HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HH = f"{{{HH_NS}}}"

HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
HC = f"{{{HC_NS}}}"

HS_NS = "http://www.hancom.co.kr/hwpml/2011/section"
HS = f"{{{HS_NS}}}"

# HWPML 2016 (for namespace registration only)
HP10_NS = "http://www.hancom.co.kr/hwpml/2016/paragraph"
HS10_NS = "http://www.hancom.co.kr/hwpml/2016/section"
HC10_NS = "http://www.hancom.co.kr/hwpml/2016/core"
HH10_NS = "http://www.hancom.co.kr/hwpml/2016/head"

# SPEC: e2e-namespace-2024-compat-008 -- namespaces.py 2024 상수 추가
# OWPML 2024 (owpml.org domain; normalised to 2011 at parse time)
OWPML_2024_CORE = "http://www.owpml.org/owpml/2024/core"
OWPML_2024_HEAD = "http://www.owpml.org/owpml/2024/head"
OWPML_2024_PARAGRAPH = "http://www.owpml.org/owpml/2024/paragraph"
OWPML_2024_SECTION = "http://www.owpml.org/owpml/2024/section"
OWPML_2024_MASTER_PAGE = "http://www.owpml.org/owpml/2024/master-page"
OWPML_2024_HISTORY = "http://www.owpml.org/owpml/2024/history"
OWPML_2024_VERSION = "http://www.owpml.org/owpml/2024/version"
