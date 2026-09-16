import xml.etree.ElementTree as ET

import requests

from config import TALLY_URL


def post_xml(xml_str: str, url: str = TALLY_URL, timeout: int = 120) -> str:
    resp = requests.post(
        url,
        data=xml_str.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=timeout,
    )
    resp.raise_for_status()
    return resp.text


def parse_import_response(response_text: str) -> dict:
    try:
        root = ET.fromstring(response_text)
    except ET.ParseError:
        return {
            "created": None, "altered": None, "errors": None,
            "exceptions": None, "line_errors": [], "raw": response_text,
        }

    def _int(tag):
        text = root.findtext(f".//{tag}")
        return int(text.strip()) if text and text.strip().lstrip("-").isdigit() else None

    line_errors = [e.text.strip() for e in root.iter("LINEERROR") if e.text and e.text.strip()]

    return {
        "created": _int("CREATED"),
        "altered": _int("ALTERED"),
        "errors": _int("ERRORS"),
        "exceptions": _int("EXCEPTIONS"),
        "line_errors": line_errors,
        "raw": response_text,
    }
