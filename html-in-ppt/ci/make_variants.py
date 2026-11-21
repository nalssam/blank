#!/usr/bin/env python3
"""Builds the documents for the diagnosis matrix (test only).

  D1 = our deck, our manifest id                 (the real thing)
  D2 = Microsoft's template deck, our manifest id
  D3 = our deck, control manifest id (minimal ASCII manifest)
  D4 = Microsoft's template deck, control manifest id
  D5 = Microsoft's template deck, manifest pointing at a public https page
  U1 = our deck with an add-in id that is not registered (what an unprepared PC shows)
  Z1..Z5 = our deck with the HTML of slide 1 padded to 0.5 / 1 / 1.9 / 4 / 10 MB (settings size limit)
"""
import sys
import zipfile
from pathlib import Path

OURS = "6f2b8c1e-4d7a-4e55-9a3b-2c8f5d1e7a90"
CONTROL = "0d1b6a3e-52a7-4b1c-9a39-7e2f4c1d8b55"
PUBLIC = "3c9e2f71-8a44-4d0b-b6f5-1e7a90c2d4b8"
ZERO = "00000000-0000-0000-0000-000000000000"


def rewrite(src: Path, dst: Path, old: str, new: str) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and old.encode() in data:
                data = data.replace(old.encode(), new.encode())
            zout.writestr(item, data)


def pad_html(src: Path, dst: Path, target_bytes: int) -> None:
    """Copy the deck, growing the HTML stored for slide 1 to about target_bytes (JSON size)."""
    import json
    import re
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/webextensions/webextension1.xml":
                text = data.decode("utf-8")
                m = re.search(r'<we:property name="html" value="([^"]*)"/>', text)
                raw = m.group(1).replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
                html = json.loads(raw)
                pad = max(0, target_bytes - len(json.dumps(html)))
                html = html.replace("</body>", "<!--" + "x" * pad + "--></body>")
                enc = json.dumps(html).replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")
                text = text[:m.start(1)] + enc + text[m.end(1):]
                data = text.encode("utf-8")
            zout.writestr(item, data)


def manifest(addin: str, url: str, name: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="ContentApp">
  <Id>{addin}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Live HTML test</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="{name}"/>
  <Description DefaultValue="Test content add-in"/>
  <IconUrl DefaultValue="https://localhost:3000/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/icon-64.png"/>
  <SupportUrl DefaultValue="https://localhost:3000/index.html"/>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="{url}"/>
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
"""


def main() -> None:
    deck, template, out = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
    out.mkdir(parents=True, exist_ok=True)
    rewrite(deck, out / "D1.pptx", "\0never\0", "")
    rewrite(template, out / "D2.pptx", ZERO, OURS)
    rewrite(deck, out / "D3.pptx", OURS, CONTROL)
    rewrite(template, out / "D4.pptx", ZERO, CONTROL)
    rewrite(template, out / "D5.pptx", ZERO, PUBLIC)
    rewrite(deck, out / "U1.pptx", OURS, "9a8b7c6d-5e4f-4a3b-8c2d-1e0f9a8b7c6d")      # an id nobody registered
    for label, size in (("Z1", 500_000), ("Z2", 1_000_000), ("Z3", 1_900_000), ("Z4", 4_000_000), ("Z5", 10_000_000)):
        pad_html(deck, out / f"{label}.pptx", size)
    (out / "manifest-control.xml").write_text(
        manifest(CONTROL, "https://localhost:3000/index.html?m=control", "Live HTML control"), encoding="utf-8")
    (out / "manifest-public.xml").write_text(
        manifest(PUBLIC, "https://example.com/", "Live HTML public page"), encoding="utf-8")
    print("variants written to", out)


if __name__ == "__main__":
    main()
