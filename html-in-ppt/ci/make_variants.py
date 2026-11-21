#!/usr/bin/env python3
"""Builds the documents for the diagnosis matrix (test only).

  D1 = our deck, our manifest id                 (the real thing)
  D2 = Microsoft's template deck, our manifest id
  D3 = our deck, control manifest id (minimal ASCII manifest)
  D4 = Microsoft's template deck, control manifest id
  D5 = Microsoft's template deck, manifest pointing at a public https page
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
    (out / "manifest-control.xml").write_text(
        manifest(CONTROL, "https://localhost:3000/index.html?m=control", "Live HTML control"), encoding="utf-8")
    (out / "manifest-public.xml").write_text(
        manifest(PUBLIC, "https://example.com/", "Live HTML public page"), encoding="utf-8")
    print("variants written to", out)


if __name__ == "__main__":
    main()
