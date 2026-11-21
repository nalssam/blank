#!/usr/bin/env python3
"""HTML 한 파일을 '슬라이드 안에서 실행되는' PowerPoint 파일로 바꾼다.

    python build.py demo/deck.html --slides 5 -o demo/live-demo.pptx

슬라이드마다 Live HTML 콘텐츠 추가 기능을 하나씩 놓고, HTML 원문을 그 추가 기능의
설정값(pptx 내부 webextension 파트)에 저장한다. 추가 기능이 없는 PC에서는 같은 자리에
미리 찍어 둔 그림이 보인다.
"""
import argparse
import json
import re
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
ADDIN_ID = "6f2b8c1e-4d7a-4e55-9a3b-2c8f5d1e7a90"  # addin/manifest.xml 의 <Id> 와 같아야 한다
ADDIN_VERSION = "1.0.0.0"

NS_WE = "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
RT_WEBEXT = "http://schemas.microsoft.com/office/2011/relationships/webextension"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
CT_WEBEXT = "application/vnd.ms-office.webextension+xml"


def webextension_xml(props: dict, snapshot_rid: str) -> bytes:
    root = etree.Element(f"{{{NS_WE}}}webextension", nsmap={"we": NS_WE})
    root.set("id", "{%s}" % str(uuid.uuid4()).upper())
    ref = etree.SubElement(root, f"{{{NS_WE}}}reference")
    ref.set("id", ADDIN_ID)
    ref.set("version", ADDIN_VERSION)
    ref.set("store", "developer")
    ref.set("storeType", "Registry")
    etree.SubElement(root, f"{{{NS_WE}}}alternateReferences")
    pr = etree.SubElement(root, f"{{{NS_WE}}}properties")
    for k, v in props.items():
        p = etree.SubElement(pr, f"{{{NS_WE}}}property")
        p.set("name", k)
        p.set("value", json.dumps(v))  # Office 설정값은 JSON 문자열로 저장된다
    etree.SubElement(root, f"{{{NS_WE}}}bindings")
    snap = etree.SubElement(root, f"{{{NS_WE}}}snapshot", nsmap={"r": NS_R})
    snap.set(f"{{{NS_R}}}embed", snapshot_rid)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def frame_xml(shape_id: int, rid: str, cx: int, cy: int) -> etree._Element:
    xml = f"""
<mc:AlternateContent xmlns:mc="{NS_MC}" xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">
  <mc:Choice xmlns:we="{NS_WE}" Requires="we">
    <p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="{shape_id}" name="Live HTML {shape_id}" title="Live HTML"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="{NS_WE}"><we:webextensionref r:id="{rid}"/></a:graphicData>
      </a:graphic>
    </p:graphicFrame>
  </mc:Choice>
  <mc:Fallback/>
</mc:AlternateContent>"""
    return etree.fromstring(xml)


def write_manifest(host_url: str) -> None:
    base = host_url.rstrip("/")
    tpl = (HERE / "addin" / "manifest.template.xml").read_text(encoding="utf-8")
    out = tpl.replace("{{HOST}}", base).replace("{{ID}}", ADDIN_ID).replace("{{VERSION}}", ADDIN_VERSION)
    (HERE / "addin" / "manifest.xml").write_text(out, encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("html", type=Path, help="슬라이드 덱 HTML (한 파일)")
    ap.add_argument("--slides", type=int, help="슬라이드 수 (생략하면 <section class=\"slide\"> 개수)")
    ap.add_argument("-o", "--out", type=Path, help="출력 pptx 경로")
    ap.add_argument("--size", default="1920x1080", help="HTML 기준 해상도 (기본 1920x1080)")
    ap.add_argument("--host-url", help="추가 기능 페이지가 올라간 https 주소. 주면 addin/manifest.xml 을 다시 만든다")
    args = ap.parse_args()

    html = args.html.read_text(encoding="utf-8")
    n = args.slides or len(re.findall(r"<section[^>]*class=\"[^\"]*\bslide\b", html))
    if n < 1:
        ap.error("슬라이드 수를 알 수 없습니다. --slides 로 지정하세요.")
    w, h = (int(x) for x in args.size.lower().split("x"))
    out = args.out or args.html.with_suffix(".pptx")
    if args.host_url:
        write_manifest(args.host_url)

    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["node", str(HERE / "snap.js"), str(args.html), tmp, str(n), str(w), str(h)],
                       check=True, cwd=HERE)

        prs = Presentation()
        prs.slide_width = Emu(12192000)
        prs.slide_height = Emu(round(12192000 * h / w))
        blank = prs.slide_layouts[6]

        for i in range(1, n + 1):
            slide = prs.slides.add_slide(blank)
            # 1) 미리보기 그림 — 추가 기능을 못 쓰는 환경에서 보이는 대체 화면
            pic = slide.shapes.add_picture(str(Path(tmp) / f"slide{i}.png"), 0, 0,
                                           prs.slide_width, prs.slide_height)
            img_rid = pic._element.blipFill.blip.rEmbed
            image_part = slide.part.related_part(img_rid)

            # 2) webextension 파트 — HTML 원문이 여기에 들어간다
            uri = PackURI(f"/ppt/webextensions/webextension{i}.xml")
            we_part = Part(uri, CT_WEBEXT, prs.part.package, b"")
            snap_rid = we_part.relate_to(image_part, RT_IMAGE)
            we_part._blob = webextension_xml({"html": html, "slide": i, "width": w, "height": h}, snap_rid)
            we_rid = slide.part.relate_to(we_part, RT_WEBEXT)

            # 3) 그림을 "추가 기능 프레임 / 대체 그림" 묶음으로 감싼다
            pic_el = pic._element
            alt = frame_xml(pic.shape_id + 100, we_rid, prs.slide_width, prs.slide_height)
            pic_el.addprevious(alt)
            alt.find(f"{{{NS_MC}}}Fallback").append(pic_el)

        prs.save(str(out))

    print(f"완료: {out}  (슬라이드 {n}장, HTML {len(html.encode('utf-8')) / 1024:.0f}KB × {n})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
