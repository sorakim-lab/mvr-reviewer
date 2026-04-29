"""검토 결과 출력 모듈.

3가지 포맷 지원:
  - JSON : UI/다른 도구가 읽을 데이터
  - Markdown : 사람이 빠르게 훑는 리포트
  - DOCX with comments : 원본에 코멘트 삽입 (Word에서 그대로 검토)
"""

import json
import shutil
from collections import Counter
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# 1. JSON
# ============================================================

def export_json(violations, doc_path, out_path):
    """위반 사항을 JSON으로."""
    sev_count = Counter(v.severity for v in violations)
    rule_count = Counter(v.rule_id for v in violations)

    data = {
        "metadata": {
            "document": Path(doc_path).name,
            "reviewed_at": datetime.now().isoformat(timespec="seconds"),
            "total_violations": len(violations),
            "by_severity": dict(sev_count),
            "rules_triggered": len(rule_count),
        },
        "violations": [asdict(v) for v in violations],
    }

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return out_path


# ============================================================
# 2. Markdown 리포트
# ============================================================

def export_markdown(violations, doc_path, out_path):
    """위반 사항을 Markdown 리포트로."""
    sev_count = Counter(v.severity for v in violations)
    sev_order = {"required": 0, "recommended": 1, "info": 2}
    sev_label = {
        "required": "🔴 필수",
        "recommended": "🟡 권고",
        "info": "🔵 참고",
    }

    lines = []
    lines.append(f"# MVR 검토 리포트")
    lines.append("")
    lines.append(f"- **문서**: {Path(doc_path).name}")
    lines.append(f"- **검토 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"- **총 위반**: {len(violations)}건")
    lines.append("")
    lines.append("## 요약")
    lines.append("")
    lines.append("| Severity | 건수 |")
    lines.append("|---|---|")
    for sev in ["required", "recommended", "info"]:
        lines.append(f"| {sev_label[sev]} | {sev_count.get(sev, 0)} |")
    lines.append("")

    # severity → rule_id 별로 그룹핑
    sorted_v = sorted(
        violations,
        key=lambda v: (sev_order[v.severity], v.rule_id),
    )

    current_sev = None
    for v in sorted_v:
        if v.severity != current_sev:
            current_sev = v.severity
            lines.append(f"## {sev_label[current_sev]}")
            lines.append("")

        lines.append(f"### `{v.rule_id}` — {v.rule_name}")
        lines.append("")
        lines.append(f"- **위치**: {v.location}")
        lines.append(f"- **메시지**: {v.message}")
        if v.matched_text:
            text = v.matched_text.replace("|", "\\|")
            lines.append(f"- **발견**: `{text}`")
        lines.append("")

    if not violations:
        lines.append("위반 사항 없음 ✅")
        lines.append("")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return out_path


# ============================================================
# 3. DOCX 코멘트 삽입 (가장 강력)
# ============================================================

def _make_comment_xml(comment_id, author, text, date_str):
    """comment 한 개의 XML 생성."""
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    return (
        f'<w:comment xmlns:w="{w}" w:id="{comment_id}" '
        f'w:author="{author}" w:date="{date_str}" w:initials="MVR">'
        f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        f'</w:comment>'
    )


def export_docx_with_comments(violations, doc_path, out_path):
    """원본 docx에 코멘트 삽입.

    docx는 ZIP이므로 직접 풀고 XML 수정 후 다시 압축.
    이게 python-docx 내부 API보다 안정적.
    """
    import zipfile
    import re as _re
    import tempfile
    from xml.etree import ElementTree as ET

    marker = {
        "required": "[필수] ",
        "recommended": "[권고] ",
        "info": "[참고] ",
    }

    # 위반을 단락 인덱스별로 그룹화
    by_paragraph = {}
    global_violations = []
    for v in violations:
        loc = v.location
        if loc.startswith("본문 단락 "):
            try:
                idx = int(loc.replace("본문 단락 ", "").strip())
                by_paragraph.setdefault(idx, []).append(v)
                continue
            except ValueError:
                pass
        global_violations.append(v)

    # 코멘트 텍스트 만들기
    def violation_text(v):
        text = (
            f"{marker[v.severity]}{v.rule_id} — {v.rule_name}\n"
            f"{v.message}"
        )
        if v.matched_text:
            text += f"\n발견: {v.matched_text}"
        return text

    def xml_escape(s):
        return (
            s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
        )

    # 코멘트 ID 부여
    comments_to_add = []  # [(comment_id, paragraph_index, text), ...]
    next_id = 0

    for idx, vlist in sorted(by_paragraph.items()):
        for v in vlist:
            comments_to_add.append((next_id, idx, violation_text(v)))
            next_id += 1

    # 전역 위반은 마지막에 종합 코멘트 1개로 (첫 본문 단락에 붙임)
    global_paragraph_idx = None
    if global_violations:
        summary_lines = ["[종합 검토 결과 — 전역/표 위반]"]
        for v in global_violations:
            summary_lines.append(
                f"{marker[v.severity]}{v.rule_id} ({v.location}): {v.message}"
            )
        comments_to_add.append((next_id, 0, "\n".join(summary_lines)))
        global_paragraph_idx = 0
        next_id += 1

    # ZIP 풀기
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        with zipfile.ZipFile(doc_path, "r") as zin:
            zin.extractall(tmp)

        # 1. word/document.xml 수정 — lxml 사용 (namespace 보존)
        from lxml import etree as LET

        doc_xml_path = tmp / "word" / "document.xml"
        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        parser = LET.XMLParser(remove_blank_text=False)
        tree = LET.parse(str(doc_xml_path), parser).getroot()
        body = tree.find(f"{{{ns_w}}}body")
        if body is None:
            return out_path

        # 텍스트가 있는 단락만 카운트 (review_engine의 paragraphs 인덱스와 일치시키기 위함)
        text_paragraphs = []
        for child in body:
            if child.tag == f"{{{ns_w}}}p":
                # 단락에 텍스트가 있는지
                has_text = False
                for t in child.iter(f"{{{ns_w}}}t"):
                    if t.text and t.text.strip():
                        has_text = True
                        break
                if has_text:
                    text_paragraphs.append(child)

        # 인덱스 → 코멘트 ID들 매핑
        idx_to_cids = {}
        for cid, pidx, _text in comments_to_add:
            idx_to_cids.setdefault(pidx, []).append(cid)

        for pidx, cids in idx_to_cids.items():
            if pidx >= len(text_paragraphs):
                continue
            p_elem = text_paragraphs[pidx]

            for cid in cids:
                # commentRangeStart (단락 첫 자식 앞에 삽입, pPr 뒤)
                crs = LET.Element(f"{{{ns_w}}}commentRangeStart")
                crs.set(f"{{{ns_w}}}id", str(cid))

                # commentRangeEnd
                cre = LET.Element(f"{{{ns_w}}}commentRangeEnd")
                cre.set(f"{{{ns_w}}}id", str(cid))

                # commentReference run
                ref_run = LET.Element(f"{{{ns_w}}}r")
                rpr = LET.SubElement(ref_run, f"{{{ns_w}}}rPr")
                rstyle = LET.SubElement(rpr, f"{{{ns_w}}}rStyle")
                rstyle.set(f"{{{ns_w}}}val", "CommentReference")
                cref = LET.SubElement(ref_run, f"{{{ns_w}}}commentReference")
                cref.set(f"{{{ns_w}}}id", str(cid))

                # 삽입 위치: pPr 다음
                insert_pos = 0
                for i, child in enumerate(list(p_elem)):
                    if child.tag != f"{{{ns_w}}}pPr":
                        insert_pos = i
                        break
                else:
                    insert_pos = len(p_elem)
                p_elem.insert(insert_pos, crs)
                p_elem.append(cre)
                p_elem.append(ref_run)

        # 새 document.xml 쓰기 (lxml — namespace 보존)
        new_doc_xml = LET.tostring(
            tree,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        ).decode("utf-8")
        doc_xml_path.write_text(new_doc_xml, encoding="utf-8")

        # 2. word/comments.xml 생성
        date_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
        comment_xmls = []
        for cid, _pidx, text in comments_to_add:
            text_escaped = xml_escape(text)
            # 줄바꿈은 <w:br/>로 변환
            parts = text_escaped.split("\n")
            run_xml = '<w:r>'
            for i, part in enumerate(parts):
                if i > 0:
                    run_xml += '<w:br/>'
                run_xml += f'<w:t xml:space="preserve">{part}</w:t>'
            run_xml += '</w:r>'

            comment_xmls.append(
                f'<w:comment w:id="{cid}" w:author="MVR Reviewer" '
                f'w:date="{date_str}" w:initials="MVR">'
                f'<w:p>{run_xml}</w:p>'
                f'</w:comment>'
            )

        comments_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:comments xmlns:w="{ns_w}">'
            + "".join(comment_xmls)
            + '</w:comments>'
        )
        (tmp / "word" / "comments.xml").write_text(comments_xml, encoding="utf-8")

        # 3. word/_rels/document.xml.rels 수정 — comments 관계 추가
        rels_path = tmp / "word" / "_rels" / "document.xml.rels"
        rels_xml = rels_path.read_text(encoding="utf-8")
        # 기존 relationship ID 중 가장 큰 거 + 1
        existing_ids = _re.findall(r'Id="rId(\d+)"', rels_xml)
        new_rid = max((int(x) for x in existing_ids), default=0) + 1
        new_rel = (
            f'<Relationship Id="rId{new_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
            f'Target="comments.xml"/>'
        )
        rels_xml = rels_xml.replace("</Relationships>", new_rel + "</Relationships>")
        rels_path.write_text(rels_xml, encoding="utf-8")

        # 4. [Content_Types].xml 수정 — comments content type 등록
        ct_path = tmp / "[Content_Types].xml"
        ct_xml = ct_path.read_text(encoding="utf-8")
        comments_ct = (
            '<Override PartName="/word/comments.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.comments+xml"/>'
        )
        if comments_ct not in ct_xml:
            ct_xml = ct_xml.replace("</Types>", comments_ct + "</Types>")
        ct_path.write_text(ct_xml, encoding="utf-8")

        # 5. 다시 ZIP으로 묶기
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for path in tmp.rglob("*"):
                if path.is_file():
                    arcname = str(path.relative_to(tmp)).replace("\\", "/")
                    zout.write(path, arcname)

    return out_path


# ============================================================
# 통합 실행기
# ============================================================

def export_all(violations, doc_path, out_dir):
    """3가지 포맷 다 생성."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(doc_path).stem

    json_path = export_json(violations, doc_path, out_dir / f"{stem}_review.json")
    md_path = export_markdown(violations, doc_path, out_dir / f"{stem}_review.md")
    docx_path = export_docx_with_comments(
        violations, doc_path, out_dir / f"{stem}_reviewed.docx"
    )

    return {
        "json": json_path,
        "markdown": md_path,
        "docx": docx_path,
    }


# ============================================================
# 실행
# ============================================================

if __name__ == "__main__":
    from review_engine import review_document

    base = Path(__file__).parent
    docx_path = base / "sample_MVR_with_violations.docx"
    rules_path = base / "MVR_rules.yaml"

    violations, _ = review_document(docx_path, rules_path)
    out_dir = base / "output"

    paths = export_all(violations, docx_path, out_dir)

    print("=" * 60)
    print("출력 완료")
    print("=" * 60)
    for fmt, p in paths.items():
        size = Path(p).stat().st_size
        print(f"  {fmt:10s} {p.name}  ({size:,} bytes)")
