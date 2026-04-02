---
name: hanguldoc
description: >
  한국 문서(HWP, HWPX, PDF) 통합 읽기/쓰기/편집 스킬.
  'hwpx', 'hwp', 'pdf', '한글 문서', '공문서', '보고서', '회의록', '제안서' 관련 요청,
  문서 파싱·비교·양식 추출, HWPX 생성·편집·레이아웃 복원 요청 시 항상 사용.
  kordoc(파싱 엔진)과 jun-hwpx2hwpx(XML 생성 엔진)의 강점을 결합한 통합 스킬.
---

# hanguldoc — 한국 문서 통합 스킬

**읽기(파싱)는 kordoc, 쓰기(생성)는 XML-first.** 두 엔진을 상황에 맞게 선택한다.

---

## 0. 작업 유형 판별 (가장 먼저 실행)

요청을 받으면 아래 표에서 해당 유형을 찾아 지정된 워크플로우로 이동한다.

| 요청 유형 | 사용 엔진 | 이동 |
|---|---|---|
| HWP/HWPX/PDF 내용 읽기·추출·요약 | kordoc | [WF-1 읽기](#wf-1-읽기--파싱) |
| 두 문서 비교·신구대조 | kordoc | [WF-2 비교](#wf-2-문서-비교-diff) |
| 공문서 양식 필드 추출 (성명·소속 등) | kordoc | [WF-3 양식 인식](#wf-3-양식-인식) |
| HWPX 첨부 있음 → 내용 수정·재작성 | XML-first | [WF-4 레퍼런스 복원](#wf-4-레퍼런스-복원-우선-편집) |
| HWPX 첨부 없음 → 새 문서 생성 | XML-first | [WF-5 신규 생성](#wf-5-신규-hwpx-생성) |
| 기존 HWPX 구조 직접 편집 | XML-first | [WF-6 unpack-edit-pack](#wf-6-unpack--edit--pack) |
| Markdown → HWPX 변환 | kordoc or XML | [WF-7 역변환](#wf-7-markdown--hwpx-역변환) |

> **HWP 5.x(바이너리) 파일** 읽기는 kordoc만 지원한다. 생성·편집은 HWPX만 가능하므로, 사용자에게 한글에서 HWPX로 다시 저장하도록 안내한다.

---

## 환경

```bash
# kordoc (읽기 엔진) — Node.js 18+
npx kordoc <파일>                         # CLI 즉시 사용
npm install kordoc                        # 라이브러리로 설치

# PDF 파싱 추가 시
npm install pdfjs-dist

# XML-first 엔진 (쓰기) — Python + lxml
pip install lxml --break-system-packages

SKILL_DIR="/home/claude/skills/kordoc-hwpx"   # 이 스킬 디렉토리
```

---

## WF-1: 읽기 / 파싱

HWP·HWPX·PDF를 Markdown + 구조화 데이터(IRBlock[])로 변환한다.

### CLI (빠른 확인)

```bash
# 마크다운 출력
npx kordoc 문서.hwpx
npx kordoc 보고서.hwp
npx kordoc 공문.pdf

# 파일로 저장
npx kordoc 문서.hwpx -o 문서.md

# JSON (blocks + metadata 포함)
npx kordoc 문서.hwpx --format json

# 페이지 범위
npx kordoc 문서.hwpx --pages 1-3

# 일괄 변환
npx kordoc *.hwpx -d ./변환결과/
```

### TypeScript API

```typescript
import { parse } from "kordoc"
import { readFileSync } from "fs"

const buffer = readFileSync("문서.hwpx").buffer
const result = await parse(buffer)

if (result.success) {
  console.log(result.markdown)    // 마크다운 텍스트
  console.log(result.blocks)      // IRBlock[] 구조화 데이터
  console.log(result.metadata)    // { title, author, createdAt, pageCount, ... }
  console.log(result.outline)     // 헤딩 트리
  console.log(result.warnings)    // 파싱 경고 (이미지 스킵 등)
}
```

### Python (lxml 직접 파싱 — kordoc 설치 없이)

HWPX만 해당. 간단한 텍스트 추출 시 사용.

```bash
python3 "$SKILL_DIR/scripts/text_extract.py" 문서.hwpx
python3 "$SKILL_DIR/scripts/text_extract.py" 문서.hwpx --include-tables
python3 "$SKILL_DIR/scripts/text_extract.py" 문서.hwpx --format markdown
```

### 포맷 자동 감지 원리

kordoc는 확장자가 아닌 **매직 바이트**로 포맷을 감지한다.

| 포맷 | 매직 바이트 | 비고 |
|---|---|---|
| HWPX | `PK\x03\x04` (ZIP) | 한컴 2020+ |
| HWP 5.x | `\xD0\xCF\x11\xE0` (OLE2) | 레거시 바이너리 |
| PDF | `%PDF` | pdfjs-dist 필요 |

---

## WF-2: 문서 비교 (Diff)

두 문서를 IR 블록 단위로 비교해 신구대조표를 생성한다. **HWP ↔ HWPX 크로스 포맷 비교 가능.**

```typescript
import { compare } from "kordoc"
import { readFileSync } from "fs"

const bufA = readFileSync("구버전.hwpx").buffer
const bufB = readFileSync("신버전.hwp").buffer   // 크로스 포맷 OK

const diff = await compare(bufA, bufB)
console.log(diff.stats)   // { added: 3, removed: 1, modified: 5, unchanged: 42 }
console.log(diff.diffs)   // BlockDiff[] — 테이블은 CellDiff[][] 포함
```

### 결과 해석

```typescript
for (const d of diff.diffs) {
  if (d.type === "modified" && d.cellDiffs) {
    // 테이블 셀 단위 변경 확인
    d.cellDiffs.forEach((row, r) =>
      row.forEach((cell, c) => {
        if (cell.type !== "unchanged")
          console.log(`[${r},${c}] ${cell.before} → ${cell.after}`)
      })
    )
  }
}
```

---

## WF-3: 양식 인식

공문서 테이블에서 label-value 쌍을 자동 추출한다.

```typescript
import { parse, extractFormFields } from "kordoc"

const result = await parse(buffer)
if (result.success) {
  const form = extractFormFields(result.blocks)
  // form.fields → [{ label: "성명", value: "홍길동", row: 0, col: 0 }, ...]
  // form.confidence → 0.85
}
```

### 감지 가능 필드

성명·이름·소속·직위·직급·부서·전화번호·이메일·주소·생년월일·일시·날짜·장소·목적·금액·수량·합계 등 한국 공문서 표준 레이블.

---

## WF-4: 레퍼런스 복원 우선 편집

**사용자가 HWPX를 첨부한 경우 기본 워크플로우.** 원본 레이아웃을 99% 복원하고 내용만 교체한다.

### 핵심 원칙

- `charPrIDRef`, `paraPrIDRef`, `borderFillIDRef` 참조 체계 유지
- 표의 `rowCnt`, `colCnt`, `colSpan`, `rowSpan`, `cellSz`, `cellMargin` 동일
- 쪽수는 레퍼런스와 반드시 동일 (`page_guard.py` 필수 통과)
- 변경은 사용자 요청 범위(텍스트·값)로만 제한

### 실행 순서

```bash
# 1. 레퍼런스 심층 분석
python3 "$SKILL_DIR/scripts/analyze_template.py" reference.hwpx \
  --extract-header /tmp/ref_header.xml \
  --extract-section /tmp/ref_section.xml

# 2. /tmp/ref_section.xml 복제 → 텍스트만 수정 → /tmp/new_section0.xml 저장

# 3. 복원 빌드
python3 "$SKILL_DIR/scripts/build_hwpx.py" \
  --header /tmp/ref_header.xml \
  --section /tmp/new_section0.xml \
  --output result.hwpx

# 4. 무결성 검증
python3 "$SKILL_DIR/scripts/validate.py" result.hwpx

# 5. 쪽수 드리프트 가드 (필수 — 실패 시 재수정)
python3 "$SKILL_DIR/scripts/page_guard.py" \
  --reference reference.hwpx \
  --output result.hwpx
```

### 분석 출력 항목

| 항목 | 내용 |
|---|---|
| charPr | 글꼴 크기·이름·색상·볼드/이탤릭/밑줄 |
| paraPr | 정렬·줄간격·여백·들여쓰기·borderFillIDRef |
| borderFill | 테두리 타입/두께·배경색 |
| 표 구조 | 행×열·열너비·셀 span·margin·vertAlign |
| 문서 구조 | 페이지 크기·여백·섹션 설정 |

---

## WF-5: 신규 HWPX 생성

첨부 레퍼런스 없이 처음부터 만들 때 사용.

### 템플릿 선택

| 템플릿 | 용도 |
|---|---|
| `base` | 빈 문서 (최소 구조) |
| `gonmun` | 공문 (기관명·수신·제목·서명) |
| `report` | 보고서 (섹션 헤더·체크항목·들여쓰기) |
| `minutes` | 회의록 |
| `proposal` | 제안서/사업개요 (색상 헤더바·번호 배지) |

```bash
# 빈 문서
python3 "$SKILL_DIR/scripts/build_hwpx.py" --output result.hwpx

# 템플릿 사용
python3 "$SKILL_DIR/scripts/build_hwpx.py" --template gonmun --output result.hwpx

# 커스텀 section0.xml 오버라이드
python3 "$SKILL_DIR/scripts/build_hwpx.py" \
  --template report \
  --section my_section0.xml \
  --output result.hwpx

# 메타데이터 포함
python3 "$SKILL_DIR/scripts/build_hwpx.py" \
  --template report \
  --section my.xml \
  --title "제목" --creator "작성자" \
  --output result.hwpx
```

### section0.xml 작성 핵심 패턴

**필수 구조** — 첫 문단 첫 run에 `secPr` + `colPr` 포함:

```xml
<hp:p id="1000000001" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0">
    <hp:secPr ...><!-- 페이지 크기·여백 --></hp:secPr>
    <hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>
  </hp:run>
  <hp:run charPrIDRef="0"><hp:t/></hp:run>
</hp:p>
```

**일반 문단:**
```xml
<hp:p id="고유ID" paraPrIDRef="문단스타일ID" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="글자스타일ID"><hp:t>텍스트</hp:t></hp:run>
</hp:p>
```

**빈 줄:**
```xml
<hp:p id="고유ID" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="0"><hp:t/></hp:run>
</hp:p>
```

> 상세 XML 패턴(표·병합셀·혼합서식) → `$SKILL_DIR/references/hwpx-format.md` 참조

### 템플릿별 스타일 ID 요약

**gonmun (공문)**

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 22pt 볼드 함초롬바탕 (기관명·제목) |
| charPr 8 | 글자 | 16pt 볼드 함초롬바탕 (서명자) |
| charPr 9 | 글자 | 8pt 함초롬바탕 (하단 연락처) |
| charPr 10 | 글자 | 10pt 볼드 (표 헤더) |
| paraPr 20 | 문단 | CENTER 160% |
| paraPr 21 | 문단 | CENTER 130% (표 셀) |
| borderFill 3 | 테두리 | SOLID 0.12mm 4면 |
| borderFill 4 | 테두리 | SOLID 0.12mm + #D6DCE4 배경 |

**report (보고서)**

| ID | 유형 | 설명 |
|----|------|------|
| charPr 7 | 글자 | 20pt 볼드 (문서 제목) |
| charPr 8 | 글자 | 14pt 볼드 (소제목) |
| charPr 13 | 글자 | 12pt 볼드 함초롬돋움 (섹션 헤더) |
| paraPr 24 | 문단 | JUSTIFY left 600 (□ 항목) |
| paraPr 25 | 문단 | JUSTIFY left 1200 (①②③ 하위) |
| paraPr 27 | 문단 | LEFT 상하단 테두리선 (섹션 헤더용) |

> 전체 스타일 ID 맵 → `$SKILL_DIR/references/hwpx-format.md`

---

## WF-6: Unpack → Edit → Pack

기존 HWPX의 XML을 직접 수정할 때 사용.

```bash
# 1. 압축 해제 (XML pretty-print)
python3 "$SKILL_DIR/scripts/office/unpack.py" document.hwpx ./unpacked/

# 2. XML 직접 편집
#    본문: ./unpacked/Contents/section0.xml
#    스타일: ./unpacked/Contents/header.xml

# 3. 재패키징
python3 "$SKILL_DIR/scripts/office/pack.py" ./unpacked/ edited.hwpx

# 4. 검증
python3 "$SKILL_DIR/scripts/validate.py" edited.hwpx
```

---

## WF-7: Markdown → HWPX 역변환

AI가 생성한 마크다운 문서를 바로 한글 공문서로 변환한다.

### kordoc API (권장 — 테이블 자동 변환 포함)

```typescript
import { markdownToHwpx } from "kordoc"
import { writeFileSync } from "fs"

const md = `# 제목\n\n본문 내용\n\n| 이름 | 직급 |\n| --- | --- |\n| 홍길동 | 과장 |`
const hwpxBuffer = await markdownToHwpx(md)
writeFileSync("출력.hwpx", Buffer.from(hwpxBuffer))
```

### XML-first (서식·레이아웃 정밀 제어 시)

마크다운을 section0.xml로 수동 변환 후 WF-5 실행. 공문서 레이아웃 정확도가 높아야 할 때 선택.

---

## WF-8: MCP 서버 (Claude Desktop 연동)

Claude Desktop에서 HWP·HWPX·PDF를 직접 읽어 분석할 수 있다.

```json
{
  "mcpServers": {
    "kordoc": {
      "command": "npx",
      "args": ["-y", "kordoc-mcp"]
    }
  }
}
```

**7개 MCP 도구:**

| 도구 | 설명 |
|---|---|
| `parse_document` | 문서 → 마크다운 (메타데이터 포함) |
| `detect_format` | 포맷 감지 |
| `parse_metadata` | 메타데이터만 빠르게 |
| `parse_pages` | 특정 페이지 범위만 |
| `parse_table` | N번째 테이블만 추출 |
| `compare_documents` | 두 문서 비교 |
| `parse_form` | 양식 필드 JSON 추출 |

---

## 검증 (모든 WF 공통)

생성·편집 결과는 반드시 검증 후 제출한다.

```bash
# 1. 구조 검증 (ZIP·XML well-formedness)
python3 "$SKILL_DIR/scripts/validate.py" result.hwpx

# 2. 쪽수 드리프트 가드 (레퍼런스 기반 작업 시 필수)
python3 "$SKILL_DIR/scripts/page_guard.py" \
  --reference reference.hwpx \
  --output result.hwpx

# 3. 내용 확인 (kordoc로 역파싱)
npx kordoc result.hwpx
```

---

## 단위 변환 (HWPUNIT)

| 값 | HWPUNIT | 비고 |
|----|---------|------|
| 1pt | 100 | |
| 1mm | 283.5 | |
| 1cm | 2835 | |
| A4 폭 | 59528 | 210mm |
| A4 높이 | 84186 | 297mm |
| 좌우여백 | 8504 | 30mm |
| 본문폭 | 42520 | 150mm |

---

## Critical Rules

### 공통
1. **작업 유형 판별 먼저**: 읽기/비교/양식 → kordoc, 생성/편집 → XML-first
2. **HWP 5.x는 읽기 전용**: 생성·편집 불가. 한글에서 HWPX로 저장 후 작업 안내
3. **검증 필수**: 생성·편집 후 `validate.py` 통과 확인

### XML-first (WF-4·5·6)
4. **레퍼런스 첨부 시 WF-4 강제**: `analyze_template.py` 기반 복원
5. **쪽수 동일 필수**: 레퍼런스 기반 작업 → `page_guard.py` 필수 통과
6. **구조 변경 제한**: 사용자 요청 없이 `hp:p`, `hp:tbl`, `rowCnt`, `colCnt`, `pageBreak` 변경 금지
7. **치환 우선 편집**: 텍스트 노드 치환 > 문단·표 추가/삭제
8. **secPr 필수**: section0.xml 첫 문단 첫 run에 반드시 secPr + colPr 포함
9. **mimetype 순서**: HWPX 패키징 시 mimetype은 첫 번째 ZIP 엔트리, ZIP_STORED
10. **네임스페이스 보존**: `hp:`, `hs:`, `hh:`, `hc:` 접두사 유지
11. **itemCnt 정합성**: header.xml의 charProperties/paraProperties/borderFills itemCnt = 실제 자식 수
12. **ID 참조 정합성**: section0.xml의 charPrIDRef/paraPrIDRef ↔ header.xml 정의 일치
13. **page_guard 최종 게이트**: validate.py 통과 ≠ 완료. page_guard.py도 통과해야 완료

### kordoc (WF-1·2·3)
14. **Node.js 18+ 필요**: kordoc CLI·API 실행 환경 확인
15. **PDF는 pdfjs-dist 필요**: `npm install pdfjs-dist` 선행 설치
16. **보안 내장**: ZIP bomb·경로 순회·숨김 텍스트(프롬프트 인젝션) 자동 방어

---

## 참조 문서

| 파일 | 내용 |
|---|---|
| `$SKILL_DIR/references/hwpx-format.md` | OWPML XML 요소 상세 레퍼런스 |
| `$SKILL_DIR/references/kordoc-api.md` | kordoc 전체 API·CLI 가이드 |
