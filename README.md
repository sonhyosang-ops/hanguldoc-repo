# hanguldoc

**한국 문서(HWP · HWPX · PDF) 통합 읽기/쓰기/편집 Claude 스킬**

`kordoc` 파싱 엔진과 XML-first 생성 엔진의 강점을 결합한 통합 스킬입니다.

---

## 기능

| 기능 | 설명 |
|---|---|
| **읽기 / 파싱** | HWP 5.x · HWPX · PDF → Markdown + 구조화 데이터 |
| **문서 비교** | 두 문서 블록 단위 신구대조 (크로스 포맷 지원) |
| **양식 인식** | 공문서 테이블에서 성명·소속 등 label-value 자동 추출 |
| **HWPX 생성** | 템플릿 기반 공문 · 보고서 · 회의록 · 제안서 생성 |
| **레퍼런스 복원** | 기존 HWPX 레이아웃 99% 복원 후 내용만 교체 |
| **Markdown → HWPX** | AI 생성 텍스트를 바로 한글 공문서로 변환 |
| **MCP 연동** | Claude Desktop에서 문서 직접 읽기 (kordoc MCP) |

---

## 구조

```
hanguldoc/
├── SKILL.md                  # Claude 스킬 본체
├── scripts/
│   ├── build_hwpx.py         # 템플릿 + XML → HWPX 조립
│   ├── analyze_template.py   # HWPX 심층 분석
│   ├── validate.py           # HWPX 구조 검증
│   ├── page_guard.py         # 쪽수 드리프트 가드
│   ├── text_extract.py       # 텍스트 추출
│   ├── md2hwpx.py            # Markdown → HWPX
│   └── office/
│       ├── unpack.py         # HWPX → XML 디렉토리
│       └── pack.py           # XML 디렉토리 → HWPX
├── templates/
│   ├── base/                 # 기본 빈 문서
│   ├── gonmun/               # 공문
│   ├── report/               # 보고서
│   ├── minutes/              # 회의록
│   └── proposal/             # 제안서
└── references/
    ├── hwpx-format.md        # OWPML XML 요소 레퍼런스
    └── kordoc-api.md         # kordoc API 가이드
```

---

## 설치

### 읽기 엔진 (kordoc)

```bash
npm install kordoc
# PDF 파싱 추가 시
npm install pdfjs-dist
```

### 쓰기 엔진 (Python)

```bash
pip install lxml
```

---

## 빠른 시작

### 문서 읽기

```bash
npx kordoc 문서.hwpx
npx kordoc 보고서.hwp
npx kordoc 공문.pdf -o 공문.md
```

### HWPX 생성

```bash
python3 scripts/build_hwpx.py --template report --output 결과.hwpx
```

### 문서 비교

```typescript
import { compare } from "kordoc"
const diff = await compare(구버전Buffer, 신버전Buffer)
console.log(diff.stats) // { added, removed, modified, unchanged }
```

---

## 요구 사항

- Node.js 18+
- Python 3.8+
- lxml (`pip install lxml`)

---

## 원본 출처 및 감사의 글

hanguldoc은 두 오픈소스 프로젝트의 소스 코드와 설계를 기반으로 만들어졌습니다.

### kordoc

- **제작자**: [chrisryugj](https://github.com/chrisryugj)
- **저장소**: [https://github.com/chrisryugj/kordoc](https://github.com/chrisryugj/kordoc)
- **npm**: [https://www.npmjs.com/package/kordoc](https://www.npmjs.com/package/kordoc)
- **라이선스**: MIT
- **기여한 부분**: HWP 5.x · HWPX · PDF 파싱 엔진, 문서 비교(Diff), 양식 인식, MCP 서버 설계
- **제작 배경**: 대한민국 지방공무원(광진구청)으로 7년간 HWP 파일과 씨름하며 직접 만든 실전 파싱 라이브러리

> kordoc의 소스 코드는 MIT 라이선스 하에 제공됩니다.  
> Copyright (c) 2026 chrisryugj

### jun-hwpx2hwpx

- **제작자**: jun (Claude skill 제작자)
- **기여한 부분**: XML-first HWPX 생성 워크플로우, 레퍼런스 복원 엔진, 쪽수 가드(page_guard), 템플릿 시스템(공문·보고서·회의록·제안서), Python 스크립트 전반(`build_hwpx.py`, `analyze_template.py`, `validate.py` 등)

### hanguldoc 통합 작업

위 두 프로젝트의 강점을 하나의 Claude 스킬로 통합한 작업은 [sonhyosang-ops](https://github.com/sonhyosang-ops)이 수행했습니다.

- kordoc → **읽기(파싱) 엔진** 담당
- jun-hwpx2hwpx → **쓰기(생성·편집) 엔진** 담당
- 두 엔진을 작업 유형에 따라 자동 라우팅하는 통합 워크플로우(WF-0~8) 설계

---



## 라이선스

MIT

원본 프로젝트(kordoc)의 MIT 라이선스 조건에 따라, 이 저장소의 소스 코드에도 동일하게 MIT 라이선스가 적용됩니다.
