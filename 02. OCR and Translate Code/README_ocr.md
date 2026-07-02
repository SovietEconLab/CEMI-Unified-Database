# CEMI Archival PDF Processor

OCR과 구조화를 통해 ЦЭМИ АН СССР 인사 기록 (АРАН Ф.1959) 스캔 PDF를 다중 시트 Excel 워크북으로 변환합니다.

이전에 수동으로 만든 워크북과 같은 품질을 목표로 합니다 — `_Summary`, `_CrossReferences`, 리스트별 시트, 5색 강조, ISO 날짜 변환, 역사적 발견 등.

## 작동 방식

1. **파일 검색**: 지정된 디렉토리(생략 시 현재 디렉토리)에서 파일명에 `1959`가 포함된 모든 PDF를 찾습니다.
2. **사용자 선택**: 메뉴에서 처리할 PDF 하나를 선택합니다.
3. **API 키 입력**: 한 번 입력 (PDF 한 개당 한 번 — 다음 PDF 처리 시 다시 실행해야 함).
4. **2단계 처리**:
   - 1단계: 각 페이지를 Claude vision API로 OCR (러시아어 원문 보존, 손글씨/줄긋기 표시)
   - 2단계: 전체 OCR 텍스트를 Claude로 구조화 (도메인 컨텍스트 임베드)
5. **Excel 생성**: 같은 디렉토리에 `cemi_<delo>_<year>_full.xlsx` 저장
6. **종료**: 다음 PDF는 스크립트를 다시 실행

## 설치

```bash
pip install -r requirements.txt
```

요구 사항:
- Python 3.10+
- `anthropic` (Claude API 클라이언트)
- `PyMuPDF` (PDF 래스터화 — 외부 의존성 없는 순수 Python)
- `openpyxl` (Excel 빌드)
- `Pillow` (이미지 처리)

## 실행

### 기본 사용

```bash
# 현재 디렉토리 스캔
python cemi_processor.py

# 특정 디렉토리 지정
python cemi_processor.py /path/to/archive

# 다른 디렉토리 (Windows 예시)
python cemi_processor.py "C:\Users\me\CEMI\PDFs"
```

### 옵션

```bash
python cemi_processor.py --help

  --model MODEL    Anthropic 모델 ID (기본: claude-opus-4-5)
  --dpi DPI        래스터화 해상도 (기본: 150)
  --force          캐시된 OCR 무시하고 모든 페이지 재OCR
```

### 워크플로우 예시

```
$ python cemi_processor.py ~/cemi_archive

PDFs found (filename contains '1959'):
  [ 1] 1959_1_10.pdf  (13796 KB)
  [ 2] 1959_1_19.pdf  (2695 KB)
  [ 3] 1959_1_30.pdf  (1377 KB)
  ...
  [ 0] cancel

Select a PDF by number: 2

Use ANTHROPIC_API_KEY from environment? [Y/n] n
Enter ANTHROPIC_API_KEY (input hidden): ********

=== Processing: 1959_1_19.pdf ===
  rasterising at 150 DPI …
  9 pages.
  page   1: OCR …
  page   2: OCR …
  ...
  page   9: OCR …
  structuring 9 pages (year hint: 1962) …
  writing /home/user/cemi_archive/cemi_1959_1_19_1962_full.xlsx …
  done.

✓ Saved: /home/user/cemi_archive/cemi_1959_1_19_1962_full.xlsx
  Re-run the script (without arguments) to process another PDF.
```

## API 키 관리

세 가지 방식 지원:

1. **매 실행 시 직접 입력 (기본)**: 입력은 화면에 보이지 않음 (`getpass`)
2. **환경 변수**: `ANTHROPIC_API_KEY=sk-ant-...` 설정 시 사용 여부 확인 후 진행
3. **`.env` 파일**: 환경 변수에 직접 export 권장

요청된 동작 그대로 — **PDF 한 개 처리 시 한 번** 키 요구.

## 캐시

각 PDF의 OCR 결과는 다음 경로에 저장됩니다:

```
<PDF 디렉토리>/.cemi_cache/<PDF 파일명>/
  ├── images/        # 래스터화된 페이지 (page_001.jpg, ...)
  ├── ocr/           # OCR 결과 텍스트 (page_001.txt, ...)
  └── structured.json # 최종 구조화 데이터
```

이를 통해:
- 부분 실패 시 재실행하면 캐시된 페이지는 건너뛰고 이어서 처리
- `--force` 플래그로 강제 재OCR 가능
- 원하면 OCR 텍스트 파일을 수동 편집 후 다시 빌드 가능 (구조화 단계는 캐시 안 됨)

## 출력 Excel 구조

수동으로 만들었던 워크북과 동일한 품질 목표:

| 시트 | 내용 |
|---|---|
| `_Summary` | 표지·메타데이터·시트 목록·핵심 발견 (8+ 단락) |
| `_CrossReferences` | 핵심 인물 cross-trail (★/★★/★★★ 표시) |
| `List_*`, `Roster_*`, `Form_*` 등 | 각 리스트의 행 데이터 + ISO 날짜 변환 |

**색상 코드**:
- 🟩 녹색 — 소장 (Nemchinov, Federenko 등)
- 🟦 파랑 — 박사·후보 학위 보유자
- 🟪 분홍 — КПСС 회원
- 🟨 노랑 — 핵심 인물 (Suvorov, Shatalin 등)
- 🟧 오렌지 — 손글씨 추가 항목
- ⬜ 회색 — 줄긋기/취소 항목

## 도메인 컨텍스트

스크립트 내부 (`CEMI_DOMAIN_CONTEXT` 상수)에 다음 임베드:

- ЦЭМИ 기관사 (1958-1987)
- 핵심 인물 명단 (Nemchinov, Federenko, Makarov, Shatalin, Katsenelinboigen, Lurie, Mints, Vainstein, Vishnev, Suvorov, Ovsienko, Faerman, Mednitsky, Volkonsky, Gavrilets, Dadayan, Kossov, Glaziev, Polterovich 등)
- 양식 코드 시대별 (10-НР / 10-НПР / 5-нк)
- 리스트 번호 체계 시대별 (1962 / 1963 / 1965+)
- 날짜 형식 규칙 (로마숫자 월, УП=VII 타자기 치환, 슬래시/점/하이픈 변형)

## 알려진 한계

1. **수기 텍스트의 정확도**: Claude vision은 인쇄체에서 매우 정확하지만, 1962-1965년 손글씨 키릴체는 일부 [?] 표시로 남을 수 있음. 출력 후 수동 교정 권장.

2. **줄긋기 복원 불가**: X-마크로 완전히 가려진 텍스트는 복구 불가 (예: д.48 page 1 entry #1).

3. **시리즈 외 패킷**: 도메인 컨텍스트는 ЦЭМИ에 특화됨. 다른 기관 문서에는 적합하지 않음.

4. **비용**: 24페이지 PDF 한 개 처리 시 대략 OCR 24회 + 구조화 1회 = 25회 API 호출. Opus 4.5 기준 약 $3-8 추정.

5. **출력 토큰 한계**: 구조화 단계가 16K 토큰으로 제한됨. 매우 큰 패킷(300+ 항목)은 출력이 잘릴 가능성 — 이 경우 별도 처리 필요.

6. **JSON 파싱**: 모델이 가끔 JSON 외 텍스트를 추가하면 자동 정리 시도. 실패 시 `cemi_structure_raw.txt`에 원본 응답 저장.

## 트러블슈팅

### "Failed to parse structuring JSON"
모델이 잘못된 JSON 형식으로 응답한 경우. `cemi_structure_raw.txt` 파일을 열어 수동으로 JSON 부분만 추출하거나, `--force` 없이 재실행하면 캐시된 OCR을 사용해 구조화만 다시 시도.

### "PDF가 ZIP 아카이브"
1980년대 일부 패킷이 ZIP으로 래핑된 PDF임. PyMuPDF가 자동 처리.

### API 호출 실패
지수 백오프로 3회 자동 재시도. 그래도 실패하면 캐시된 페이지는 보존되므로 다시 실행 시 이어서 처리.

### "No PDFs containing '1959' in filename found"
디렉토리에 파일명에 `1959`가 들어가는 PDF가 없음. 다른 디렉토리 지정 필요.

## 출력 구조 커스터마이징

`CEMI_DOMAIN_CONTEXT` 문자열을 편집해 도메인 지식을 추가/변경하면 다른 종류의 러시아어 아카이브에도 적용 가능. 핵심 인물 명단·양식 코드·리스트 번호·날짜 형식 규칙을 갱신하세요.
