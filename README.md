# Paper Translator

영어 학술 논문·책 장(PDF/DOCX/MD)을 Claude API로 한국어로 옮기는 개인용 번역기 (Windows).

## 기능
- 텍스트 PDF: 단락 복원·페이지 경계 병합, 도식/표/수식 이미지 클립 삽입
- 스캔 PDF·혼합 PDF: 텍스트 레이어 없는 쪽은 페이지 이미지로 비전 번역
- `[원서 p. N]` 쪽 마커(인쇄 쪽번호 자동 감지), 각주 3모드(strip / commentary / all), 장 끝 미주 → Word 각주
- 용어집 자동 추출 + 사용자 용어집(`glossaries/*.txt`, Options > 9에서 선택·생성·편집) 우선 적용
- DOCX / PDF(Word 변환) / MD 출력, 실시간·배치(50% 할인) 모드, 자동 QA(`last_qa.txt`)

## 설치
```
python -m pip install -r requirements.txt
copy extract.cfg.example extract.cfg
```
사용자 용어집은 `glossary_example.txt` 형식으로 `glossaries/` 폴더에 만듭니다.
`Translator.bat` 실행 → Options → API key 에서 키 입력 (`api_key.cfg`, git 제외).

## 파일
| 파일 | 역할 |
|---|---|
| `Translator.bat` | 메뉴(실행·옵션) |
| `translate_paper.py` | 실시간 번역 |
| `batch_submit.py` / `batch_download.py` | 배치 제출 / 결과 수집 |
| `pdf_extract.py` | PDF 추출(단락·쪽번호·각주·미주·도식, 혼합 PDF 판정) |
| `paper_utils.py` | 설정·프롬프트·청크·QA·저장 |
| `doc_render.py` | DOCX 렌더(서식·Word 각주), DOCX→PDF |

`api_key.cfg`, `extract.cfg`, `model.cfg`, 개인 용어집(`glossaries/`), 실행 로그와 `_local/`(샘플·이전 결과)는 `.gitignore` 로 제외됩니다.
