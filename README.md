# Paper Translator

영어 학술 논문·책 장(PDF / DOCX / MD)을 Anthropic Claude API로 한국어로 옮기는 Windows용 번역 도구입니다.
메뉴형 배치 파일(`Translator.bat`)로 실행합니다.

## 주요 기능
- **본문 전체 번역**: 요약 없이 단락 단위로 옮기며, 원문보다 번역이 지나치게 짧은 구간(요약·누락 의심)은 자동으로 한 번 더 번역합니다.
- **스캔 PDF 지원**: 텍스트 레이어가 없는 쪽은 페이지 이미지를 모델에 보내 번역합니다. 일부 쪽만 스캔인 혼합 PDF도 쪽별로 판정합니다.
- **원서 쪽 표시**: 본문에 `[원서 p. N]`을 넣습니다. 인쇄된 쪽번호를 자동으로 찾고, 실패하면 첫 쪽 번호를 직접 지정할 수 있습니다.
- **각주 처리 3가지**: 모두 생략 / 논평성 각주만 번역(서지 정보는 영문 유지) / 전체 번역. 번역한 각주는 Word 각주로 들어가며, 장 끝 미주(Notes)도 본문 위치에 연결합니다.
- **그림·표·수식**: 원본에서 잘라 해당 위치에 이미지로 넣습니다.
- **용어 일관성**: 번역 전에 용어집을 자동으로 만들고, 직접 만든 사용자 용어집을 우선 적용합니다.
- **출력**: DOCX / PDF / DOCX+PDF / Markdown. 파일명은 `<원본>_pp<시작>-<끝>_한국어번역.<확장자>`입니다.
- **자동 점검(QA)**: 쪽 표시 누락, 각주 짝, 쪽별 번역 분량을 검사해 `last_qa.txt`로 남깁니다.
- **배치 모드**: Anthropic Batch API로 제출해 비용을 50% 줄입니다. 결과는 나중에 받습니다.

## 준비물
| 항목 | 내용 |
|---|---|
| Windows 10/11 | 배치 파일과 Word 연동이 Windows 기준입니다. |
| Python 3.10 이상 | [python.org](https://www.python.org/downloads/)에서 설치. 설치 첫 화면에서 **"Add python.exe to PATH"를 반드시 체크**하세요. |
| Anthropic API 키 | [console.anthropic.com](https://console.anthropic.com)에서 발급하고 크레딧을 충전해야 합니다. 비용은 키 주인에게 청구됩니다. |
| Microsoft Word (선택) | 있으면 PDF를 Word로 변환해 각주가 페이지 하단에 들어갑니다. 없으면 LibreOffice, 그것도 없으면 내장 렌더러를 쓰며 이때 각주는 문서 끝 "원주"에 모입니다. |

## 설치
1. 이 저장소를 받습니다. Git이 없다면 위쪽 **Code → Download ZIP** 후 원하는 폴더에 압축을 풉니다.
2. `Translator.bat`을 더블클릭합니다. 처음 실행할 때 필요한 파이썬 패키지(`requirements.txt`)를 자동으로 설치합니다.
   - 자동 설치가 실패하면 명령 프롬프트에서 직접 실행하세요: `python -m pip install -r requirements.txt`
3. 메뉴에서 **4 Options → 8 API key → 1**을 골라 API 키를 붙여 넣습니다.
   - 키는 같은 폴더의 `api_key.cfg`에 평문으로 저장됩니다. 이 파일은 남에게 주지 마세요.
   - 환경 변수 `ANTHROPIC_API_KEY`가 있으면 그것을 우선 사용합니다.

## 사용법
메인 메뉴에서 번호를 고르고, 번역할 파일 경로를 붙여 넣습니다(탐색기에서 파일을 창으로 끌어다 놓아도 됩니다). 결과는 **원본 파일과 같은 폴더**에 저장됩니다.

| 메뉴 | 설명 |
|---|---|
| 1 Translate now | 바로 번역합니다. 진행 상황이 창에 표시됩니다. |
| 2 Batch submit | 배치로 제출합니다(50% 할인). 제출 후 컴퓨터를 꺼도 됩니다. |
| 3 Batch download | 제출한 배치의 상태를 확인하고, 끝났으면 결과를 받아 저장합니다. 결과는 제출 후 약 29일간 다시 받을 수 있습니다. |
| 4 Options | 아래 설정 |

### 설정(Options)
| 번호 | 항목 | 설명 |
|---|---|---|
| 1 | Model | 기본값은 `claude-sonnet-5-5`. 품질을 높이려면 `claude-opus-5-5`, 최고 품질은 `claude-fable-5-1`. |
| 2 | Thinking effort | 번역에는 `low`면 충분합니다. 높일수록 비용이 늘어납니다. |
| 3 | Footnotes | `strip`(생략) / `commentary`(논평성만) / `all`(전체) |
| 4 | Page markers | 쪽 표시 켜기·끄기, 첫 쪽 번호 수동 지정 |
| 5 | Extract options | 참고문헌 이후 자르기, 그림·표 추출 |
| 6 | Auto glossary | 번역 전 용어집 자동 추출(API 1회 추가) |
| 7 | Output format | DOCX / PDF / DOCX+PDF / MD |
| 8 | API key | 키 입력·삭제 |
| 9 | User glossary | 사용자 용어집 선택·새로 만들기·편집·끄기 |

설정은 `extract.cfg`와 `model.cfg`에 저장됩니다. 처음에는 파일이 없어도 기본값으로 동작하며, 견본은 `extract.cfg.example`입니다.

### 사용자 용어집
- **9 User glossary → n**으로 새 용어집을 만들면 `glossaries\` 폴더에 견본(`glossary_example.txt`) 형식의 파일이 생기고 메모장이 열립니다.
- 한 줄에 하나씩 `English term = 한국어`로 적습니다. `#`으로 시작하는 줄은 주석입니다.
- 수업·주제별로 파일을 나눠 두고 번역 전에 골라 쓰는 것을 권합니다. 군사용어 용어집을 다른 분야 글에 적용하면 오히려 오역이 생길 수 있습니다.
- 파일 이름은 영문으로 지으세요(메뉴 화면이 한글 파일명을 제대로 표시하지 못할 수 있습니다).

## 비용 (참고)
2026년 10월 기준 가격으로 추정한 값입니다. 가격은 바뀔 수 있으니 [공식 가격표](https://platform.claude.com/docs/en/about-claude/pricing)를 확인하세요.

| 모델 | 입력 / 출력 (100만 토큰당) | 30쪽 논문(약 2.5만 단어) |
|---|---|---|
| claude-sonnet-5-5 (기본) | $2 / $10 | 약 $0.6 |
| claude-opus-5-5 | $4 / $20 | 약 $1.2 |
| claude-fable-5-1 | $10 / $50 | 약 $3 |

- 배치 모드는 위 금액의 절반입니다. 실측 예: 128쪽 책 장(스캔 79쪽 포함, Sonnet, 배치) 약 $1.3.
- 스캔 쪽은 텍스트 쪽보다 입력 토큰이 많이 듭니다.
- 실행이 끝나면 비용 요약이 창에 표시되고 `last_cost.txt`에도 남습니다.

## 결과 확인과 문제 해결
- 번역이 끝나면 **QA 점검** 결과가 표시됩니다. "누락/요약 의심" 쪽이 있으면 원문과 대조해 보세요. "번역문이 없는 쪽"은 장 끝 미주나 그림만 있는 쪽이면 정상입니다.
- **결과 파일이 다른 프로그램에 열려 있으면** 덮어쓰지 않고 이름 뒤에 날짜·시각을 붙여 새로 저장합니다.
- **`python`을 찾을 수 없다**는 오류: Python 설치 시 PATH 옵션을 빠뜨린 경우입니다. Python을 다시 설치하며 "Add python.exe to PATH"를 체크하세요.
- **PDF 각주가 문서 끝에 모였다**: Word(또는 LibreOffice)를 찾지 못해 내장 렌더러를 쓴 경우입니다. DOCX 출력은 항상 실제 Word 각주로 들어갑니다.
- **쪽 표시가 맞지 않는다**: Options → 4에서 원서 첫 쪽 번호를 직접 지정하세요.
- 마지막 실행 기록은 `last_translate.log`(실시간), `last_submit.log` / `last_download.log`(배치)에 남습니다. 문제를 문의할 때 이 파일을 함께 보내 주세요(API 키는 들어 있지 않습니다).

## 이용 시 유의사항
- 번역할 문서의 내용은 Anthropic API로 전송됩니다. 외부로 보내면 안 되는 문서에는 쓰지 마세요.
- 번역 결과물은 **개인 학습·연구용**으로만 쓰고 재배포하지 마세요. 원서의 저작권은 저자와 출판사에 있습니다.
- 기계 번역이므로 인용·발표 전에는 반드시 원문과 대조하세요.

## 파일 구성
| 파일 | 역할 |
|---|---|
| `Translator.bat` | 메뉴(실행·설정) |
| `translate_paper.py` | 실시간 번역 |
| `batch_submit.py` / `batch_download.py` | 배치 제출 / 결과 수집 |
| `pdf_extract.py` | PDF 추출(단락·쪽번호·각주·미주·그림, 스캔 쪽 판정) |
| `paper_utils.py` | 설정·프롬프트·분할·용어집·QA·저장 |
| `doc_render.py` | DOCX 서식·Word 각주, DOCX→PDF 변환 |
| `glossary_example.txt` | 사용자 용어집 견본 |
| `extract.cfg.example` | 설정 파일 견본 |

`api_key.cfg`, `extract.cfg`, `model.cfg`, `glossaries/`, 실행 로그는 개인 파일이라 `.gitignore`로 저장소에서 제외됩니다.

## 라이선스
MIT License — `LICENSE` 파일을 참고하세요.
