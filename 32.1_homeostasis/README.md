# 32.1 Homeostasis — 발표 자료

일반생물학 32장 1절 **"Homeostasis: Why and How Do Animals Regulate Their Internal Environment? (항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?)"** 발표용 슬라이드입니다.
영어 슬라이드에 교재(한국어판) 용어를 병기했습니다. 두 가지 버전이 있습니다.

| 파일 | 내용 |
|---|---|
| [`32.1_Homeostasis_short.pptx`](32.1_Homeostasis_short.pptx) | **발표용 짧은 버전 (15장).** 교재의 필수 내용만 담았고, 슬라이드 노트에 그대로 읽을 수 있는 발표 대본이 들어 있습니다. |
| [`32.1_Homeostasis_short_script.md`](32.1_Homeostasis_short_script.md) | 짧은 버전의 발표 대본만 모은 파일 (연습·출력용, 그대로 읽으면 약 12분) |
| [`32.1_Homeostasis_full.pptx`](32.1_Homeostasis_full.pptx) | 전체 버전 (29장). 사이드 박스, 퀴즈, 토론 슬라이드까지 포함한 참고용입니다. 노트는 이전 형식(시간·참고 메모 포함)입니다. |

글꼴: 제목 Cambria, 본문 Calibri, 한글 맑은 고딕 (Windows·Office 기본 글꼴)

## 짧은 버전 구성 (15장)

| 구간 | 슬라이드 | 내용 |
|---|---|---|
| 도입 | 1–2 | 표지, Case Study "Overheated(과도한 열)" + 이 절에서 답할 질문 4가지 |
| Part 1 · What & Why | 3–4 | 항상성의 정의(Bernard, Cannon, 조절 대상, 동적 평형), 효소와 항상성 |
| Part 2 · Body Temperature | 5–6 | 외온동물 vs 내온동물(온혈·냉혈이 오해인 이유), 동물들의 체온 조절 방식 |
| Part 3 · Negative Feedback | 7–9 | 피드백 두 종류와 음성피드백의 3요소, **그림 32-2 재구성 도식**, 시상하부의 체온 조절 |
| Part 4 · Positive Feedback | 10–13 | 열사병의 악순환, 양성피드백과 그것이 드문 이유, 출산, 북극의 양성피드백 |
| 마무리 | 14–15 | 핵심개념 요약, 참고 자료·이미지 출처 |

전체 버전에서 뺀 것: 목차, 열사병 후 장기 손상, Have You Ever Wondered(과량의 물), Case Study Revisited, OpenStax 음성피드백 그림(1.10), 퀴즈·정답, 토론 질문. 여러 장에 걸쳐 있던 정의·체온 조절·음성피드백·북극 내용은 각각 한 장으로 합쳤습니다.

## 발표 전 확인할 것

- **출산 그림(짧은 버전 12번)**: OpenStax 그림은 "뇌하수체(pituitary)가 옥시토신 분비"라고 되어 있고, 우리 교재는 "시상하부가 옥시토신 방출을 일으킨다"고 씁니다. 대본에서 두 내용을 이어서 설명하도록 써 두었습니다.
- **한국어판 번역 주의**: 영어 원문 "The body rises to the occasion with increased respiratory rate…"(운동 예시)를 한국어판은 "…몸의 온도는 상승한다"로 옮겼습니다. 원문 뜻은 "몸이 운동의 요구에 대응한다"이고, 대본은 원문 뜻대로 썼습니다.
- 사진과 OpenStax 그림에는 출처 표기가 필요합니다(슬라이드와 마지막 장에 들어 있음). 지우지 마세요.

## 참고 자료 및 이미지 출처

- 주 교재: Audesirk, Audesirk & Byers, *Biology: Life on Earth*, Ch. 32 (영문판 pp. 604–609, 617–618 / 한국어판 pp. 602–607, 615–616)
- 그림: OpenStax *Anatomy and Physiology 2e* Fig. 1.10, 1.11 · OpenStax *Biology 2e* Fig. 33.23 — CC BY 4.0
- 사진 (Wikimedia Commons)
  - 사막 pupfish — P. V. Loiselle, CC BY-SA 3.0
  - 벌새(cinnamon hummingbird) — Charles J. Sharp, CC BY-SA 4.0
  - 이구아나 — Cayambe, CC BY-SA 3.0
  - 북극곰 — Andreas Weith, CC BY-SA 4.0
  - Claude Bernard — 퍼블릭 도메인 / Walter B. Cannon — Bachrach 1934, CC BY 4.0
- 도식: 음성피드백 고리(그림 32-2 재구성), 열사병 악순환, 북극 얼음 고리, 개념 그래프는 교재 본문을 바탕으로 새로 그렸습니다. 모두 PowerPoint 도형이라 직접 수정할 수 있습니다.

## 수정하는 방법

- **간단한 수정**: PowerPoint에서 `.pptx`를 바로 고치면 됩니다.
- **스크립트로 다시 만들기**: [`source/`](source) 폴더에 생성 스크립트와 이미지가 있습니다.
  - 슬라이드 디자인: `source/build.js`
  - 발표 대본: `source/notes_short.js`(짧은 버전), `source/notes_full.js`(전체 버전)

  Node.js 18 이상에서 아래 명령을 실행하면 두 버전과 대본 파일이 다시 만들어집니다. (PowerPoint에서 직접 고친 내용은 덮어쓰이니 주의하세요.)

  ```bash
  cd 32.1_homeostasis/source
  npm install
  npm run build
  ```
