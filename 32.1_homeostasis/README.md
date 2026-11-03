# 32.1 Homeostasis — 발표 자료

일반생물학 32장 1절 **"Homeostasis: Why and How Do Animals Regulate Their Internal Environment? (항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?)"** 발표용 슬라이드입니다.
영어 슬라이드에 교재(한국어판)의 용어를 병기했고, 모든 슬라이드에 한국어 발표 대본(슬라이드 노트)이 들어 있습니다.

- 발표 파일: [`32.1_Homeostasis.pptx`](32.1_Homeostasis.pptx) (16:9, 29장)
- 글꼴: 제목 Cambria, 본문 Calibri, 한글 맑은 고딕 (Windows·Office 기본 글꼴)

## 슬라이드 구성

| 구간 | 슬라이드 | 내용 | 예상 시간 |
|---|---|---|---|
| 도입 | 1–3 | 표지, Case Study "Overheated(과도한 열)", 목차 | 약 3.5분 |
| Part 1 · What & Why | 4–8 | 내부 불변성 → 항상성, 정의, 동적 평형, 효소와 항상성, 열사병 후 장기 손상 | 약 8분 |
| Part 2 · Body Temperature | 9–13 | 온혈/냉혈의 오해, 외온동물 vs 내온동물, 외온동물의 체온 변화와 조절, 내온동물의 이점과 비용 | 약 7.5분 |
| Part 3 · Negative Feedback | 14–18 | 피드백의 두 종류, 음성피드백의 3요소, **그림 32-2 재구성 도식**, 시상하부의 체온 조절, Have You Ever Wondered(과량의 물) | 약 9.5분 |
| Part 4 · Positive Feedback | 19–23 | 열사병의 악순환, 양성피드백과 그것이 드문 이유, 출산, Earth Watch(북극의 양성피드백) | 약 8.5분 |
| 마무리 | 24–29 | Case Study Revisited, 핵심개념 요약, 복습 퀴즈와 정답, 토론 질문, 출처 | 약 9–11분 |

전체는 토론을 포함해 약 45분입니다. 35분 정도로 줄이려면 18(과량의 물), 24(Case Study Revisited), 26–27(퀴즈), 28(토론)을 빼거나 짧게 하면 됩니다.

## 발표 전 확인할 것

- **퀴즈 정답(27번 슬라이드)**은 교재 본문을 근거로 정리했습니다. 교재 뒤 Answers와 한 번 대조해 주세요.
- **출산 그림(21번, OpenStax)**에는 "뇌하수체(pituitary)가 옥시토신을 분비"한다고 되어 있습니다. 우리 교재는 "시상하부가 옥시토신 방출을 일으킨다"고 서술합니다. 노트에 설명을 넣어 두었습니다.
- **한국어판 번역 주의(6번)**: 영어 원문 "The body rises to the occasion with increased respiratory rate…"를 한국어판은 "…몸의 온도는 상승한다"로 옮겼습니다. 원문 뜻은 "몸이 (운동의 요구에) 대응한다"입니다.
- 사진과 OpenStax 그림에는 출처 표기가 필요합니다(이미 슬라이드와 29번 슬라이드에 들어 있음). 지우지 마세요.

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

- **간단한 수정**: PowerPoint에서 `32.1_Homeostasis.pptx`를 바로 고치면 됩니다.
- **스크립트로 다시 만들기**: [`source/`](source) 폴더에 생성 스크립트와 이미지가 있습니다. Node.js 18 이상에서

  ```bash
  cd 32.1_homeostasis/source
  npm install
  npm run build
  ```

  을 실행하면 `32.1_Homeostasis.pptx`가 다시 만들어집니다. (PowerPoint에서 직접 고친 내용은 덮어쓰이니 주의하세요.)
