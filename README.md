# school-project

Cloudflare Pages app for the 26수능 합격예측 분석기 웹 버전.

## What it does

- Enter scores directly in the browser without uploading Excel
- Reproduce the 26수능 analyzer formulas from extracted workbook data
- Recalculate results immediately as inputs change
- Browse 이과 / 문과 모집단위를 검색하고 상태별로 확인

## Current flow

1. Open the deployed web app
2. 국어, 수학, 영어, 한국사, 탐구, 제2외국어 점수를 직접 입력
3. 필요하면 대학별 내신 보정값을 조정
4. 결과 표에서 지원 가능 상태와 기준 점수를 바로 확인

## Important files

- `index.html`: direct-input analyzer UI
- `app.js`: formula parser, evaluator, UI binding, result renderer
- `data/analyzer-26.json`: extracted 26수능 workbook data used at runtime
- `scripts/extract-26-data.js`: extractor that rebuilds `data/analyzer-26.json`
- `wrangler.toml`: Cloudflare Pages configuration

## Local preview

- `npx wrangler pages dev .`

## Notes

- The current direct-input calculator is based on the 26수능 workbook.
- Validation against workbook cached values is effectively matched except for tiny floating-point rounding noise on 2 formulas.
