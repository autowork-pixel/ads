# Ads Workbench

브라우저에서 메타 DA, 네이버 SA, 쿠팡 SA 원본을 읽어 주간 광고 성과 HTML 보고서를 만듭니다. 원본은 서버에 업로드되지 않습니다.

## 사용

1. 배포 주소 `https://autowork-pixel.github.io/ads/`를 엽니다.
2. 고객명과 매체를 선택하고 금주 원본을 넣습니다. 전주 비교가 필요하면 전주 원본도 넣습니다.
3. **보고서 생성**을 누릅니다. 화면에서 확인하거나 HTML 파일로 내려받습니다.

메타 원본 하나에 여러 보고 기간이 있다면 날짜 입력란에 금주·전주 시작일과 종료일을 지정하세요. 전주 원본이 없으면 비교 수치를 만들지 않습니다.

지원 원본: `.xlsx`, `.xlsb`, `.csv`. Excel 파싱에는 [SheetJS Community Edition](https://docs.sheetjs.com/docs/getting-started/installation/standalone/) 0.20.3 브라우저 빌드를 사용합니다. 첫 접속 시 CDN 연결이 필요합니다.

## 지표 기준

- ROAS = 전환매출 ÷ 광고비, CTR = 클릭 ÷ 노출, CPC = 광고비 ÷ 클릭, CPA = 광고비 ÷ 전환수.
- 메타: 공유 항목 구매와 구매 전환값. 쿠팡: 14일 주문·전환매출. 네이버: 원본의 전환수·전환매출.
- 개선사항은 지표 급변에 따른 검토 제안입니다. 고객 목표 ROAS, 상품 마진, 계정 설정을 반영한 자동 예산 집행은 하지 않습니다.

## 배포

`main` 브랜치의 루트에서 GitHub Pages로 발행합니다. 저장소 **Settings → Pages → Build and deployment → Source: Deploy from a branch**, Branch `main`, Folder `/ (root)`를 선택하면 됩니다. 이후 `main`에 변경이 들어가면 자동 반영됩니다.

