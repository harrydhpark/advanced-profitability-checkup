# LGE Europe TV 선행수익성 분석 대시보드 (Advance Profitability Dashboard)

본 프로젝트는 LGE Europe TV 선행수익성 점검 데이터(`선행점검양식 0722 취합 2.xlsx`)의 원천 시트(`통합DB`, `실적`, `이동`, `판가인하`, `전략제품`, `EU TTL`)를 기반으로 엑셀 연산 수식을 100% JavaScript로 재구현하여 동적 법인별 수익성 분석 웹 시스템을 제공합니다.

---

## 📌 주요 특징 및 핵심 개발 지침

### 1. 엑셀 수식 100% JS 완전 재구현 엔진 (`build_from_source_sheets.js`)
- `SheetJS` 라이브러리의 엑셀 저장 시점 캐시(독일 `DG` 고정) 한계를 극복하기 위해 RAW 시트에서 SUMIFS 및 연산식을 직접 수행합니다.
- **하이브리드 데이터 매핑 규칙**:
  - **상반기 (H1, 1~6월)**: `실적` 시트의 2026년 1~6월 실적 Raw 데이터 집계
  - **하반기 (H2, 7~12월)**: `통합DB` 시트의 2026년 7~12월 사업계획 Raw 데이터 집계
  - **전년 (2025년)**: `실적` 시트의 2025년 1~12월 실적 Raw 데이터 집계
  - **이동연간**: `이동` 시트 데이터 (상반기 이동 = 상반기 올해 수량 동일 매핑 적용)

### 2. 주요 지표 연산 공식
- **수량 / G.ASP / 제상품**: 이동비 = `올해 값 - 이동 값` (연간 수량/G.ASP/제상품 이동비 1:1 완벽 일치)
- **차감율 (%)**: `1 - (제상품 / 총매출)`, 차감 이동비 = `(올해 차감 - 이동 차감) * 100` (%p 단위)
- **판가인하 (%)**: `판가인하` 시트 VLOOKUP 연동 (`Col AI` 올해, `Col A` 전년, `Col X` 이동), 연간 판가인하 = `상반기(-1.17%) + 하반기(-15.83%)` = `-17.0%`, 전년비 = `23.0p`
- **COI / COI(%)**: 시산 매출액 대비 COI 비율 연산 및 전년비 소수점 한자리 표기
- **MP(%)**: 시산 매출액 대비 MP 비율 및 `전년비` 서브 행 연산

### 3. UI/UX 디자인 및 대시보드 렌더링 규칙 (`index.html`)
- **사이드바 메인 모듈 구성**:
  1. **Executive Overview (탭 1)**: 유럽 전체(`EU TTL`) 최상단 포함 및 지표별 선행 점검표 표출
  2. **법인별 수익성 분석 (탭 2)**: 법인 드롭다운 알파벳 순서 정렬(AG, BN, CK, DG, ES, FS, HS, IS, LA, MK, PL, PT, RO, SW, UK, Swiss), PnL Summary 점검표 및 BLU 세그먼트 분석
  3. **전략제품 판매 연간 예상 (탭 3)**: OLED (Rows 4~20), MRGB (Rows 24~40), QNED (Rows 44~60) 3개 카드의 연간 예상 판매 분석
- **레이아웃 & 타이틀 스타일링**:
  - 상단 컨트롤 헤더 높이(`h-14`) 및 배너 카드 패딩(`px-4 py-2.5`) 50% 슬림화
  - 모든 점검표 제목 앞 **검은색 세워진 바 포인트 아이콘 (`w-1 h-4 bg-slate-900 rounded-sm`)** 일괄 적용
  - 불필요한 부제목 영문 구문(`(PnL Table 1)`, `(BLU Table 2)`) 삭제 및 깔끔한 한국어 타이틀 적용
  - BP/SG비 100% 미만 분홍색 하이라이트(`bg-rose-100`), BP/SG비(Q) 미달 항목 노란색 하이라이트(`bg-amber-300`), ↓/Δ 마이너스 항목 빨간색 강조

---

## 🛠 실행 및 배포 명령어

```powershell
# 1. 데이터 집계 및 파싱 실행 (RAW 시트 재계산 & advance_profitability_data.json / .js 생성)
node build_from_source_sheets.js

# 2. 로컬 웹 서버 실행
node server.js

# 3. Firebase Hosting 운영 배포
npx --yes firebase-tools deploy --only hosting --project lge-advance-profitability-2026
```

---

## 🌐 라이브 웹 배포 서비스 URL

- **메인 파이어베이스 웹 호스팅 URL**: **https://lge-advance-profitability-2026.web.app**
- **보조 도메인 URL**: **https://lge-advance-profitability-2026.firebaseapp.com**
- **Firebase 관리 콘솔**: https://console.firebase.google.com/project/lge-advance-profitability-2026/overview
