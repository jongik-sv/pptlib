const pptxgen = require('pptxgenjs');
const html2pptx = require('../.claude/skills/pptx/scripts/html2pptx');
const fs = require('fs');
const path = require('path');

// 색상 (# 없이)
const colors = {
  deepGreen: '22523B',
  darkGreen: '153325',
  midGreen: '183C2B',
  lightGreen: '479374',
  paleGreen: '8BAFA2',
  accentGreen: '6F886A',
  textDark: '333333',
  textGray: '767171',
  white: 'FFFFFF',
  riskHigh: 'c0392b',
  riskMid: 'e67e22',
  riskLow: '27ae60'
};

// 슬라이드 HTML 생성 함수
function createSlideHTML(slideNum, content) {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body {
  width: 720pt; height: 405pt; margin: 0; padding: 0;
  font-family: Arial, sans-serif; background: #FFFFFF;
  display: flex; flex-direction: column;
}
${content.styles || ''}
</style>
</head>
<body>
${content.body}
</body>
</html>`;
}

// 공통 헤더/푸터 스타일
const commonStyles = `
.slide-header { background: #22523B; height: 55pt; display: flex; align-items: center; padding: 0 35pt; }
.slide-header h1 { color: #FFFFFF; font-size: 18pt; }
.page-num { color: #FFFFFF; font-size: 10pt; opacity: 0.8; margin-left: auto; }
.slide-body { padding: 12pt 35pt; flex: 1; }
.slide-footer { display: flex; justify-content: space-between; font-size: 7pt; color: #767171; padding: 8pt 35pt; }
.sub-title { color: #22523B; font-size: 11pt; font-weight: bold; margin-bottom: 8pt; }
.sub-title-wrap { margin-bottom: 8pt; }
.sub-title-bar { height: 2pt; background: #8BAFA2; margin-top: 4pt; }
.two-cols { display: flex; gap: 20pt; }
.col { flex: 1; }
.placeholder { background: #f0f0f0; }
`;

// 슬라이드 1: 타이틀
const slide1 = {
  styles: `
.header-bar { background: #22523B; color: #FFFFFF; padding: 8pt 40pt; border-radius: 25pt; font-size: 11pt; font-weight: bold; margin-top: 80pt; align-self: center; }
.main-title { font-size: 30pt; color: #22523B; font-weight: bold; text-align: center; margin-top: 40pt; line-height: 1.4; }
.sub-info { color: #767171; font-size: 9pt; text-align: center; margin-top: 15pt; line-height: 1.8; }
.presenter { color: #22523B; font-size: 11pt; text-align: center; margin-top: 30pt; }
.divider { width: 500pt; height: 1pt; background: #153325; margin: 15pt auto 0 auto; }
.copyright { color: #333333; font-size: 6pt; text-align: right; margin-top: auto; margin-bottom: 10pt; margin-right: 15pt; }
`,
  body: `
<div class="header-bar"><p>2025년 스마트 물류관리 시스템 구축</p></div>
<div class="divider"></div>
<h1 class="main-title">스마트 물류관리 시스템<br>구축 프로젝트 수행계획서</h1>
<p class="sub-info">발주기관: (주)글로벌물류 | 수행사: (주)테크솔루션<br>계약기간: 2025-01-06 ~ 2025-12-31 (12개월) | 계약금액: 15억원</p>
<div class="divider"></div>
<p class="presenter">경영지원팀 김철수 차장 (PM)</p>
<p class="copyright">DOC-PRJ-2025-001-001 | v1.0</p>
`
};

// 슬라이드 2: 목차
const slide2 = {
  styles: `
.section-title { font-size: 18pt; color: #22523B; font-weight: bold; margin: 60pt 0 0 50pt; }
.contents-list { margin: 20pt 50pt; }
.contents-item { display: flex; align-items: flex-start; margin-bottom: 12pt; }
.contents-num { font-size: 16pt; color: #22523B; font-weight: bold; min-width: 45pt; }
.contents-text { flex: 1; }
.contents-text h3 { font-size: 12pt; color: #22523B; margin-bottom: 3pt; }
.contents-text p { font-size: 8pt; color: #22523B; }
`,
  body: `
<h1 class="section-title">CONTENTS</h1>
<div class="contents-list">
  <div class="contents-item"><div class="contents-num"><p>01</p></div><div class="contents-text"><h3>프로젝트 개요</h3><p>프로젝트 정보, 배경 및 목적, 범위</p></div></div>
  <div class="contents-item"><div class="contents-num"><p>02</p></div><div class="contents-text"><h3>추진 전략 및 방법론</h3><p>추진 전략, 수행 방법론, 기술 아키텍처</p></div></div>
  <div class="contents-item"><div class="contents-num"><p>03</p></div><div class="contents-text"><h3>프로젝트 조직 및 역할</h3><p>조직별 R&R, 투입 인력 현황</p></div></div>
  <div class="contents-item"><div class="contents-num"><p>04</p></div><div class="contents-text"><h3>프로젝트 일정 및 산출물</h3><p>마스터 일정, 마일스톤, 산출물 목록</p></div></div>
  <div class="contents-item"><div class="contents-num"><p>05</p></div><div class="contents-text"><h3>품질/보안/위험 관리</h3><p>품질 보증, 보안 관리, 위험 대응</p></div></div>
  <div class="contents-item"><div class="contents-num"><p>06</p></div><div class="contents-text"><h3>교육 및 의사소통</h3><p>교육 계획, 기술이전, 보고 체계</p></div></div>
</div>
`
};

// 슬라이드 3: 프로젝트 개요
const slide3 = {
  styles: commonStyles + `
.highlight-box { background: #22523B; color: #FFFFFF; padding: 12pt 15pt; border-radius: 6pt; margin-bottom: 12pt; }
.highlight-box h2 { font-size: 13pt; margin-bottom: 5pt; }
.highlight-box p { font-size: 9pt; opacity: 0.95; }
.card { background: #FFFFFF; border: 1pt solid #8BAFA2; border-left: 3pt solid #22523B; padding: 8pt; margin-bottom: 8pt; }
.card ul { padding-left: 12pt; font-size: 8pt; line-height: 1.5; }
.card-green { background: #e8f5e9; border: 1pt solid #8BAFA2; padding: 8pt; }
.card-green ul { padding-left: 12pt; font-size: 8pt; line-height: 1.5; }
`,
  body: `
<div class="slide-header"><h1>1. 프로젝트 개요</h1><p class="page-num">3 / 15</p></div>
<div class="slide-body">
  <div class="highlight-box"><h2>프로젝트 비전</h2><p>AI 기반 스마트 물류관리 시스템 구축을 통한 물류 처리 효율 30% 향상, 재고 정확도 99.5% 달성</p></div>
  <div class="two-cols">
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">프로젝트 정보</h3><div class="sub-title-bar"></div></div>
      <div id="table1" class="placeholder" style="width: 295pt; height: 120pt;"></div>
    </div>
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">현황 및 문제점</h3><div class="sub-title-bar"></div></div>
      <div class="card"><ul><li>10년 이상 운영된 레거시 시스템</li><li>처리 속도 저하로 인한 업무 지연</li><li>실시간 재고 파악 불가</li><li>타 시스템 연동 한계</li></ul></div>
      <div class="sub-title-wrap"><h3 class="sub-title">기대 효과</h3><div class="sub-title-bar"></div></div>
      <div class="card-green"><ul><li>물류 처리 효율 <b>30% 향상</b></li><li>재고 정확도 <b>99.5% 달성</b></li><li>운영 비용 <b>20% 절감</b></li></ul></div>
    </div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    if (t1) {
      slide.addTable([
        [{ text: '프로젝트명', options: { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true } }, '스마트 물류관리 시스템 구축'],
        [{ text: '발주기관', options: { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true } }, '(주)글로벌물류'],
        [{ text: '수행사', options: { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true } }, '(주)테크솔루션'],
        [{ text: '계약금액', options: { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true } }, '15억원 (VAT 별도)'],
        [{ text: '계약기간', options: { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true } }, '2025-01-06 ~ 2025-12-31'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.1, 2.0] });
    }
  }
};

// 슬라이드 4: 프로젝트 범위
const slide4 = {
  styles: commonStyles + `
.slide-body { padding: 10pt 35pt; }
.card-grid-4 { display: flex; gap: 10pt; margin-bottom: 12pt; }
.card { flex: 1; background: #FFFFFF; border: 1pt solid #8BAFA2; border-radius: 5pt; text-align: center; }
.card-header { background: #22523B; color: #FFFFFF; padding: 5pt; border-radius: 4pt 4pt 0 0; font-size: 9pt; font-weight: bold; }
.card h3 { font-size: 9pt; color: #22523B; margin: 8pt 0 3pt; }
.card p { font-size: 7pt; color: #333333; margin-bottom: 8pt; }
`,
  body: `
<div class="slide-header"><h1>1. 프로젝트 범위</h1><p class="page-num">4 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">시스템 범위</h3><div class="sub-title-bar"></div></div>
  <div class="card-grid-4">
    <div class="card"><div class="card-header"><p>WMS</p></div><h3>창고관리시스템</h3><p>입출고, 재고, 피킹/패킹</p></div>
    <div class="card"><div class="card-header"><p>TMS</p></div><h3>배송관리시스템</h3><p>배차, 경로최적화, 추적</p></div>
    <div class="card"><div class="card-header"><p>OMS</p></div><h3>주문관리시스템</h3><p>주문접수, 할당, 정산</p></div>
    <div class="card"><div class="card-header"><p>Dashboard</p></div><h3>통합모니터링</h3><p>실시간 현황, KPI, 리포트</p></div>
  </div>
  <div class="two-cols">
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">포함 범위</h3><div class="sub-title-bar"></div></div>
      <div id="table1" class="placeholder" style="width: 295pt; height: 95pt;"></div>
    </div>
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">제외 범위</h3><div class="sub-title-bar"></div></div>
      <div id="table2" class="placeholder" style="width: 295pt; height: 95pt;"></div>
    </div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const t2 = placeholders.find(p => p.id === 'table2');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '구분', options: headerOpts }, { text: '범위', options: headerOpts }],
        ['입고/출고 관리', '4개 핵심 모듈'],
        ['재고 관리', '실시간 추적'],
        ['배송 추적', 'GPS 연동'],
        ['정산 시스템', '자동화'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.5, 1.6] });
    }
    if (t2) {
      slide.addTable([
        [{ text: '구분', options: headerOpts }, { text: '비고', options: headerOpts }],
        ['회계 시스템', '인터페이스만 개발'],
        ['HR 시스템', '인터페이스만 개발'],
        ['기존 ERP', '인터페이스만 개발'],
      ], { x: t2.x, y: t2.y, w: t2.w, h: t2.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.5, 1.6] });
    }
  }
};

// 슬라이드 5: 추진 전략
const slide5 = {
  styles: commonStyles + `
.slide-body { padding: 10pt 35pt; }
.highlight-box { background: #22523B; color: #FFFFFF; padding: 10pt 15pt; border-radius: 6pt; margin-bottom: 10pt; }
.highlight-box h2 { font-size: 12pt; margin-bottom: 4pt; }
.highlight-box p { font-size: 9pt; }
.card-grid-3 { display: flex; gap: 12pt; margin-bottom: 10pt; }
.card { flex: 1; background: #FFFFFF; border: 1pt solid #8BAFA2; border-radius: 5pt; padding: 10pt; text-align: center; }
.card h3 { font-size: 10pt; color: #22523B; margin-bottom: 6pt; }
.card p { font-size: 8pt; color: #333333; line-height: 1.4; }
.timeline { display: flex; justify-content: space-between; position: relative; margin-top: 12pt; }
.timeline-item { text-align: center; flex: 1; }
.timeline-dot { width: 12pt; height: 12pt; background: #22523B; border-radius: 50%; margin: 0 auto 8pt; }
.timeline-content { background: #FFFFFF; border: 1pt solid #22523B; border-radius: 5pt; padding: 8pt; margin: 0 5pt; }
.timeline-content h4 { font-size: 9pt; color: #22523B; margin-bottom: 3pt; }
.timeline-content p { font-size: 7pt; color: #767171; }
.timeline-line { position: absolute; top: 6pt; left: 60pt; right: 60pt; height: 2pt; background: #8BAFA2; z-index: -1; }
`,
  body: `
<div class="slide-header"><h1>2. 추진 전략 및 방법론</h1><p class="page-num">5 / 15</p></div>
<div class="slide-body">
  <div class="highlight-box"><h2>핵심 가치</h2><p>"안정적 전환"과 "혁신적 기능 구현"의 균형</p></div>
  <div class="card-grid-3">
    <div class="card"><h3>단계적 전환 전략</h3><p>Big Bang 방식이 아닌<br>모듈별 순차 전환으로<br><b>리스크 최소화</b></p></div>
    <div class="card"><h3>애자일 기반 개발</h3><p>2주 단위 스프린트로<br>빠른 피드백과<br><b>유연한 대응</b></p></div>
    <div class="card"><h3>현업 밀착 협업</h3><p>주 2회 현업 리뷰를 통한<br><b>요구사항 정합성 확보</b></p></div>
  </div>
  <div class="sub-title-wrap"><h3 class="sub-title">수행 방법론: Agile-Waterfall 하이브리드</h3></div>
  <div class="timeline">
    <div class="timeline-line"></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>분석</h4><p>요구사항 정의<br>현행 분석</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>설계</h4><p>아키텍처/UI<br>DB 설계</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>개발</h4><p>기능 구현<br>단위 테스트</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>테스트/이행</h4><p>통합/인수 테스트<br>시스템 전환</p></div></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`
};

// 슬라이드 6: 기술 아키텍처
const slide6 = {
  styles: commonStyles + `
.slide-body { padding: 10pt 35pt; }
.tech-stack { display: flex; flex-wrap: wrap; gap: 6pt; margin-top: 10pt; }
.tech-badge { background: #22523B; color: #FFFFFF; padding: 5pt 10pt; border-radius: 12pt; font-size: 8pt; }
.tech-badge-alt { background: #479374; color: #FFFFFF; padding: 5pt 10pt; border-radius: 12pt; font-size: 8pt; }
`,
  body: `
<div class="slide-header"><h1>2. 기술 아키텍처</h1><p class="page-num">6 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">기술 스택</h3><div class="sub-title-bar"></div></div>
  <div id="table1" class="placeholder" style="width: 650pt; height: 115pt;"></div>
  <div class="sub-title-wrap" style="margin-top: 15pt;"><h3 class="sub-title">기술 스택 Overview</h3><div class="sub-title-bar"></div></div>
  <div class="tech-stack">
    <div class="tech-badge"><p>Vue.js 3</p></div>
    <div class="tech-badge"><p>TypeScript</p></div>
    <div class="tech-badge-alt"><p>Spring Boot</p></div>
    <div class="tech-badge-alt"><p>Java 17</p></div>
    <div class="tech-badge"><p>PostgreSQL</p></div>
    <div class="tech-badge"><p>Redis</p></div>
    <div class="tech-badge-alt"><p>AWS EKS</p></div>
    <div class="tech-badge-alt"><p>AWS RDS</p></div>
    <div class="tech-badge"><p>Python</p></div>
    <div class="tech-badge"><p>TensorFlow</p></div>
    <div class="tech-badge-alt"><p>Docker</p></div>
    <div class="tech-badge-alt"><p>Kubernetes</p></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '영역', options: headerOpts }, { text: '기술/제품', options: headerOpts }, { text: '버전', options: headerOpts }, { text: '라이선스', options: headerOpts }, { text: '비고', options: headerOpts }],
        [{ text: 'Frontend', options: labelOpts }, 'Vue.js 3, TypeScript', '3.4', 'MIT', 'SPA 구조'],
        [{ text: 'Backend', options: labelOpts }, 'Spring Boot, Java', '3.2 / 17', 'Apache 2.0', 'MSA 구조'],
        [{ text: 'Database', options: labelOpts }, 'PostgreSQL, Redis', '16 / 7', 'PostgreSQL / BSD', 'RDBMS + Cache'],
        [{ text: 'Infra', options: labelOpts }, 'AWS (EKS, RDS, S3)', '-', 'Commercial', '클라우드 네이티브'],
        [{ text: 'AI/ML', options: labelOpts }, 'Python, TensorFlow', '3.11 / 2.15', 'PSF / Apache 2.0', '수요예측, 경로최적화'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 7, border: { pt: 0.5, color: 'DDDDDD' }, colW: [0.9, 2.0, 0.9, 1.5, 1.5] });
    }
  }
};

// 슬라이드 7: 프로젝트 조직
const slide7 = {
  styles: commonStyles + `
.highlight-box { background: #8BAFA2; color: #153325; padding: 10pt 15pt; border-radius: 6pt; margin-top: 10pt; }
.highlight-box p { font-size: 10pt; }
`,
  body: `
<div class="slide-header"><h1>3. 프로젝트 조직 및 역할</h1><p class="page-num">7 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">조직별 역할 및 책임 (R&R)</h3><div class="sub-title-bar"></div></div>
  <div id="table1" class="placeholder" style="width: 650pt; height: 165pt;"></div>
  <div class="highlight-box"><p><b>총 투입 인력:</b> 8명 / <b>총 투입 공수:</b> 78MM</p></div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '역할', options: headerOpts }, { text: '담당자', options: headerOpts }, { text: '주요 책임', options: headerOpts }, { text: '비고', options: headerOpts }],
        [{ text: '발주기관 PM', options: labelOpts }, '정민호 부장', '프로젝트 총괄 관리, 의사결정', '최종 승인권자'],
        [{ text: '수행사 PM', options: labelOpts }, '김철수 차장', '프로젝트 수행 총괄, 품질/위험 관리', '현장 대리인'],
        [{ text: 'PMO', options: labelOpts }, '이영희 과장', '프로젝트 관리 지원, 품질 검토', '테크솔루션 PMO'],
        [{ text: 'AA (아키텍트)', options: labelOpts }, '박준형 책임', '시스템 설계, 아키텍처 검토', 'MSA 전문가'],
        [{ text: '개발 리더', options: labelOpts }, '최수진 선임', '개발 총괄, 기술 이슈 해결', '풀스택'],
        [{ text: 'QA 담당', options: labelOpts }, '한미영 대리', '품질 관리, 테스트 총괄', 'ISTQB 자격'],
        [{ text: 'DBA', options: labelOpts }, '오현우 대리', 'DB 설계 및 튜닝', 'OCP 자격'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 7, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.2, 1.0, 3.0, 1.2] });
    }
  }
};

// 슬라이드 8: 투입 인력 현황
const slide8 = {
  styles: commonStyles,
  body: `
<div class="slide-header"><h1>3. 투입 인력 현황</h1><p class="page-num">8 / 15</p></div>
<div class="slide-body">
  <div id="table1" class="placeholder" style="width: 650pt; height: 250pt;"></div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '구분', options: headerOpts }, { text: '분야', options: headerOpts }, { text: '등급', options: headerOpts }, { text: '성명', options: headerOpts }, { text: '투입기간', options: headerOpts }, { text: '투입률', options: headerOpts }],
        [{ text: 'PM', options: labelOpts }, '관리', '특급', '김철수', '12개월', '100%'],
        [{ text: 'AA', options: labelOpts }, '설계', '특급', '박준형', '10개월', '83%'],
        [{ text: 'PL', options: labelOpts }, '개발', '고급', '최수진', '12개월', '100%'],
        [{ text: 'Dev', options: labelOpts }, 'Backend', '고급', '임동현', '10개월', '83%'],
        [{ text: 'Dev', options: labelOpts }, 'Frontend', '중급', '강서연', '10개월', '83%'],
        [{ text: 'Dev', options: labelOpts }, 'Backend', '중급', '윤재호', '10개월', '83%'],
        [{ text: 'DBA', options: labelOpts }, 'DB', '고급', '오현우', '8개월', '67%'],
        [{ text: 'QA', options: labelOpts }, '테스트', '중급', '한미영', '6개월', '50%'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [0.8, 1.0, 0.8, 0.9, 1.0, 0.8], align: 'center' });
    }
  }
};

// 슬라이드 9: 프로젝트 일정
const slide9 = {
  styles: commonStyles + `
.slide-body { padding: 10pt 35pt; }
.progress-bar { display: flex; height: 22pt; border-radius: 5pt; overflow: hidden; margin-top: 10pt; }
.progress-bar div { display: flex; align-items: center; justify-content: center; color: #FFFFFF; font-size: 7pt; }
`,
  body: `
<div class="slide-header"><h1>4. 프로젝트 일정</h1><p class="page-num">9 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">마스터 일정</h3><div class="sub-title-bar"></div></div>
  <div id="table1" class="placeholder" style="width: 650pt; height: 140pt;"></div>
  <div class="sub-title-wrap" style="margin-top: 12pt;"><h3 class="sub-title">단계별 진행 비율</h3><div class="sub-title-bar"></div></div>
  <div class="progress-bar">
    <div style="background: #1a3a2a; width: 8%;"><p>착수 8%</p></div>
    <div style="background: #22523B; width: 17%;"><p>분석 17%</p></div>
    <div style="background: #2d6b4f; width: 17%;"><p>설계 17%</p></div>
    <div style="background: #479374; width: 33%;"><p>개발 33%</p></div>
    <div style="background: #6F886A; width: 12%;"><p>테스트 12%</p></div>
    <div style="background: #8BAFA2; width: 13%; color: #333;"><p>이행 13%</p></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '단계', options: headerOpts }, { text: '시작일', options: headerOpts }, { text: '종료일', options: headerOpts }, { text: '기간', options: headerOpts }, { text: '주요 활동', options: headerOpts }],
        [{ text: '착수', options: labelOpts }, '2025-01-06', '2025-01-31', '4주', '킥오프, 환경구축, 표준정의'],
        [{ text: '분석', options: labelOpts }, '2025-02-01', '2025-03-31', '8주', '요구사항 확정, 현행 분석'],
        [{ text: '설계', options: labelOpts }, '2025-04-01', '2025-05-31', '8주', '아키텍처/화면/DB 설계'],
        [{ text: '개발', options: labelOpts }, '2025-06-01', '2025-09-30', '16주', '코딩, 단위테스트'],
        [{ text: '테스트', options: labelOpts }, '2025-10-01', '2025-11-15', '6주', '통합/성능/인수 테스트'],
        [{ text: '이행/안정화', options: labelOpts }, '2025-11-16', '2025-12-31', '6주', '오픈, 운영이관, 안정화'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 7, border: { pt: 0.5, color: 'DDDDDD' }, colW: [0.9, 1.1, 1.1, 0.6, 2.7] });
    }
  }
};

// 슬라이드 10: 마일스톤
const slide10 = {
  styles: commonStyles + `
.slide-body { padding: 8pt 35pt; }
.milestone { display: flex; align-items: center; margin-bottom: 6pt; }
.milestone-marker { width: 22pt; height: 22pt; background: #22523B; color: #FFFFFF; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 8pt; font-weight: bold; margin-right: 10pt; }
.milestone-info h4 { font-size: 9pt; color: #22523B; }
.milestone-info p { font-size: 8pt; color: #767171; }
.card-grid-3 { display: flex; gap: 10pt; }
.card { flex: 1; background: #FFFFFF; border: 1pt solid #8BAFA2; border-radius: 5pt; }
.card-header { background: #22523B; color: #FFFFFF; padding: 5pt; border-radius: 4pt 4pt 0 0; font-size: 8pt; font-weight: bold; }
.card ul { padding: 6pt 6pt 6pt 18pt; font-size: 7pt; line-height: 1.5; }
`,
  body: `
<div class="slide-header"><h1>4. 주요 마일스톤</h1><p class="page-num">10 / 15</p></div>
<div class="slide-body">
  <div class="two-cols" style="margin-bottom: 8pt;">
    <div class="col">
      <div class="milestone"><div class="milestone-marker"><p>M1</p></div><div class="milestone-info"><h4>착수보고</h4><p>2025-01-10 | 수행계획서</p></div></div>
      <div class="milestone"><div class="milestone-marker"><p>M2</p></div><div class="milestone-info"><h4>요구사항 확정</h4><p>2025-03-31 | 요구사항정의서</p></div></div>
      <div class="milestone"><div class="milestone-marker"><p>M3</p></div><div class="milestone-info"><h4>설계 완료</h4><p>2025-05-31 | 설계서 일체</p></div></div>
    </div>
    <div class="col">
      <div class="milestone"><div class="milestone-marker"><p>M4</p></div><div class="milestone-info"><h4>개발 완료</h4><p>2025-09-30 | 소스코드</p></div></div>
      <div class="milestone"><div class="milestone-marker"><p>M5</p></div><div class="milestone-info"><h4>통합테스트 완료</h4><p>2025-11-15 | 테스트결과서</p></div></div>
      <div class="milestone"><div class="milestone-marker"><p>M6</p></div><div class="milestone-info"><h4>시스템 오픈</h4><p>2025-12-01 | 오픈보고서</p></div></div>
    </div>
  </div>
  <div class="sub-title-wrap"><h3 class="sub-title">산출물 주요 목록</h3><div class="sub-title-bar"></div></div>
  <div class="card-grid-3">
    <div class="card"><div class="card-header"><p>착수/분석</p></div><ul><li>착수보고서</li><li>수행계획서</li><li>요구사항정의서</li><li>현행시스템분석서</li></ul></div>
    <div class="card"><div class="card-header"><p>설계/개발</p></div><ul><li>아키텍처설계서</li><li>데이터베이스설계서</li><li>화면설계서</li><li>소스코드</li></ul></div>
    <div class="card"><div class="card-header"><p>테스트/완료</p></div><ul><li>통합테스트결과서</li><li>운영자매뉴얼</li><li>완료보고서</li></ul></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`
};

// 슬라이드 11: 품질 및 보안 관리
const slide11 = {
  styles: commonStyles,
  body: `
<div class="slide-header"><h1>5. 품질 및 보안 관리</h1><p class="page-num">11 / 15</p></div>
<div class="slide-body">
  <div class="two-cols">
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">품질 보증 활동</h3><div class="sub-title-bar"></div></div>
      <div id="table1" class="placeholder" style="width: 295pt; height: 110pt;"></div>
    </div>
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">보안 관리 계획</h3><div class="sub-title-bar"></div></div>
      <div id="table2" class="placeholder" style="width: 295pt; height: 110pt;"></div>
    </div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const t2 = placeholders.find(p => p.id === 'table2');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '활동', options: headerOpts }, { text: '시기', options: headerOpts }, { text: '방법', options: headerOpts }],
        [{ text: '산출물 검토', options: labelOpts }, '단계별 말', '발주 담당자 검토/승인'],
        [{ text: '코드 리뷰', options: labelOpts }, '개발 기간', 'GitHub PR, 2인 이상 승인'],
        [{ text: '정적 분석', options: labelOpts }, '개발 기간', 'SonarQube 품질 게이트'],
        [{ text: '테스트', options: labelOpts }, '통합/인수', '시나리오 기반 테스트'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.0, 0.9, 1.8] });
    }
    if (t2) {
      slide.addTable([
        [{ text: '구분', options: headerOpts }, { text: '보안 대책', options: headerOpts }],
        [{ text: '인적 보안', options: labelOpts }, '보안서약서, 보안 교육'],
        [{ text: '물리 보안', options: labelOpts }, '통제구역 출입 관리'],
        [{ text: '기술 보안', options: labelOpts }, 'OWASP Top 10 점검'],
        [{ text: 'PC 보안', options: labelOpts }, '백신, DLP 솔루션 적용'],
      ], { x: t2.x, y: t2.y, w: t2.w, h: t2.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.0, 2.6] });
    }
  }
};

// 슬라이드 12: 위험 관리
const slide12 = {
  styles: commonStyles + `
.slide-body { padding: 8pt 35pt; }
.timeline { display: flex; justify-content: space-between; position: relative; margin-top: 8pt; }
.timeline-item { text-align: center; flex: 1; }
.timeline-dot { width: 10pt; height: 10pt; background: #22523B; border-radius: 50%; margin: 0 auto 6pt; }
.timeline-content { background: #FFFFFF; border: 1pt solid #22523B; border-radius: 4pt; padding: 6pt; margin: 0 3pt; }
.timeline-content h4 { font-size: 8pt; color: #22523B; margin-bottom: 2pt; }
.timeline-content p { font-size: 6pt; color: #767171; }
.timeline-line { position: absolute; top: 5pt; left: 50pt; right: 50pt; height: 2pt; background: #8BAFA2; z-index: -1; }
`,
  body: `
<div class="slide-header"><h1>5. 위험 관리</h1><p class="page-num">12 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">주요 위험 및 대응</h3><div class="sub-title-bar"></div></div>
  <div id="table1" class="placeholder" style="width: 650pt; height: 115pt;"></div>
  <div class="sub-title-wrap" style="margin-top: 8pt;"><h3 class="sub-title">변경 통제 절차</h3><div class="sub-title-bar"></div></div>
  <div class="timeline">
    <div class="timeline-line"></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>1. 변경 요청</h4><p>CR 접수<br>(Jira 티켓)</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>2. 영향도 분석</h4><p>일정/비용/품질<br>(3일 내)</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>3. 승인/반려</h4><p>CCB 의사결정<br>(주 1회)</p></div></div>
    <div class="timeline-item"><div class="timeline-dot"></div><div class="timeline-content"><h4>4. 변경 이행</h4><p>반영 및<br>산출물 현행화</p></div></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '위험 요소', options: headerOpts }, { text: '영향도', options: headerOpts }, { text: '발생확률', options: headerOpts }, { text: '대응 방안', options: headerOpts }, { text: '담당', options: headerOpts }],
        [{ text: '요구사항 변경', options: labelOpts }, { text: '상', options: { color: colors.riskHigh, bold: true } }, { text: '상', options: { color: colors.riskHigh, bold: true } }, '변경관리위원회(CCB), 영향도 분석 필수', 'PM'],
        [{ text: '일정 지연', options: labelOpts }, { text: '상', options: { color: colors.riskHigh, bold: true } }, { text: '중', options: { color: colors.riskMid, bold: true } }, '주간 진척관리, Critical Path 집중 모니터링', 'PM'],
        [{ text: '기술적 난제', options: labelOpts }, { text: '중', options: { color: colors.riskMid, bold: true } }, { text: '중', options: { color: colors.riskMid, bold: true } }, 'PoC 선행, 외부 전문가 자문', 'AA'],
        [{ text: '인력 이탈', options: labelOpts }, { text: '상', options: { color: colors.riskHigh, bold: true } }, { text: '하', options: { color: colors.riskLow, bold: true } }, '대체 인력 풀 확보, 기술 문서화 강화', 'PM'],
        [{ text: '데이터 마이그레이션', options: labelOpts }, { text: '상', options: { color: colors.riskHigh, bold: true } }, { text: '중', options: { color: colors.riskMid, bold: true } }, '사전 검증, 롤백 계획 수립', 'DBA'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 7, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.3, 0.6, 0.7, 3.1, 0.6], align: 'center' });
    }
  }
};

// 슬라이드 13: 교육 및 기술이전
const slide13 = {
  styles: commonStyles + `
.slide-body { padding: 8pt 35pt; }
.card { background: #FFFFFF; border: 1pt solid #8BAFA2; border-left: 3pt solid #22523B; padding: 8pt; }
.card ul { padding-left: 12pt; font-size: 8pt; line-height: 1.6; }
`,
  body: `
<div class="slide-header"><h1>6. 교육 및 기술이전</h1><p class="page-num">13 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">교육 훈련 계획</h3><div class="sub-title-bar"></div></div>
  <div id="table1" class="placeholder" style="width: 650pt; height: 80pt;"></div>
  <div class="two-cols" style="margin-top: 10pt;">
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">기술이전 계획</h3><div class="sub-title-bar"></div></div>
      <div class="card"><ul><li><b>이전 대상:</b> 시스템 소스코드, 설계 산출물, 운영 매뉴얼 일체</li><li><b>이전 방법:</b> 담당자 1:1 멘토링 및 인계인수서 작성</li><li><b>지원 기간:</b> 시스템 안정화 기간 (오픈 후 3개월)</li></ul></div>
    </div>
    <div class="col">
      <div class="sub-title-wrap"><h3 class="sub-title">협조 요청 사항</h3><div class="sub-title-bar"></div></div>
      <div class="card"><ul><li><b>환경:</b> 프로젝트 룸, VPN, 출입증 (10석)</li><li><b>데이터:</b> 테스트용 샘플 데이터 (비식별화)</li><li><b>인프라:</b> AWS 계정 및 권한</li><li><b>참여:</b> 요구사항 인터뷰, UAT 지원</li></ul></div>
    </div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`,
  tables: (slide, placeholders) => {
    const t1 = placeholders.find(p => p.id === 'table1');
    const headerOpts = { fill: { color: colors.deepGreen }, color: colors.white, bold: true };
    const labelOpts = { fill: { color: colors.paleGreen }, color: colors.darkGreen, bold: true };
    if (t1) {
      slide.addTable([
        [{ text: '교육명', options: headerOpts }, { text: '대상', options: headerOpts }, { text: '시기', options: headerOpts }, { text: '내용', options: headerOpts }],
        [{ text: '사용자 교육', options: labelOpts }, '현업 담당자 30명', '오픈 2주 전', '시스템 주요 기능 및 사용법'],
        [{ text: '운영자 교육', options: labelOpts }, 'IT 운영팀 5명', '오픈 3주 전', '시스템 구성, 백업/복구, 장애대응'],
        [{ text: '관리자 교육', options: labelOpts }, '관리자 10명', '오픈 1주 전', '권한 관리, 리포트 활용'],
      ], { x: t1.x, y: t1.y, w: t1.w, h: t1.h, fontSize: 8, border: { pt: 0.5, color: 'DDDDDD' }, colW: [1.2, 1.5, 1.1, 2.8] });
    }
  }
};

// 슬라이드 14: 의사소통
const slide14 = {
  styles: commonStyles + `
.slide-body { padding: 8pt 35pt; }
.card-grid-4 { display: flex; gap: 10pt; margin-bottom: 12pt; }
.card { flex: 1; background: #FFFFFF; border: 1pt solid #8BAFA2; border-radius: 5pt; text-align: center; padding-bottom: 8pt; }
.card-header { background: #22523B; color: #FFFFFF; padding: 5pt; border-radius: 4pt 4pt 0 0; font-size: 8pt; font-weight: bold; margin-bottom: 6pt; }
.card p { font-size: 8pt; margin: 2pt 0; }
.card .time { font-weight: bold; color: #22523B; }
.card .desc { font-size: 7pt; color: #767171; margin-top: 4pt; }
.card-grid-3 { display: flex; gap: 10pt; }
.card-green { flex: 1; background: #e8f5e9; border: 1pt solid #8BAFA2; border-radius: 5pt; padding: 10pt; text-align: center; }
.card-green h3 { font-size: 9pt; color: #22523B; margin-bottom: 4pt; }
.card-green p { font-size: 8pt; color: #333333; }
`,
  body: `
<div class="slide-header"><h1>6. 의사소통 및 보고 체계</h1><p class="page-num">14 / 15</p></div>
<div class="slide-body">
  <div class="sub-title-wrap"><h3 class="sub-title">보고 및 회의 체계</h3><div class="sub-title-bar"></div></div>
  <div class="card-grid-4">
    <div class="card"><div class="card-header"><p>일일 스탠드업</p></div><p class="time">매일 09:30</p><p>개발팀</p><p class="desc">진행 현황, 블로커 공유</p></div>
    <div class="card"><div class="card-header"><p>주간보고</p></div><p class="time">매주 금 14:00</p><p>PM, 발주PM</p><p class="desc">주간 실적/계획, 이슈</p></div>
    <div class="card"><div class="card-header"><p>월간보고</p></div><p class="time">매월 말</p><p>경영진</p><p class="desc">마일스톤, 리스크 보고</p></div>
    <div class="card"><div class="card-header"><p>스프린트 리뷰</p></div><p class="time">격주 금</p><p>전체</p><p class="desc">데모, 피드백 수렴</p></div>
  </div>
  <div class="sub-title-wrap"><h3 class="sub-title">의사소통 채널</h3><div class="sub-title-bar"></div></div>
  <div class="card-grid-3">
    <div class="card-green"><h3>공식 문서</h3><p>Confluence, SharePoint</p></div>
    <div class="card-green"><h3>일상 커뮤니케이션</h3><p>Slack, Teams</p></div>
    <div class="card-green"><h3>이슈 관리</h3><p>Jira, GitHub Issues</p></div>
  </div>
</div>
<div class="slide-footer"><p>(주)테크솔루션</p><p>스마트 물류관리 시스템 구축 수행계획서</p></div>
`
};

// 슬라이드 15: 마무리
const slide15 = {
  styles: `
body { background: #22523B; justify-content: center; align-items: center; text-align: center; }
h1 { color: #FFFFFF; font-size: 28pt; margin-bottom: 15pt; }
.subtitle { color: #8BAFA2; font-size: 12pt; margin-bottom: 25pt; }
.contact-info { color: #FFFFFF; font-size: 10pt; line-height: 2; }
`,
  body: `
<h1>감사합니다</h1>
<p class="subtitle">스마트 물류관리 시스템 구축 프로젝트</p>
<div class="contact-info">
  <p><b>수행사:</b> (주)테크솔루션</p>
  <p><b>PM:</b> 김철수 차장</p>
  <p><b>문서번호:</b> DOC-PRJ-2025-001-001 v1.0</p>
  <p><b>작성일:</b> 2025-01-03</p>
</div>
`
};

// 모든 슬라이드
const slides = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12, slide13, slide14, slide15];

async function main() {
  const workspaceDir = path.join(__dirname, 'slides');

  // 슬라이드 HTML 파일 생성
  for (let i = 0; i < slides.length; i++) {
    const htmlContent = createSlideHTML(i + 1, slides[i]);
    const filePath = path.join(workspaceDir, `slide${i + 1}.html`);
    fs.writeFileSync(filePath, htmlContent, 'utf8');
    console.log(`Created: slide${i + 1}.html`);
  }

  // PPTX 생성
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = '스마트 물류관리 시스템 구축 프로젝트 수행계획서';
  pptx.author = '(주)테크솔루션';
  pptx.company = '(주)테크솔루션';
  pptx.subject = '프로젝트 수행계획서';

  for (let i = 0; i < slides.length; i++) {
    const htmlFile = path.join(workspaceDir, `slide${i + 1}.html`);
    console.log(`Converting: slide${i + 1}.html`);
    try {
      const { slide, placeholders } = await html2pptx(htmlFile, pptx);
      // 테이블 추가
      if (slides[i].tables) {
        slides[i].tables(slide, placeholders);
      }
    } catch (err) {
      console.error(`Error on slide ${i + 1}:`, err.message);
    }
  }

  const outputPath = path.join(__dirname, '..', 'output', 'project-plan-slides.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log(`\nPPTX created: ${outputPath}`);
}

main().catch(console.error);
