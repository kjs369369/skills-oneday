const docx = require("docx");
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, HeadingLevel, ShadingType,
  TableBorders, convertInchesToTwip
} = docx;

// 색상 팔레트
const COLORS = {
  primary: "1a56db",    // 파란색
  dark: "1f2937",       // 본문
  gray: "6b7280",       // 보조 텍스트
  lightBg: "f0f5ff",    // 연한 파랑 배경
  white: "ffffff",
  accent: "dc2626",     // 강조 빨강
  star: "f59e0b",       // 별 노랑
  tableBorder: "d1d5db",
  tableHeader: "1e40af",
  tableHeaderText: "ffffff",
  sectionBg: "f9fafb",
};

const noBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 },
};

const thinBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.tableBorder },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.tableBorder },
  left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.tableBorder },
  right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.tableBorder },
};

function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({
    heading: level,
    spacing: { before: level === HeadingLevel.HEADING_1 ? 400 : 240, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: level === HeadingLevel.HEADING_1 ? 36 : level === HeadingLevel.HEADING_2 ? 28 : 24,
        color: level === HeadingLevel.HEADING_1 ? COLORS.primary : COLORS.dark,
        font: "맑은 고딕",
      }),
    ],
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after || 80 },
    alignment: opts.align || AlignmentType.LEFT,
    indent: opts.indent ? { left: convertInchesToTwip(0.3) } : undefined,
    children: [
      new TextRun({
        text,
        size: opts.size || 20,
        color: opts.color || COLORS.dark,
        bold: opts.bold || false,
        italics: opts.italics || false,
        font: "맑은 고딕",
      }),
    ],
  });
}

function quoteBlock(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: convertInchesToTwip(0.3) },
    border: { left: { style: BorderStyle.SINGLE, size: 6, color: COLORS.primary } },
    shading: { type: ShadingType.CLEAR, fill: COLORS.lightBg },
    children: [
      new TextRun({
        text,
        size: 22,
        color: COLORS.primary,
        bold: true,
        font: "맑은 고딕",
      }),
    ],
  });
}

function emptyLine() {
  return new Paragraph({ spacing: { after: 80 }, children: [] });
}

function separator() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.tableBorder } },
    children: [],
  });
}

function makeTableCell(text, opts = {}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    shading: opts.header
      ? { type: ShadingType.CLEAR, fill: COLORS.tableHeader }
      : opts.shading
        ? { type: ShadingType.CLEAR, fill: opts.shading }
        : undefined,
    borders: thinBorders,
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({
            text,
            size: opts.size || 18,
            bold: opts.bold || opts.header || false,
            color: opts.header ? COLORS.tableHeaderText : (opts.color || COLORS.dark),
            font: "맑은 고딕",
          }),
        ],
      }),
    ],
  });
}

function makeTable(headers, rows, colWidths) {
  const headerRow = new TableRow({
    children: headers.map((h, i) =>
      makeTableCell(h, { header: true, width: colWidths ? colWidths[i] : undefined })
    ),
  });
  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((cell, ci) =>
        makeTableCell(cell, {
          width: colWidths ? colWidths[ci] : undefined,
          shading: ri % 2 === 1 ? COLORS.sectionBg : undefined,
        })
      ),
    })
  );
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}

function codeBlock(lines) {
  return lines.map(line =>
    new Paragraph({
      spacing: { after: 20 },
      shading: { type: ShadingType.CLEAR, fill: "f3f4f6" },
      indent: { left: convertInchesToTwip(0.2), right: convertInchesToTwip(0.2) },
      children: [
        new TextRun({
          text: line,
          size: 18,
          font: "Consolas",
          color: COLORS.dark,
        }),
      ],
    })
  );
}

function bulletItem(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 60 },
    indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.2) },
    children: [
      new TextRun({
        text: (opts.marker || "•") + " " + text,
        size: opts.size || 20,
        color: opts.color || COLORS.dark,
        bold: opts.bold || false,
        font: "맑은 고딕",
      }),
    ],
  });
}

// 본문 생성
const children = [];

// === 표지 영역 ===
children.push(emptyLine(), emptyLine(), emptyLine());
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 80 },
  children: [new TextRun({ text: "PROMO KIT", size: 20, color: COLORS.gray, font: "맑은 고딕", bold: true })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 40 },
  children: [new TextRun({ text: "모객 특강 홍보 키트", size: 48, color: COLORS.primary, bold: true, font: "맑은 고딕" })],
}));
children.push(emptyLine());
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 40 },
  children: [new TextRun({ text: "강의일: 2026년 3월 14일(금)  |  온라인(Zoom)  |  무료", size: 22, color: COLORS.dark, font: "맑은 고딕" })],
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "강사: AICLab 김진수", size: 22, color: COLORS.gray, font: "맑은 고딕" })],
}));
children.push(separator());

// === 1. 홍보 브리프 ===
children.push(heading("1. 홍보 브리프 (종합)"));
children.push(heading("강의 기본 정보", HeadingLevel.HEADING_3));
children.push(makeTable(
  ["항목", "내용"],
  [
    ["강의명", "당신의 보고서를 바꿀 AI 비밀무기 3가지"],
    ["부제", "노트북LM 필살기 + 클로드 Show Me + 실시간 데모"],
    ["일시", "2026년 3월 14일(금), 시간 TBD"],
    ["형태", "온라인 (Zoom 라이브)"],
    ["참가비", "무료"],
    ["소요시간", "2시간"],
    ["강사", "AICLab 김진수"],
    ["대상", "보고서·제안서를 작성하는 모든 직장인, 강사, 소상공인"],
    ["준비물", "노트북, 구글 계정, 클로드 무료 계정"],
  ],
  [25, 75]
));
children.push(emptyLine());

children.push(heading("핵심 메시지 (Key Message)", HeadingLevel.HEADING_3));
children.push(quoteBlock('"3일 걸리던 보고서, AI로 3시간이면 됩니다. 그 비밀을 2시간 안에 공개합니다."'));
children.push(emptyLine());

children.push(heading("타겟 오디언스", HeadingLevel.HEADING_3));
children.push(makeTable(
  ["우선순위", "대상", "핵심 니즈"],
  [
    ["★★★", "기업 기획/전략/마케팅 실무자", "보고서·제안서 작업 시간 단축"],
    ["★★★", "1인지식사업자(강사)", "강의자료·콘텐츠 퀄리티 향상"],
    ["★★☆", "소상공인", "사업계획서·홍보물 직접 제작"],
    ["★★☆", "컨설턴트/프리랜서", "클라이언트 제안서 차별화"],
    ["★★☆", "공공기관 담당자", "정책 보고서 시각화"],
  ],
  [15, 35, 50]
));
children.push(emptyLine());

children.push(heading("참석자 혜택", HeadingLevel.HEADING_3));
children.push(bulletItem("실습 자료 PDF 무료 제공", { marker: "✅" }));
children.push(bulletItem("본과정 얼리버드 할인 안내", { marker: "✅" }));
children.push(bulletItem("강의 중 만든 결과물 본인 소유", { marker: "✅" }));
children.push(emptyLine());

children.push(heading("커리큘럼 요약", HeadingLevel.HEADING_3));
children.push(makeTable(
  ["시간", "내용", "핵심 체험"],
  [
    ["25분", "보고서의 종말과 부활", "Before/After 충격 사례 3연발"],
    ["30분", "실습: 노트북LM 숨겨진 필살기", "오디오 오버뷰, 교차질문, 원클릭 브리핑"],
    ["10분", "쉬는시간", ""],
    ["30분", "실습: 클로드 스킬 라이브 — 맞춤 슬라이드 제작", "양식 분석 → 스킬 제작 → 결과물 즉석 생성 체험"],
    ["20분", "라이브 데모 + 결과물 갤러리", "4개 도구 조합 보고서 완성 + 스킬 결과물 5종 공개"],
    ["5분", "마무리 & 본과정 안내", "얼리버드 혜택"],
  ],
  [12, 40, 48]
));

children.push(separator());

// === 2. SNS 카피 세트 ===
children.push(heading("2. SNS 카피 세트"));

children.push(heading("A. 인스타그램 / 페이스북", HeadingLevel.HEADING_3));
children.push(bodyText("▎ 버전 1: 충격형", { bold: true, size: 22, color: COLORS.primary }));
children.push(...codeBlock([
  "보고서 만드는 데 며칠씩 쓰고 계신가요?",
  "",
  "이 보고서, 만드는 데 5분 걸렸습니다.",
  "(AI가 만들었거든요)",
  "",
  "3월 14일(금), 무료 온라인 특강에서",
  "AI 보고서의 비밀무기 3가지를 공개합니다.",
  "",
  "✅ 노트북LM 숨겨진 필살기",
  "✅ 클로드로 차트·다이어그램 자동 생성",
  "✅ 4개 AI 도구 조합 라이브 데모",
  "",
  "📌 대상: 보고서·제안서 쓰는 모든 분",
  "(기획자, 마케터, 강사, 소상공인, 컨설턴트)",
  "",
  "🆓 참가비 무료 | 온라인 Zoom",
  "🔗 신청 링크: [링크]",
  "",
  "#AI보고서 #노트북LM #클로드 #데이터시각화",
  "#AI활용 #직장인스킬업 #무료특강",
]));
children.push(emptyLine());

children.push(bodyText("▎ 버전 2: 질문형", { bold: true, size: 22, color: COLORS.primary }));
children.push(...codeBlock([
  "혹시 이런 고민 있으신가요?",
  "",
  "❌ 자료 수집에 시간의 80%를 쓴다",
  "❌ PPT 디자인 감각이 없다",
  "❌ 데이터를 차트로 바꾸는 게 귀찮다",
  "",
  "3월 14일, 이 고민을 한 번에 해결할",
  "AI 도구 조합법을 알려드립니다.",
  "",
  "노트북LM × 제미나이 × 클로드 × 안티그래비티",
  "→ 보고서 자동 완성 워크플로우",
  "",
  "📍 3/14(금) 온라인 무료 특강",
  "🔗 신청: [링크]",
  "",
  "#AI시각화 #보고서자동화 #무료강의",
  "#1인사업자 #소상공인 #직장인필수",
]));
children.push(emptyLine());

children.push(heading("B. 블로그 / 네이버 카페", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "[무료 온라인 특강] 당신의 보고서를 바꿀 AI 비밀무기 3가지",
  "",
  "안녕하세요, AICLab 김진수입니다.",
  "",
  "보고서, 제안서, 기획서...",
  "만드는 데 며칠씩 걸리시나요?",
  "",
  "저는 AI 도구 4개를 조합해서",
  "보고서 1편을 30분 안에 완성합니다.",
  "",
  "이번 무료 특강에서 그 방법을 공개합니다.",
  "",
  "■ 특강 정보",
  "- 일시: 2026년 3월 14일(금)",
  "- 형태: 온라인 (Zoom 라이브)",
  "- 참가비: 무료",
  "- 대상: 보고서를 쓰는 모든 분",
  "  (기획자, 마케터, 강사, 소상공인, 컨설턴트, 공무원)",
  "",
  "■ 이런 분께 추천합니다",
  "✔ 보고서 작성에 시간을 너무 많이 쓰는 분",
  "✔ PPT 디자인 감각이 없어 고민인 분",
  "✔ AI 도구를 써보고 싶은데 뭘 써야 할지 모르는 분",
  "✔ 혼자서 리서치·분석·보고서를 다 해야 하는 1인 사업자",
  "✔ 홍보물·사업계획서를 직접 만들어야 하는 소상공인",
  "",
  "■ 커리큘럼",
  "1. [충격 오프닝] 이 보고서, 5분 만에 만들었습니다",
  "2. [실습] 노트북LM 숨겨진 필살기 3가지",
  "   - 오디오 오버뷰: 자료가 팟캐스트로 변신",
  "   - 교차 질문 폭격: 숨은 인사이트 발굴",
  "   - 원클릭 브리핑: 1시간 요약을 30초에",
  "3. [실습] 클로드 Show Me + 스킬 2.0",
  "   - 텍스트 → 다이어그램 자동 변환",
  "   - 엑셀 없이 인터랙티브 차트 생성",
  "   - 프롬프트 하나로 보고서 연쇄 생성",
  "4. [라이브 데모] 5분의 기적",
  "   - 4개 AI 도구를 조합해 보고서 1편 실시간 완성",
  "",
  "■ 준비물",
  "- 노트북 (태블릿 가능하나 노트북 권장)",
  "- 구글 계정 (노트북LM 접속용)",
  "- 클로드 무료 계정",
  "",
  "■ 신청",
  "[신청 링크]",
  "",
  "이 워크플로우를 아는 사람은 아직 1%뿐입니다.",
  "먼저 시작하세요.",
]));
children.push(emptyLine());

children.push(heading("C. 카카오톡 / 문자", HeadingLevel.HEADING_3));
children.push(bodyText("▎ 초단문 (카톡 공유용)", { bold: true, size: 22, color: COLORS.primary }));
children.push(...codeBlock([
  "[무료 특강] AI로 보고서 만드는 비밀무기 3가지",
  "",
  "📅 3/14(금) 온라인 Zoom",
  "💰 무료",
  "👤 보고서 쓰는 모든 분 (직장인·강사·소상공인)",
  "",
  "👉 신청: [링크]",
]));
children.push(emptyLine());

children.push(bodyText("▎ 중문 (카카오 채널 / 단체방용)", { bold: true, size: 22, color: COLORS.primary }));
children.push(...codeBlock([
  "매번 보고서에 며칠씩 쓰시나요?",
  "",
  "AI 도구 4개를 조합하면",
  "보고서 1편을 30분에 완성할 수 있습니다.",
  "",
  "3/14(금) 무료 온라인 특강에서",
  "노트북LM 필살기 + 클로드 시각화 +",
  "라이브 데모까지 직접 체험해보세요.",
  "",
  "✅ 무료 | 온라인 Zoom",
  "✅ 직장인·강사·소상공인 누구나",
  "📎 신청: [링크]",
]));

children.push(separator());

// === 3. 포스터 / 배너 텍스트 ===
children.push(heading("3. 포스터 / 배너 텍스트"));

children.push(heading("메인 포스터", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "[헤드라인]",
  "당신의 보고서를 바꿀",
  "AI 비밀무기 3가지",
  "",
  "[서브 헤드라인]",
  "노트북LM 필살기 + 클로드 Show Me + 실시간 라이브 데모",
  "",
  "[본문 불릿]",
  "• 3일 걸리던 보고서 → AI로 3시간",
  "• 실습으로 직접 체험하는 2시간",
  "• 4개 AI 도구 조합 워크플로우 공개",
  "",
  "[정보 영역]",
  "2026. 3. 14 (금)",
  "온라인 Zoom 라이브",
  "참가비 무료",
  "",
  "[CTA]",
  "지금 신청하기 →",
  "",
  "[강사]",
  "AICLab 김진수",
  "",
  "[하단 태그]",
  "#노트북LM #제미나이 #클로드 #안티그래비티",
]));
children.push(emptyLine());

children.push(heading("가로 배너 (웹/이메일용)", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "[왼쪽 텍스트]",
  "AI로 보고서 만드는 비밀무기 3가지",
  "3/14(금) 무료 온라인 특강",
  "",
  "[오른쪽 CTA 버튼]",
  "무료 신청 →",
]));
children.push(emptyLine());

children.push(heading("정사각형 배너 (SNS 광고용)", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "[상단]",
  "무료 온라인 특강",
  "",
  "[중앙 - 대형 텍스트]",
  "이 보고서,",
  "5분 만에",
  "만들었습니다",
  "",
  "[하단]",
  "3/14(금) | AI 비밀무기 3가지 공개",
  "신청 → [링크]",
]));
children.push(emptyLine());

children.push(heading("스토리 / 릴스 텍스트 (세로형)", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "[화면 1] 보고서 만드는 데 며칠씩 쓰고 있다면?",
  "[화면 2] AI 도구 4개를 조합하면 30분이면 됩니다",
  "[화면 3] 3/14(금) 무료 특강 - 비밀무기 3가지 공개",
  "[화면 4] 지금 신청 → 프로필 링크",
]));

children.push(separator());

// === 4. 이메일 뉴스레터용 ===
children.push(heading("4. 이메일 뉴스레터용"));

children.push(heading("제목 후보", HeadingLevel.HEADING_3));
children.push(makeTable(
  ["#", "제목", "스타일"],
  [
    ["1", "[무료 특강] 이 보고서, 5분 만에 만들었습니다", "충격형"],
    ["2", "보고서 때문에 야근하는 당신에게", "공감형"],
    ["3", "AI 도구 4개 조합법, 무료로 공개합니다", "직접형"],
  ],
  [5, 70, 25]
));
children.push(emptyLine());

children.push(heading("본문", HeadingLevel.HEADING_3));
children.push(...codeBlock([
  "안녕하세요, [이름]님",
  "",
  "이 보고서를 만드는 데 얼마나 걸렸을까요?",
  "정답: 5분.",
  "",
  "노트북LM, 제미나이, 클로드, 안티그래비티.",
  "이 4개의 AI 도구를 조합하면 가능합니다.",
  "",
  "이번 금요일, 무료 온라인 특강에서",
  "그 비밀무기 3가지를 직접 체험해보세요.",
  "",
  "━━━━━━━━━━━━━━━━━",
  "📅 일시: 3월 14일(금)",
  "💻 형태: 온라인 Zoom",
  "💰 참가비: 무료",
  "━━━━━━━━━━━━━━━━━",
  "",
  "▶ 이런 걸 배웁니다",
  "1. 노트북LM 필살기 – 자료가 팟캐스트로 변신",
  "2. 클로드 시각화 – 텍스트가 차트·다이어그램으로",
  "3. 라이브 데모 – 4개 도구 조합으로 보고서 실시간 완성",
  "",
  "▶ 이런 분께 추천합니다",
  "• 보고서·제안서에 시간을 많이 쓰는 직장인",
  "• 강의자료 퀄리티를 높이고 싶은 강사",
  "• 홍보물·사업계획서를 직접 만드는 소상공인",
  "• AI 도구를 써보고 싶은데 시작이 어려운 분",
  "",
  "[무료 신청하기 →]",
  "",
  "이 워크플로우를 아는 사람은 아직 1%뿐입니다.",
  "",
  "AICLab 김진수 드림",
]));

children.push(separator());
children.push(bodyText("홍보팀 전달용  |  작성일: 2026-03-13", { color: COLORS.gray, size: 18, align: AlignmentType.RIGHT }));

// Document 생성
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "맑은 고딕", size: 20 },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(0.8),
          bottom: convertInchesToTwip(0.8),
          left: convertInchesToTwip(0.9),
          right: convertInchesToTwip(0.9),
        },
      },
    },
    children,
  }],
});

const outputPath = "c:/project/skills oneday/docs/lecture-plan/promo-kit.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("✅ docx 생성 완료:", outputPath);
});
