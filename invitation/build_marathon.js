// Builds invitation-marathon.docx: A5 card in the dark, bold "marathon poster" style.
// The background (runners, lime accents, the word دعوة and the date) is a full-page image
// behind the text; everything else is editable RTL text.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, AlignmentType,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign,
} = require('docx');

const YELLOW = 'F2D64B';
const WHITE = 'FFFFFF';
const LIGHT = 'D9D5C9';
const DARK = '1C1A17';
const FONT = 'Segoe UI';

const run = (text, { size = 11, bold = false, color = WHITE } = {}) =>
  new TextRun({
    text,
    rightToLeft: true,
    bold, boldComplexScript: bold,
    size: size * 2, sizeComplexScript: size * 2,
    color,
    font: { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT },
  });

const para = (children, { before = 0, after = 0, line = 240, align = AlignmentType.CENTER } = {}) =>
  new Paragraph({
    bidirectional: true,
    alignment: align,
    spacing: { before, after, line },
    children: Array.isArray(children) ? children : [children],
  });

const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
const yellowLine = { style: BorderStyle.SINGLE, size: 12, color: YELLOW };

const cell = (width, children, { fill, borders = noBorders, margins = { top: 80, bottom: 80, left: 120, right: 120 } } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    margins,
    verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
    children,
  });

const table = (widths, cells, opts = {}) => new Table({
  visuallyRightToLeft: true,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  columnWidths: widths,
  borders: noBorders,
  rows: cells.map((row) => new TableRow({ children: row })),
  ...opts,
});

const header = new Header({
  children: [new Paragraph({
    children: [new ImageRun({
      type: 'jpg',
      data: fs.readFileSync(path.join(__dirname, 'bg_marathon.jpg')),
      transformation: { width: 559, height: 794 },
      floating: {
        horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
        verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
        behindDocument: true,
        allowOverlap: true,
        wrap: { type: TextWrappingType.NONE },
      },
    })],
  })],
});

const W = 6391; // 8391 page - 2 x 1000 margins

const infoBox = table([3000, 3391], [[
  cell(3000, [
    para(run('الاثنين 12 أكتوبر 2026', { size: 13, bold: true, color: YELLOW })),
    para(run('على الساعة 09:00 صباحًا', { size: 12, bold: true })),
  ], { borders: { top: yellowLine, bottom: yellowLine, right: yellowLine, left: none } }),
  cell(3391, [
    para(run('المكان', { size: 9.5, bold: true, color: YELLOW })),
    para(run('دار الذكاء الاصطناعي', { size: 11.5, bold: true })),
    para(run('جامعة قاصدي مرباح – ورقلة', { size: 10, color: LIGHT })),
  ], { borders: { top: yellowLine, bottom: yellowLine, left: yellowLine, right: none } }),
]]);

const pill = table([2600], [[
  cell(2600, [para(run('محاور اللقاء:', { size: 11, bold: true, color: DARK }))], { fill: YELLOW, margins: { top: 40, bottom: 40, left: 120, right: 120 } }),
]], { alignment: AlignmentType.RIGHT });

const bullet = (text) => para(run('• ' + text, { size: 10, bold: true }), { align: AlignmentType.RIGHT });
const tight = { top: 30, bottom: 30, left: 120, right: 120 };
const topics = table([3195, 3196], [
  [cell(3195, [bullet('التكوين وسوق العمل')], { margins: tight }), cell(3196, [bullet('البحث والتظاهرات الرياضية')], { margins: tight })],
  [cell(3195, [bullet('التربصات والإدماج المهني')], { margins: tight }), cell(3196, [bullet('دعم المؤسسات الناشئة')], { margins: tight })],
]);

const footer = table([2500, 3891], [[
  cell(2500, [para(run('نتشرّف بحضوركم', { size: 12, bold: true, color: DARK }))], { fill: YELLOW }),
  cell(3891, [
    para(run('المدير المساعد المكلف بما بعد التدرج', { size: 8.5, bold: true })),
    para(run('والبحث العلمي والعلاقات الخارجية', { size: 8.5, bold: true })),
  ], { fill: '3A3731' }),
]]);

const doc = new Document({
  creator: 'معهد علوم وتقنيات النشاطات البدنية والرياضية',
  title: 'دعوة – لقاء الشركاء الاقتصاديين والاجتماعيين',
  sections: [{
    properties: {
      page: {
        size: { width: 8391, height: 11906 },
        margin: { top: 420, bottom: 500, left: 1000, right: 1000, header: 0, footer: 0 },
      },
    },
    headers: { default: header },
    children: [
      para(run('الجمهورية الجزائرية الديمقراطية الشعبية', { size: 9, bold: true })),
      para(run('وزارة التعليم العالي والبحث العلمي', { size: 8, color: LIGHT })),
      para(run('جامعة قاصدي مرباح – ورقلة', { size: 9, bold: true })),
      para(run('معهد علوم وتقنيات النشاطات البدنية والرياضية', { size: 9.5, bold: true, color: YELLOW }), { after: 3150 }),
      para(run('لقاء الشركاء الاقتصاديين والاجتماعيين', { size: 13, bold: true }), { after: 200 }),
      para([
        run('السيد(ة): ', { size: 12, bold: true }),
        run('…………………………………', { size: 12, color: YELLOW }),
        run(' المحترم(ة)', { size: 12, bold: true }),
      ], { after: 200 }),
      para(run('يتشرّف المدير المساعد المكلف بما بعد التدرج والبحث العلمي والعلاقات الخارجية بمعهد علوم وتقنيات النشاطات البدنية والرياضية بدعوتكم لحضور اللقاء المخصّص لـ', { size: 10.5, color: LIGHT }), { line: 280 }),
      para(run('إنشاء وتنصيب اللجنة المشتركة للتعاون', { size: 14, bold: true, color: YELLOW }), { before: 100, after: 40 }),
      para(run('بين المعهد وشركائه الاقتصاديين والاجتماعيين', { size: 10.5, color: LIGHT }), { after: 240 }),
      infoBox,
      para(run('', { size: 6 }), { after: 160 }),
      pill,
      para(run('', { size: 4 }), { after: 60 }),
      topics,
      para(run('', { size: 6 }), { after: 200 }),
      footer,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, 'invitation-marathon.docx'), buf);
  console.log('wrote invitation-marathon.docx');
});
