// Builds invitation-modern.docx: A5 card, sports-themed background image in the header (behind text),
// all wording as editable Arabic (RTL) text so the invitee's name can be typed per copy.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, AlignmentType,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign,
} = require('docx');

const NAVY = '0A1E3F';
const TEAL = '0F7C86';
const ORANGE = 'E8820C';
const LIGHT_ORANGE = 'FFB547';
const WHITE = 'FFFFFF';
const INK = '2B2F36';
const QUIET = '6B7280';
const BODY_FONT = 'Segoe UI';
const HEAD_FONT = 'Segoe UI';

const run = (text, { size = 14, bold = false, color = INK, font = BODY_FONT, italics = false } = {}) =>
  new TextRun({
    text,
    rightToLeft: true,
    bold, boldComplexScript: bold,
    italics, italicsComplexScript: italics,
    size: size * 2, sizeComplexScript: size * 2,
    color,
    font: { ascii: font, hAnsi: font, cs: font, eastAsia: font },
  });

const para = (children, { before = 0, after = 0, line = 240 } = {}) =>
  new Paragraph({
    bidirectional: true,
    alignment: AlignmentType.CENTER,
    spacing: { before, after, line },
    children: Array.isArray(children) ? children : [children],
  });

const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

// A5 = 148 x 210 mm = 8391 x 11906 DXA; image size in px at 96 dpi.
const bg = fs.readFileSync(path.join(__dirname, 'background_modern.jpg'));
const header = new Header({
  children: [new Paragraph({
    children: [new ImageRun({
      type: 'jpg',
      data: bg,
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


const W = 6391; // text width: 8391 page - 2 x 1000 margins
const cellBorder = (top) => ({
  top: top ? { style: BorderStyle.SINGLE, size: 24, color: ORANGE } : none,
  bottom: none,
  left: { style: BorderStyle.SINGLE, size: 24, color: WHITE },
  right: { style: BorderStyle.SINGLE, size: 24, color: WHITE },
});
const infoCell = (width, label, value) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  borders: cellBorder(true),
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'EEF6F7' },
  margins: { top: 100, bottom: 100, left: 80, right: 80 },
  verticalAlign: VerticalAlign.CENTER,
  children: [
    para(run(label, { size: 9.5, bold: true, color: TEAL }), { after: 20 }),
    para(run(value, { size: 11, bold: true, color: NAVY })),
  ],
});
const cols = [2400, 1250, 2741];
const details = new Table({
  visuallyRightToLeft: true,
  width: { size: W, type: WidthType.DXA },
  columnWidths: cols,
  borders: noBorders,
  rows: [new TableRow({
    children: [
      infoCell(cols[0], 'اليوم', 'الاثنين 12 أكتوبر 2026'),
      infoCell(cols[1], 'الساعة', '09:00 صباحًا'),
      infoCell(cols[2], 'المكان', 'دار الذكاء الاصطناعي – جامعة ورقلة'),
    ],
  })],
});

const signature = new Table({
  visuallyRightToLeft: true,
  width: { size: W, type: WidthType.DXA },
  columnWidths: [2791, 3600],
  borders: noBorders,
  rows: [new TableRow({
    children: [
      new TableCell({ width: { size: 2791, type: WidthType.DXA }, borders: noBorders, children: [para(run(''))] }),
      new TableCell({
        width: { size: 3600, type: WidthType.DXA },
        borders: noBorders,
        children: [
          para(run('المدير المساعد المكلف بما بعد التدرج', { size: 9.5, bold: true, color: NAVY })),
          para(run('والبحث العلمي والعلاقات الخارجية', { size: 9.5, bold: true, color: NAVY })),
        ],
      }),
    ],
  })],
});

const doc = new Document({
  creator: 'معهد علوم وتقنيات النشاطات البدنية والرياضية',
  title: 'دعوة – لقاء الشركاء الاقتصاديين والاجتماعيين',
  sections: [{
    properties: {
      page: {
        size: { width: 8391, height: 11906 },
        margin: { top: 560, bottom: 1815, left: 1000, right: 1000, header: 0, footer: 0 },
      },
    },
    headers: { default: header },
    children: [
      para(run('الجمهورية الجزائرية الديمقراطية الشعبية', { size: 10, bold: true, color: WHITE })),
      para(run('وزارة التعليم العالي والبحث العلمي', { size: 9, color: WHITE })),
      para(run('جامعة قاصدي مرباح – ورقلة', { size: 10, bold: true, color: WHITE })),
      para(run('معهد علوم وتقنيات النشاطات البدنية والرياضية', { size: 11, bold: true, color: LIGHT_ORANGE }), { after: 1650 }),
      para(run('دعـوة', { size: 40, bold: true, color: NAVY }), { after: 200, line: 240 }),
      para(run('لقاء تشاوري مع الشركاء الاقتصاديين والاجتماعيين', { size: 12, bold: true, color: TEAL }), { after: 420 }),
      para([
        run('السيد(ة): ', { size: 13, bold: true, color: NAVY }),
        run('………………………………………', { size: 13, color: QUIET }),
        run(' المحترم(ة)', { size: 13, bold: true, color: NAVY }),
      ], { after: 320 }),
      para(run('يتشرّف المدير المساعد المكلف بما بعد التدرج والبحث العلمي والعلاقات الخارجية بمعهد علوم وتقنيات النشاطات البدنية والرياضية بدعوتكم لحضور اللقاء المخصّص لـ', { size: 11.5 }), { line: 300 }),
      para(run('إنشاء وتنصيب اللجنة المشتركة للتعاون', { size: 14, bold: true, color: ORANGE }), { before: 100, after: 40 }),
      para(run('بين المعهد وشركائه الاقتصاديين والاجتماعيين', { size: 11.5 }), { after: 420 }),
      details,
      para(run('حضوركم إضافة نوعية لبناء شراكة فعّالة بين الجامعة ومحيطها', { size: 10.5, color: QUIET }), { before: 380, after: 300 }),
      signature,
      para(run('يُرجى تأكيد الحضور مسبقًا', { size: 8.5, color: QUIET }), { before: 300 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, 'invitation-modern.docx'), buf);
  console.log('wrote invitation-modern.docx');
});
