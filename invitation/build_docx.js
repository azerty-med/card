// Builds invitation.docx: A5 card, coloured background image in the header (behind text),
// all wording as editable Arabic (RTL) text so the invitee's name can be typed per copy.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, AlignmentType,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign,
} = require('docx');

const GREEN = '0B5A3A';
const GOLD = 'A9822A';
const INK = '1D1D1B';
const QUIET = '5A5A55';
const BODY_FONT = 'Traditional Arabic';
const HEAD_FONT = 'Arial';

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
const gold = { style: BorderStyle.DOUBLE, size: 6, color: GOLD };

// A5 = 148 x 210 mm = 8391 x 11906 DXA; image size in px at 96 dpi.
const bg = fs.readFileSync(path.join(__dirname, 'background.jpg'));
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

const detailsWidth = 5400;
const details = new Table({
  visuallyRightToLeft: true,
  alignment: AlignmentType.CENTER,
  width: { size: detailsWidth, type: WidthType.DXA },
  columnWidths: [detailsWidth],
  borders: { top: gold, bottom: gold, left: gold, right: gold, insideHorizontal: none, insideVertical: none },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: detailsWidth, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FBF3DC' },
      margins: { top: 120, bottom: 120, left: 120, right: 120 },
      children: [
        para([run('اليوم: ', { size: 13, bold: true, color: GREEN, font: HEAD_FONT }), run('الاثنين 12 أكتوبر 2026', { size: 13, bold: true, font: HEAD_FONT })], { after: 60 }),
        para([run('الساعة: ', { size: 13, bold: true, color: GREEN, font: HEAD_FONT }), run('التاسعة صباحًا (09:00)', { size: 13, bold: true, font: HEAD_FONT })], { after: 60 }),
        para([run('المكان: ', { size: 13, bold: true, color: GREEN, font: HEAD_FONT }), run('دار الذكاء الاصطناعي – جامعة ورقلة', { size: 13, bold: true, font: HEAD_FONT })]),
      ],
    })],
  })],
});

const signWidth = 6151; // page width minus side margins
const signature = new Table({
  visuallyRightToLeft: true,
  width: { size: signWidth, type: WidthType.DXA },
  columnWidths: [2551, 3600],
  borders: noBorders,
  rows: [new TableRow({
    children: [
      new TableCell({ width: { size: 2551, type: WidthType.DXA }, borders: noBorders, children: [para(run(''))] }),
      new TableCell({
        width: { size: 3600, type: WidthType.DXA },
        borders: noBorders,
        verticalAlign: VerticalAlign.CENTER,
        children: [
          para(run('المدير المساعد المكلف بما بعد التدرج', { size: 10, bold: true, color: GREEN, font: HEAD_FONT })),
          para(run('والبحث العلمي والعلاقات الخارجية', { size: 10, bold: true, color: GREEN, font: HEAD_FONT })),
        ],
      }),
    ],
  })],
});

const ornament = para(run('◆  ◆  ◆', { size: 9, color: GOLD, font: HEAD_FONT }), { before: 60, after: 120 });

const doc = new Document({
  creator: 'معهد علوم وتقنيات النشاطات البدنية والرياضية',
  title: 'دعوة – لقاء الشركاء الاقتصاديين والاجتماعيين',
  sections: [{
    properties: {
      page: {
        size: { width: 8391, height: 11906 },
        margin: { top: 1250, bottom: 900, left: 1120, right: 1120, header: 0, footer: 0 },
      },
    },
    headers: { default: header },
    children: [
      para(run('الجمهورية الجزائرية الديمقراطية الشعبية', { size: 10.5, bold: true, font: HEAD_FONT })),
      para(run('وزارة التعليم العالي والبحث العلمي', { size: 9.5, font: HEAD_FONT })),
      para(run('جامعة قاصدي مرباح – ورقلة', { size: 10.5, bold: true, font: HEAD_FONT })),
      para(run('معهد علوم وتقنيات النشاطات البدنية والرياضية', { size: 11, bold: true, color: GREEN, font: HEAD_FONT }), { after: 40 }),
      ornament,
      para(run('دعـــــوة', { size: 54, bold: true, color: GREEN }), { before: 0, after: 120, line: 420 }),
      para(run('لقاء تشاوري مع الشركاء الاقتصاديين والاجتماعيين', { size: 11.5, bold: true, color: GOLD, font: HEAD_FONT }), { before: 60, after: 400 }),
      para([
        run('السيد(ة): ', { size: 16, bold: true }),
        run('……………………………………', { size: 16, color: QUIET }),
        run(' المحترم(ة)', { size: 16, bold: true }),
      ], { after: 300 }),
      para(run('يتشرّف المدير المساعد المكلف بما بعد التدرج والبحث العلمي والعلاقات الخارجية بمعهد علوم وتقنيات النشاطات البدنية والرياضية بدعوتكم لحضور اللقاء المخصّص لـ', { size: 16 }), { line: 340 }),
      para(run('إنشاء وتنصيب اللجنة المشتركة للتعاون', { size: 14, bold: true, color: GREEN, font: HEAD_FONT }), { before: 120, after: 60 }),
      para(run('بين المعهد وشركائه الاقتصاديين والاجتماعيين', { size: 16 }), { after: 360 }),
      details,
      para(run('حضوركم إضافة نوعية لبناء شراكة فعّالة بين الجامعة ومحيطها', { size: 15, color: QUIET }), { before: 360, after: 360 }),
      signature,
      para(run('يُرجى تأكيد الحضور مسبقًا', { size: 9, color: QUIET, font: HEAD_FONT }), { before: 420 }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, 'invitation.docx'), buf);
  console.log('wrote invitation.docx');
});
