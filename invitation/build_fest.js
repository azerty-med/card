// Builds invitation-fest.docx: A5 card in the illustrated "sports fest" style.
// The background (sports equipment + the word دعوة) is a full-page image behind the text;
// everything else is editable RTL text so the invitee's name can be typed per copy.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, AlignmentType,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
} = require('docx');

const RED = 'C8362B';
const GREEN = '2E7D3A';
const BROWN = '3B2A20';
const FONT = 'Segoe UI';

const run = (text, { size = 12, bold = false, color = BROWN } = {}) =>
  new TextRun({
    text,
    rightToLeft: true,
    bold, boldComplexScript: bold,
    size: size * 2, sizeComplexScript: size * 2,
    color,
    font: { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT },
  });

const para = (children, { before = 0, after = 0, line = 240 } = {}) =>
  new Paragraph({
    bidirectional: true,
    alignment: AlignmentType.CENTER,
    spacing: { before, after, line },
    children: Array.isArray(children) ? children : [children],
  });

const header = new Header({
  children: [new Paragraph({
    children: [new ImageRun({
      type: 'jpg',
      data: fs.readFileSync(path.join(__dirname, 'bg_fest.jpg')),
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

const doc = new Document({
  creator: 'معهد علوم وتقنيات النشاطات البدنية والرياضية',
  title: 'دعوة – لقاء الشركاء الاقتصاديين والاجتماعيين',
  sections: [{
    properties: {
      page: {
        size: { width: 8391, height: 11906 },
        margin: { top: 760, bottom: 1700, left: 1560, right: 1560, header: 0, footer: 0 },
      },
    },
    headers: { default: header },
    children: [
      para(run('الجمهورية الجزائرية الديمقراطية الشعبية', { size: 9.5, bold: true })),
      para(run('وزارة التعليم العالي والبحث العلمي', { size: 8.5 })),
      para(run('جامعة قاصدي مرباح – ورقلة', { size: 9.5, bold: true })),
      para(run('معهد علوم وتقنيات النشاطات البدنية والرياضية', { size: 10, bold: true, color: GREEN }), { after: 2350 }),
      para(run('لقاء الشركاء الاقتصاديين والاجتماعيين', { size: 13, bold: true, color: RED }), { after: 300 }),
      para([
        run('السيد(ة): ', { size: 12.5, bold: true }),
        run('…………………………………', { size: 12.5 }),
        run(' المحترم(ة)', { size: 12.5, bold: true }),
      ], { after: 260 }),
      para(run('يتشرّف المدير المساعد المكلف بما بعد التدرج والبحث العلمي والعلاقات الخارجية بمعهد علوم وتقنيات النشاطات البدنية والرياضية بدعوتكم لحضور اللقاء المخصّص لـ', { size: 11 }), { line: 290 }),
      para(run('إنشاء وتنصيب اللجنة المشتركة للتعاون', { size: 14, bold: true, color: GREEN }), { before: 140, after: 60 }),
      para(run('بين المعهد وشركائه الاقتصاديين والاجتماعيين', { size: 11 }), { after: 360 }),
      para(run('الاثنين 12 أكتوبر 2026 – 09:00 صباحًا', { size: 13, bold: true, color: RED }), { after: 60 }),
      para(run('دار الذكاء الاصطناعي – جامعة قاصدي مرباح، ورقلة', { size: 11.5, bold: true, color: RED }), { after: 360 }),
      para(run('في انتظاركم!', { size: 24, bold: true, color: RED }), { after: 260 }),
      para(run('المدير المساعد المكلف بما بعد التدرج', { size: 9, bold: true })),
      para(run('والبحث العلمي والعلاقات الخارجية', { size: 9, bold: true })),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, 'invitation-fest.docx'), buf);
  console.log('wrote invitation-fest.docx');
});
