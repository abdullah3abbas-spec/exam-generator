import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  ImageRun, PageOrientation
} from 'docx';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));

const SOLID = BorderStyle.SINGLE;
const bdrThick  = { style: SOLID, size: 12, color: '000000' };
const bdrThin   = { style: SOLID, size: 4,  color: '888888' };
const bdrNone   = { style: BorderStyle.NIL, size: 0, color: 'FFFFFF' };
const bordersThick = { top: bdrThick, bottom: bdrThick, left: bdrThick, right: bdrThick };
const bordersThin  = { top: bdrThin,  bottom: bdrThin,  left: bdrThin,  right: bdrThin  };
const bordersNone  = { top: bdrNone,  bottom: bdrNone,  left: bdrNone,  right: bdrNone  };

// A4 content width in DXA: 11906 - 1440 (margins) = 9000 (approx)
const PAGE_W = 9000;

function rtl(text, opts = {}) {
  return new TextRun({ text, font: 'Times New Roman', size: 22, ...opts });
}

function para(children, align = AlignmentType.RIGHT, spacing = {}) {
  return new Paragraph({
    bidirectional: true,
    alignment: align,
    children: Array.isArray(children) ? children : [rtl(children)],
    spacing: { before: 60, after: 60, ...spacing }
  });
}

function headerPara(text, size = 26, align = AlignmentType.CENTER) {
  return para([rtl(text, { bold: true, size })], align, { before: 80, after: 80 });
}

function tableCell(children, opts = {}) {
  const { shade, bold = true, align = AlignmentType.CENTER, borders = bordersThin, colspan } = opts;
  const cellChildren = typeof children === 'string'
    ? [new Paragraph({
        bidirectional: true, alignment: align,
        children: [rtl(children, { bold, size: 20 })],
        spacing: { before: 40, after: 40 }
      })]
    : children;
  const cell = {
    borders,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: cellChildren,
    ...(shade ? { shading: { fill: shade, type: ShadingType.CLEAR } } : {}),
    ...(colspan ? { columnSpan: colspan } : {})
  };
  return new TableCell(cell);
}

function makeTable(rows, widths) {
  return new Table({
    width: { size: PAGE_W, type: WidthType.DXA },
    columnWidths: widths,
    rows
  });
}

function emptyPara(before = 120) {
  return new Paragraph({ children: [], spacing: { before, after: 0 } });
}

// ── Build the score table ──────────────────────────────────────────────────
function buildScoreTable(mcqs, essays) {
  const colW = [1600, 1300, 1300, 1300, 1800, 1700]; // sum = 9000
  const hdr = ['الأسئلة', 'رقم السؤال', 'درجة السؤال', 'درجة الطالب', 'المصحح', 'المراجع'];

  const rows = [
    new TableRow({
      children: hdr.map(h => tableCell(h, { shade: 'D9D9D9', bold: true, borders: bordersThick }))
    }),
    new TableRow({
      children: [
        tableCell('الأسئلة الموضوعية', { shade: 'F2F2F2', borders: bordersThick }),
        tableCell(`1 – ${mcqs.length}`, { borders: bordersThick }),
        tableCell(`${mcqs.length}`, { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
      ]
    }),
    ...essays.map((q, i) => new TableRow({
      children: [
        tableCell(i === 0 ? 'الأسئلة المقالية' : '', { shade: 'F2F2F2', borders: bordersThick }),
        tableCell(`${mcqs.length + i + 1}`, { borders: bordersThick }),
        tableCell(`${q.marks || 15}`, { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
      ]
    })),
    new TableRow({
      children: [
        tableCell('', { borders: bordersThick }),
        tableCell('المجموع', { shade: 'F2F2F2', borders: bordersThick }),
        tableCell(`${mcqs.reduce((s,_)=>s+1,0) + essays.reduce((s,q)=>s+(q.marks||15),0)} درجة`, { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
        tableCell('', { borders: bordersThick }),
      ]
    }),
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 6, borders: bordersThick,
          shading: { fill: 'F2F2F2', type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({
            bidirectional: true, alignment: AlignmentType.RIGHT,
            children: [rtl('الدرجة بالحروف: ............................................', { bold: true, size: 20 })]
          })]
        })
      ]
    }),
  ];
  return makeTable(rows, colW);
}

// ── Build instructions box ─────────────────────────────────────────────────
function buildInstructions(totalQ) {
  const instrs = [
    'يجب استخدام القلم الرصاص للإجابة عن أسئلة الاختيار من متعدد كما يمكن استخدامه في الرسومات.',
    'يجب استخدام القلم الحبر في الإجابة عن الأسئلة المقالية.',
    'تم إعداد أسئلة الاختبار باللغة العربية.',
    'بعض أسئلة الاختبار هي أسئلة اختيار من متعدد. والبعض يتطلب منك إجابة قصيرة.',
    'أسئلة الاختيار من متعدد تتضمن أربعة اختيارات للإجابة. قم بتحديد إجابتك في المربع المقابل للاختيار الصحيح.',
    'قم بتحديد إجابة واحدة فقط بالنسبة لكل سؤال اختيار من متعدد. إذا رغبت في تغيير إجابتك قم بتظليل مربع الإجابة التي لا تريدها بشكل تام.',
    'أجب عن جميع الأسئلة. حتى إذا كنت غير متأكد منها، حيث أنه لا يتم خصم درجات على الإجابات غير الصحيحة.',
    'لا تضيع وقتاً طويلاً في الإجابة على سؤال واحد إذا وجدت سؤالاً صعباً. انتقل للإجابة عن الأسئلة الأخرى ثم عد إليه.',
    'سيتم تذكيرك بالوقت المتبقي للاختبار عند منتصف الوقت وقبل نهايته بـ 30 دقيقة.',
  ];

  return makeTable([
    new TableRow({
      children: [
        new TableCell({
          borders: bordersThick,
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            para([rtl('بسم الله الرحمن الرحيم', { bold: true, size: 26 })], AlignmentType.CENTER, { before: 60, after: 100 }),
            para([rtl(`عدد أسئلة اختبار ${totalQ} سؤالاً`, { bold: true, size: 24 })], AlignmentType.CENTER, { before: 60, after: 120 }),
            para([rtl('الإرشادات العامة:', { bold: true, size: 22 })], AlignmentType.RIGHT, { before: 60, after: 80 }),
            ...instrs.map(t =>
              para([rtl(`-  ${t}`, { bold: true, size: 20 })], AlignmentType.RIGHT, { before: 40, after: 40 })
            )
          ]
        })
      ]
    })
  ], [PAGE_W]);
}

// ── Build MCQ section ──────────────────────────────────────────────────────
function buildMCQ(mcqs) {
  const items = [];
  const letters = ['أ', 'ب', 'ج', 'د'];

  // Section instruction box
  items.push(makeTable([
    new TableRow({
      children: [
        new TableCell({
          columnSpan: 2, borders: bordersThick,
          shading: { fill: 'E8E8E8', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          children: [
            para([rtl(`اختر الإجابة الصحيحة لكل من الأسئلة من 1 إلى ${mcqs.length}، وذلك بوضع علامة × داخل المربع المجاور للإجابة الصحيحة`, { bold: true, size: 22 })], AlignmentType.RIGHT)
          ]
        })
      ]
    })
  ], [PAGE_W]));

  items.push(emptyPara(120));

  mcqs.forEach((q, i) => {
    // Question number + text
    items.push(para(
      [rtl(`${i + 1}-  `, { bold: true, size: 22 }), rtl(q.q, { bold: true, size: 22 })],
      AlignmentType.RIGHT, { before: 140, after: 80 }
    ));

    // 4 options in 2-column RTL table
    const colW2 = [4500, 4500];
    const optRows = [];
    for (let r = 0; r < 2; r++) {
      const cells = [];
      // RTL: col 0 = right side (options 1,3), col 1 = left side (options 0,2)
      for (let c = 1; c >= 0; c--) {
        const oi = r * 2 + c;
        cells.push(new TableCell({
          borders: bordersNone,
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          width: { size: 4500, type: WidthType.DXA },
          children: [para(
            [rtl(`□  ${letters[oi]}) `, { bold: false, size: 22 }), rtl(q.options[oi] || '', { size: 22 })],
            AlignmentType.RIGHT, { before: 40, after: 40 }
          )]
        }));
      }
      optRows.push(new TableRow({ children: cells }));
    }
    items.push(makeTable(optRows, colW2));
  });

  return items;
}

// ── Build Essay section ────────────────────────────────────────────────────
function buildEssay(mcqs, essays) {
  const items = [];

  essays.forEach((q, i) => {
    const qNum = mcqs.length + i + 1;

    // Question header table
    items.push(emptyPara(200));
    items.push(makeTable([
      new TableRow({
        children: [
          new TableCell({
            width: { size: 1400, type: WidthType.DXA },
            borders: bordersThick,
            shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            verticalAlign: VerticalAlign.CENTER,
            children: [para([rtl(`${q.marks || 15}`, { bold: true, size: 22 })], AlignmentType.CENTER)]
          }),
          new TableCell({
            width: { size: 7600, type: WidthType.DXA },
            borders: bordersThick,
            shading: { fill: 'EFEFEF', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER,
            children: [para([rtl(`السؤال رقم ${qNum}`, { bold: true, size: 24 })], AlignmentType.CENTER)]
          }),
        ]
      })
    ], [1400, 7600]));

    // Question content
    items.push(para(
      [rtl(`${qNum} –  `, { bold: true, size: 22 }), rtl(q.q, { bold: true, size: 22 })],
      AlignmentType.RIGHT, { before: 120, after: 80 }
    ));

    // Answer lines (dashes)
    for (let l = 0; l < 8; l++) {
      items.push(new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.RIGHT,
        children: [rtl('_'.repeat(90), { size: 18 })],
        spacing: { before: 100, after: 40 }
      }));
    }
  });

  return items;
}

// ── Main handler ───────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  try {
    const { examData, schoolName = 'مدرسة الايمان الثانوية' } = req.body;
    const mcqs   = examData.mcq   || [];
    const essays = examData.essay || [];

    // Load logo
    let logoImage = null;
    try {
      const logoData = readFileSync(join(__dirname, 'logo.png'));
      logoImage = new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            data: logoData,
            transformation: { width: 198, height: 40 },
            type: 'png'
          })
        ],
        spacing: { before: 0, after: 100 }
      });
    } catch (_) {
      logoImage = headerPara(schoolName, 30);
    }

    const doc = new Document({
      styles: {
        default: {
          document: { run: { font: 'Times New Roman', size: 22 } }
        }
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 720, right: 900, bottom: 720, left: 900 }
          }
        },
        children: [
          // ── Header ────────────────────────────────────────────
          logoImage,
          headerPara(schoolName, 30),
          headerPara('الاختبار التجريبي للشهادة الثانوية', 26),
          headerPara('الفصل الدراسي الثاني للعام الدراسي 2024/2025م', 24),
          headerPara('مادة: الكيمياء                مسار: العلمي', 24),
          headerPara(`زمن الاختبار: ${examData.duration || 'ساعتان'}`, 24),

          emptyPara(120),

          // ── Score Table ────────────────────────────────────────
          buildScoreTable(mcqs, essays),

          emptyPara(120),

          // Coordinator line
          para(
            [rtl('المنسق / قائد الطاولة :  ................................................  التوقيع :  .......................', { size: 22 })],
            AlignmentType.RIGHT, { before: 100, after: 200 }
          ),

          // ── Instructions ──────────────────────────────────────
          buildInstructions(mcqs.length + essays.length),

          emptyPara(240),

          // ── MCQ Questions ─────────────────────────────────────
          ...buildMCQ(mcqs),

          emptyPara(240),

          // ── Essay Questions ───────────────────────────────────
          ...buildEssay(mcqs, essays),

          emptyPara(400),

          // ── End ───────────────────────────────────────────────
          para([rtl('انتهت جميع الأسئلة', { bold: true, size: 26 })], AlignmentType.CENTER, { before: 400 }),
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = encodeURIComponent('امتحان-كيمياء.docx');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    res.send(buffer);

  } catch (err) {
    console.error('download error:', err);
    res.status(500).json({ error: err.message });
  }
}
