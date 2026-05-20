const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, WidthType,
  LevelFormat
} = require("docx");

// ── Colors ──
const WARM    = "E07C5A"; // brand warm
const GRAY    = "999999"; // example text
const DARK    = "333333"; // body text
const LIGHT_BG = "FDF0EB"; // warm light bg
const BORDER  = "E0D4CA";

// ── Helpers ──
function heading(text) {
  return new Paragraph({
    spacing: { before: 360, after: 160 },
    children: [
      new TextRun({ text, font: "Microsoft YaHei", size: 28, bold: true, color: WARM }),
    ],
  });
}

function subheading(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [
      new TextRun({ text, font: "Microsoft YaHei", size: 24, bold: true, color: DARK }),
    ],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text, font: "Microsoft YaHei", size: 21, color: opts.color || DARK, ...opts }),
    ],
  });
}

function hint(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: "💡 ", font: "Microsoft YaHei", size: 18 }),
      new TextRun({ text, font: "Microsoft YaHei", size: 18, color: GRAY, italics: true }),
    ],
  });
}

// Fill-in line with gray example
function fillLine(label, example) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({ text: label + "：", font: "Microsoft YaHei", size: 21, bold: true, color: DARK }),
      new TextRun({ text: example || "（请填写）", font: "Microsoft YaHei", size: 21, color: GRAY }),
    ],
  });
}

function fillArea(label, example) {
  return [
    new Paragraph({
      spacing: { before: 100, after: 40 },
      children: [
        new TextRun({ text: label, font: "Microsoft YaHei", size: 21, bold: true, color: DARK }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      indent: { left: 240 },
      children: [
        new TextRun({ text: example || "（请在此处填写）", font: "Microsoft YaHei", size: 21, color: GRAY }),
      ],
    }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 140, right: 140 };

function infoRow(label, example) {
  return new TableRow({
    children: [
      new TableCell({
        borders, margins: cellMargins,
        width: { size: 2200, type: WidthType.DXA },
        shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [new TextRun({ text: label, font: "Microsoft YaHei", size: 21, bold: true, color: WARM })],
        })],
      }),
      new TableCell({
        borders, margins: cellMargins,
        width: { size: 6880, type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text: example || "（请填写）", font: "Microsoft YaHei", size: 21, color: GRAY })],
        })],
      }),
    ],
  });
}

// ── Document ──
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Microsoft YaHei", size: 21 } },
    },
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1300, bottom: 1440, left: 1300 },
      },
    },
    children: [

      // ── TITLE ──
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: "🎨 咔乐艺术 · 教师个人简介", font: "Microsoft YaHei", size: 36, bold: true, color: WARM }),
      ]}),
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: "请按模板逐项填写，灰色文字为参考示例，填写时替换为您的真实信息即可。", font: "Microsoft YaHei", size: 20, color: GRAY }),
      ]}),
      new Paragraph({ spacing: { after: 360 }, children: [
        new TextRun({ text: "提交方式：填写完成后发给 [负责人姓名] ，如有疑问请在教师群内沟通。", font: "Microsoft YaHei", size: 20, color: GRAY }),
      ]}),

      // ══════ ① ══════
      heading("① 个人照片"),
      hint("请提供一张半身照或形象照，建议暖色调背景。照片请单独作为附件发送（不要贴在文档里），文件名为「姓名-照片」。"),
      body("照片要求：正面或微侧身、画面清晰、服装得体、背景简洁。"),

      // ══════ ② ══════
      heading("② 基本信息"),
      new Table({
        width: { size: 9080, type: WidthType.DXA },
        columnWidths: [2200, 6880],
        rows: [
          infoRow("姓名", "张梦琪"),
          infoRow("一句话标签", "陪孩子用色彩讲故事的人"),
          infoRow("毕业院校 + 专业", "安徽师范大学 · 美术学"),
          infoRow("学历", "本科"),
          infoRow("教师资格证", "高级中学美术"),
          infoRow("从教年限", "8年"),
          infoRow("执教课程 / 班级", "少儿水彩启蒙班、素描进阶班"),
        ],
      }),

      // ══════ ③ ══════
      heading("③ 教育格言"),
      hint("一句凝练的个人信条，放在简介开头或姓名下方，让家长一眼记住你。"),
      new Paragraph({
        spacing: { before: 80, after: 80 }, indent: { left: 240 },
        children: [
          new TextRun({ text: "「教育不是灌满一桶水，而是点燃一把火。」", font: "Microsoft YaHei", size: 21, color: GRAY }),
        ],
      }),
      body("您的格言："),

      // ══════ ④ ══════
      heading("④ 教学理念"),
      hint("用 1-2 句话展开你的教学理念。口语化、有温度，家长读的不是理论，是共鸣。"),
      ...fillArea("理念阐述", "我从不教孩子「应该怎么画」，而是帮他们发现「原来我可以这样画」。每个孩子的笔触都是独一无二的语言，老师要做的不是纠正，而是看见和引导。"),

      // ══════ ⑤ ══════
      heading("⑤ 擅长领域"),
      hint("用标签形式罗列，便于家长快速匹配。用逗号或斜杠分隔。"),
      new Paragraph({
        spacing: { before: 80, after: 80 }, indent: { left: 240 },
        children: [
          new TextRun({ text: "水彩 / 儿童创意美术 / 素描基础 / 绘本创作 / 国画入门", font: "Microsoft YaHei", size: 21, color: GRAY }),
        ],
      }),
      body("您的擅长领域："),

      // ══════ ⑥ ══════
      heading("⑥ 教学成果"),
      hint("选择 2-3 项最亮眼的成果，用数据说话（获奖人次、通过率、参展级别等）。每项一行。"),
      ...fillArea("成果 1", "辅导学生获合肥市少儿绘画大赛一等奖 12 人次"),
      ...fillArea("成果 2", "所带班级美术考级通过率 100%"),
      ...fillArea("成果 3", "个人水彩作品《徽州印象》入选安徽省青年美术作品展"),

      // ══════ ⑦ ══════
      heading("⑦ 写给家长的一段话"),
      hint("150 字以内，温暖、真诚、口语化。想象你正面对一位来接孩子的家长，你会说什么？"),
      ...fillArea("",
        "我始终相信，艺术教育不是培养画家，而是培养一个完整的人。在我的课堂上，" +
        "孩子们学会的不仅是技法，更是观察世界的方式、表达自己的勇气。当您的孩子举着作品" +
        "跑出教室、眼睛亮晶晶地喊「妈妈你看」—— 那就是我最幸福的时刻。期待与您一起，" +
        "守护孩子心中那份对美的本能热爱。"),
      body("（字数参考：上面这段约 150 字）", { color: GRAY, size: 18 }),

      // ══════ ⑧ ══════
      heading("⑧ 家长好评截图"),
      hint("请提供 2-3 张微信聊天截图或朋友圈家长好评，马赛克掉家长头像和姓名等隐私信息。"),
      hint("截图请单独作为附件发送（不要贴在文档里），文件名为「姓名-好评1」「姓名-好评2」等。"),
      body("建议选择：表达感谢的、提到孩子进步的、推荐给其他家长的 —— 这三种类型各一张。", { color: GRAY }),

      // ── FOOTER ──
      new Paragraph({ spacing: { before: 600, after: 100 }, border: { top: { style: BorderStyle.SINGLE, size: 2, color: WARM, space: 12 } }, children: [
        new TextRun({ text: "感谢您的配合！请于 ____ 年 ____ 月 ____ 日前提交。", font: "Microsoft YaHei", size: 20, color: GRAY }),
      ]}),
      new Paragraph({ children: [
        new TextRun({ text: "合肥咔乐艺术培训中心", font: "Microsoft YaHei", size: 18, color: GRAY }),
      ]}),

    ],
  }],
});

// ── Output ──
Packer.toBuffer(doc).then(buf => {
  const outPath = "c:/Users/Administrator/Desktop/123/teacher-profile-template.docx";
  fs.writeFileSync(outPath, buf);
  console.log("✅ 模板已生成: " + outPath);
});
