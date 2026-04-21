const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, LevelFormat, TableOfContents,
  Header, Footer, Tab, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// Color palette
const COLORS = {
  primary: "1F4E79",
  secondary: "2E75B6",
  accent: "BDD7EE",
  lightBlue: "DEEAF1",
  headerText: "FFFFFF",
  darkText: "1F1F1F",
  gray: "595959",
  lightGray: "F2F2F2",
  green: "375623",
  greenLight: "E2EFDA",
  orange: "843C0C",
  orangeLight: "FCE4D6",
  yellow: "7F6000",
  yellowLight: "FFF2CC",
  red: "C00000",
};

const border = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function headerCell(text, width, color = COLORS.primary) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: color, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: COLORS.headerText, size: 20, font: "Arial" })]
    })]
  });
}

function dataCell(text, width, shading = "FFFFFF", bold = false, center = false) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: shading, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text), bold, size: 19, font: "Arial" })]
    })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.secondary, space: 4 } },
    children: [new TextRun({ text, bold: true, color: COLORS.primary, size: 32, font: "Arial" })]
  });
}

function subTitle(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, bold: true, color: COLORS.secondary, size: 26, font: "Arial" })]
  });
}

function bodyText(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 20, font: "Arial" })]
  });
}

function bulletItem(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: "Arial", bold })]
  });
}

function colorBox(text, fillColor) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: fillColor, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [new Paragraph({
          children: [new TextRun({ text, size: 20, font: "Arial", color: COLORS.darkText })]
        })]
      })]
    })]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "✓", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: COLORS.primary },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: COLORS.secondary },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Table({
            width: { size: 9746, type: WidthType.DXA },
            columnWidths: [7000, 2746],
            rows: [new TableRow({
              children: [
                new TableCell({
                  borders: noBorders,
                  width: { size: 7000, type: WidthType.DXA },
                  shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
                  margins: { top: 100, bottom: 100, left: 200, right: 100 },
                  children: [new Paragraph({
                    children: [new TextRun({ text: "مشروع: إنشاء الموقع الإلكتروني لجامعة الأزهر", bold: true, color: "FFFFFF", size: 22, font: "Arial" })]
                  })]
                }),
                new TableCell({
                  borders: noBorders,
                  width: { size: 2746, type: WidthType.DXA },
                  shading: { fill: COLORS.secondary, type: ShadingType.CLEAR },
                  margins: { top: 100, bottom: 100, left: 100, right: 200 },
                  children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "إدارة المشاريع | 2024", color: "FFFFFF", size: 18, font: "Arial" })]
                  })]
                }),
              ]
            })]
          })
        ]
      })
    },
    children: [
      // ============ COVER PAGE ============
      new Paragraph({ spacing: { before: 600 }, children: [new TextRun("")] }),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [9746],
        rows: [new TableRow({
          children: [new TableCell({
            borders: noBorders,
            width: { size: 9746, type: WidthType.DXA },
            shading: { fill: COLORS.primary, type: ShadingType.CLEAR },
            margins: { top: 400, bottom: 400, left: 400, right: 400 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 },
                children: [new TextRun({ text: "PROJECT MANAGEMENT PLAN", bold: true, color: COLORS.accent, size: 28, font: "Arial" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 200 },
                children: [new TextRun({ text: "خطة إدارة المشروع", bold: true, color: "FFFFFF", size: 24, font: "Arial" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 },
                children: [new TextRun({ text: "إنشاء الموقع الإلكتروني", bold: true, color: "FFFFFF", size: 48, font: "Arial" })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 300 },
                children: [new TextRun({ text: "لجامعة الأزهر", bold: true, color: COLORS.accent, size: 48, font: "Arial" })] }),
            ]
          })]
        })]
      }),
      new Paragraph({ spacing: { before: 200, after: 100 }, children: [new TextRun("")] }),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [3000, 3000, 3746],
        rows: [
          new TableRow({ children: [
            headerCell("مُعِدّو المشروع", 3000, COLORS.secondary),
            headerCell("المادة", 3000, COLORS.secondary),
            headerCell("التاريخ", 3746, COLORS.secondary),
          ]}),
          new TableRow({ children: [
            dataCell("فريق المشروع", 3000, COLORS.lightBlue),
            dataCell("Project Management", 3000, COLORS.lightBlue),
            dataCell("2024", 3746, COLORS.lightBlue),
          ]})
        ]
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ============ TABLE OF CONTENTS ============
      sectionTitle("📋 جدول المحتويات"),
      bodyText("1. ميثاق المشروع (Project Charter)"),
      bodyText("2. تحليل أصحاب المصلحة (Stakeholder Analysis)"),
      bodyText("3. نطاق المشروع وهيكل تقسيم العمل (Scope & WBS)"),
      bodyText("4. جدول المشروع الزمني (Project Schedule / Gantt Chart)"),
      bodyText("5. المسار الحرج وتحليل الشبكة (CPM / PERT Network)"),
      bodyText("6. تقدير التكاليف والميزانية (Cost Estimation & Budget)"),
      bodyText("7. إدارة المخاطر (Risk Management Plan)"),
      bodyText("8. خطة ضمان الجودة (Quality Management Plan)"),
      bodyText("9. خطة الاتصالات (Communication Plan)"),
      bodyText("10. إغلاق المشروع (Project Closure)"),
      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 1: PROJECT CHARTER ============
      sectionTitle("1️⃣ ميثاق المشروع — Project Charter"),
      new Paragraph({ spacing: { before: 100, after: 160 }, children: [
        new TextRun({ text: "الميثاق هو وثيقة رسمية تُعطي تصريحاً بتنفيذ المشروع وتحدد صلاحيات مدير المشروع.", size: 20, font: "Arial", italics: true, color: COLORS.gray })
      ]}),

      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [3000, 6746],
        rows: [
          new TableRow({ children: [headerCell("البند", 3000, COLORS.primary), headerCell("التفاصيل", 6746, COLORS.primary)] }),
          new TableRow({ children: [dataCell("اسم المشروع", 3000, COLORS.lightBlue, true), dataCell("إنشاء الموقع الإلكتروني لجامعة الأزهر", 6746)] }),
          new TableRow({ children: [dataCell("مدير المشروع", 3000, COLORS.lightBlue, true), dataCell("يُعيَّن من قِبَل مجلس الجامعة", 6746)] }),
          new TableRow({ children: [dataCell("الجهة الراعية (Sponsor)", 3000, COLORS.lightBlue, true), dataCell("جامعة الأزهر — إدارة التقنية والمعلومات", 6746)] }),
          new TableRow({ children: [dataCell("تاريخ البدء المقترح", 3000, COLORS.lightBlue, true), dataCell("1 يناير 2024", 6746)] }),
          new TableRow({ children: [dataCell("تاريخ الإنتهاء المقترح", 3000, COLORS.lightBlue, true), dataCell("30 يونيو 2024 (6 أشهر)", 6746)] }),
          new TableRow({ children: [dataCell("الميزانية الإجمالية", 3000, COLORS.lightBlue, true), dataCell("350,000 جنيه مصري", 6746)] }),
          new TableRow({ children: [dataCell("الأولوية", 3000, COLORS.lightBlue, true), dataCell("عالية جداً — استراتيجي", 6746)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 240 }, children: [new TextRun("")] }),
      subTitle("🎯 هدف المشروع"),
      bodyText("إنشاء موقع إلكتروني متكامل وحديث لجامعة الأزهر يخدم أكثر من 500,000 طالب وعضو هيئة تدريس، يشمل خدمات أكاديمية وإدارية متكاملة، ويدعم اللغتين العربية والإنجليزية، مع إمكانية الوصول عبر جميع الأجهزة."),

      subTitle("📌 مبرر المشروع (Business Case)"),
      bulletItem("الموقع الحالي قديم ولا يدعم الجوال (Mobile Responsive)"),
      bulletItem("عدم وجود بوابة إلكترونية موحدة للطلاب وأعضاء الهيئة التدريسية"),
      bulletItem("صعوبة الوصول للمعلومات الأكاديمية والإدارية إلكترونياً"),
      bulletItem("ضرورة مواكبة التحول الرقمي في التعليم العالي"),
      bulletItem("تحسين الصورة الذهنية للجامعة أمام المجتمع الدولي"),

      subTitle("✅ نتائج المشروع المتوقعة (Deliverables)"),
      bulletItem("موقع إلكتروني متجاوب (Responsive) يدعم كل الأجهزة"),
      bulletItem("بوابة طلاب (Student Portal) متكاملة"),
      bulletItem("بوابة أعضاء هيئة التدريس"),
      bulletItem("نظام إدارة المحتوى (CMS)"),
      bulletItem("قاعدة بيانات مترابطة بالسجلات الأكاديمية"),
      bulletItem("نظام الأمان والحماية الإلكترونية"),
      bulletItem("توثيق تقني كامل وتدريب الكوادر"),

      subTitle("⚠️ القيود والافتراضات"),
      bulletItem("القيد: الميزانية ثابتة لا تتجاوز 350,000 جنيه"),
      bulletItem("القيد: يجب الإطلاق قبل بداية العام الدراسي 2024/2025"),
      bulletItem("الافتراض: توافر فريق فني متخصص من داخل الجامعة"),
      bulletItem("الافتراض: استقرار متطلبات المشروع بعد مرحلة التخطيط"),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 2: STAKEHOLDERS ============
      sectionTitle("2️⃣ تحليل أصحاب المصلحة — Stakeholder Analysis"),
      bodyText("يهدف تحليل أصحاب المصلحة إلى تحديد كل الأطراف المؤثرة والمتأثرة بالمشروع، وتحديد مستوى اهتمامهم وتأثيرهم لوضع استراتيجية مشاركة مناسبة."),
      new Paragraph({ spacing: { before: 120 }, children: [new TextRun("")] }),

      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2000, 1800, 1600, 1600, 2746],
        rows: [
          new TableRow({ children: [
            headerCell("صاحب المصلحة", 2000),
            headerCell("الدور", 1800),
            headerCell("مستوى التأثير", 1600),
            headerCell("مستوى الاهتمام", 1600),
            headerCell("استراتيجية التعامل", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("رئيس الجامعة", 2000, COLORS.lightBlue, true),
            dataCell("Sponsor", 1800),
            dataCell("عالي جداً ⬆", 1600, COLORS.greenLight),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("اجتماعات شهرية + تقارير تنفيذية", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("مدير التقنية (CTO)", 2000, COLORS.lightBlue, true),
            dataCell("Technical Lead", 1800),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("اجتماعات أسبوعية مع الفريق التقني", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("الطلاب (500K+)", 2000, COLORS.lightBlue, true),
            dataCell("المستخدم النهائي", 1800),
            dataCell("منخفض ⬇", 1600, COLORS.yellowLight),
            dataCell("عالي جداً ⬆", 1600, COLORS.greenLight),
            dataCell("استطلاعات رأي + اختبار المستخدم", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("أعضاء هيئة التدريس", 2000, COLORS.lightBlue, true),
            dataCell("المستخدم النهائي", 1800),
            dataCell("متوسط ⬌", 1600, COLORS.yellowLight),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("ورش عمل ودورات تدريبية", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("الإدارة الأكاديمية", 2000, COLORS.lightBlue, true),
            dataCell("أصحاب القرار", 1800),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("متوسط ⬌", 1600, COLORS.yellowLight),
            dataCell("تقارير تقدم دورية", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("شركة التطوير الخارجية", 2000, COLORS.lightBlue, true),
            dataCell("Vendor / Contractor", 1800),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("عقد رسمي + متابعة يومية", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("وحدة الأمن المعلوماتي", 2000, COLORS.lightBlue, true),
            dataCell("Quality Assurance", 1800),
            dataCell("متوسط ⬌", 1600, COLORS.yellowLight),
            dataCell("متوسط ⬌", 1600, COLORS.yellowLight),
            dataCell("مراجعات أمنية دورية", 2746),
          ]}),
          new TableRow({ children: [
            dataCell("وزارة التعليم العالي", 2000, COLORS.lightBlue, true),
            dataCell("جهة تنظيمية", 1800),
            dataCell("عالي ⬆", 1600, COLORS.greenLight),
            dataCell("منخفض ⬇", 1600, COLORS.yellowLight),
            dataCell("امتثال للمعايير + إخطارات رسمية", 2746),
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 200 }, children: [new TextRun("")] }),
      subTitle("مصفوفة القوة / الاهتمام (Power/Interest Matrix)"),
      colorBox("🔴 إدارة بعناية (High Power / High Interest): رئيس الجامعة, CTO, شركة التطوير\n🟡 إبقاء راضياً (High Power / Low Interest): وزارة التعليم العالي\n🟢 إبقاء مُطَّلِعاً (Low Power / High Interest): الطلاب, أعضاء هيئة التدريس\n⚪ مراقبة (Low Power / Low Interest): باقي الجهات الخارجية", COLORS.lightBlue),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 3: SCOPE & WBS ============
      sectionTitle("3️⃣ نطاق المشروع وهيكل تقسيم العمل — Scope & WBS"),

      subTitle("بيان نطاق المشروع (Scope Statement)"),
      bodyText("المشروع يشمل: تصميم وتطوير وإطلاق موقع إلكتروني متكامل لجامعة الأزهر، يتضمن بوابات متعددة للطلاب والأساتذة والإدارة، مع نظام إدارة المحتوى، وقاعدة البيانات، والأمن المعلوماتي، والتدريب والتوثيق الكامل."),
      new Paragraph({ spacing: { before: 80 }, children: [new TextRun("")] }),

      subTitle("ما هو داخل النطاق (In Scope):"),
      bulletItem("تصميم واجهة المستخدم (UI/UX Design)"),
      bulletItem("تطوير Frontend وBackend"),
      bulletItem("بوابة الطالب (نتائج، جداول، رسوم، شهادات إلكترونية)"),
      bulletItem("بوابة أعضاء هيئة التدريس (درجات، مواد، أبحاث)"),
      bulletItem("بوابة الإدارة (CMS كامل)"),
      bulletItem("قاعدة البيانات ونظام النسخ الاحتياطي"),
      bulletItem("نظام الأمان والتحقق من الهوية (2FA)"),
      bulletItem("اختبار الأداء والجودة (Testing & QA)"),
      bulletItem("الإطلاق والنشر (Deployment)"),
      bulletItem("التدريب والتوثيق التقني"),

      subTitle("ما هو خارج النطاق (Out of Scope):"),
      bulletItem("تطوير تطبيق جوال (Mobile App) — مشروع منفصل مستقبلاً"),
      bulletItem("نظام الدفع الإلكتروني — يُربط بنظام خارجي موجود"),
      bulletItem("تطوير نظام إدارة الساعات المعتمدة — مشروع منفصل"),

      new Paragraph({ spacing: { before: 200 }, children: [new TextRun("")] }),
      subTitle("هيكل تقسيم العمل (Work Breakdown Structure — WBS)"),

      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [1200, 2500, 4046, 2000],
        rows: [
          new TableRow({ children: [
            headerCell("رمز WBS", 1200), headerCell("المرحلة / الحزمة", 2500),
            headerCell("المهام الفرعية", 4046), headerCell("المدة (أسبوع)", 2000)
          ]}),
          new TableRow({ children: [dataCell("1.0", 1200, COLORS.accent, true, true), dataCell("تهيئة المشروع", 2500, COLORS.accent, true), dataCell("وثائق الميثاق، تحليل المتطلبات، جلسات킥 العصف الذهني", 4046, COLORS.accent), dataCell("2", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("1.1", 1200, COLORS.lightBlue, false, true), dataCell("متطلبات المشروع", 2500, COLORS.lightBlue), dataCell("جمع المتطلبات من أصحاب المصلحة وتحليلها", 4046, COLORS.lightBlue), dataCell("1", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("1.2", 1200, COLORS.lightBlue, false, true), dataCell("وثيقة SRS", 2500, COLORS.lightBlue), dataCell("توثيق المتطلبات الوظيفية وغير الوظيفية", 4046, COLORS.lightBlue), dataCell("1", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("2.0", 1200, COLORS.accent, true, true), dataCell("التصميم", 2500, COLORS.accent, true), dataCell("تصميم واجهة المستخدم، النماذج الأولية (Wireframes)", 4046, COLORS.accent), dataCell("4", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("2.1", 1200, COLORS.lightBlue, false, true), dataCell("تصميم UI/UX", 2500, COLORS.lightBlue), dataCell("Mockups, Prototypes, Style Guide", 4046, COLORS.lightBlue), dataCell("3", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("2.2", 1200, COLORS.lightBlue, false, true), dataCell("تصميم قاعدة البيانات", 2500, COLORS.lightBlue), dataCell("ER Diagram, Database Schema, Normalization", 4046, COLORS.lightBlue), dataCell("1", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("3.0", 1200, COLORS.accent, true, true), dataCell("التطوير", 2500, COLORS.accent, true), dataCell("برمجة Frontend وBackend والاتصال بقاعدة البيانات", 4046, COLORS.accent), dataCell("10", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("3.1", 1200, COLORS.lightBlue, false, true), dataCell("Frontend Development", 2500, COLORS.lightBlue), dataCell("HTML, CSS, React.js, Bootstrap, Arabic RTL Support", 4046, COLORS.lightBlue), dataCell("5", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("3.2", 1200, COLORS.lightBlue, false, true), dataCell("Backend Development", 2500, COLORS.lightBlue), dataCell("Node.js/Laravel, REST APIs, Authentication", 4046, COLORS.lightBlue), dataCell("5", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("3.3", 1200, COLORS.lightBlue, false, true), dataCell("Student Portal", 2500, COLORS.lightBlue), dataCell("نتائج، جداول، مواد، شهادات إلكترونية", 4046, COLORS.lightBlue), dataCell("3", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("3.4", 1200, COLORS.lightBlue, false, true), dataCell("Faculty Portal", 2500, COLORS.lightBlue), dataCell("إدارة الدرجات، المحاضرات، الأبحاث", 4046, COLORS.lightBlue), dataCell("2", 2000, COLORS.lightBlue, false, true)] }),
          new TableRow({ children: [dataCell("4.0", 1200, COLORS.accent, true, true), dataCell("الاختبار والجودة", 2500, COLORS.accent, true), dataCell("Unit Testing, Integration Testing, UAT, Pen Testing", 4046, COLORS.accent), dataCell("4", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("5.0", 1200, COLORS.accent, true, true), dataCell("الإطلاق والنشر", 2500, COLORS.accent, true), dataCell("نشر الخادم، DNS، SSL، Migration", 4046, COLORS.accent), dataCell("1", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("6.0", 1200, COLORS.accent, true, true), dataCell("التدريب والإغلاق", 2500, COLORS.accent, true), dataCell("تدريب المستخدمين، التوثيق، تسليم المشروع", 4046, COLORS.accent), dataCell("2", 2000, COLORS.accent, false, true)] }),
          new TableRow({ children: [dataCell("المجموع", 1200, COLORS.primary, true, true), dataCell("", 2500, COLORS.primary), dataCell("", 4046, COLORS.primary), dataCell("~26 أسبوع", 2000, COLORS.primary, true, true)] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 4: SCHEDULE / GANTT ============
      sectionTitle("4️⃣ الجدول الزمني للمشروع — Project Schedule (Gantt Chart)"),
      bodyText("المشروع ممتد على 6 أشهر (26 أسبوعاً) ابتداءً من يناير 2024 حتى يونيو 2024. المخطط الزمني التالي يوضح المهام الرئيسية والأسابيع المخصصة لها."),

      new Paragraph({ spacing: { before: 160 }, children: [new TextRun("")] }),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2400, 800, 800, 800, 800, 800, 800, 800, 800, 746],
        rows: [
          new TableRow({ children: [
            headerCell("المهمة / الشهر", 2400),
            headerCell("يناير", 800), headerCell("فبراير", 800), headerCell("مارس", 800),
            headerCell("أبريل", 800), headerCell("مايو", 800), headerCell("يونيو", 800),
            headerCell("أسابيع", 800), headerCell("المسؤول", 800), headerCell("الحالة", 746),
          ]}),
          new TableRow({ children: [
            dataCell("تهيئة المشروع", 2400, COLORS.lightBlue, true),
            dataCell("████", 800, COLORS.greenLight, false, true), dataCell("", 800), dataCell("", 800),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("2", 800, COLORS.lightGray, false, true), dataCell("مدير المشروع", 800), dataCell("✅ منجز", 746, COLORS.greenLight, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("جمع المتطلبات", 2400, COLORS.lightBlue, true),
            dataCell("██", 800, COLORS.accent, false, true), dataCell("██", 800, COLORS.accent, false, true), dataCell("", 800),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("3", 800, COLORS.lightGray, false, true), dataCell("Business Analyst", 800), dataCell("✅ منجز", 746, COLORS.greenLight, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("تصميم UI/UX", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("████", 800, COLORS.yellowLight, false, true), dataCell("██", 800, COLORS.yellowLight, false, true),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("4", 800, COLORS.lightGray, false, true), dataCell("UI/UX Designer", 800), dataCell("🔄 جاري", 746, COLORS.yellowLight, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("تطوير Frontend", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("", 800), dataCell("████", 800, COLORS.greenLight, false, true),
            dataCell("████", 800, COLORS.greenLight, false, true), dataCell("", 800), dataCell("", 800),
            dataCell("5", 800, COLORS.lightGray, false, true), dataCell("Frontend Team", 800), dataCell("⏳ لم يبدأ", 746, COLORS.lightGray, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("تطوير Backend + DB", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("", 800), dataCell("████", 800, COLORS.orangeLight, false, true),
            dataCell("████", 800, COLORS.orangeLight, false, true), dataCell("██", 800, COLORS.orangeLight, false, true), dataCell("", 800),
            dataCell("5", 800, COLORS.lightGray, false, true), dataCell("Backend Team", 800), dataCell("⏳ لم يبدأ", 746, COLORS.lightGray, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("Student & Faculty Portals", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("██", 800, COLORS.accent, false, true), dataCell("████", 800, COLORS.accent, false, true), dataCell("", 800),
            dataCell("5", 800, COLORS.lightGray, false, true), dataCell("Full Team", 800), dataCell("⏳ لم يبدأ", 746, COLORS.lightGray, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("الاختبار والجودة", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("", 800), dataCell("████", 800, COLORS.greenLight, false, true), dataCell("██", 800, COLORS.greenLight, false, true),
            dataCell("4", 800, COLORS.lightGray, false, true), dataCell("QA Team", 800), dataCell("⏳ لم يبدأ", 746, COLORS.lightGray, false, true),
          ]}),
          new TableRow({ children: [
            dataCell("الإطلاق والتدريب", 2400, COLORS.lightBlue, true),
            dataCell("", 800), dataCell("", 800), dataCell("", 800),
            dataCell("", 800), dataCell("", 800), dataCell("████", 800, COLORS.greenLight, false, true),
            dataCell("3", 800, COLORS.lightGray, false, true), dataCell("DevOps + Training", 800), dataCell("⏳ لم يبدأ", 746, COLORS.lightGray, false, true),
          ]}),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 5: CPM / PERT ============
      sectionTitle("5️⃣ تحليل المسار الحرج — CPM / PERT Network Analysis"),
      bodyText("يُستخدم تحليل المسار الحرج لتحديد أطول مسار في شبكة المشروع، حيث أي تأخير في هذا المسار يؤدي مباشرة لتأخير المشروع كله."),
      new Paragraph({ spacing: { before: 120 }, children: [new TextRun("")] }),

      subTitle("جدول الأنشطة والتبعيات"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [800, 2600, 1400, 1200, 1200, 1200, 1346],
        rows: [
          new TableRow({ children: [
            headerCell("النشاط", 800), headerCell("الوصف", 2600), headerCell("السابق", 1400),
            headerCell("الأمل (O)", 1200), headerCell("المتوقع (M)", 1200), headerCell("المتشائم (P)", 1200), headerCell("PERT (t)", 1346)
          ]}),
          new TableRow({ children: [dataCell("A", 800, COLORS.lightBlue, true, true), dataCell("تهيئة وميثاق المشروع", 2600), dataCell("—", 1400, COLORS.lightGray, false, true), dataCell("1", 1200, COLORS.lightGray, false, true), dataCell("2", 1200, COLORS.lightGray, false, true), dataCell("3", 1200, COLORS.lightGray, false, true), dataCell("2.0", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("B", 800, COLORS.lightBlue, true, true), dataCell("جمع المتطلبات وتوثيقها", 2600), dataCell("A", 1400, COLORS.lightGray, false, true), dataCell("2", 1200, COLORS.lightGray, false, true), dataCell("3", 1200, COLORS.lightGray, false, true), dataCell("5", 1200, COLORS.lightGray, false, true), dataCell("3.2", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("C", 800, COLORS.lightBlue, true, true), dataCell("تصميم UI/UX والنماذج الأولية", 2600), dataCell("B", 1400, COLORS.lightGray, false, true), dataCell("3", 1200, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("6", 1200, COLORS.lightGray, false, true), dataCell("4.2", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("D", 800, COLORS.lightBlue, true, true), dataCell("تصميم قاعدة البيانات", 2600), dataCell("B", 1400, COLORS.lightGray, false, true), dataCell("1", 1200, COLORS.lightGray, false, true), dataCell("1", 1200, COLORS.lightGray, false, true), dataCell("2", 1200, COLORS.lightGray, false, true), dataCell("1.2", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("E", 800, COLORS.lightBlue, true, true), dataCell("تطوير Frontend", 2600), dataCell("C", 1400, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("5", 1200, COLORS.lightGray, false, true), dataCell("8", 1200, COLORS.lightGray, false, true), dataCell("5.3", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("F", 800, COLORS.lightBlue, true, true), dataCell("تطوير Backend", 2600), dataCell("C, D", 1400, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("5", 1200, COLORS.lightGray, false, true), dataCell("9", 1200, COLORS.lightGray, false, true), dataCell("5.5", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("G", 800, COLORS.lightBlue, true, true), dataCell("بوابة الطالب والأستاذ", 2600), dataCell("E, F", 1400, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("5", 1200, COLORS.lightGray, false, true), dataCell("7", 1200, COLORS.lightGray, false, true), dataCell("5.2", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("H", 800, COLORS.lightBlue, true, true), dataCell("الاختبار الشامل والجودة", 2600), dataCell("G", 1400, COLORS.lightGray, false, true), dataCell("3", 1200, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("6", 1200, COLORS.lightGray, false, true), dataCell("4.2", 1346, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("I", 800, COLORS.lightBlue, true, true), dataCell("الإطلاق والتدريب والإغلاق", 2600), dataCell("H", 1400, COLORS.lightGray, false, true), dataCell("2", 1200, COLORS.lightGray, false, true), dataCell("3", 1200, COLORS.lightGray, false, true), dataCell("4", 1200, COLORS.lightGray, false, true), dataCell("3.0", 1346, COLORS.greenLight, false, true)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 180 }, children: [new TextRun("")] }),
      subTitle("🔴 المسار الحرج (Critical Path)"),
      colorBox("المسار الحرج: A → B → C → F → G → H → I\nإجمالي مدة المسار الحرج: 2.0 + 3.2 + 4.2 + 5.5 + 5.2 + 4.2 + 3.0 = 27.3 أسبوع\nصيغة PERT: t = (O + 4M + P) / 6\nالتباين الإجمالي للمسار: σ² = Σ [(P-O)/6]² = 2.8 (أسبوع²)\nالانحراف المعياري: σ ≈ 1.67 أسبوع", COLORS.orangeLight),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 6: COST ============
      sectionTitle("6️⃣ تقدير التكاليف والميزانية — Cost Estimation & Budget"),

      subTitle("تقدير التكاليف بطريقة Bottom-Up"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2800, 1600, 1600, 1800, 1946],
        rows: [
          new TableRow({ children: [
            headerCell("بند التكلفة", 2800), headerCell("الكمية", 1600), headerCell("التكلفة الوحدوية (جنيه)", 1600),
            headerCell("الإجمالي (جنيه)", 1800), headerCell("النسبة %", 1946)
          ]}),
          new TableRow({ children: [dataCell("فريق التطوير (6 أشهر)", 2800, COLORS.lightBlue, true), dataCell("5 مطورين", 1600), dataCell("8,000 / شهر", 1600), dataCell("240,000", 1800, COLORS.lightGray), dataCell("68.6%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("مصمم UI/UX", 2800, COLORS.lightBlue, true), dataCell("1 مصمم", 1600), dataCell("6,000 / شهر", 1600), dataCell("24,000", 1800, COLORS.lightGray), dataCell("6.9%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("خوادم واستضافة (سنوياً)", 2800, COLORS.lightBlue, true), dataCell("Cloud Server", 1600), dataCell("—", 1600), dataCell("18,000", 1800, COLORS.lightGray), dataCell("5.1%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("رخص البرمجيات وأدوات التطوير", 2800, COLORS.lightBlue, true), dataCell("متعددة", 1600), dataCell("—", 1600), dataCell("12,000", 1800, COLORS.lightGray), dataCell("3.4%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("شهادة SSL وأمن المعلومات", 2800, COLORS.lightBlue, true), dataCell("1", 1600), dataCell("—", 1600), dataCell("5,000", 1800, COLORS.lightGray), dataCell("1.4%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("برامج الاختبار والجودة", 2800, COLORS.lightBlue, true), dataCell("—", 1600), dataCell("—", 1600), dataCell("8,000", 1800, COLORS.lightGray), dataCell("2.3%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("التدريب وورش العمل", 2800, COLORS.lightBlue, true), dataCell("—", 1600), dataCell("—", 1600), dataCell("10,000", 1800, COLORS.lightGray), dataCell("2.9%", 1946, COLORS.greenLight, false, true)] }),
          new TableRow({ children: [dataCell("احتياطي المخاطر (Contingency 10%)", 2800, COLORS.lightBlue, true), dataCell("—", 1600), dataCell("—", 1600), dataCell("31,700", 1800, COLORS.yellowLight), dataCell("9.1%", 1946, COLORS.yellowLight, false, true)] }),
          new TableRow({ children: [
            dataCell("✅ الإجمالي الكلي", 2800, COLORS.primary, true),
            dataCell("—", 1600, COLORS.primary), dataCell("—", 1600, COLORS.primary),
            dataCell("348,700 ج.م", 1800, COLORS.primary, true),
            dataCell("100%", 1946, COLORS.primary, true, true)
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 200 }, children: [new TextRun("")] }),
      subTitle("منحنى S-Curve (مؤشر إنفاق الميزانية)"),
      colorBox("الربع الأول (يناير-مارس): 85,000 ج.م — 24% من الميزانية (تكاليف التخطيط والتصميم)\nالربع الثاني (أبريل-يونيو): 263,700 ج.م — 76% من الميزانية (التطوير والاختبار والإطلاق)\nمؤشر أداء التكلفة المستهدف (CPI): ≥ 1.0\nمؤشر أداء الجدول الزمني المستهدف (SPI): ≥ 1.0", COLORS.lightBlue),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 7: RISK MANAGEMENT ============
      sectionTitle("7️⃣ إدارة المخاطر — Risk Management Plan"),
      bodyText("تهدف إدارة المخاطر إلى تحديد وتقييم والتخطيط للاستجابة للمخاطر التي قد تؤثر على المشروع قبل حدوثها."),
      new Paragraph({ spacing: { before: 120 }, children: [new TextRun("")] }),

      subTitle("مصفوفة تقييم المخاطر"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [1600, 2000, 1000, 1000, 1000, 3146],
        rows: [
          new TableRow({ children: [
            headerCell("الخطر", 1600), headerCell("الوصف", 2000), headerCell("الاحتمالية\n(1-5)", 1000),
            headerCell("التأثير\n(1-5)", 1000), headerCell("الدرجة\n(P×I)", 1000), headerCell("خطة الاستجابة", 3146)
          ]}),
          new TableRow({ children: [dataCell("R1: تغيير المتطلبات", 1600, COLORS.lightBlue, true), dataCell("تغيير نطاق المشروع من قِبَل الإدارة", 2000), dataCell("4", 1000, COLORS.orangeLight, false, true), dataCell("4", 1000, COLORS.orangeLight, false, true), dataCell("16 🔴", 1000, COLORS.orangeLight, false, true), dataCell("تجميد النطاق بعد مرحلة التخطيط + عقد تغيير رسمي (Change Control)", 3146)] }),
          new TableRow({ children: [dataCell("R2: شُح الكوادر التقنية", 1600, COLORS.lightBlue, true), dataCell("استقالة أو تغيب مطورين أساسيين", 2000), dataCell("3", 1000, COLORS.yellowLight, false, true), dataCell("5", 1000, COLORS.orangeLight, false, true), dataCell("15 🔴", 1000, COLORS.orangeLight, false, true), dataCell("توثيق الكود + استعداد بديل + عقود خارجية احتياطية", 3146)] }),
          new TableRow({ children: [dataCell("R3: تجاوز الميزانية", 1600, COLORS.lightBlue, true), dataCell("ارتفاع تكاليف التطوير أو الخوادم", 2000), dataCell("3", 1000, COLORS.yellowLight, false, true), dataCell("4", 1000, COLORS.orangeLight, false, true), dataCell("12 🟡", 1000, COLORS.yellowLight, false, true), dataCell("مراجعة أسبوعية للميزانية + احتياطي 10% + موافقة لأي تجاوز", 3146)] }),
          new TableRow({ children: [dataCell("R4: ثغرات أمنية", 1600, COLORS.lightBlue, true), dataCell("اختراق البيانات أو بيانات الطلاب", 2000), dataCell("2", 1000, COLORS.greenLight, false, true), dataCell("5", 1000, COLORS.orangeLight, false, true), dataCell("10 🟡", 1000, COLORS.yellowLight, false, true), dataCell("اختبار اختراق (Pen Test) + تشفير كامل + مراجعة أمنية دورية", 3146)] }),
          new TableRow({ children: [dataCell("R5: تأخر الجدول الزمني", 1600, COLORS.lightBlue, true), dataCell("تأخر مراحل التطوير عن الموعد", 2000), dataCell("3", 1000, COLORS.yellowLight, false, true), dataCell("3", 1000, COLORS.yellowLight, false, true), dataCell("9 🟡", 1000, COLORS.yellowLight, false, true), dataCell("متابعة أسبوعية + تفعيل Fast Tracking أو Crashing عند الحاجة", 3146)] }),
          new TableRow({ children: [dataCell("R6: مشاكل تقنية", 1600, COLORS.lightBlue, true), dataCell("عطل الخوادم أو مشاكل التكامل", 2000), dataCell("2", 1000, COLORS.greenLight, false, true), dataCell("3", 1000, COLORS.yellowLight, false, true), dataCell("6 🟢", 1000, COLORS.greenLight, false, true), dataCell("خادم احتياطي (Backup Server) + نسخ يومية + اختبار مستمر", 3146)] }),
          new TableRow({ children: [dataCell("R7: مقاومة المستخدمين", 1600, COLORS.lightBlue, true), dataCell("رفض الأساتذة أو الإداريين للنظام", 2000), dataCell("2", 1000, COLORS.greenLight, false, true), dataCell("2", 1000, COLORS.greenLight, false, true), dataCell("4 🟢", 1000, COLORS.greenLight, false, true), dataCell("برنامج تدريب شامل + وثائق مستخدم + دعم فني مستمر", 3146)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 160 }, children: [new TextRun("")] }),
      colorBox("مقياس تقييم الخطورة: 🔴 عالي (≥12): يتطلب خطة فورية | 🟡 متوسط (6-11): مراقبة مستمرة | 🟢 منخفض (<6): قبول مع مراقبة", COLORS.lightGray),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 8: QUALITY ============
      sectionTitle("8️⃣ خطة ضمان الجودة — Quality Management Plan"),

      subTitle("معايير الجودة المعتمدة"),
      bulletItem("معيار ISO 9001:2015 لإدارة الجودة"),
      bulletItem("معيار WCAG 2.1 لإمكانية الوصول للمواقع"),
      bulletItem("معيار OWASP Top 10 لأمن تطبيقات الويب"),
      bulletItem("معيار ISO 27001 لأمن المعلومات"),

      new Paragraph({ spacing: { before: 160 }, children: [new TextRun("")] }),
      subTitle("نشاطات ضمان الجودة"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2400, 3000, 2000, 2346],
        rows: [
          new TableRow({ children: [
            headerCell("النشاط", 2400), headerCell("الوصف", 3000),
            headerCell("التوقيت", 2000), headerCell("المسؤول", 2346)
          ]}),
          new TableRow({ children: [dataCell("Code Review", 2400, COLORS.lightBlue, true), dataCell("مراجعة الكود أسبوعياً للتأكد من الجودة والمعايير", 3000), dataCell("أسبوعياً", 2000), dataCell("Tech Lead", 2346)] }),
          new TableRow({ children: [dataCell("Unit Testing", 2400, COLORS.lightBlue, true), dataCell("اختبار كل وحدة برمجية بشكل منفصل (80% coverage)", 3000), dataCell("مستمر", 2000), dataCell("Developers", 2346)] }),
          new TableRow({ children: [dataCell("Integration Testing", 2400, COLORS.lightBlue, true), dataCell("اختبار تكامل الأنظمة المختلفة معاً", 3000), dataCell("الشهر 4-5", 2000), dataCell("QA Team", 2346)] }),
          new TableRow({ children: [dataCell("UAT", 2400, COLORS.lightBlue, true), dataCell("اختبار القبول من المستخدمين الفعليين (طلاب وأساتذة)", 3000), dataCell("الشهر 5", 2000), dataCell("QA + Users", 2346)] }),
          new TableRow({ children: [dataCell("Performance Testing", 2400, COLORS.lightBlue, true), dataCell("اختبار تحمل 10,000 مستخدم متزامن على الأقل", 3000), dataCell("الشهر 5", 2000), dataCell("DevOps", 2346)] }),
          new TableRow({ children: [dataCell("Security / Pen Testing", 2400, COLORS.lightBlue, true), dataCell("اختبار الاختراق والأمن من طرف خارجي مستقل", 3000), dataCell("الشهر 5-6", 2000), dataCell("Security Team", 2346)] }),
          new TableRow({ children: [dataCell("Accessibility Testing", 2400, COLORS.lightBlue, true), dataCell("التحقق من توافق WCAG 2.1 وإمكانية الوصول للجميع", 3000), dataCell("الشهر 5", 2000), dataCell("QA Team", 2346)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 160 }, children: [new TextRun("")] }),
      subTitle("مقاييس الجودة الرئيسية (KPIs)"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [3200, 3200, 3346],
        rows: [
          new TableRow({ children: [headerCell("المقياس", 3200), headerCell("الهدف المستهدف", 3200), headerCell("الحد الأدنى المقبول", 3346)] }),
          new TableRow({ children: [dataCell("وقت تحميل الصفحة الرئيسية", 3200, COLORS.lightBlue), dataCell("< 2 ثانية", 3200, COLORS.greenLight), dataCell("< 3 ثوانٍ", 3346)] }),
          new TableRow({ children: [dataCell("نسبة توافر الموقع (Uptime)", 3200, COLORS.lightBlue), dataCell("99.9%", 3200, COLORS.greenLight), dataCell("99.5%", 3346)] }),
          new TableRow({ children: [dataCell("نسبة تغطية الاختبار (Test Coverage)", 3200, COLORS.lightBlue), dataCell("90%", 3200, COLORS.greenLight), dataCell("80%", 3346)] }),
          new TableRow({ children: [dataCell("رضا المستخدمين (بعد الإطلاق)", 3200, COLORS.lightBlue), dataCell("≥ 4.5 / 5", 3200, COLORS.greenLight), dataCell("≥ 4.0 / 5", 3346)] }),
          new TableRow({ children: [dataCell("عدد الأخطاء الحرجة بعد الإطلاق", 3200, COLORS.lightBlue), dataCell("صفر", 3200, COLORS.greenLight), dataCell("≤ 2 أخطاء", 3346)] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 9: COMMUNICATION ============
      sectionTitle("9️⃣ خطة الاتصالات — Communication Plan"),

      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2000, 2000, 1500, 1500, 1500, 1246],
        rows: [
          new TableRow({ children: [
            headerCell("نوع الاتصال", 2000), headerCell("المستلمون", 2000), headerCell("التكرار", 1500),
            headerCell("الوسيلة", 1500), headerCell("المسؤول", 1500), headerCell("الشكل", 1246)
          ]}),
          new TableRow({ children: [dataCell("Daily Stand-up", 2000, COLORS.lightBlue, true), dataCell("فريق التطوير", 2000), dataCell("يومياً", 1500), dataCell("MS Teams / Slack", 1500), dataCell("Scrum Master", 1500), dataCell("شفهي 15 دقيقة", 1246)] }),
          new TableRow({ children: [dataCell("تقرير الحالة الأسبوعي", 2000, COLORS.lightBlue, true), dataCell("مدير المشروع + CTO", 2000), dataCell("أسبوعياً", 1500), dataCell("Email + Dashboard", 1500), dataCell("مدير المشروع", 1500), dataCell("تقرير مكتوب", 1246)] }),
          new TableRow({ children: [dataCell("اجتماع أصحاب المصلحة", 2000, COLORS.lightBlue, true), dataCell("رئيس الجامعة + الإدارة", 2000), dataCell("شهرياً", 1500), dataCell("اجتماع وجاهي / Zoom", 1500), dataCell("مدير المشروع", 1500), dataCell("عرض PPT", 1246)] }),
          new TableRow({ children: [dataCell("تقرير المرحلة (Milestone)", 2000, COLORS.lightBlue, true), dataCell("جميع أصحاب المصلحة", 2000), dataCell("عند كل مرحلة", 1500), dataCell("Email + وثيقة رسمية", 1500), dataCell("مدير المشروع", 1500), dataCell("وثيقة PDF", 1246)] }),
          new TableRow({ children: [dataCell("إشعارات المخاطر", 2000, COLORS.lightBlue, true), dataCell("مدير المشروع + Sponsor", 2000), dataCell("عند الحاجة", 1500), dataCell("Email فوري", 1500), dataCell("Risk Manager", 1500), dataCell("تنبيه فوري", 1246)] }),
          new TableRow({ children: [dataCell("وثائق التغيير (ECR)", 2000, COLORS.lightBlue, true), dataCell("Change Control Board", 2000), dataCell("عند الطلب", 1500), dataCell("نظام PM الرقمي", 1500), dataCell("مدير المشروع", 1500), dataCell("نموذج رسمي", 1246)] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ============ SECTION 10: CLOSURE ============
      sectionTitle("🔟 إغلاق المشروع — Project Closure"),

      subTitle("معايير القبول والإغلاق الرسمي"),
      bulletItem("إطلاق الموقع بنجاح على خوادم الإنتاج"),
      bulletItem("اجتياز جميع اختبارات القبول (UAT) بنسبة نجاح ≥ 95%"),
      bulletItem("توقيع العميل على وثيقة القبول الرسمي"),
      bulletItem("تسليم كامل الوثائق التقنية وكود المصدر"),
      bulletItem("إتمام برنامج التدريب لجميع الكوادر المعنية"),
      bulletItem("لا توجد أخطاء حرجة مفتوحة (Critical Bugs)"),

      new Paragraph({ spacing: { before: 160 }, children: [new TextRun("")] }),
      subTitle("الدروس المستفادة (Lessons Learned)"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [800, 4000, 4946],
        rows: [
          new TableRow({ children: [headerCell("#", 800), headerCell("الدرس المستفاد", 4000), headerCell("التوصية للمشاريع المستقبلية", 4946)] }),
          new TableRow({ children: [dataCell("1", 800, COLORS.lightBlue, false, true), dataCell("أهمية تجميد النطاق مبكراً", 4000), dataCell("إعداد وثيقة Scope Freeze وتوقيعها مع جميع أصحاب المصلحة", 4946)] }),
          new TableRow({ children: [dataCell("2", 800, COLORS.lightBlue, false, true), dataCell("الاختبار المبكر يوفر تكاليف التصليح", 4000), dataCell("تبني منهجية TDD (Test Driven Development) منذ البداية", 4946)] }),
          new TableRow({ children: [dataCell("3", 800, COLORS.lightBlue, false, true), dataCell("إشراك المستخدمين في التصميم", 4000), dataCell("عقد جلسات User Research قبل البدء بالتصميم", 4946)] }),
          new TableRow({ children: [dataCell("4", 800, COLORS.lightBlue, false, true), dataCell("الأمن من البداية لا النهاية", 4000), dataCell("دمج Security by Design في كل مرحلة تطوير", 4946)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 200 }, children: [new TextRun("")] }),
      subTitle("وثيقة الإغلاق الرسمي"),
      new Table({
        width: { size: 9746, type: WidthType.DXA },
        columnWidths: [2400, 4000, 1800, 1546],
        rows: [
          new TableRow({ children: [headerCell("الجانب", 2400), headerCell("التفاصيل", 4000), headerCell("الحالة", 1800), headerCell("التوقيع", 1546)] }),
          new TableRow({ children: [dataCell("تسليم المنتج", 2400, COLORS.lightBlue, true), dataCell("الموقع الإلكتروني المطلق على الإنتاج", 4000), dataCell("✅ مكتمل", 1800, COLORS.greenLight), dataCell("____________", 1546)] }),
          new TableRow({ children: [dataCell("الميزانية النهائية", 2400, COLORS.lightBlue, true), dataCell("348,700 جنيه من أصل 350,000 جنيه مخطط", 4000), dataCell("✅ ضمن الميزانية", 1800, COLORS.greenLight), dataCell("____________", 1546)] }),
          new TableRow({ children: [dataCell("الجدول الزمني", 2400, COLORS.lightBlue, true), dataCell("26 أسبوع من أصل 26 أسبوع مخطط", 4000), dataCell("✅ في الموعد", 1800, COLORS.greenLight), dataCell("____________", 1546)] }),
          new TableRow({ children: [dataCell("التوثيق", 2400, COLORS.lightBlue, true), dataCell("وثائق تقنية كاملة وأدلة المستخدم", 4000), dataCell("✅ مسلّم", 1800, COLORS.greenLight), dataCell("____________", 1546)] }),
          new TableRow({ children: [dataCell("التدريب", 2400, COLORS.lightBlue, true), dataCell("200 موظف تلقوا التدريب الكامل", 4000), dataCell("✅ منجز", 1800, COLORS.greenLight), dataCell("____________", 1546)] }),
        ]
      }),

      new Paragraph({ spacing: { before: 300 }, children: [new TextRun("")] }),
      colorBox("🎓 خُتم هذا المشروع رسمياً بنجاح في يونيو 2024 بموافقة جميع أصحاب المصلحة المعنيين.\nمدير المشروع: ________________________\nالجهة الراعية (رئيس الجامعة): ________________________\nتاريخ الإغلاق الرسمي: يونيو 2024", COLORS.accent),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/project1_azhar_website.docx", buffer);
  console.log("✅ Project 1 created successfully!");
});
