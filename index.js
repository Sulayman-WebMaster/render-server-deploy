const express = require("express");
const cors = require("cors");
const multer = require("multer");
const fs = require("fs");
const XLSX = require("xlsx");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  PageBreak,
} = require("docx");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(cors());
app.use(express.json());

// === Endpoint 1: Top Sheet Generator ===
app.post("/generate", upload.single("excel"), async (req, res) => {
  try {
    const filePath = req.file.path;
    const subjectCode = req.body.subjectCode;
    const absentRolls = JSON.parse(req.body.absentRolls || "[]");

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const filtered = data.filter((row) =>
      Object.values(row).includes(parseInt(subjectCode))
    );

    const allRolls = filtered
      .filter((row) => row.roll !== undefined && row.roll !== null)
      .map((row) => row.roll.toString());

    // Group into 200 present students
    const groups = [];
    let presentCount = 0;
    let i = 0;
    let currentGroup = [];
    let currentGroupAbsent = [];

    while (i < allRolls.length) {
      const roll = allRolls[i];
      const isAbsent = absentRolls.includes(roll);

      currentGroup.push(roll);
      if (isAbsent) {
        currentGroupAbsent.push(roll);
      } else {
        presentCount++;
      }

      if (presentCount === 200) {
        groups.push({
          fullGroup: [...currentGroup],
          absents: [...currentGroupAbsent],
        });
        currentGroup = [];
        currentGroupAbsent = [];
        presentCount = 0;
      }

      i++;
    }

    if (currentGroup.length > 0) {
      groups.push({
        fullGroup: [...currentGroup],
        absents: [...currentGroupAbsent],
      });
    }

    const sections = groups.map((group, index) => {
      const present = group.fullGroup.filter(
        (r) => !group.absents.includes(r)
      );

      const rollRanges = [];
      let i = 0;
      while (i < present.length) {
        const start = present[i];
        let end = start;
        let count = 1;

        while (
          i + 1 < present.length &&
          parseInt(present[i + 1]) === parseInt(present[i]) + 1
        ) {
          end = present[i + 1];
          count++;
          i++;
        }

        if (start === end) {
          rollRanges.push(`${start}`);
        } else {
          rollRanges.push(`${start}---${end}=${count}`);
        }

        i++;
      }

      const rollRangeText = rollRanges.join(", ");
      const absentText = group.absents.length
        ? group.absents.join(", ")
        : "0";

      const children = [
        new Paragraph({
          children: [
            new TextRun({
              text: `Group ${index + 1}`,
              bold: true,
              size: 28,
            }),
          ],
        }),
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun(`Roll Range: ${rollRangeText}`)],
        }),
        new Paragraph({
          children: [new TextRun(`Absent: ${absentText}`)],
        }),
      ];

      if (index < groups.length - 1) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
      }

      return { children };
    });

    const doc = new Document({
      creator: "Top Sheet Generator",
      title: "Student Top Sheet",
      description: "200-present-per-group layout",
      sections,
    });

    const buffer = await Packer.toBuffer(doc);
    fs.unlinkSync(filePath);

    res.setHeader("Content-Disposition", "attachment; filename=TopSheet.docx");
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating top sheet");
  }
});

// === Endpoint 2: Subject-wise Roll Column Sheet ===
app.post("/generate-subject-rolls", upload.single("excel"), async (req, res) => {
  try {
    const subjectCode = req.body.subjectCode;
    const filePath = req.file.path;

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const rolls = data
      .filter((row) => Object.values(row).includes(parseInt(subjectCode)))
      .map((row) => row.roll?.toString())
      .filter(Boolean);

    const columns = 6;
    const maxRows = 48;
    const rollsPerPage = columns * maxRows;

    const pages = [];
    for (let i = 0; i < rolls.length; i += rollsPerPage) {
      pages.push(rolls.slice(i, i + rollsPerPage));
    }

    const sections = pages.map((pageRolls, pageIndex) => {
      const rows = [];

      const colHeight = Math.ceil(pageRolls.length / columns);
      const columnData = [];

      for (let col = 0; col < columns; col++) {
        columnData.push(
          pageRolls.slice(col * colHeight, (col + 1) * colHeight)
        );
      }

      for (let row = 0; row < colHeight; row++) {
        const line = [];

        for (let col = 0; col < columns; col++) {
          const val = columnData[col][row] || "";
          line.push(val.padEnd(8, " "));
        }

        rows.push(
          new Paragraph({
            children: [new TextRun({ text: line.join("   "), font: "Courier New", size: 24 })],
          })
        );
      }

      if (pageIndex === pages.length - 1) {
        rows.push(
          new Paragraph({
            spacing: { before: 300 },
            children: [
              new TextRun({
                text: `Total: ${rolls.length} students`,
                bold: true,
                size: 28,
              }),
            ],
          })
        );
      }

      if (pageIndex !== pages.length - 1) {
        rows.push(new Paragraph({ children: [new PageBreak()] }));
      }

      return {
        children: [
          new Paragraph({
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: `Rolls for Subject: ${subjectCode}`,
                bold: true,
                size: 30,
              }),
            ],
          }),
          ...rows,
        ],
      };
    });

    const doc = new Document({ sections });
    const buffer = await Packer.toBuffer(doc);
    fs.unlinkSync(filePath);

    res.setHeader("Content-Disposition", `attachment; filename=Subject-${subjectCode}.docx`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating subject-wise roll list");
  }
});

app.listen(5000, () => {
  console.log("🚀 Server started on http://localhost:5000");
});
