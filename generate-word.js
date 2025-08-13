const { Document, Packer, Paragraph, TextRun } = require("docx");
const express = require("express");
const app = express();
const cors = require("cors");

app.use(cors());
app.use(express.json());

app.post("/api/generate-word", async (req, res) => {
    const data = req.body;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ text: "تقرير زيارة إشرافية", bold: true, size: 32 }),
                    ],
                }),
                new Paragraph({ text: `اسم المعلمة: ${data.teacherName || ""}` }),
                new Paragraph({ text: `المادة: ${data.subject || ""}` }),
                new Paragraph({ text: `التاريخ: ${data.date || ""}` }),
                new Paragraph({ text: `الملاحظات: ${data.notes || ""}` }),
            ],
        }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Disposition", "attachment; filename=visit-report.docx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.send(buffer);
});

module.exports = app;
