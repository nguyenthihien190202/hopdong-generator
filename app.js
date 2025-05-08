const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const mammoth = require("mammoth");

const app = express();
const port = process.env.PORT || 3000;


// Cho phép truy cập file tĩnh (HTML form)
app.use(express.static("public"));

// Cấu hình multer upload
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, "uploads/"),
    filename: (req, file, cb) => cb(null, file.fieldname + "-" + Date.now() + ".docx"),
});
const upload = multer({ storage: storage });

// Trích xuất dữ liệu từ data.docx (dạng key: value)
async function extractDataFromDocx(filePath) {
    const result = await mammoth.extractRawText({ path: filePath });
    const lines = result.value.split("\n");
    const data = {};
    lines.forEach((line) => {
        const [key, ...rest] = line.split(":");
        if (key && rest.length > 0) {
            data[key.trim()] = rest.join(":").trim();
        }
    });
    return data;
}

// Route tạo hợp đồng
app.post("/generate", upload.fields([{ name: "template" }, { name: "data" }]), async (req, res) => {
    try {
        const templatePath = req.files.template[0].path;
        const dataPath = req.files.data[0].path;

        // Đọc template
        const templateContent = fs.readFileSync(templatePath, "binary");
        let zip, doc;
        try {
            zip = new PizZip(templateContent);
            doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });
        } catch (err) {
            console.error("❌ Không đọc được template .docx:", err);
            return res.status(400).send("File template không hợp lệ.");
        }

        // Trích xuất dữ liệu từ file data.docx
        const data = await extractDataFromDocx(dataPath);
        doc.setData(data);

        try {
            doc.render();
        } catch (err) {
            console.error("❌ Lỗi khi render hợp đồng:", err);
            return res.status(500).send("Lỗi khi chèn dữ liệu vào hợp đồng.");
        }

        // Ghi ra file output
        const buffer = doc.getZip().generate({ type: "nodebuffer" });
        const outputName = "hop_dong_" + Date.now() + ".docx";
        const outputPath = path.join(__dirname, "output", outputName);
        fs.writeFileSync(outputPath, buffer);

        // Trả file kết quả
        res.download(outputPath, "hop_dong_hoan_chinh.docx");
    } catch (error) {
        console.error("❌ Lỗi không xác định:", error);
        res.status(500).send("Đã xảy ra lỗi khi xử lý hợp đồng.");
    }
});

app.listen(port, () => {
    console.log(`🔥 Server đang chạy tại http://localhost:${port}`);
});
