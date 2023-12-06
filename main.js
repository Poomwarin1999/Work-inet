const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
const port = 3000;

// กำหนดที่เก็บไฟล์ที่อัปโหลด
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// const XLSX = require("xlsx");
const fs = require("fs");

app.post("/upload-excel-check", upload.array("excelFiles", 2), (req, res) => {
  try {
    const excelFiles = req.files;
    const excelFile1 = [];
    const excelFile2 = [];

    excelFiles.forEach((excelFile, index) => {
      const workbook = XLSX.read(excelFile.buffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const excelData = XLSX.utils.sheet_to_json(sheet);

      if (index === 0) {
        excelFile1.push(excelData);
      } else if (index === 1) {
        excelFile2.push(excelData);
      }
    });

    let missingEmployeeCodes = [];
    let missingAccountIdCodes = [];
    let additionalData = [];

    for (let dataArray of excelFile1) {
      for (let data of dataArray) {
        const trimmedEmployee1 = data.employee.trim();
        let foundEmployee = false;
        let foundAccountId = false;

        for (let dataArray2 of excelFile2) {
          for (let data2 of dataArray2) {
            if (trimmedEmployee1 == data2.รหัส) {
              foundEmployee = true;
            }

            if (data.accountId == data2["Account ID"]) {
              foundAccountId = true;
            }

            if (data.accountId == data2["Account ID"] && trimmedEmployee1 != data2.รหัส) {
              additionalData.push(data);
            }

            if (foundEmployee && foundAccountId) {
              break;
            }
          }

          if (foundEmployee && foundAccountId) {
            break;
          }
        }

        if (!foundEmployee) {
          missingEmployeeCodes.push(data.employee);
        }

        if (!foundAccountId) {
          missingAccountIdCodes.push(data.accountId);
        }
      }
    }

    // สร้าง Worksheet สำหรับ missingEmployeeCodes
    const wsEmployee = XLSX.utils.json_to_sheet(
      missingEmployeeCodes.map((code) => ({ employee: code }))
    );

    // สร้าง Worksheet สำหรับ missingAccountIdCodes
    const wsAccountId = XLSX.utils.json_to_sheet(
      missingAccountIdCodes.map((code) => ({ accountId: code }))
    );

    const wsCheck = XLSX.utils.json_to_sheet(
      additionalData.map((code) => ({ check: code }))
    );

    // สร้าง Workbook และเพิ่ม Worksheet
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsEmployee, "รหัสพนักงานที่ไม่เจอ");
    XLSX.utils.book_append_sheet(wb, wsAccountId, "accountIdที่ไม่เจอ");
    XLSX.utils.book_append_sheet(wb, wsCheck, "Check");

    // สร้างชื่อไฟล์ที่ไม่ซ้ำ (เช่น timestamp)
    const timestamp = Date.now();
    const filename = `missing_data_${timestamp}.xlsx`;
    const filePath = `./uploads/${filename}`;

    // เขียนไฟล์ Excel
    XLSX.writeFile(wb, filePath);

    // ส่งพาธไฟล์ในการตอบกลับ
    res.status(200).json({
      success: true,
      countEmployee: missingEmployeeCodes.length,
      dataEmployee: missingEmployeeCodes,
      countAccountId: missingAccountIdCodes.length,
      dataAccountId: missingAccountIdCodes,
      excelFilePath: filePath,
    });
  } catch (error) {
    console.error(error);
    res
      .status(500)
      .json({ success: false, error: "เกิดข้อผิดพลาดในการประมวลผล" });
  }
});

// เริ่มต้นเซิร์ฟเวอร์
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
