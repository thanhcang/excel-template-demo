import XlsxTemplate from "xlsx-template";
import * as fs from "fs";
import * as path from "path";

function inspect() {
  try {
    const templatePath = path.join(__dirname, "../COA.xlsx");

    if (!fs.existsSync(templatePath)) {
      console.error(`Template file not found at: ${templatePath}`);
      return;
    }

    console.log("Loading template:", templatePath);
    const buffer = fs.readFileSync(templatePath);
    const template: any = new (XlsxTemplate as any)(buffer);

    console.log("\n=== Template Structure ===");
    console.log("Sheets:", template.sheets);
    console.log("Sheet:", template.sheet);

    if (template.sheets) {
      console.log("\nSheets array:");
      template.sheets.forEach((sheet: any, idx: number) => {
        console.log(`  [${idx}]:`, {
          name: sheet.name || sheet.$,
          id: sheet.id,
          entries: sheet.entries,
          column: sheet.column,
        });
      });
    }

    console.log("\n=== Trying substituteAll ===");
    const allData = {
      invoice_no: "FG-DF-0309-01",
      container_no: "FBIU5605008",
      serial_no: "SITR486555",
      date: "March 30, 2026",
      batch_no: "20003/0021",
      production_date: "21/4/2026",
      expired_date: "20/4/28",
    };

    template.substituteAll(allData);
    const output = template.generate();

    const outputPath = path.join(__dirname, "../COA_Output.xlsx");
    fs.writeFileSync(outputPath, output, "binary");

    console.log("✓ Successfully generated with substituteAll!");
    console.log("Output file:", outputPath);
  } catch (error) {
    console.error("Error:", error);
  }
}

inspect();
