import XlsxTemplate from "xlsx-template";
import * as fs from "fs";
import * as path from "path";

interface CoAData {
  invoice_no: string;
  container_no: string;
  serial_no: string;
  date: string;
  batch_no: string;
  produciont_date: string;
  expired_date: string;
}

const data: CoAData = {
  invoice_no: "FG-DF-0309-01",
  container_no: "FBIU5605008",
  serial_no: "SITR486555",
  date: "March 30, 2026",
  batch_no: "20003/0021",
  produciont_date: "21/4/2026",
  expired_date: "20/4/28",
};

async function generateCoA() {
  try {
    const templatePath = path.join(__dirname, "../COA.xlsx");
    const outputPath = path.join(__dirname, "../COA_Output.xlsx");

    // Check if template file exists
    if (!fs.existsSync(templatePath)) {
      console.error(`Template file not found at: ${templatePath}`);
      return;
    }

    console.log("Loading template:", templatePath);

    // Create a new XlsxTemplate instance from the template file
    const template: any = new (XlsxTemplate as any)(
      fs.readFileSync(templatePath),
    );

    // Use substituteAll to fill placeholders in all sheets
    // The template uses {{key}} placeholders that get replaced with values
    template.substituteAll(data);

    // Generate the output file
    const output = template.generate();

    // Write the output to file
    fs.writeFileSync(outputPath, output, "binary");

    console.log("✓ Excel file generated successfully!");
    console.log("Output file:", outputPath);
    console.log("\nData populated:");
    Object.entries(data).forEach(([key, value]) => {
      console.log(`  ${key}: ${value}`);
    });
  } catch (error) {
    console.error("Error generating Excel file:", error);
    console.error(
      "\nNote: Make sure your COA.xlsx template has placeholders like:",
    );
    console.error("  {{invoice_no}}, {{container_no}}, {{serial_no}}, etc.");
    console.error(
      "\nXLSX-Template requires placeholders in cells to work properly.",
    );
    process.exit(1);
  }
}

generateCoA();
