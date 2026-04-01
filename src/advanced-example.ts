import XlsxTemplate from "xlsx-template";
import * as fs from "fs";
import * as path from "path";

/**
 * Advanced example showing how to extend the base project
 * with multiple records, arrays, and conditional data
 */

interface LineItem {
  test_name: string;
  result: string;
  specification: string;
  unit: string;
}

interface AdvancedCoAData {
  invoice_no: string;
  container_no: string;
  serial_no: string;
  date: string;
  batch_no: string;
  production_date: string;
  expired_date: string;
  test_results?: LineItem[];
  total_tests?: number;
  approved_by?: string;
  approval_date?: string;
}

function generateAdvancedCoA() {
  try {
    const templatePath = path.join(__dirname, "../COA.xlsx");
    const outputPath = path.join(__dirname, "../COA_Advanced_Output.xlsx");

    if (!fs.existsSync(templatePath)) {
      console.error(`Template file not found at: ${templatePath}`);
      return;
    }

    console.log("Loading template for advanced export...");

    const template: any = new (XlsxTemplate as any)(
      fs.readFileSync(templatePath),
    );

    // Extended data with arrays for repeating rows
    const advancedData: AdvancedCoAData = {
      invoice_no: "FG-DF-0309-01",
      container_no: "FBIU5605008",
      serial_no: "SITR486555",
      date: "March 30, 2026",
      batch_no: "20003/0021",
      production_date: "21/4/2026",
      expired_date: "20/4/28",
      approved_by: "John Smith",
      approval_date: "March 31, 2026",
      total_tests: 3,
      test_results: [
        {
          test_name: "Microbial Count",
          result: "< 1000",
          specification: "< 10000",
          unit: "CFU/g",
        },
        {
          test_name: "pH Level",
          result: "6.5",
          specification: "6.0-7.0",
          unit: "",
        },
        {
          test_name: "Moisture",
          result: "12.3%",
          specification: "< 13.0%",
          unit: "%",
        },
      ],
    };

    // Substitute all data
    template.substituteAll(advancedData);

    // Generate and save
    const output = template.generate();
    fs.writeFileSync(outputPath, output, "binary");

    console.log("✓ Advanced Excel file generated successfully!");
    console.log("Output file:", outputPath);
    console.log("\nNote: This example shows how to structure data for tables.");
    console.log("Array data (test_results) can populate rows in Excel tables.");
  } catch (error) {
    console.error("Error:", error);
  }
}

// Uncomment to run:
// generateAdvancedCoA();

export { generateAdvancedCoA, AdvancedCoAData };
