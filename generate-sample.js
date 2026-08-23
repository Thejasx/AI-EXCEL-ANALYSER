import XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sampleData = [
  { "ID": 101, "Date": "2026-01-05", "Region": "North", "Sales Rep": "Alice", "Product": "Laptop Pro", "Category": "Electronics", "Units": 12, "Unit Price": 1200, "Total Revenue": 14400, "Status": "Completed" },
  { "ID": 102, "Date": "2026-01-07", "Region": "South", "Sales Rep": "Bob", "Product": "Smart Monitor 27\"", "Category": "Electronics", "Units": 25, "Unit Price": 350, "Total Revenue": 8750, "Status": "Completed" },
  { "ID": 103, "Date": "2026-01-10", "Region": "East", "Sales Rep": "Charlie", "Product": "Ergo Chair", "Category": "Furniture", "Units": 8, "Unit Price": 450, "Total Revenue": 3600, "Status": "Pending" },
  { "ID": 104, "Date": "2026-01-12", "Region": "West", "Sales Rep": "Diana", "Product": "Mechanical Keyboard", "Category": "Electronics", "Units": 45, "Unit Price": 110, "Total Revenue": 4950, "Status": "Completed" },
  { "ID": 105, "Date": "2026-01-15", "Region": "North", "Sales Rep": "Alice", "Product": "Wireless Mouse", "Category": "Electronics", "Units": 60, "Unit Price": 45, "Total Revenue": 2700, "Status": "Completed" },
  { "ID": 106, "Date": "2026-01-18", "Region": "South", "Sales Rep": "Bob", "Product": "Standing Desk", "Category": "Furniture", "Units": 15, "Unit Price": 750, "Total Revenue": 11250, "Status": "Completed" },
  { "ID": 107, "Date": "2026-01-20", "Region": "East", "Sales Rep": "Eve", "Product": "USB-C Hub", "Category": "Electronics", "Units": 100, "Unit Price": 30, "Total Revenue": 3000, "Status": "Shipped" },
  { "ID": 108, "Date": "2026-01-22", "Region": "West", "Sales Rep": "Diana", "Product": "UltraWide Monitor", "Category": "Electronics", "Units": 18, "Unit Price": 850, "Total Revenue": 15300, "Status": "Completed" },
  { "ID": 109, "Date": "2026-01-25", "Region": "North", "Sales Rep": "Frank", "Product": "Noise Cancelling Headphones", "Category": "Electronics", "Units": 30, "Unit Price": 250, "Total Revenue": 7500, "Status": "Pending" },
  { "ID": 110, "Date": "2026-01-28", "Region": "South", "Sales Rep": "Bob", "Product": "Webcam 4K", "Category": "Electronics", "Units": 40, "Unit Price": 120, "Total Revenue": 4800, "Status": "Completed" }
];

function generateExcel() {
  const worksheet = XLSX.utils.json_to_sheet(sampleData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Data");
  
  const outputPath = path.join(__dirname, 'sales_data.xlsx');
  XLSX.writeFile(workbook, outputPath);
  console.log(`✅ Sample Excel file successfully generated at: ${outputPath}`);
}

generateExcel();
