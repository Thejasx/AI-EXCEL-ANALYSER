import express from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import { GoogleGenAI, Type } from '@google/genai';
import dotenv from 'dotenv';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS and JSON parsing
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Configure Multer for in-memory file uploads
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 } // 20 MB limit
});

// Initialize Gemini Client using @google/genai SDK
const getGeminiClient = () => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey || apiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is missing or invalid in environment variables.');
  }
  return new GoogleGenAI({ apiKey });
};

/**
 * Endpoint: POST /api/upload
 * Reads uploaded Excel file (.xlsx, .xls, .csv) and converts sheet data to JSON.
 */
app.post('/api/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No Excel file provided.' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetNames = workbook.SheetNames;

    if (!sheetNames.length) {
      return res.status(400).json({ error: 'The uploaded file has no sheets.' });
    }

    const targetSheetName = sheetNames[0];
    const worksheet = workbook.Sheets[targetSheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

    res.json({
      success: true,
      fileName: req.file.originalname,
      sheetNames,
      activeSheet: targetSheetName,
      rowCount: data.length,
      data
    });
  } catch (error) {
    console.error('Error parsing Excel file:', error);
    res.status(500).json({ error: `Failed to parse Excel file: ${error.message}` });
  }
});

/**
 * Endpoint: POST /api/analyze
 * Accepts JSON dataset + user instructions. Uses gemini-2.5-flash with structured output schema.
 */
app.post('/api/analyze', async (req, res) => {
  try {
    const { data, prompt } = req.body;

    if (!data || !Array.isArray(data) || data.length === 0) {
      return res.status(400).json({ error: 'Dataset is missing or empty.' });
    }

    if (!prompt || typeof prompt !== 'string' || !prompt.trim()) {
      return res.status(400).json({ error: 'Analysis instruction prompt is required.' });
    }

    const ai = getGeminiClient();

    const systemPrompt = `You are a world-class data scientist and Excel automation expert.
You are given a JSON array representing rows of an Excel spreadsheet.
Your task is to execute the user's natural language instruction precisely on the dataset.

User Instruction: "${prompt.trim()}"

Dataset JSON:
${JSON.stringify(data, null, 2)}

Respond strictly in valid JSON with the following structure:
{
  "summary": "Detailed summary of the data transformations, sorting, filtering, or aggregations performed, along with key statistical observations and insights.",
  "processedData": [
    // Array of row objects reflecting the requested sorting, filtering, aggregations, or column calculations. Preserve all column field names and include any newly calculated fields.
  ]
}`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.6-flash',
      contents: systemPrompt,
      config: {
        responseMimeType: 'application/json'
      }
    });

    const text = response.text;
    const resultJson = JSON.parse(text);

    res.json({
      success: true,
      summary: resultJson.summary || 'Data analysis complete.',
      processedData: resultJson.processedData || []
    });
  } catch (error) {
    console.error('Error during AI analysis:', error);
    res.status(500).json({
      error: error.message || 'An error occurred while communicating with Gemini API.'
    });
  }
});

/**
 * Endpoint: POST /api/export
 * Converts processed JSON data back into an Excel sheet (.xlsx) for download.
 */
app.post('/api/export', (req, res) => {
  try {
    const { processedData, fileName } = req.body;

    if (!processedData || !Array.isArray(processedData) || processedData.length === 0) {
      return res.status(400).json({ error: 'No processed data available to export.' });
    }

    const worksheet = XLSX.utils.json_to_sheet(processedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Analyzed Data');

    const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    const downloadFileName = fileName 
      ? `Analyzed_${path.basename(fileName, path.extname(fileName))}.xlsx` 
      : 'Analyzed_Excel_Data.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadFileName}"`);
    res.send(excelBuffer);
  } catch (error) {
    console.error('Error exporting Excel file:', error);
    res.status(500).json({ error: `Failed to generate Excel file: ${error.message}` });
  }
});

if (process.env.NODE_ENV !== 'production' || !process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`🚀 Real-Time Excel Sheet Analyzer server running at http://localhost:${PORT}`);
  });
}

export default app;
