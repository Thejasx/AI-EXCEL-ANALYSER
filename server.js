import express from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import vm from 'vm';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS and generous JSON parsing limits
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Configure Multer for in-memory file uploads (up to 50 MB)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
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
 * Robust JSON cleaning and parsing helper to handle markdown fences, 
 * conversational prefixes, smart quotes, and trailing commas.
 */
function cleanAndParseJson(rawText) {
  if (!rawText || typeof rawText !== 'string') {
    throw new Error('Empty or non-string response received from Gemini API.');
  }

  let text = rawText.trim();

  // Remove markdown code fences if present (e.g. ```json ... ``` or ``` ...)
  if (text.includes('```')) {
    text = text.replace(/```(?:json|javascript|js)?/gi, '').replace(/```/g, '').trim();
  }

  // Extract content between the first '{' and last '}'
  const firstBrace = text.indexOf('{');
  const lastBrace = text.lastIndexOf('}');

  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    text = text.substring(firstBrace, lastBrace + 1);
  }

  try {
    return JSON.parse(text);
  } catch (firstErr) {
    try {
      // Attempt minor sanitization for smart quotes and trailing commas
      const sanitized = text
        .replace(/,\s*([\}\]])/g, '$1')
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/[\u2018\u2019]/g, "'");
      return JSON.parse(sanitized);
    } catch (secondErr) {
      console.error('Failed to parse Gemini output as JSON. Raw text was:\n', rawText);
      throw new Error(`Invalid JSON format from AI response: ${firstErr.message}`);
    }
  }
}

/**
 * Executes a JavaScript transformation string safely in a sandboxed VM environment.
 */
function executeTransformationCode(codeStr, dataset) {
  let cleanCode = codeStr.trim();
  if (cleanCode.startsWith('```')) {
    cleanCode = cleanCode.replace(/```(?:javascript|js)?/gi, '').replace(/```/g, '').trim();
  }

  // Prepare sandboxed execution context with standard JS built-ins
  const context = {
    data: dataset,
    Math,
    Array,
    Object,
    String,
    Number,
    Boolean,
    Date,
    RegExp,
    console: { log: () => {} },
    result: null
  };

  vm.createContext(context);

  let scriptCode;
  if (cleanCode.startsWith('(') || cleanCode.startsWith('data =>') || cleanCode.startsWith('function')) {
    scriptCode = `result = (${cleanCode})(data);`;
  } else {
    scriptCode = `const transformFn = ${cleanCode};\nresult = transformFn(data);`;
  }

  const script = new vm.Script(scriptCode);
  script.runInContext(context, { timeout: 10000 }); // 10 second timeout

  if (!Array.isArray(context.result)) {
    throw new Error('The transformation code did not return a valid array of row objects.');
  }

  return context.result;
}

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

    if (!sheetNames || !sheetNames.length) {
      return res.status(400).json({ error: 'The uploaded file contains no sheets.' });
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
 * Generates JS transformation code using Gemini based on dataset schema & sample rows.
 * Supports lightweight metadata payload (< 5 KB) to bypass Vercel serverless body limits (4.5 MB).
 */
app.post('/api/analyze', async (req, res) => {
  try {
    const { data, columns: reqColumns, sampleRows: reqSampleRows, totalRows: reqTotalRows, prompt } = req.body;

    if (!prompt || typeof prompt !== 'string' || !prompt.trim()) {
      return res.status(400).json({ error: 'Analysis instruction prompt is required.' });
    }

    let columns = [];
    let sampleRows = [];
    let totalRows = 0;
    let fullDataset = null;

    if (Array.isArray(data) && data.length > 0) {
      fullDataset = data;
      totalRows = data.length;
      columns = Object.keys(data[0] || {});
      sampleRows = data.slice(0, Math.min(10, totalRows));
    } else if (Array.isArray(reqColumns) && Array.isArray(reqSampleRows)) {
      columns = reqColumns;
      sampleRows = reqSampleRows;
      totalRows = reqTotalRows || reqSampleRows.length;
    } else {
      return res.status(400).json({ error: 'Dataset metadata or rows are missing.' });
    }

    const ai = getGeminiClient();

    const systemPrompt = `You are an expert data scientist and JavaScript spreadsheet automation engine.
You are given metadata and sample rows from an Excel dataset with ${totalRows} total rows.

Dataset Column Names: ${JSON.stringify(columns)}

Sample Data (First ${sampleRows.length} rows):
${JSON.stringify(sampleRows, null, 2)}

User Instruction: "${prompt.trim()}"

Your task:
1. Provide a clear, professional summary of the transformation, sorting, filtering, aggregations, or column calculations requested.
2. Write a clean, high-performance JavaScript arrow function that takes the full dataset array 'data' and returns the transformed array 'processedData'.

Requirements for the JS code:
- The function signature MUST be: \`data => { /* transform logic */ return processedData; }\`
- Do NOT use external libraries or network requests.
- Use standard JS features (Array.prototype.filter, map, sort, reduce, Object.values, Math, Date, RegExp, etc.).
- Preserve existing column fields while adding any new calculated columns.
- Ensure null/undefined safety when reading row properties.

Respond STRICTLY in valid JSON matching this exact structure:
{
  "summary": "Detailed human-readable explanation of transformations and observations.",
  "code": "data => data.filter(...).map(...)"
}`;

    const modelName = process.env.GEMINI_MODEL || 'gemini-3.6-flash';

    const response = await ai.models.generateContent({
      model: modelName,
      contents: systemPrompt,
      config: {
        responseMimeType: 'application/json'
      }
    });

    const responseText = response.text;
    const resultJson = cleanAndParseJson(responseText);

    let processedData = null;

    if (fullDataset && resultJson.code && typeof resultJson.code === 'string') {
      try {
        processedData = executeTransformationCode(resultJson.code, fullDataset);
      } catch (codeErr) {
        console.warn('VM Execution of Gemini JS code failed, falling back:', codeErr.message);
        processedData = resultJson.processedData || fullDataset;
      }
    }

    res.json({
      success: true,
      summary: resultJson.summary || 'Data analysis and transformation complete.',
      code: resultJson.code || null,
      rowCount: processedData ? processedData.length : totalRows,
      processedData
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

    const baseName = fileName ? path.basename(fileName, path.extname(fileName)) : 'Data';
    const downloadFileName = `Analyzed_${baseName}.xlsx`;

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
