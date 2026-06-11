require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;
const security = require('./security');

// Initialize Supabase Client if env variables exist
const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;
let supabase = null;

if (SUPABASE_URL && SUPABASE_KEY) {
  supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
  console.log('Supabase connection initialized successfully.');
} else {
  console.log('Using local JSON database (results.json). Set SUPABASE_URL and SUPABASE_KEY to use Supabase.');
}

app.use(cors());
app.use(bodyParser.json());

// Decryption and Encryption Middleware
app.use((req, res, next) => {
  // Enforce encrypted body on all /api/ requests
  if (req.path.startsWith('/api')) {
    if (!req.body || !req.body.data) {
      return res.status(401).json({ error: 'وصول غير مصرح به. البيانات مشفرة مطلوبة.' });
    }
    const decrypted = security.decrypt(req.body.data);
    if (!decrypted) {
      return res.status(401).json({ error: 'فشل فك تشفير البيانات. التوقيع غير صالح.' });
    }
    try {
      req.body = JSON.parse(decrypted);
    } catch (e) {
      return res.status(400).json({ error: 'حملة بيانات غير صالحة' });
    }

    // Anti-replay: Verify timestamp exists and is within 5 minutes (300,000 ms)
    const timestamp = req.body.timestamp;
    if (!timestamp) {
      return res.status(400).json({ error: 'طلب غير صالح. الطابع الزمني مفقود.' });
    }
    const diff = Math.abs(Date.now() - timestamp);
    if (diff > 5 * 60 * 1000) {
      return res.status(400).json({ error: 'طلب منتهي الصلاحية.' });
    }
  }

  // Override res.json to automatically encrypt all responses
  const originalJson = res.json;
  res.json = function (body) {
    if (body && body.data) {
      // Already encrypted
      return originalJson.call(this, body);
    }
    const encryptedData = security.encrypt(JSON.stringify(body));
    return originalJson.call(this, { data: encryptedData });
  };

  next();
});

const EXCEL_PATH = path.join(__dirname, 'employees.xlsx');
const RESULTS_PATH = path.join(__dirname, 'results.json');

// Self-healing: Generate default Excel sheet if it does not exist
if (!fs.existsSync(EXCEL_PATH)) {
  console.log('Generating default employees.xlsx file...');
  const sampleData = [
    {
      ServiceNumber: '1001',
      FullName: 'أحمد محمد',
      MotherName: 'فاطمة',
      NationalID: '12345678901',
      Language1: 'English',
      Level1: 'B1',
      Language2: 'French',
      Level2: 'A2',
      Language3: '',
      Level3: ''
    },
    {
      ServiceNumber: '1002',
      FullName: 'خالد علي',
      MotherName: 'مريم',
      NationalID: '09876543210',
      Language1: 'English',
      Level1: 'C1',
      Language2: '',
      Level2: '',
      Language3: '',
      Level3: ''
    },
    {
      ServiceNumber: '1003',
      FullName: 'سارة أحمد',
      MotherName: 'أمينة',
      NationalID: '11223344556',
      Language1: 'French',
      Level1: 'B2',
      Language2: 'Russian',
      Level2: 'B1',
      Language3: '',
      Level3: ''
    },
    {
      ServiceNumber: '1004',
      FullName: 'فهد سالم',
      MotherName: 'عائشة',
      NationalID: '55667788990',
      Language1: '',
      Level1: '',
      Language2: '',
      Level2: '',
      Language3: '',
      Level3: ''
    }
  ];

  const ws = xlsx.utils.json_to_sheet(sampleData);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Employees');
  xlsx.writeFile(wb, EXCEL_PATH);
  console.log('employees.xlsx generated successfully.');
}

// Read employees data from Excel
function getEmployees() {
  try {
    const workbook = xlsx.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
  } catch (error) {
    console.error('Error reading Excel file:', error);
    return [];
  }
}

// Read results data (from Supabase or local JSON)
async function getResults(serviceNumber = null) {
  if (supabase) {
    try {
      let query = supabase.from('results').select('*');
      if (serviceNumber) {
        query = query.eq('serviceNumber', String(serviceNumber));
      }
      const { data, error } = await query;
      if (error) throw error;
      return data || [];
    } catch (error) {
      console.error('Error reading results from Supabase:', error);
      return [];
    }
  }

  if (!fs.existsSync(RESULTS_PATH)) {
    return [];
  }
  try {
    const data = fs.readFileSync(RESULTS_PATH, 'utf8');
    const results = JSON.parse(data);
    if (serviceNumber) {
      return results.filter(r => String(r.serviceNumber) === String(serviceNumber));
    }
    return results;
  } catch (error) {
    console.error('Error reading results database:', error);
    return [];
  }
}

// Save result (to Supabase or local JSON)
async function saveResult(result) {
  if (supabase) {
    try {
      const { error } = await supabase
        .from('results')
        .insert([
          {
            serviceNumber: String(result.serviceNumber),
            language: result.language,
            score: Number(result.score),
            correctCount: Number(result.correctCount),
            incorrectCount: Number(result.incorrectCount),
            timeTaken: Number(result.timeTaken)
          }
        ]);
      if (error) throw error;
      return true;
    } catch (error) {
      console.error('Error saving result to Supabase:', error);
      return false;
    }
  }

  const results = await getResults();
  results.push({
    ...result,
    timestamp: new Date().toISOString()
  });
  try {
    fs.writeFileSync(RESULTS_PATH, JSON.stringify(results, null, 2), 'utf8');
    return true;
  } catch (error) {
    console.error('Error saving result:', error);
    return false;
  }
}

// API: Verify Employee Login
app.post('/api/login', async (req, res) => {
  const { serviceNumber, fullName, motherName, nationalId } = req.body;
  
  if (!serviceNumber || !fullName || !motherName || !nationalId) {
    return res.status(400).json({ error: 'يرجى ملء كافة الحقول المطلوبة' });
  }

  const employees = getEmployees();
  
  // Find employee matching all fields (trimming whitespaces)
  const employee = employees.find(emp => 
    String(emp.ServiceNumber).trim() === String(serviceNumber).trim() &&
    String(emp.FullName).trim() === String(fullName).trim() &&
    String(emp.MotherName).trim() === String(motherName).trim() &&
    String(emp.NationalID).trim() === String(nationalId).trim()
  );

  if (!employee) {
    return res.status(401).json({ error: 'المعلومات المدخلة غير صحيحة أو غير متطابقة' });
  }

  // Parse language levels: Language1, Level1, Language2, Level2...
  const languagesList = [];
  let index = 1;
  while (true) {
    const langKey = `Language${index}`;
    const lvlKey = `Level${index}`;
    if (employee[langKey] !== undefined) {
      const langName = String(employee[langKey]).trim();
      const langLevel = employee[lvlKey] ? String(employee[lvlKey]).trim() : 'A1';
      if (langName) {
        languagesList.push({ name: langName, level: langLevel });
      }
    } else {
      break;
    }
    index++;
  }

  // Get completed exams for this service number
  const employeeResults = await getResults(serviceNumber);
  const completedExams = employeeResults.map(r => r.language);

  res.json({
    message: 'تم التحقق بنجاح',
    employee: {
      serviceNumber: employee.ServiceNumber,
      fullName: employee.FullName,
      languages: languagesList,
      completedExams: completedExams,
      allResults: employeeResults
    }
  });
});

// API: Submit Exam Results
app.post('/api/submit-exam', async (req, res) => {
  const { serviceNumber, language, score, correctCount, incorrectCount, timeTaken } = req.body;

  if (!serviceNumber || !language) {
    return res.status(400).json({ error: 'بيانات غير مكتملة لتسجيل النتيجة' });
  }

  // Check if exam is already submitted for this language and user
  const results = await getResults(serviceNumber);
  const alreadyTaken = results.some(r => 
    r.language === language
  );

  if (alreadyTaken) {
    return res.status(400).json({ error: 'لقد قمت بإجراء هذا الاختبار مسبقاً، لا يمكنك إعادتة.' });
  }

  const success = await saveResult({
    serviceNumber,
    language,
    score,
    correctCount,
    incorrectCount,
    timeTaken
  });

  if (success) {
    res.json({ message: 'تم حفظ النتيجة بنجاح' });
  } else {
    res.status(500).json({ error: 'حدث خطأ أثناء حفظ النتيجة' });
  }
});

// API: Get Exam Results by Service Number
app.post('/api/results', async (req, res) => {
  const { serviceNumber } = req.body;
  if (!serviceNumber) {
    return res.status(400).json({ error: 'الرقم الوظيفي مطلوب لجلب النتائج' });
  }

  const employeeResults = await getResults(serviceNumber);
  
  res.json({
    results: employeeResults
  });
});

const QUESTIONS = {
  english: [
    {
      id: 1,
      text: "The committee decided to investigate the matter further before making a decision.",
      keyword: "investigate",
      options: ["يستثمر", "يتجاهل", "يتحقق في", "يستجوب"],
      correctOptionIndex: 2
    },
    {
      id: 2,
      text: "She was reluctant to sign the contract without reading the details.",
      keyword: "reluctant",
      options: ["متحمس", "متردد", "مستعد", "عازم"],
      correctOptionIndex: 1
    },
    {
      id: 3,
      text: "The company needs to acquire more assets to expand its operations.",
      keyword: "acquire",
      options: ["يبيع", "يتخلص من", "يستحوذ على", "يخسر"],
      correctOptionIndex: 2
    },
    {
      id: 4,
      text: "His explanation was very clear and concise, saving us a lot of time.",
      keyword: "concise",
      options: ["موجز", "طويل", "غامض", "ممل"],
      correctOptionIndex: 0
    },
    {
      id: 5,
      text: "The project was delayed due to unforeseen circumstances.",
      keyword: "unforeseen",
      options: ["متوقعة", "مخططة", "غير متوقعة", "بسيطة"],
      correctOptionIndex: 2
    }
  ],
  french: [
    {
      id: 1,
      text: "Le gouvernement a décidé d'investir massivement dans les énergies renouvelables.",
      keyword: "d'investir",
      options: ["الاستثمار", "التحقيق", "التجاهل", "الإلغاء"],
      correctOptionIndex: 0
    },
    {
      id: 2,
      text: "Il est indispensable de maintenir une communication étroite entre les départements.",
      keyword: "maintenir",
      options: ["تدمير", "تجاهل", "الحفاظ على", "تغيير"],
      correctOptionIndex: 2
    },
    {
      id: 3,
      text: "Elle a refusé de participer aux négociations sans conditions préalables.",
      keyword: "participer",
      options: ["المشاركة", "الانسحاب", "التنظيم", "الرفض"],
      correctOptionIndex: 0
    },
    {
      id: 4,
      text: "Le rapport met en évidence les défis majeurs auxquels nous sommes confrontés.",
      keyword: "défis",
      options: ["التحديات", "النجاحات", "الخطط", "النتائج"],
      correctOptionIndex: 0
    },
    {
      id: 5,
      text: "Nous devons optimiser nos ressources pour atteindre nos objectifs annuels.",
      keyword: "optimiser",
      options: ["تبديد", "تحسين/تحسين استغلال", "تجاهل", "تغيير"],
      correctOptionIndex: 1
    }
  ],
  russian: [
    {
      id: 1,
      text: "Стороны обсудили перспективы дальнейшего развития двусторонних отношений.",
      keyword: "развития",
      options: ["تطوير", "إنهاء", "تراجع", "تجاهل"],
      correctOptionIndex: 0
    },
    {
      id: 2,
      text: "Реализация этого проекта требует значительных финансовых вложений.",
      keyword: "требует",
      options: ["يرفض", "يطلب/يتطلب", "يلغي", "يسهل"],
      correctOptionIndex: 1
    },
    {
      id: 3,
      text: "Эффективное управление ресурсами является ключом к успеху компании.",
      keyword: "управление",
      options: ["إدارة", "إهدار", "توزيع", "شراء"],
      correctOptionIndex: 0
    },
    {
      id: 4,
      text: "Необходимо принять срочные меры для стабилизации экономической ситуации.",
      keyword: "меры",
      options: ["خطط", "إجراءات/تدابير", "تعديلات", "قوانين"],
      correctOptionIndex: 1
    },
    {
      id: 5,
      text: "Новая стратегия направлена на укрепление доверия между партнерами.",
      keyword: "укрепление",
      options: ["إضعاف", "تعزيز/تقوية", "تغيير", "دراسة"],
      correctOptionIndex: 1
    }
  ]
};

function generateQuestionsExcelIfNeeded(filePath, lang) {
  if (fs.existsSync(filePath)) {
    return;
  }
  
  console.log(`Generating default questions Excel file at ${filePath} for language ${lang}...`);
  
  const base = QUESTIONS[lang] || QUESTIONS['english'];
  const levels = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2'];
  const data = [];
  for (let i = 0; i < 100; i++) {
    const template = base[i % base.length];
    // Assign levels round-robin or randomly so each level gets approximately equal share
    const level = levels[i % levels.length];
    data.push({
      QuestionText: template.text,
      Keyword: template.keyword,
      Option1: template.options[0] || '',
      Option2: template.options[1] || '',
      Option3: template.options[2] || '',
      Option4: template.options[3] || '',
      CorrectOption: template.correctOptionIndex + 1, // 1-based for Excel users
      Level: level
    });
  }
  
  const ws = xlsx.utils.json_to_sheet(data);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Questions');
  xlsx.writeFile(wb, filePath);
  console.log(`Successfully generated ${filePath}`);
}

// API: Get Exam Questions by Language (e.g. english)
app.post('/api/questions', (req, res) => {
  const { language, serviceNumber } = req.body;
  const lang = String(language || '').toLowerCase().trim();
  const filePath = path.join(__dirname, `${lang}.xlsx`);
  
  let baseLang = 'english';
  if (lang.startsWith('french')) {
    baseLang = 'french';
  } else if (lang.startsWith('russian')) {
    baseLang = 'russian';
  }
  
  // Lookup employee level for the requested language
  let userLevel = 'A1'; // default fallback
  if (serviceNumber) {
    const employees = getEmployees();
    const employee = employees.find(emp => String(emp.ServiceNumber).trim() === String(serviceNumber).trim());
    if (employee) {
      let index = 1;
      while (true) {
        const langKey = `Language${index}`;
        const lvlKey = `Level${index}`;
        if (employee[langKey] !== undefined) {
          const langName = String(employee[langKey]).trim().toLowerCase();
          if (langName === lang) {
            userLevel = employee[lvlKey] ? String(employee[lvlKey]).trim().toUpperCase() : 'A1';
            break;
          }
        } else {
          break;
        }
        index++;
      }
    }
  }

  try {
    generateQuestionsExcelIfNeeded(filePath, baseLang);
    
    // Read from the language-specific Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    
    // Filter questions by user level
    const filteredRows = rows.filter(row => {
      const rowLevel = row.Level ? String(row.Level).trim().toUpperCase() : 'A1';
      return rowLevel === userLevel;
    });

    // If no questions match this level, fallback to A1 or all rows to prevent empty exam
    const rowsToUse = filteredRows.length > 0 ? filteredRows : rows;

    const rawQuestions = rowsToUse.map((row, index) => {
      const options = [];
      if (row.Option1 !== undefined) options.push(String(row.Option1));
      if (row.Option2 !== undefined) options.push(String(row.Option2));
      if (row.Option3 !== undefined) options.push(String(row.Option3));
      if (row.Option4 !== undefined) options.push(String(row.Option4));
      
      let correctOptionIndex = 0;
      if (row.CorrectOption !== undefined) {
        const parsed = parseInt(row.CorrectOption);
        if (parsed >= 1 && parsed <= 4) {
          correctOptionIndex = parsed - 1;
        } else if (parsed >= 0 && parsed <= 3) {
          correctOptionIndex = parsed;
        }
      }
      
      return {
        id: index + 1,
        text: row.QuestionText || row.text || '',
        keyword: row.Keyword || row.keyword || '',
        options: options,
        correctOptionIndex: correctOptionIndex,
        level: row.Level || 'A1'
      };
    });
    
    if (rawQuestions.length === 0) {
      return res.status(404).json({ error: 'لم يتم العثور على أسئلة في الملف' });
    }
    
    // Shuffle (Fisher-Yates)
    for (let i = rawQuestions.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [rawQuestions[i], rawQuestions[j]] = [rawQuestions[j], rawQuestions[i]];
    }
    
    // Pick 25 questions (or less if not enough questions are available for this level)
    const countToTake = Math.min(25, rawQuestions.length);
    const selectedQuestions = rawQuestions.slice(0, countToTake).map((q, idx) => {
      return {
        ...q,
        id: idx + 1
      };
    });
    
    res.json({ questions: selectedQuestions, level: userLevel });
  } catch (error) {
    console.error('Error loading questions from Excel:', error);
    res.status(500).json({ error: 'حدث خطأ أثناء تحميل الأسئلة من ملف Excel' });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Backend server running on port ${PORT} (listening on all interfaces)`);
});
