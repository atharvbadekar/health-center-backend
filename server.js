require('dotenv').config();

const dns = require('dns');
const express = require('express');
const cors = require('cors');
const axios = require('axios');
const multer = require('multer');
const xlsx = require('xlsx');
const { Resend } = require('resend');
const pool = require('./db');

// ============================================================================
// 1. FORCE IPV4
// ============================================================================

if (dns.setDefaultResultOrder) {
  dns.setDefaultResultOrder('ipv4first');
}

const app = express();
const PORT = process.env.PORT || 5000;

// Initialize Resend HTTPS API Client
const resend = new Resend(process.env.RESEND_API_KEY || '');

// ============================================================================
// 2. CORS & MIDDLEWARE
// ============================================================================

app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: [
    'Content-Type',
    'Authorization',
    'Accept',
    'X-Requested-With'
  ],
  credentials: true
}));

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header(
    'Access-Control-Allow-Methods',
    'GET, POST, PUT, DELETE, OPTIONS'
  );
  res.header(
    'Access-Control-Allow-Headers',
    'Content-Type, Authorization, Accept, X-Requested-With'
  );

  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }

  next();
});

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// ============================================================================
// 3. FILE UPLOAD
// ============================================================================

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024
  }
});

// ============================================================================
// 4. EMAIL SERVICE STATUS
// ============================================================================

if (!process.env.RESEND_API_KEY) {
  console.warn('⚠️ RESEND_API_KEY is missing from environment variables. Email service disabled.');
} else {
  console.log('✅ Resend HTTPS API client initialized.');
}

// ============================================================================
// 5. HTML ESCAPE FUNCTION
// ============================================================================

const escapeHtml = (value) => {
  if (value === null || value === undefined) {
    return '—';
  }

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
};

// ============================================================================
// 6. EMAIL VALIDATION
// ============================================================================

const isValidEmail = (email) => {
  if (!email) return false;

  const cleanEmail = String(email).trim();

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail);
};

// ============================================================================
// 7. SEND CONSULTATION EMAIL (via HTTPS Port 443)
// ============================================================================

const sendConsultationEmail = async (
  studentEmail,
  studentName,
  details
) => {
  if (!process.env.RESEND_API_KEY) {
    console.error('❌ Email cannot be sent: RESEND_API_KEY is missing.');
    return false;
  }

  const targetEmail = String(studentEmail || '').trim().toLowerCase();

  if (!targetEmail || !isValidEmail(targetEmail)) {
    console.error(`❌ Invalid student email address: ${targetEmail}`);
    return false;
  }

  console.log('--------------------------------------------------');
  console.log('📧 Preparing consultation email via Resend API (HTTPS)');
  console.log(`📥 TO:   ${targetEmail}`);
  console.log('--------------------------------------------------');

  const safeStudentName = escapeHtml(studentName || 'Student');
  const safeDoctorName = escapeHtml(details.doctor_name || 'Medical Officer');
  const safeSymptoms = escapeHtml(details.symptoms || '—');
  const safeTreatment = escapeHtml(details.treatment || '—');
  const safePrescription = escapeHtml(details.prescription || '—');
  const safeAdditionalNotes = escapeHtml(details.additional_notes || '—');

  const currentDate = new Date();
  const formattedDate = currentDate.toLocaleString('en-IN', {
    dateStyle: 'medium',
    timeStyle: 'short'
  });

  const htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CURAJ Health Centre</title>
</head>
<body style="margin:0;padding:0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;">
  <div style="max-width:650px;margin:30px auto;background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;">
    
    <div style="background:#1e3a8a;color:#ffffff;padding:25px;text-align:center;">
      <h2 style="margin:0;font-size:22px;">Central University of Rajasthan</h2>
      <p style="margin:7px 0 0;font-size:14px;color:#bfdbfe;">University Health Centre</p>
    </div>

    <div style="padding:30px;color:#334155;line-height:1.6;">
      <p>Dear <strong>${safeStudentName}</strong>,</p>
      <p>Your medical consultation has been successfully recorded. Please find your consultation and prescription details below.</p>

      <table style="width:100%;border-collapse:collapse;margin-top:20px;font-size:14px;">
        <tr>
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #e2e8f0;width:38%;background:#f8fafc;">Date & Time</td>
          <td style="padding:12px;border-bottom:1px solid #e2e8f0;">${formattedDate}</td>
        </tr>
        <tr>
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #e2e8f0;background:#ffffff;">Attending Doctor</td>
          <td style="padding:12px;border-bottom:1px solid #e2e8f0;">Dr. ${safeDoctorName}</td>
        </tr>
        <tr>
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #e2e8f0;background:#f8fafc;">Symptoms</td>
          <td style="padding:12px;border-bottom:1px solid #e2e8f0;">${safeSymptoms}</td>
        </tr>
        <tr>
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #e2e8f0;background:#ffffff;">Diagnosis / Advice</td>
          <td style="padding:12px;border-bottom:1px solid #e2e8f0;">${safeTreatment}</td>
        </tr>
        <tr style="background:#ecfdf5;">
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #bbf7d0;color:#065f46;">Prescription</td>
          <td style="padding:12px;border-bottom:1px solid #bbf7d0;font-weight:bold;color:#065f46;">${safePrescription}</td>
        </tr>
        ${details.additional_notes ? `
        <tr>
          <td style="padding:12px;font-weight:bold;border-bottom:1px solid #e2e8f0;background:#f8fafc;">Special Notes</td>
          <td style="padding:12px;border-bottom:1px solid #e2e8f0;">${safeAdditionalNotes}</td>
        </tr>` : ''}
      </table>

      <div style="margin-top:25px;padding:15px;background:#f8fafc;border-radius:8px;font-size:12px;color:#64748b;">
        <strong>Important:</strong> This is an automated medical record summary. Please consult the University Health Centre if symptoms persist.
      </div>

      <p style="margin-top:25px;font-size:13px;color:#475569;">
        Regards,<br>
        <strong>CURAJ Health Centre</strong><br>
        Central University of Rajasthan
      </p>
    </div>
  </div>
</body>
</html>`;

  try {
    const { data, error } = await resend.emails.send({
      from: 'CURAJ Health Centre <onboarding@resend.dev>',
      to: [targetEmail],
      subject: `Medical Prescription & Consultation Summary - ${currentDate.toLocaleDateString('en-IN')}`,
      html: htmlContent
    });

    if (error) {
      console.error('❌ Resend API Error:', error);
      return false;
    }

    console.log('--------------------------------------------------');
    console.log('✅ EMAIL SENT SUCCESSFULLY (via HTTPS)');
    console.log(`📥 TO:   ${targetEmail}`);
    console.log(`🆔 Message ID: ${data.id}`);
    console.log('--------------------------------------------------');
    return true;

  } catch (error) {
    console.error('❌ Resend Dispatch Failed:', error.message);
    return false;
  }
};

// ============================================================================
// 8. DATABASE INITIALIZATION
// ============================================================================

const initDatabase = async () => {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS doctors (
        id SERIAL PRIMARY KEY,
        name VARCHAR(100) NOT NULL,
        department VARCHAR(100) NOT NULL,
        created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS students (
        id SERIAL PRIMARY KEY,
        college_id VARCHAR(50) UNIQUE NOT NULL,
        full_name VARCHAR(100),
        email VARCHAR(100),
        mobile_number VARCHAR(20),
        hostel_name VARCHAR(50),
        created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS consultations (
        id SERIAL PRIMARY KEY,
        doctor_id INTEGER REFERENCES doctors(id) ON DELETE SET NULL,
        patient_id VARCHAR(50) NOT NULL,
        is_student BOOLEAN DEFAULT true,
        symptoms TEXT,
        treatment TEXT,
        prescription TEXT,
        additional_notes TEXT,
        consultation_date TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      );
    `);

    // Ensure columns exist
    await pool.query(`
      ALTER TABLE students ADD COLUMN IF NOT EXISTS full_name VARCHAR(100);
      ALTER TABLE students ADD COLUMN IF NOT EXISTS email VARCHAR(100);
      ALTER TABLE students ADD COLUMN IF NOT EXISTS mobile_number VARCHAR(20);
      ALTER TABLE students ADD COLUMN IF NOT EXISTS hostel_name VARCHAR(50);
    `);

    // Remove legacy NOT NULL constraints dynamically
    await pool.query(`
      DO $$
      DECLARE
        rec RECORD;
      BEGIN
        FOR rec IN
          SELECT column_name
          FROM information_schema.columns
          WHERE table_name = 'students'
          AND column_name NOT IN ('id', 'college_id')
          AND is_nullable = 'NO'
        LOOP
          EXECUTE 'ALTER TABLE students ALTER COLUMN "' || rec.column_name || '" DROP NOT NULL;';
        END LOOP;
      END $$;
    `);

    // Seed doctors
    const docCheck = await pool.query('SELECT COUNT(*) FROM doctors');
    if (parseInt(docCheck.rows[0].count, 10) === 0) {
      await pool.query(
        `INSERT INTO doctors (name, department) VALUES ($1, $2), ($3, $4)`,
        ['Dr. Sharma', 'General Medicine', 'Dr. Verma', 'Emergency & OPD']
      );
    }

    console.log('✅ PostgreSQL database initialized successfully.');
  } catch (error) {
    console.error('❌ Database initialization error:', error.message);
  }
};

initDatabase();

// ============================================================================
// 9. HEALTH CHECK
// ============================================================================

app.get('/', (req, res) => {
  res.status(200).send('CURAJ Health Portal API is active and running.');
});

// ============================================================================
// 10. ADMIN LOGIN
// ============================================================================

app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  const adminUser = process.env.ADMIN_USER || 'admin';
  const adminPass = process.env.ADMIN_PASS || 'curaj123';

  if (username === adminUser && password === adminPass) {
    return res.json({
      success: true,
      message: 'Admin login successful',
      role: 'admin'
    });
  }

  return res.status(401).json({
    success: false,
    message: 'Invalid Administrator credentials.'
  });
});

// ============================================================================
// 11. OFFICE / WARDEN LOGIN
// ============================================================================

const handleOfficeLogin = (req, res) => {
  const { username, password } = req.body;
  const validUser = process.env.OFFICE_USER || process.env.WARDEN_USER || 'warden';
  const validPass = process.env.OFFICE_PASS || process.env.WARDEN_PASS || 'hostel123';

  if (username === validUser && password === validPass) {
    return res.json({
      success: true,
      message: 'Login successful',
      role: 'office'
    });
  }

  return res.status(401).json({
    success: false,
    message: 'Invalid Office/Warden credentials.'
  });
};

app.post('/api/office/login', handleOfficeLogin);
app.post('/api/warden/login', handleOfficeLogin);

// ============================================================================
// 12. DOCTOR MANAGEMENT
// ============================================================================

app.get('/api/doctors', async (req, res) => {
  try {
    const doctors = await pool.query('SELECT * FROM doctors ORDER BY id ASC');
    res.json(doctors.rows);
  } catch (error) {
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

app.post('/api/doctors', async (req, res) => {
  const { name, department } = req.body;
  if (!name || !department) {
    return res.status(400).json({
      success: false,
      message: 'Doctor name and department are required.'
    });
  }

  try {
    const newDoctor = await pool.query(
      `INSERT INTO doctors (name, department) VALUES ($1, $2) RETURNING *`,
      [name.trim(), department.trim()]
    );
    res.json({
      success: true,
      doctor: newDoctor.rows[0]
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

app.delete('/api/doctors/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const result = await pool.query(
      `DELETE FROM doctors WHERE id = $1 RETURNING *`,
      [id]
    );

    if (result.rowCount === 0) {
      return res.status(404).json({
        success: false,
        message: 'Doctor record not found.'
      });
    }

    res.json({
      success: true,
      message: `Dr. ${result.rows[0].name} removed from system.`
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

// ============================================================================
// 13. MOBILE NUMBER SANITIZATION
// ============================================================================

const sanitizeMobileNumber = (input) => {
  if (input === null || input === undefined) {
    return null;
  }

  let rawStr = String(input).trim();
  if (rawStr.includes('.')) {
    rawStr = rawStr.split('.')[0];
  }

  let digits = rawStr.replace(/\D/g, '');
  if (digits.length === 12 && digits.startsWith('91')) {
    digits = digits.slice(2);
  } else if (digits.length === 11 && digits.startsWith('0')) {
    digits = digits.slice(1);
  }

  return digits.length === 10 ? digits : null;
};

// ============================================================================
// 14. EXCEL STUDENT UPLOAD
// ============================================================================

const handleExcelUpload = async (req, res) => {
  if (!req.file) {
    return res.status(400).json({
      success: false,
      message: 'No spreadsheet file received.'
    });
  }

  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet, { raw: false, defval: '' });

    if (!rawData || rawData.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'Uploaded spreadsheet is empty.'
      });
    }

    let successCount = 0;
    let skippedCount = 0;
    const errors = [];

    for (let index = 0; index < rawData.length; index++) {
      const row = rawData[index];
      const rowNumber = index + 2;

      const normalized = {};
      Object.keys(row).forEach(key => {
        const cleanKey = key.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        normalized[cleanKey] = String(row[key]).trim();
      });

      const collegeId =
        normalized.collegeid ||
        normalized.id ||
        normalized.enrollmentno ||
        normalized.enrollment ||
        normalized.rollno ||
        normalized.studentid;

      const fullName =
        normalized.fullname ||
        normalized.name ||
        normalized.studentname ||
        normalized.student;

      const email =
        normalized.email ||
        normalized.emailid ||
        normalized.mail ||
        normalized.studentemail;

      const rawMobile =
        normalized.mobilenumber ||
        normalized.mobile ||
        normalized.phone ||
        normalized.phonenumber ||
        normalized.contact;

      const hostelName =
        normalized.hostelname ||
        normalized.hostel ||
        normalized.hostelid ||
        normalized.room ||
        null;

      const cleanMobile = sanitizeMobileNumber(rawMobile);
      const cleanEmail = String(email || '').trim().toLowerCase();

      if (!collegeId || !fullName || !cleanEmail) {
        skippedCount++;
        errors.push(`Row ${rowNumber}: Missing ID, Name, or Email`);
        continue;
      }

      if (!isValidEmail(cleanEmail)) {
        skippedCount++;
        errors.push(`Row ${rowNumber} (${collegeId}): Invalid email '${email}'`);
        continue;
      }

      if (!cleanMobile) {
        skippedCount++;
        errors.push(`Row ${rowNumber} (${collegeId}): Invalid mobile number '${rawMobile}'`);
        continue;
      }

      try {
        await pool.query(
          `INSERT INTO students (college_id, full_name, email, mobile_number, hostel_name)
           VALUES ($1, $2, $3, $4, $5)
           ON CONFLICT (college_id)
           DO UPDATE SET
             full_name = EXCLUDED.full_name,
             email = EXCLUDED.email,
             mobile_number = EXCLUDED.mobile_number,
             hostel_name = EXCLUDED.hostel_name`,
          [
            String(collegeId).trim().toUpperCase(),
            fullName,
            cleanEmail,
            cleanMobile,
            hostelName || null
          ]
        );
        successCount++;
      } catch (dbError) {
        skippedCount++;
        errors.push(`Row ${rowNumber} DB Error: ${dbError.message}`);
      }
    }

    if (successCount === 0 && skippedCount > 0) {
      return res.status(400).json({
        success: false,
        message: `Upload failed. All ${skippedCount} rows were invalid. First error: ${errors[0]}`
      });
    }

    return res.status(200).json({
      success: true,
      message: `Successfully processed ${successCount} students. (${skippedCount} skipped/invalid)`,
      errors: errors.slice(0, 5)
    });

  } catch (error) {
    console.error('Spreadsheet parse fatal error:', error);
    return res.status(500).json({
      success: false,
      message: `Failed to read spreadsheet: ${error.message}`
    });
  }
};

// ============================================================================
// 15. TEST EMAIL ROUTE (via Resend HTTPS)
// ============================================================================

app.get('/api/test-email', async (req, res) => {
  const targetEmail = String(req.query.to || '').trim().toLowerCase();

  if (!targetEmail) {
    return res.status(400).json({
      success: false,
      error: 'Please provide an email query parameter, e.g. ?to=user@example.com'
    });
  }

  if (!isValidEmail(targetEmail)) {
    return res.status(400).json({
      success: false,
      error: `Invalid email address: ${targetEmail}`
    });
  }

  console.log(`📧 Test email requested -> ${targetEmail}`);

  try {
    const { data, error } = await resend.emails.send({
      from: 'CURAJ Health Centre <onboarding@resend.dev>',
      to: [targetEmail],
      subject: 'Test Email from CURAJ Health Centre',
      html: `
        <div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:auto;padding:30px;border:1px solid #e2e8f0;border-radius:8px;">
          <h2 style="color:#1e3a8a;">CURAJ Health Centre</h2>
          <p>This is a test email sent via <strong>Resend HTTPS API (Port 443)</strong>.</p>
          <p>Your email service on Render is fully functional and ready!</p>
        </div>`
    });

    if (error) {
      return res.status(400).json({ success: false, error });
    }

    console.log(`✅ Test email accepted for ${targetEmail}`);

    res.json({
      success: true,
      message: `Email successfully sent to ${targetEmail}`,
      messageId: data.id
    });
  } catch (error) {
    console.error('❌ Test email failed:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ============================================================================
// 16. STUDENT MANAGEMENT
// ============================================================================

app.post('/api/students', async (req, res) => {
  const { college_id, full_name, email, mobile_number, hostel_name } = req.body;
  const cleanMobile = sanitizeMobileNumber(mobile_number);
  const cleanEmail = String(email || '').trim().toLowerCase();

  if (!cleanMobile) {
    return res.status(400).json({
      success: false,
      message: 'Invalid 10-digit mobile number.'
    });
  }

  if (!college_id || !full_name || !cleanEmail) {
    return res.status(400).json({
      success: false,
      message: 'All student fields are required.'
    });
  }

  if (!isValidEmail(cleanEmail)) {
    return res.status(400).json({
      success: false,
      message: 'Invalid email address.'
    });
  }

  try {
    const newStudent = await pool.query(
      `INSERT INTO students (college_id, full_name, email, mobile_number, hostel_name)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING *`,
      [
        college_id.trim().toUpperCase(),
        full_name.trim(),
        cleanEmail,
        cleanMobile,
        hostel_name || null
      ]
    );

    res.json({
      success: true,
      student: newStudent.rows[0]
    });
  } catch (error) {
    if (error.code === '23505') {
      return res.status(400).json({
        success: false,
        message: 'Student College ID already exists.'
      });
    }

    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

// ============================================================================
// 17. GET STUDENTS
// ============================================================================

app.get('/api/students', async (req, res) => {
  try {
    const students = await pool.query('SELECT * FROM students ORDER BY id ASC');
    res.json(students.rows);
  } catch (error) {
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

// ============================================================================
// 18. EXCEL UPLOAD ROUTES
// ============================================================================

app.post('/api/office/upload-students', upload.single('file'), handleExcelUpload);
app.post('/api/warden/upload-students', upload.single('file'), handleExcelUpload);

// ============================================================================
// 19. OTP VERIFICATION
// ============================================================================

app.post('/api/send-otp', async (req, res) => {
  const { identifier, isStudent } = req.body;
  let targetMobile = identifier;
  let patientDetails = null;

  try {
    if (isStudent) {
      const student = await pool.query(
        `SELECT college_id, full_name, email, mobile_number, hostel_name
         FROM students
         WHERE UPPER(college_id) = UPPER($1)`,
        [String(identifier).trim()]
      );

      if (student.rows.length === 0) {
        return res.status(404).json({
          success: false,
          message: 'Student record not found in database.'
        });
      }

      targetMobile = student.rows[0].mobile_number;
      patientDetails = student.rows[0];
    }

    const otp = Math.floor(1000 + Math.random() * 9000);
    console.log(`[MOCK OTP] Sent code ${otp} to target: ${targetMobile}`);

    if (process.env.FAST2SMS_API_KEY && targetMobile) {
      try {
        await axios.get('https://www.fast2sms.com/dev/bulkV2', {
          params: {
            authorization: process.env.FAST2SMS_API_KEY,
            variables_values: otp.toString(),
            route: 'otp',
            numbers: targetMobile
          }
        });
      } catch (smsError) {
        console.error('SMS delivery error:', smsError.message);
      }
    }

    res.json({
      success: true,
      message: 'OTP dispatched',
      mockOtp: otp,
      patientDetails
    });
  } catch (error) {
    console.error('OTP dispatch error:', error.message);
    res.status(500).json({
      success: false,
      message: `Server error: ${error.message}`
    });
  }
});

// ============================================================================
// 20. MEDICAL CONSULTATIONS
// ============================================================================

app.post('/api/consultations', async (req, res) => {
  const {
    doctor_id,
    patient_id,
    is_student,
    symptoms,
    treatment,
    prescription,
    additional_notes
  } = req.body;

  try {
    const newConsultation = await pool.query(
      `INSERT INTO consultations
      (doctor_id, patient_id, is_student, symptoms, treatment, prescription, additional_notes)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING *`,
      [
        doctor_id || null,
        patient_id,
        is_student ?? true,
        symptoms || null,
        treatment || null,
        prescription || null,
        additional_notes || null
      ]
    );

    let studentEmail = null;
    let studentName = null;
    let doctorName = null;

    if (doctor_id) {
      const docRes = await pool.query(
        `SELECT name FROM doctors WHERE id = $1`,
        [doctor_id]
      );
      if (docRes.rows.length > 0) {
        doctorName = docRes.rows[0].name;
      }
    }

    if (is_student) {
      const cleanPatientId = String(patient_id || '').trim();
      const studentRes = await pool.query(
        `SELECT college_id, email, full_name
         FROM students
         WHERE UPPER(TRIM(college_id)) = UPPER(TRIM($1))
         LIMIT 1`,
        [cleanPatientId]
      );

      if (studentRes.rows.length > 0) {
        studentEmail = String(studentRes.rows[0].email || '').trim().toLowerCase();
        studentName = studentRes.rows[0].full_name;
      }
    }

    let emailDelivered = false;
    if (studentEmail) {
      emailDelivered = await sendConsultationEmail(
        studentEmail,
        studentName,
        {
          doctor_name: doctorName,
          symptoms,
          treatment,
          prescription,
          additional_notes
        }
      );
    }

    res.json({
      success: true,
      consultation: newConsultation.rows[0],
      emailDispatched: emailDelivered,
      targetEmail: studentEmail
    });

  } catch (error) {
    console.error('❌ Save consultation error:', error);
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

// ============================================================================
// 21. GET CONSULTATIONS
// ============================================================================

app.get('/api/consultations', async (req, res) => {
  try {
    const consultations = await pool.query(`
      SELECT
        c.*,
        COALESCE(d.name, 'Unassigned') AS doctor_name,
        COALESCE(d.department, 'General') AS department
      FROM consultations c
      LEFT JOIN doctors d ON c.doctor_id = d.id
      ORDER BY c.consultation_date DESC
    `);
    res.json(consultations.rows);
  } catch (error) {
    res.status(500).json({
      success: false,
      message: `Database error: ${error.message}`
    });
  }
});

// ============================================================================
// 22. START SERVER
// ============================================================================

app.listen(PORT, () => {
  console.log(`🚀 CURAJ Health Portal API server running on port ${PORT}`);
});