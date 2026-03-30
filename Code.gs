// ============================================================
//  CONFIG
// ============================================================
const SHEET_ID = 'ENTER_SHEET_ID';
// All 3 Tabs should be in the Same Sheet

// GIDs for the three tabs
const GID_PHONE   = XXXX;  // Phone number verification list  (Col A)
const GID_INTERIM = XXXX;   // Interim table — new entries appended here
const GID_ALUMNI  = XXXX;  // Alumni final list (degree, year, email, tokens…)

// Interim sheet columns (all start from A = 1)
const IC_ROW_NUM    = 1;   // A — sequential row number
const IC_NAME       = 2;   // B
const IC_PHONE      = 3;   // C
const IC_CC         = 4;   // D — country code
const IC_ISI_CENTER = 5;   // E
const IC_COURSE     = 6;   // F
const IC_YEAR_START = 7;   // G
const IC_YEAR_END   = 8;   // H
const IC_EMAIL      = 9;   // I
const IC_COUNTRY    = 10;  // J — country of residence
const IC_TIMESTAMP  = 11;  // K
const IC_TAG        = 12;  // L — Validated / Partially_Validated / Not_Validated
const IC_ALUMNI_ROWS= 13;  // M — matched alumni row numbers (comma-separated)

// Highlight colours
const COL_VALIDATED   = '#1E7145';  // deep green
const COL_PARTIAL     = '#FFC000';  // amber/yellow
const COL_NOT_VAL     = '#C00000';  // deep red
const COL_FONT_LIGHT  = '#FFFFFF';  // white font for dark bg

// ============================================================
//  Course mapping: Form value → Alumni list degree string(s)
//  One form course may map to several alumni degree spellings
// ============================================================
const COURSE_MAP = {
  'BSDS'       : ['BSDS'],
  'BMath'      : ['BMath'],
  'BStat'      : ['BStat'],
  'MMath'      : ['MMath'],
  'MSLIB'      : ['MSLIB'],
  'MSQE'       : ['MSQE'],
  'MSQMS'      : ['MSQMS'],
  'MStat'      : ['MStat'],
  'MTech CrS'  : ['MTech CrS'],
  'MTech CS'   : ['MTech CS', 'MTech CS by External Exam'],
  'MTech QROR' : ['MTech QROR'],
  'PhD'        : ['PhD'],
  'PG Diploma' : ['PGDARSMA','PGDARSMA.','PGDAS.','PGDCA.','PGDSMA.','PGDSTAT.',
                  'Advanced Diploma CS','Diploma CS','Diploma Demography',
                  'Diploma Econometrics and Planning','Diploma SQC','Diploma SQCOR',
                  'Part-time Diploma SQCOR','Punched Card Data Processing Diploma'],
};

// Non-alumni courses: name match alone is never Validated, always Partially_Validated.
// Faculty/Staff/Other have no alumni record so degree+year match is impossible.
const NON_ALUMNI_COURSES = ['Faculty', 'Staff / Intern', 'Other'];

// ============================================================
//  Helper: find sheet tab by numeric gid
// ============================================================
function getSheetByGid(ss, gid) {
  for (const s of ss.getSheets()) {
    if (s.getSheetId() === gid) return s;
  }
  return null;
}

// ============================================================
//  Fuzzy match logic
//  submitted : string entered by user   e.g. "Rajanala Samyak"
//  stored    : full name built from tokens e.g. "RAJANALA SAMYAK"
//
//  Rules (all case-insensitive, punctuation-stripped):
//  1. Every word in submitted must appear somewhere in stored tokens  → match
//  2. Every word in stored must appear somewhere in submitted tokens  → match
//  3. Any single word appears in both sets (≥4 chars)               → partial
//  Returns: 'full' | 'partial' | 'none'
// ============================================================
function tokenise(name) {
  return name
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .split(/\s+/)
    .map(t => t.trim())
    .filter(t => t.length > 0);
}

function fuzzyMatch(submitted, stored) {
  const sToks = tokenise(submitted);   // user input tokens
  const rToks = tokenise(stored);      // alumni record tokens

  if (sToks.length === 0 || rToks.length === 0) return 'none';

  // Rule 1: every submitted token is contained in record tokens
  const rule1 = sToks.every(t => rToks.includes(t));
  // Rule 2: every record token is contained in submitted tokens
  const rule2 = rToks.every(t => sToks.includes(t));

  if (rule1 || rule2) return 'full';

  // Rule 3: significant overlap — any shared token ≥ 4 chars
  const shared = sToks.filter(t => t.length >= 4 && rToks.includes(t));
  if (shared.length > 0) return 'partial';

  return 'none';
}

// ============================================================
//  Build full name string from alumni row tokens
// ============================================================
function buildAlumniName(row) {
  // row = [degree, year, email, token_1, token_2, token_3, token_4]
  //  idx:    0      1     2       3         4         5        6
  return [row[3], row[4], row[5], row[6]]
    .map(t => (t || '').trim())
    .filter(t => t.length > 0)
    .join(' ');
}

// ============================================================
//  Entry point
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Form')
    .setTitle('ISI Alumni Community Verification')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setFaviconUrl('https://raw.githubusercontent.com/0r-gh0/ISI-Community-Verify/refs/heads/main/ISI_LOGO.ico');
}

// ============================================================
//  Main submit handler
// ============================================================
function submitForm(data) {
  // ── 1. Validate mandatory fields ─────────────────────────
  const required = ['name', 'phone', 'isiCenter', 'course', 'yearStart', 'yearEnd'];
  for (const key of required) {
    if (!data[key] || String(data[key]).trim() === '') {
      return { success: false, message: 'All mandatory fields must be filled.' };
    }
  }

  const inputPhone = String(data.phone).trim().replace(/\D/g, '');
  if (inputPhone.length < 5) {
    return { success: false, message: 'Please enter a valid phone number.' };
  }

  // ── 2. Open spreadsheet ───────────────────────────────────
  let ss;
  try { ss = SpreadsheetApp.openById(SHEET_ID); }
  catch (e) { return { success: false, message: 'Could not open spreadsheet.' }; }

  const phoneSheet   = getSheetByGid(ss, GID_PHONE);
  const interimSheet = getSheetByGid(ss, GID_INTERIM);
  const alumniSheet  = getSheetByGid(ss, GID_ALUMNI);

  if (!phoneSheet)   return { success: false, message: 'Phone verification sheet not found.' };
  if (!interimSheet) return { success: false, message: 'Interim sheet not found.' };
  if (!alumniSheet)  return { success: false, message: 'Alumni list sheet not found.' };

  // ── 3. Phone number verification ─────────────────────────
  const phoneLastRow = phoneSheet.getLastRow();
  if (phoneLastRow < 1) {
    return { success: false, message: 'Phone list is empty.' };
  }

  const phoneValues = phoneSheet.getRange(1, 1, phoneLastRow, 1).getValues();
  let phoneFound = false;
  for (let i = 0; i < phoneValues.length; i++) {
    const stored = String(phoneValues[i][0]).trim().replace(/\D/g, '');
    if (stored.length > 0 && (stored === inputPhone || stored.endsWith(inputPhone))) {
      phoneFound = true;
      break;
    }
  }

  if (!phoneFound) {
    return {
      success: false,
      message: 'Please enter the phone number registered in the ISI WhatsApp Community.'
    };
  }

  // ── 4. Alumni fuzzy-match ─────────────────────────────────
  // Load full alumni list: [degree, year, email, t1, t2, t3, t4]
  const alumniLastRow = alumniSheet.getLastRow();
  // Skip header row (row 1) — data starts row 2
  const alumniData = alumniSheet.getRange(2, 1, alumniLastRow - 1, 7).getValues();

  const submittedName   = data.name.trim();
  const submittedCourse = data.course.trim();
  const submittedYearEnd= parseInt(data.yearEnd);

  // Alumni degree strings that correspond to this form course
  const mappedDegrees   = COURSE_MAP[submittedCourse] || [];
  // If the submitted course is Faculty/Staff/Other, it has no alumni record.
  // Degree+year match is structurally impossible → cap at Partially_Validated.
  const isNonAlumniRole = NON_ALUMNI_COURSES.includes(submittedCourse);

  const fullMatchRows    = [];
  const partialMatchRows = [];
  const notMatchRows     = [];  // name found but neither degree nor year matches

  for (let i = 0; i < alumniData.length; i++) {
    const row        = alumniData[i];
    const alDegree   = String(row[0]).trim();
    const alYear     = parseInt(row[1]);
    const alFullName = buildAlumniName(row);

    if (!alFullName) continue;

    const nameResult = fuzzyMatch(submittedName, alFullName);
    if (nameResult === 'none') continue;

    // For non-alumni roles, degree match is impossible (no COURSE_MAP entry).
    // Year match is intentionally ignored too — a Staff member from 1962
    // should NOT get Validated just because 1962 appears in the alumni list.
    const degreeMatch = !isNonAlumniRole && mappedDegrees.includes(alDegree);
    const yearMatch   = !isNonAlumniRole &&
                        !isNaN(alYear) && !isNaN(submittedYearEnd) &&
                        alYear === submittedYearEnd;

    const alumniRowNum = i + 2; // 1-based, header is row 1

    if (!isNonAlumniRole && nameResult === 'full' && degreeMatch && yearMatch) {
      // Name (full) + Degree + Year all match → Validated
      fullMatchRows.push(alumniRowNum);
    } else if ((nameResult === 'full' || nameResult === 'partial') && (degreeMatch || yearMatch)) {
      // Name found + at least one of degree OR year matches → Partially_Validated
      partialMatchRows.push(alumniRowNum);
    } else if (nameResult === 'full' || nameResult === 'partial') {
      // Name found but NEITHER degree NOR year match → Not_Validated
      // Still record the row number so admins can see where the name was found
      notMatchRows.push(alumniRowNum);
    }
    // No name match at all → Not_Validated with no rows (handled outside loop)
  }

  // ── 5. Determine validation tag ───────────────────────────
  let tag, bgColour, fontColour;

  if (fullMatchRows.length > 0) {
    // Name (full) + Degree + Year all match
    tag        = 'Validated';
    bgColour   = COL_VALIDATED;
    fontColour = COL_FONT_LIGHT;
  } else if (partialMatchRows.length > 0) {
    // Name found + at least degree OR year matches
    tag        = 'Partially_Validated';
    bgColour   = COL_PARTIAL;
    fontColour = '#000000';
  } else {
    // Name not found at all, OR name found but neither degree nor year matches
    tag        = 'Not_Validated';
    bgColour   = COL_NOT_VAL;
    fontColour = COL_FONT_LIGHT;
  }

  // Collect all row numbers where any name match was found (for admin reference)
  const matchedAlumniRows = [...new Set([
    ...fullMatchRows,
    ...partialMatchRows,
    ...notMatchRows       // include even name-only matches so admins can inspect
  ])];

  // ── 6. Append to Interim sheet ────────────────────────────
  const timestamp   = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
  const nextRow     = interimSheet.getLastRow() + 1;
  const rowNumber   = nextRow - 1;  // sequential entry number (first real entry = 1)

  const rowData = [
    rowNumber,                                          // A — row number
    data.name.trim(),                                   // B — name
    inputPhone,                                         // C — phone (digits)
    data.countryCode || '',                             // D — country code
    data.isiCenter.trim(),                              // E — ISI center
    data.course.trim(),                                 // F — course
    data.yearStart,                                     // G — start year
    data.yearEnd,                                       // H — end year
    data.email ? data.email.trim() : '',                // I — email
    (data.country && data.country.trim()) ? data.country.trim() : 'India', // J — country of residence (default: India)
    timestamp,                                          // K — timestamp
    tag,                                                // L — validation tag
    matchedAlumniRows.length > 0                        // M — matched alumni row numbers
      ? matchedAlumniRows.join(', ')
      : ''
  ];

  interimSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

  // Apply highlight to the entire appended row
  const totalCols = Math.max(interimSheet.getLastColumn(), IC_ALUMNI_ROWS);
  interimSheet.getRange(nextRow, 1, 1, totalCols)
    .setBackground(bgColour)
    .setFontColor(fontColour);

  // Bold the tag cell
  interimSheet.getRange(nextRow, IC_TAG).setFontWeight('bold');

  return {
    success: true,
    message: 'success'
  };
}
