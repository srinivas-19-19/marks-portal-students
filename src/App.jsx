import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { FileSpreadsheet, Printer, Trash2, Plus, Layout } from 'lucide-react';

const INITIAL_STUDENT = {
  rollNo: '',
  name: '',
  q1: '', q2: '', q3: '', q4: '', q5: '', q6: '',
  objective: '',
  mid1: '', mid2: '',
  assign1: '', assign2: ''
};

const WORKSPACE_STORAGE_KEY = 'svce-marks-portal-workspace-v1';

const createDefaultWorkspace = () => ({
  calculationMode: 'single',
  program: 'B.Tech',
  regulation: 'R23',
  department: 'CSE',
  year: 'II',
  semesterNum: 'II',
  examType: 'I Internal Examinations',
  examMonthYear: 'Feb - 2026',
  facultyName: '',
  courseCode: '',
  subjectName: '',
  students: [{ ...INITIAL_STUDENT }]
});

const loadSavedWorkspace = () => {
  const fallback = createDefaultWorkspace();
  try {
    const saved = JSON.parse(localStorage.getItem(WORKSPACE_STORAGE_KEY));
    if (!saved || typeof saved !== 'object') return fallback;
    return {
      ...fallback,
      ...saved,
      students: Array.isArray(saved.students) && saved.students.length
        ? saved.students.map(student => ({ ...INITIAL_STUDENT, ...student }))
        : fallback.students
    };
  } catch {
    return fallback;
  }
};

const PROGRAMS = ['B.Tech', 'M.Tech', 'MBA', 'MCA', 'PHD'];
const REGULATIONS = ['R20', 'R23', 'R24', 'R25', 'R26'];
const DEPARTMENTS = ['CSE', 'ECE', 'CIVIL', 'MEC', 'CSC', 'ETC'];
const YEARS = ['I', 'II', 'III', 'IV'];
const SEMESTERS = ['I', 'II'];
const EXAM_TYPES = ['I Internal Examinations', 'II Internal Examinations', 'Pre-Final Examinations'];
const MARK_LIMITS = {
  q1: 10, q2: 10, q3: 10, q4: 10, q5: 10, q6: 10,
  objective: 10, mid1: 25, mid2: 25, assign1: 5, assign2: 5
};
const SINGLE_MARK_FIELDS = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'objective'];
const CONSOLIDATED_MARK_FIELDS = ['mid1', 'mid2', 'assign1', 'assign2'];

const DIGIT_WORDS = ['ZERO', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE'];

function App() {
  const [savedWorkspace] = useState(loadSavedWorkspace);
  const [calculationMode, setCalculationMode] = useState(savedWorkspace.calculationMode);
  
  // V2 Exam Metadata States
  const [program, setProgram] = useState(savedWorkspace.program);
  const [regulation, setRegulation] = useState(savedWorkspace.regulation);
  const [department, setDepartment] = useState(savedWorkspace.department);
  const [year, setYear] = useState(savedWorkspace.year);
  const [semesterNum, setSemesterNum] = useState(savedWorkspace.semesterNum);
  const [examType, setExamType] = useState(savedWorkspace.examType);
  const [examMonthYear, setExamMonthYear] = useState(savedWorkspace.examMonthYear);
  
  const [facultyName, setFacultyName] = useState(savedWorkspace.facultyName);
  const [courseCode, setCourseCode] = useState(savedWorkspace.courseCode);
  const [subjectName, setSubjectName] = useState(savedWorkspace.subjectName);
  
  const [students, setStudents] = useState(savedWorkspace.students);
  const fileInputRef = useRef(null);
  const backupInputRef = useRef(null);
  const [saveStatus, setSaveStatus] = useState('Saved locally');
  const [notice, setNotice] = useState('');

  // Derived Values
  const generatedSemesterString = useMemo(() => {
    return `${year} ${program} ${semesterNum} Semester (${department}) ${regulation} ${examType} ${examMonthYear}`;
  }, [year, program, semesterNum, department, regulation, examType, examMonthYear]);

  const activeMarkFields = calculationMode === 'single' ? SINGLE_MARK_FIELDS : CONSOLIDATED_MARK_FIELDS;

  const getMarkError = (field, value) => {
    const text = String(value ?? '').trim();
    if (!text) return '';
    const number = Number(text);
    if (!Number.isFinite(number)) return 'Enter a valid number.';
    if (!Number.isInteger(number)) return 'Marks must be whole numbers.';
    if (number < 0 || number > MARK_LIMITS[field]) return `Enter a mark from 0 to ${MARK_LIMITS[field]}.`;
    return '';
  };

  const markErrors = useMemo(() => students.flatMap((student, index) => activeMarkFields
    .map(field => ({ row: index + 1, field, message: getMarkError(field, student[field]) }))
    .filter(error => error.message)), [students, activeMarkFields]);

  useEffect(() => {
    try {
      localStorage.setItem(WORKSPACE_STORAGE_KEY, JSON.stringify({
        calculationMode, program, regulation, department, year, semesterNum,
        examType, examMonthYear, facultyName, courseCode, subjectName, students
      }));
      setSaveStatus('Saved locally');
    } catch {
      setSaveStatus('Local save unavailable');
    }
  }, [calculationMode, program, regulation, department, year, semesterNum, examType, examMonthYear, facultyName, courseCode, subjectName, students]);

  const applyWorkspace = (workspace) => {
    const normalized = {
      ...createDefaultWorkspace(),
      ...workspace,
      students: Array.isArray(workspace.students) && workspace.students.length
        ? workspace.students.map(student => ({ ...INITIAL_STUDENT, ...student }))
        : [{ ...INITIAL_STUDENT }]
    };
    setCalculationMode(normalized.calculationMode);
    setProgram(normalized.program);
    setRegulation(normalized.regulation);
    setDepartment(normalized.department);
    setYear(normalized.year);
    setSemesterNum(normalized.semesterNum);
    setExamType(normalized.examType);
    setExamMonthYear(normalized.examMonthYear);
    setFacultyName(normalized.facultyName);
    setCourseCode(normalized.courseCode);
    setSubjectName(normalized.subjectName);
    setStudents(normalized.students);
  };

  const exportBackup = () => {
    const workspace = { calculationMode, program, regulation, department, year, semesterNum, examType, examMonthYear, facultyName, courseCode, subjectName, students };
    const blob = new Blob([JSON.stringify(workspace, null, 2)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `marks-workspace-${courseCode || 'backup'}.json`;
    link.click();
    URL.revokeObjectURL(link.href);
  };

  const importBackup = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const workspace = JSON.parse(evt.target.result);
        if (!workspace || typeof workspace !== 'object' || !Array.isArray(workspace.students)) throw new Error('Invalid backup');
        applyWorkspace(workspace);
        setSaveStatus('Backup restored and saved locally');
      } catch {
        setSaveStatus('Could not read that backup file');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const clearWorkspace = () => {
    if (window.confirm('Clear the current marks workspace? Export a backup first if you need these entries.')) {
      applyWorkspace(createDefaultWorkspace());
    }
  };

  const handlePrint = () => {
    if (markErrors.length) {
      setNotice(`Fix ${markErrors.length} invalid mark${markErrors.length === 1 ? '' : 's'} before printing.`);
      return;
    }
    setNotice('');
    window.print();
  };

  const MarkInput = ({ student, index, field, width }) => {
    const error = getMarkError(field, student[field]);
    return <input
      className={`marks-input${error ? ' marks-input-error' : ''}`}
      style={width ? { maxWidth: width } : undefined}
      type="number"
      min="0"
      max={MARK_LIMITS[field]}
      step="1"
      inputMode="numeric"
      value={student[field]}
      aria-label={`${field} marks for ${student.name || `student ${index + 1}`}`}
      aria-invalid={Boolean(error)}
      title={error || `Enter a whole number from 0 to ${MARK_LIMITS[field]}.`}
      onChange={(e) => updateStudentField(index, field, e.target.value)}
    />;
  };

  const calculateResult = (student) => {
    if (calculationMode === 'single') {
      const q1 = Number(student.q1) || 0;
      const q2 = Number(student.q2) || 0;
      const q3 = Number(student.q3) || 0;
      const q4 = Number(student.q4) || 0;
      const q5 = Number(student.q5) || 0;
      const q6 = Number(student.q6) || 0;
      const objective = Number(student.objective) || 0;

      const m1 = Math.max(q1, q2);
      const m2 = Math.max(q3, q4);
      const m3 = Math.max(q5, q6);
      
      const total30 = m1 + m2 + m3;
      const descriptive15 = Math.ceil(total30 / 2);
      const final25 = descriptive15 + objective;

      return { total30, descriptive15, final25 };
    } else {
      // Consolidated Mode V2 (+ Assignments)
      const m1 = Number(student.mid1) || 0;
      const m2 = Number(student.mid2) || 0;
      const a1 = Number(student.assign1) || 0;
      const a2 = Number(student.assign2) || 0;

      const maxMid = Math.max(m1, m2);
      const minMid = Math.min(m1, m2);
      
      const internalMarks25 = Math.round(maxMid * 0.8 + minMid * 0.2);
      const assignment5 = Math.ceil((a1 + a2) / 2);
      const final30 = internalMarks25 + assignment5;

      return { internalMarks25, assignment5, final30 };
    }
  };

  const handleExcelImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target.result, { type: 'array' });
        const normalizeHeader = (value) => String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
        const rollHeaders = ['rollno', 'rollnumber', 'hallticketnumber', 'htno'];
        const nameHeaders = ['name', 'nameofstudent', 'nameofthestudent', 'studentname'];
        const fieldHeaders = {
          q1: ['q1'], q2: ['q2'], q3: ['q3'], q4: ['q4'], q5: ['q5'], q6: ['q6'],
          objective: ['objective', 'obj'], mid1: ['mid1', 'midi'], mid2: ['mid2', 'midii'],
          assign1: ['assign1', 'assignment1'], assign2: ['assign2', 'assignment2']
        };

        let importedStudents = null;
        for (const sheetName of workbook.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
          const headerIndex = rows.findIndex((row, index) => index < 20 && row.some(cell => rollHeaders.includes(normalizeHeader(cell))) && row.some(cell => nameHeaders.includes(normalizeHeader(cell))));
          if (headerIndex === -1) continue;

          const headers = rows[headerIndex].map(normalizeHeader);
          const rollIndex = headers.findIndex(header => rollHeaders.includes(header));
          const nameIndex = headers.findIndex(header => nameHeaders.includes(header));
          const columns = Object.fromEntries(Object.entries(fieldHeaders).map(([field, aliases]) => [field, headers.findIndex(header => aliases.includes(header))]));
          const processed = rows.slice(headerIndex + 1)
            .filter(row => String(row[rollIndex] ?? '').trim() && String(row[nameIndex] ?? '').trim())
            .map(row => {
              const student = { ...INITIAL_STUDENT, rollNo: String(row[rollIndex]).trim(), name: String(row[nameIndex]).trim() };
              Object.entries(columns).forEach(([field, column]) => {
                if (column >= 0 && row[column] !== '') student[field] = String(row[column]).trim();
              });
              return student;
            });
          if (processed.length) {
            importedStudents = processed;
            break;
          }
        }

        if (!importedStudents) {
          setNotice('No students were imported. Use a sheet with Roll No and Name/Name of the Student headers.');
          return;
        }
        setStudents(importedStudents);
        setNotice(`${importedStudents.length} students imported successfully.`);
      } catch {
        setNotice('Could not read that file. Please upload a valid Excel or CSV file.');
      }
    };
    reader.onerror = () => setNotice('Could not read that file. Please try again.');
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const getFinalMark = (student) => {
    const result = calculateResult(student);
    return calculationMode === 'single' ? result.final25 : result.final30;
  };

  const updateStudentField = (index, field, value) => {
    const newStudents = [...students];
    newStudents[index][field] = value;
    setStudents(newStudents);
  };

  const downloadTemplate = () => {
    const data = [
      ['Roll No', 'Name'],
      ['23BFA05277', 'GERRI VYSHNAVI'],
      ['23BFA05278', 'GORRIPARTHI SANTHOSH'],
      ['23BFA05280', 'KARIKERA NAGARAJU GARI SUMANTH']
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "Student_List_Template.xlsx");
  };

  return (
    <div className="container" style={{ maxWidth: '1400px' }}>
      <div className="no-print">
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
            <h1>Student Marks Portal</h1>
            <div style={{ display: 'flex', gap: '0.5rem' }}>
              <button className="btn btn-outline" onClick={downloadTemplate} title="Download Excel Template">
                <FileSpreadsheet size={18} />
                Template
              </button>
              <button className="btn btn-outline" onClick={exportBackup} title="Download a portable copy of this workspace">
                Backup
              </button>
              <button className="btn btn-outline" onClick={() => backupInputRef.current.click()} title="Restore a workspace backup">
                Restore
              </button>
              <input type="file" ref={backupInputRef} accept="application/json,.json" style={{ display: 'none' }} onChange={importBackup} />
              <button 
                className="btn btn-outline" 
                onClick={() => setCalculationMode(calculationMode === 'single' ? 'consolidated' : 'single')}
              >
                <Layout size={18} />
                {calculationMode === 'single' ? 'Switch to Consolidated View' : 'Switch to Single Exam View'}
              </button>
            </div>
          </div>

          <div style={{ background: '#f8fafc', padding: '1.5rem', borderRadius: '1rem', border: '1px solid #e2e8f0', marginBottom: '2rem' }}>
            <h3 style={{ marginBottom: '1rem', fontSize: '1.1rem', color: '#334155' }}>Exam Metadata</h3>
            <div className="form-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))' }}>
              <div className="input-group">
                <label>Program</label>
                <select className="select-field" value={program} onChange={(e) => setProgram(e.target.value)}>
                  {PROGRAMS.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Regulation</label>
                <select className="select-field" value={regulation} onChange={(e) => setRegulation(e.target.value)}>
                  {REGULATIONS.map(r => <option key={r} value={r}>{r}</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Department</label>
                <select className="select-field" value={department} onChange={(e) => setDepartment(e.target.value)}>
                  {DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Year</label>
                <select className="select-field" value={year} onChange={(e) => setYear(e.target.value)}>
                  {(program === 'M.Tech' || program === 'MBA' ? ['I', 'II'] : YEARS).map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Semester</label>
                <select className="select-field" value={semesterNum} onChange={(e) => setSemesterNum(e.target.value)}>
                  {SEMESTERS.map(s => <option key={s} value={s}>{s} Semester</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Exam Type</label>
                <select className="select-field" value={examType} onChange={(e) => setExamType(e.target.value)}>
                  {EXAM_TYPES.map(e => <option key={e} value={e}>{e}</option>)}
                </select>
              </div>
              <div className="input-group">
                <label>Month & Year</label>
                <input className="input-field" value={examMonthYear} onChange={(e) => setExamMonthYear(e.target.value)} placeholder="e.g. Feb - 2026" />
              </div>
            </div>
            
            <hr style={{ margin: '1.5rem 0', borderColor: '#e2e8f0' }} />
            
            <div className="form-grid">
              <div className="input-group">
                <label>Faculty Name</label>
                <input className="input-field" value={facultyName} onChange={(e) => setFacultyName(e.target.value)} placeholder="Enter faculty name" />
              </div>
              <div className="input-group">
                <label>Subject Name</label>
                <input className="input-field" value={subjectName} onChange={(e) => setSubjectName(e.target.value)} placeholder="e.g. DBMS" />
              </div>
              <div className="input-group">
                <label>Course Code</label>
                <input className="input-field" value={courseCode} onChange={(e) => setCourseCode(e.target.value)} placeholder="e.g. CSEM305" />
              </div>
            </div>
          </div>

          <div className="button-row">
            <button className="btn btn-primary" onClick={() => fileInputRef.current.click()}>
              <FileSpreadsheet size={18} />
              Import Student Excel
            </button>
            <input type="file" ref={fileInputRef} accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" style={{ display: 'none' }} onChange={handleExcelImport} />
            <button className="btn btn-success" onClick={handlePrint}>
              <Printer size={18} />
              Print Award List
            </button>
            <button className="btn btn-outline" onClick={clearWorkspace}>
              Clear Workspace
            </button>
          </div>
          <p className="save-status">{saveStatus}. This browser keeps a separate local workspace; use Backup/Restore to move work between devices or users.</p>
          {notice && <p className="notice" role="status">{notice}</p>}
          {markErrors.length > 0 && <p className="notice notice-error" role="alert">{markErrors.length} mark{markErrors.length === 1 ? '' : 's'} need correction before printing.</p>}

          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', width: '120px' }}>Roll No</th>
                  <th style={{ textAlign: 'left', width: '200px' }}>Name</th>
                  {calculationMode === 'single' ? (
                    <>
                      <th>Q1</th><th>Q2</th><th>Q3</th><th>Q4</th><th>Q5</th><th>Q6</th>
                      <th>Obj</th>
                      <th>Final</th>
                    </>
                  ) : (
                    <>
                      <th>MID-I<br/>(Max:25)</th>
                      <th>MID-II<br/>(Max:25)</th>
                      <th>Internal<br/>Marks (25)</th>
                      <th>Assign 1<br/>(05)</th>
                      <th>Assign 2<br/>(05)</th>
                      <th>Assign<br/>(05)</th>
                      <th>Final<br/>(30)</th>
                    </>
                  )}
                  <th style={{ width: '50px' }}></th>
                </tr>
              </thead>
              <tbody>
                {students.map((student, index) => {
                  const res = calculateResult(student);
                  return (
                    <tr key={index}>
                      <td><input className="input-field" style={{ padding: '0.4rem', fontSize: '0.85rem' }} value={student.rollNo} onChange={(e) => updateStudentField(index, 'rollNo', e.target.value)} /></td>
                      <td><input className="input-field" style={{ textAlign: 'left', padding: '0.4rem', fontSize: '0.85rem' }} value={student.name} onChange={(e) => updateStudentField(index, 'name', e.target.value)} /></td>
                      {calculationMode === 'single' ? (
                        <>
                          <td><MarkInput student={student} index={index} field="q1" /></td>
                          <td><MarkInput student={student} index={index} field="q2" /></td>
                          <td><MarkInput student={student} index={index} field="q3" /></td>
                          <td><MarkInput student={student} index={index} field="q4" /></td>
                          <td><MarkInput student={student} index={index} field="q5" /></td>
                          <td><MarkInput student={student} index={index} field="q6" /></td>
                          <td><MarkInput student={student} index={index} field="objective" /></td>
                          <td style={{ fontWeight: 800, color: 'var(--primary)', fontSize: '1.2rem' }}>{res.final25}</td>
                        </>
                      ) : (
                        <>
                          <td><MarkInput student={student} index={index} field="mid1" width="80px" /></td>
                          <td><MarkInput student={student} index={index} field="mid2" width="80px" /></td>
                          <td style={{ fontWeight: 600, color: '#334155' }}>{res.internalMarks25}</td>
                          <td><MarkInput student={student} index={index} field="assign1" width="70px" /></td>
                          <td><MarkInput student={student} index={index} field="assign2" width="70px" /></td>
                          <td style={{ fontWeight: 600, color: '#334155' }}>{res.assignment5}</td>
                          <td style={{ fontWeight: 800, color: 'var(--primary)', fontSize: '1.2rem' }}>{res.final30}</td>
                        </>
                      )}
                      <td>
                        <button onClick={() => setStudents(students.filter((_, i) => i !== index))} style={{ color: '#ef4444', background: 'none', border: 'none', cursor: 'pointer' }}>
                          <Trash2 size={18} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn btn-outline" style={{ marginTop: '1.5rem' }} onClick={() => setStudents([...students, { ...INITIAL_STUDENT }])}>
            <Plus size={18} /> Add Row
          </button>
        </div>
      </div>

      {/* Formal Print View */}
      <div className="print-only">
        <div className="print-header">
          <h1>SV COLLEGE OF ENGINEERING</h1>
          <div style={{ fontWeight: 'bold' }}>(AUTONOMOUS)</div>
          <div style={{ fontSize: '0.8rem' }}>Karakambadi Road, Tirupati-517507</div>
          <div style={{ marginTop: '1rem', fontWeight: 600 }}>{generatedSemesterString}</div>
          <h2 style={{ marginTop: '1rem', textDecoration: 'underline', fontSize: '1.2rem' }}>Award List</h2>
        </div>

        <div className="print-info">
          <div>
            <div><strong>Name of the Subject:</strong> {subjectName || '________________'}</div>
            <div><strong>Name of the Faculty:</strong> {facultyName || '________________'}</div>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div><strong>Subject Code:</strong> {courseCode || '__________'}</div>
          </div>
        </div>

        <table className="print-table" style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid black' }}>
          <thead>
            {calculationMode === 'single' ? (
              <>
                <tr>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>S.No</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Roll Number</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Name of the Student</th>
                  <th colSpan="6" style={{ border: '1px solid black' }}>Descriptive Marks</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Total(30)</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Des(15)</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Obj(10)</th>
                  <th rowSpan="2" style={{ border: '1px solid black' }}>Final(25)</th>
                </tr>
                <tr>
                  <th style={{ border: '1px solid black' }}>Q1</th><th style={{ border: '1px solid black' }}>Q2</th><th style={{ border: '1px solid black' }}>Q3</th>
                  <th style={{ border: '1px solid black' }}>Q4</th><th style={{ border: '1px solid black' }}>Q5</th><th style={{ border: '1px solid black' }}>Q6</th>
                </tr>
              </>
            ) : (
              <tr>
                <th style={{ border: '1px solid black' }}>S.No</th>
                <th style={{ border: '1px solid black' }}>Roll Number</th>
                <th style={{ border: '1px solid black' }}>Name of the Student</th>
                <th style={{ border: '1px solid black' }}>MID-I<br/>(Max:25)</th>
                <th style={{ border: '1px solid black' }}>MID-II<br/>(Max:25)</th>
                <th style={{ border: '1px solid black' }}>Internal<br/>Marks</th>
                <th style={{ border: '1px solid black' }}>Assignment 1<br/>(05)</th>
                <th style={{ border: '1px solid black' }}>Assignment 2<br/>(05)</th>
                <th style={{ border: '1px solid black' }}>Assignment</th>
                <th style={{ border: '1px solid black' }}>Final Internal<br/>Marks (30)</th>
              </tr>
            )}
          </thead>
          <tbody>
            {students.map((student, idx) => {
              const res = calculateResult(student);
              return (
                <tr key={idx}>
                  <td style={{ border: '1px solid black' }}>{idx + 1}</td>
                  <td style={{ border: '1px solid black' }}>{student.rollNo}</td>
                  <td style={{ border: '1px solid black', textAlign: 'left', paddingLeft: '5px' }}>{student.name}</td>
                  {calculationMode === 'single' ? (
                    <>
                      <td style={{ border: '1px solid black' }}>{student.q1 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.q2 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.q3 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.q4 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.q5 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.q6 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{res.total30}</td>
                      <td style={{ border: '1px solid black' }}>{res.descriptive15}</td>
                      <td style={{ border: '1px solid black' }}>{student.objective || 0}</td>
                      <td style={{ border: '1px solid black', fontWeight: 'bold' }}>{res.final25}</td>
                    </>
                  ) : (
                    <>
                      <td style={{ border: '1px solid black' }}>{student.mid1 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.mid2 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{res.internalMarks25}</td>
                      <td style={{ border: '1px solid black' }}>{student.assign1 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{student.assign2 || '-'}</td>
                      <td style={{ border: '1px solid black' }}>{res.assignment5}</td>
                      <td style={{ border: '1px solid black', fontWeight: 'bold' }}>{res.final30}</td>
                    </>
                  )}
                </tr>
              )
            })}
          </tbody>
        </table>

        <div className="sign-row">
          <div className="sign-box">Signature of Faculty</div>
          <div className="sign-box">HoD</div>
        </div>

        <section className="print-final-sheet">
          <div className="print-header">
            <h1>SV COLLEGE OF ENGINEERING</h1>
            <div style={{ fontWeight: 'bold' }}>(AUTONOMOUS)</div>
            <div style={{ fontSize: '0.8rem' }}>Karakambadi Road, Tirupati-517507</div>
            <div style={{ marginTop: '1rem', fontWeight: 600 }}>{generatedSemesterString}</div>
            <h2 style={{ marginTop: '1rem', textDecoration: 'underline', fontSize: '1.2rem' }}>Final Marks in Words</h2>
          </div>

          <div className="print-info">
            <div>
              <div><strong>Name of the Subject:</strong> {subjectName || '________________'}</div>
              <div><strong>Name of the Faculty:</strong> {facultyName || '________________'}</div>
            </div>
            <div style={{ textAlign: 'right' }}>
              <div><strong>Subject Code:</strong> {courseCode || '__________'}</div>
            </div>
          </div>

          <table className="print-table final-words-table" style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid black' }}>
            <thead>
              <tr>
                <th rowSpan="2" style={{ border: '1px solid black' }}>S.No</th>
                <th rowSpan="2" style={{ border: '1px solid black' }}>Roll Number</th>
                <th rowSpan="2" style={{ border: '1px solid black' }}>Name of the Student</th>
                <th rowSpan="2" style={{ border: '1px solid black' }}>Total Marks<br/>({calculationMode === 'single' ? '25' : '30'})</th>
                <th colSpan="2" style={{ border: '1px solid black' }}>Marks in Words</th>
              </tr>
              <tr>
                <th style={{ border: '1px solid black' }}>First Digit in Words</th>
                <th style={{ border: '1px solid black' }}>Second Digit in Words</th>
              </tr>
            </thead>
            <tbody>
              {students.map((student, idx) => {
                const finalMark = getFinalMark(student);
                return (
                  <tr key={idx}>
                    <td style={{ border: '1px solid black' }}>{idx + 1}</td>
                    <td style={{ border: '1px solid black' }}>{student.rollNo}</td>
                    <td style={{ border: '1px solid black', textAlign: 'left', paddingLeft: '5px' }}>{student.name}</td>
                    <td style={{ border: '1px solid black', fontWeight: 'bold' }}>{finalMark}</td>
                    <td style={{ border: '1px solid black' }}>{DIGIT_WORDS[Math.floor(finalMark / 10)]}</td>
                    <td style={{ border: '1px solid black' }}>{DIGIT_WORDS[finalMark % 10]}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          <div className="sign-row">
            <div className="sign-box">Signature of Faculty</div>
            <div className="sign-box">HoD</div>
          </div>
        </section>
      </div>
    </div>
  );
}

export default App;
