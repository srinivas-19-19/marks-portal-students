import React, { useState, useRef, useMemo } from 'react';
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

const PROGRAMS = ['B.Tech', 'M.Tech', 'MBA', 'MCA', 'PHD'];
const REGULATIONS = ['R20', 'R23', 'R24', 'R25'];
const DEPARTMENTS = ['CSE', 'ECE', 'CIVIL', 'MEC', 'CSC', 'ETC'];
const YEARS = ['I', 'II', 'III', 'IV'];
const SEMESTERS = ['I', 'II'];
const EXAM_TYPES = ['I Internal Examinations', 'II Internal Examinations', 'Pre-Final Examinations'];

// Mock Faculty Data mapped by Department
const MOCK_FACULTY = {
  CSE: ['Sheshadri', 'John Doe', 'Dr. Smith', 'Prof. Alan'],
  ECE: ['Dr. Ramesh', 'Sivakumar', 'Prof. Reddy'],
  CIVIL: ['Anil Kumar', 'Dr. Sharma'],
  MEC: ['Dr. Rao', 'Vikram'],
  CSC: ['Priya', 'Dr. Venkatesh'],
  ETC: ['Sanjay', 'Dr. Kumar']
};

function App() {
  const [calculationMode, setCalculationMode] = useState('single');
  
  // V2 Exam Metadata States
  const [program, setProgram] = useState('B.Tech');
  const [regulation, setRegulation] = useState('R23');
  const [department, setDepartment] = useState('CSE');
  const [year, setYear] = useState('II');
  const [semesterNum, setSemesterNum] = useState('II');
  const [examType, setExamType] = useState('I Internal Examinations');
  const [examMonthYear, setExamMonthYear] = useState('Feb - 2026');
  
  const [facultyName, setFacultyName] = useState('');
  const [courseCode, setCourseCode] = useState('');
  const [subjectName, setSubjectName] = useState('');
  
  const [students, setStudents] = useState([INITIAL_STUDENT]);
  const fileInputRef = useRef(null);

  // Derived Values
  const availableFaculty = useMemo(() => MOCK_FACULTY[department] || [], [department]);
  
  const generatedSemesterString = useMemo(() => {
    return `${year} ${program} ${semesterNum} Semester (${department}) ${examType} ${examMonthYear}`;
  }, [year, program, semesterNum, department, examType, examMonthYear]);

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
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const processed = data
        .filter(row => row.length >= 2 && row[0] && row[1])
        .map(row => ({
          ...INITIAL_STUDENT,
          rollNo: String(row[0] || ''),
          name: String(row[1] || ''),
        }))
        .filter(s => !['roll no', 's.no', 'serial'].includes(s.rollNo.toLowerCase()) && s.rollNo.length > 2);

      setStudents(processed);
    };
    reader.readAsBinaryString(file);
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
                <select className="select-field" value={department} onChange={(e) => {
                  setDepartment(e.target.value);
                  setFacultyName(''); // Reset faculty when dept changes
                }}>
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
                <select className="select-field" value={facultyName} onChange={(e) => setFacultyName(e.target.value)}>
                  <option value="">-- Select Faculty --</option>
                  {availableFaculty.map(f => <option key={f} value={f}>{f}</option>)}
                </select>
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
            <input type="file" ref={fileInputRef} style={{ display: 'none' }} onChange={handleExcelImport} />
            <button className="btn btn-success" onClick={() => window.print()}>
              <Printer size={18} />
              Print Award List
            </button>
          </div>

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
                          <td><input className="marks-input" value={student.q1} onChange={(e) => updateStudentField(index, 'q1', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.q2} onChange={(e) => updateStudentField(index, 'q2', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.q3} onChange={(e) => updateStudentField(index, 'q3', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.q4} onChange={(e) => updateStudentField(index, 'q4', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.q5} onChange={(e) => updateStudentField(index, 'q5', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.q6} onChange={(e) => updateStudentField(index, 'q6', e.target.value)} /></td>
                          <td><input className="marks-input" value={student.objective} onChange={(e) => updateStudentField(index, 'objective', e.target.value)} /></td>
                          <td style={{ fontWeight: 800, color: 'var(--primary)', fontSize: '1.2rem' }}>{res.final25}</td>
                        </>
                      ) : (
                        <>
                          <td><input className="marks-input" style={{ maxWidth: '80px' }} value={student.mid1} onChange={(e) => updateStudentField(index, 'mid1', e.target.value)} /></td>
                          <td><input className="marks-input" style={{ maxWidth: '80px' }} value={student.mid2} onChange={(e) => updateStudentField(index, 'mid2', e.target.value)} /></td>
                          <td style={{ fontWeight: 600, color: '#334155' }}>{res.internalMarks25}</td>
                          <td><input className="marks-input" style={{ maxWidth: '70px' }} value={student.assign1} onChange={(e) => updateStudentField(index, 'assign1', e.target.value)} /></td>
                          <td><input className="marks-input" style={{ maxWidth: '70px' }} value={student.assign2} onChange={(e) => updateStudentField(index, 'assign2', e.target.value)} /></td>
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
          <div className="sign-box">Faculty Sign</div>
          <div className="sign-box">HOD Sign</div>
        </div>
      </div>
    </div>
  );
}

export default App;
