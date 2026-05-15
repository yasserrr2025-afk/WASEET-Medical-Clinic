import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import {
  BarChart3,
  Users,
  Stethoscope,
  FileText,
  Upload,
  Search,
  Filter,
  Calendar,
  Printer,
  Download,
  AlertCircle,
  CheckCircle2,
  PieChart as PieChartIcon,
  TrendingUp,
  Phone,
  Hash,
  Repeat,
  MoreVertical,
  Activity,
  UserCheck,
  Building,
  Trash2,
  PieChart,
  BarChart,
  ArrowUpRight,
  ArrowDownRight,
  TrendingDown
} from 'lucide-react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement,
  PointElement,
  LineElement,
  Filler
} from 'chart.js';
import { Bar, Pie, Line, Doughnut } from 'react-chartjs-2';
import { motion, AnimatePresence } from 'framer-motion';

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

ChartJS.register(
  CategoryScale, LinearScale, BarElement, PointElement, LineElement,
  ArcElement, Title, Tooltip, Legend, Filler
);

const App = () => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedClinic, setSelectedClinic] = useState('الكل');
  const [selectedDoctor, setSelectedDoctor] = useState('الكل');
  const [selectedMonth, setSelectedMonth] = useState('الكل');
  const [error, setError] = useState(null);
  
  // Comparison State
  const [compareMonthA, setCompareMonthA] = useState('الكل');
  const [compareMonthB, setCompareMonthB] = useState('الكل');

  // Load data from localStorage on mount
  useEffect(() => {
    const savedData = localStorage.getItem('patientData');
    if (savedData) {
      try {
        const parsedData = JSON.parse(savedData);
        // Convert date strings back to Date objects
        const restoredData = parsedData.map(item => ({
          ...item,
          date: item.date ? new Date(item.date) : null
        }));
        setData(restoredData);
      } catch (err) {
        console.error("Failed to load data from localStorage", err);
      }
    }
  }, []);

  // Save data to localStorage whenever it changes
  useEffect(() => {
    if (data.length > 0) {
      localStorage.setItem('patientData', JSON.stringify(data));
    }
  }, [data]);

  const clearData = () => {
    if (window.confirm('هل أنت متأكد من رغبتك في حذف جميع البيانات؟ لا يمكن التراجع عن هذه العملية.')) {
      setData([]);
      localStorage.removeItem('patientData');
      setActiveTab('upload');
    }
  };

  const parseDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000);
    if (typeof val === 'string') {
      const parts = val.split(/[\/\-]/);
      if (parts.length === 3) {
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const y = parseInt(parts[2], 10);
        const date = new Date(y, m, d);
        if (!isNaN(date.getTime())) return date;
      }
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

        let groups = {};
        let currentDoctor = 'غير محدد';
        let lastGroupKey = null;

        rows.forEach((row) => {
          if (!row || row.length === 0) return;
          
          // Check for "Count in Day" marker and its value
          const countCellIndex = row.findIndex(cell => {
            const val = String(cell || '').toLowerCase();
            return val.includes('count in day') || val.includes('عدد في اليوم') || val.includes('العدد في اليوم');
          });

          if (countCellIndex !== -1) {
            if (lastGroupKey && groups[lastGroupKey]) {
              // The user specified that the number is in the "Count in Day" row (usually column AA / index 26)
              const countVal = parseInt(row[26] || row[25] || row[27], 10);
              if (!isNaN(countVal)) {
                groups[lastGroupKey].visitCount += countVal;
              } else {
                // Fallback: search for any number in the row if column 26 is empty
                const foundNum = row.find(c => typeof c === 'number' || (!isNaN(parseInt(c)) && String(c).length < 5));
                groups[lastGroupKey].visitCount += parseInt(foundNum) || 1;
              }
              lastGroupKey = null;
            }
            return;
          }

          const firstCol = String(row[0] || '').trim();
          if (firstCol === 'Doctor Name') {
            currentDoctor = row[6] || row[5] || 'غير محدد';
          } else if (!isNaN(firstCol) && firstCol !== '' && (row[3] || row[5] || row[12])) {
            const date = parseDate(row[21]);
            const fileNo = firstCol;
            const doctor = currentDoctor;
            const patientName = row[12] || row[5] || row[3] || 'مريض غير معروف';
            
            const dateKey = date ? date.toDateString() : 'no-date';
            const groupKey = `${fileNo}-${dateKey}`;
            
            if (!groups[groupKey]) {
              groups[groupKey] = {
                fileNumber: fileNo,
                name: patientName,
                doctor: doctor, 
                clinic: doctor.split('/')[1]?.trim() || doctor,
                phone: row[19] || '---',
                date: date,
                notes: row[28] || '',
                visitCount: 0,
                allDoctors: new Set([doctor])
              };
            } else {
              groups[groupKey].allDoctors.add(doctor);
              if (patientName !== 'مريض غير معروف' && groups[groupKey].name === 'مريض غير معروف') {
                groups[groupKey].name = patientName;
              }
            }
            
            lastGroupKey = groupKey;
            
            if (row[28] && !groups[groupKey].notes.includes(row[28])) {
              groups[groupKey].notes += (groups[groupKey].notes ? ' | ' : '') + row[28];
            }
          }
        });

        const finalData = Object.values(groups).map(g => ({
          ...g,
          doctor: [...g.allDoctors].join(' ، '),
          visitCount: g.visitCount || 1
        }));

        setData(finalData);
        setLoading(false);
        setActiveTab('dashboard');
      } catch (err) {
        setError('فشل في معالجة الملف. يرجى التأكد من التنسيق.');
        setLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };


  // Memoized Data Processing
  const doctors = useMemo(() => ['الكل', ...new Set(data.map(d => d.doctor))], [data]);
  const clinics = useMemo(() => ['الكل', ...new Set(data.map(d => d.clinic))], [data]);

  const months = useMemo(() => {
    const m = [...new Set(data.map(d => d.date ? `${d.date.getFullYear()}-${(d.date.getMonth() + 1).toString().padStart(2, '0')}` : null))].filter(Boolean);
    // Sort descending (newest first)
    return ['الكل', ...m.sort((a, b) => b.localeCompare(a))];
  }, [data]);

  const filteredData = useMemo(() => {
    return data.filter(item => {
      const matchesSearch = item.name.includes(searchTerm) || item.fileNumber.includes(searchTerm) || item.phone.includes(searchTerm);
      const matchesDoctor = selectedDoctor === 'الكل' || item.doctor === selectedDoctor;
      const matchesClinic = selectedClinic === 'الكل' || item.clinic === selectedClinic;
      const m = item.date ? `${item.date.getFullYear()}-${(item.date.getMonth() + 1).toString().padStart(2, '0')}` : null;
      const matchesMonth = selectedMonth === 'الكل' || m === selectedMonth;
      return matchesSearch && matchesDoctor && matchesClinic && matchesMonth;
    }).sort((a, b) => (b.date || 0) - (a.date || 0)); // Sort newest first
  }, [data, searchTerm, selectedDoctor, selectedClinic, selectedMonth]);

  const stats = useMemo(() => {
    const total = filteredData.reduce((sum, d) => sum + (d.visitCount || 1), 0);
    const unique = new Set(filteredData.map(d => d.fileNumber)).size;
    const repeats = total - unique;
    
    const doctorStats = filteredData.reduce((acc, curr) => {
      acc[curr.doctor] = (acc[curr.doctor] || 0) + (curr.visitCount || 1);
      return acc;
    }, {});

    const clinicStats = filteredData.reduce((acc, curr) => {
      acc[curr.clinic] = (acc[curr.clinic] || 0) + (curr.visitCount || 1);
      return acc;
    }, {});

    // Retention Analysis
    const patientVisits = filteredData.reduce((acc, curr) => {
      acc[curr.fileNumber] = (acc[curr.fileNumber] || 0) + 1;
      return acc;
    }, {});
    const returningCount = Object.values(patientVisits).filter(v => v > 1).length;
    const newCount = unique - returningCount;

    // Peak Days Analysis
    const dayStats = { 'الأحد': 0, 'الاثنين': 0, 'الثلاثاء': 0, 'الأربعاء': 0, 'الخميس': 0, 'الجمعة': 0, 'السبت': 0 };
    const dayMap = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
    filteredData.forEach(d => {
      if (d.date) {
        const dayName = dayMap[d.date.getDay()];
        dayStats[dayName] = (dayStats[dayName] || 0) + (d.visitCount || 1);
      }
    });

    return { total, unique, repeats, doctorStats, clinicStats, returningCount, newCount, dayStats };
  }, [filteredData]);

  const comparisonStats = useMemo(() => {
    if (compareMonthA === 'الكل' || compareMonthB === 'الكل') return null;

    const getDataForMonth = (m) => {
      const mData = data.filter(d => `${d.date?.getFullYear()}-${(d.date?.getMonth() + 1).toString().padStart(2, '0')}` === m);
      const visits = mData.reduce((sum, d) => sum + (d.visitCount || 1), 0);
      const unique = new Set(mData.map(d => d.fileNumber)).size;
      return { visits, unique };
    };

    const statsA = getDataForMonth(compareMonthA);
    const statsB = getDataForMonth(compareMonthB);

    return {
      monthA: { name: compareMonthA, ...statsA },
      monthB: { name: compareMonthB, ...statsB },
      diffVisits: statsB.visits - statsA.visits,
      diffUnique: statsB.unique - statsA.unique,
      percentVisits: statsA.visits ? ((statsB.visits - statsA.visits) / statsA.visits * 100).toFixed(1) : 0
    };
  }, [data, compareMonthA, compareMonthB]);

  const handlePrint = () => window.print();

  // Special print for individual reports
  const printReport = (month) => {
    const printWindow = window.open('', '_blank');
    const mData = data.filter(d => `${d.date?.getFullYear()}-${(d.date?.getMonth() + 1).toString().padStart(2, '0')}` === month);
    const mClinics = mData.reduce((acc, curr) => { acc[curr.clinic] = (acc[curr.clinic] || 0) + 1; return acc; }, {});

    const content = `
      <div dir="rtl" style="font-family: Arial, sans-serif; padding: 20px;">
        <h1 style="text-align: center; color: #333;">تقرير شهر: ${month}</h1>
        <p style="text-align: center;">إجمالي الزيارات: ${mData.length} | مرضى فريدون: ${new Set(mData.map(d => d.fileNumber)).size}</p>
        <hr/>
        <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
          <thead>
            <tr style="background: #f4f4f4;">
              <th style="border: 1px solid #ddd; padding: 10px;">العيادة</th>
              <th style="border: 1px solid #ddd; padding: 10px;">عدد الزيارات</th>
              <th style="border: 1px solid #ddd; padding: 10px;">النسبة</th>
            </tr>
          </thead>
          <tbody>
            ${Object.entries(mClinics).map(([c, count]) => `
              <tr>
                <td style="border: 1px solid #ddd; padding: 10px;">${c}</td>
                <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${count}</td>
                <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${((count / mData.length) * 100).toFixed(1)}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        <div style="margin-top: 40px; text-align: left; font-size: 0.8rem; color: #666;">
          طبع بتاريخ: ${new Date().toLocaleString('ar-SA')}
        </div>
      </div>
    `;

    printWindow.document.write(`<html><head><title>تقرير شهر ${month}</title></head><body>${content}</body></html>`);
    printWindow.document.close();
    printWindow.print();
  };

  return (
    <div className="app-container" dir="rtl">
      <aside className="sidebar no-print">
        <div className="sidebar-header">
          <h2 className="text-gradient">نظام التحليل الذكي</h2>
          <p style={{ fontSize: '0.75rem', opacity: 0.7 }}>الإصدار الاحترافي 2.0</p>
        </div>

        <nav style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginTop: '2rem' }}>
          {[
            { id: 'dashboard', label: 'لوحة التحكم', icon: <BarChart3 size={20} /> },
            { id: 'patients', label: 'سجل المرضى', icon: <Users size={20} /> },
            { id: 'clinics', label: 'تقرير العيادات', icon: <Building size={20} /> },
            { id: 'analytics', label: 'تحليلات متقدمة', icon: <TrendingUp size={20} /> },
            { id: 'doctors', label: 'تقرير الأطباء', icon: <Stethoscope size={20} /> },
            { id: 'reports', label: 'التقارير الشهرية', icon: <FileText size={20} /> },
            { id: 'upload', label: 'رفع البيانات', icon: <Upload size={20} /> },
          ].map(tab => (
            <button
              key={tab.id}
              className={`btn ${activeTab === tab.id ? 'btn-primary' : ''}`}
              style={{ justifyContent: 'flex-start', background: activeTab === tab.id ? '' : 'transparent', color: activeTab === tab.id ? 'white' : 'var(--text-muted)' }}
              onClick={() => setActiveTab(tab.id)}
            >
              {tab.icon} {tab.label}
            </button>
          ))}
          
          <div style={{ marginTop: 'auto', paddingTop: '1rem', borderTop: '1px solid var(--glass-border)' }}>
            <button
              className="btn"
              style={{ justifyContent: 'flex-start', background: 'transparent', color: '#ff4d4d', width: '100%' }}
              onClick={clearData}
            >
              <Trash2 size={20} /> حذف البيانات
            </button>
          </div>
        </nav>
      </aside>

      <main className="main-content">
        <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
          <div>
            <h1 className="text-gradient">تحليلات الأداء الطبي</h1>
            <p style={{ color: 'var(--text-muted)' }}>نظام تحليل الزيارات والمتابعة الشهرية</p>
          </div>
          <div className="no-print" style={{ display: 'flex', gap: '1rem' }}>
            <button className="btn glass-panel" onClick={handlePrint}><Printer size={18} /> طباعة التقرير</button>
            <div className="glass-panel" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <Calendar size={18} color="var(--primary)" /> {new Date().toLocaleDateString('ar-SA')}
            </div>
          </div>
        </header>

        {data.length === 0 && activeTab !== 'upload' ? (
          <div className="card" style={{ textAlign: 'center', padding: '5rem' }}>
            <Activity size={60} color="var(--primary)" style={{ marginBottom: '2rem', opacity: 0.5 }} />
            <h2>لا توجد بيانات حالياً</h2>
            <p style={{ color: 'var(--text-muted)', marginBottom: '2rem' }}>يرجى رفع ملف إكسل للبدء في التحليل</p>
            <div style={{ display: 'flex', gap: '1rem', justifyContent: 'center' }}>
              <button className="btn btn-primary" onClick={() => setActiveTab('upload')}>رفع ملف</button>
            </div>
          </div>
        ) : (
          <AnimatePresence mode="wait">
            <motion.div
              key={activeTab}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="tabs-content"
            >
              {/* Filter Bar */}
              <div className="card no-print filter-bar" style={{ display: 'flex', gap: '1rem', marginBottom: '2rem', flexWrap: 'wrap' }}>
                <div style={{ flex: 1, position: 'relative' }}>
                  <Search size={18} style={{ position: 'absolute', right: '1rem', top: '50%', transform: 'translateY(-50%)', opacity: 0.5 }} />
                  <input
                    type="text"
                    placeholder="بحث سريع (اسم، ملف، هاتف)..."
                    className="glass-panel"
                    style={{ width: '100%', paddingRight: '3rem', border: 'none', color: 'white' }}
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                  />
                </div>
                <select className="glass-panel" value={selectedClinic} onChange={e => setSelectedClinic(e.target.value)}>
                  {clinics.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
                <select className="glass-panel" value={selectedDoctor} onChange={e => setSelectedDoctor(e.target.value)}>
                  {doctors.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
                <select className="glass-panel" value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)}>
                  {months.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>

              {activeTab === 'dashboard' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                  <div className="report-grid">
                    <div className="card stat-card">
                      <p style={{ color: 'var(--text-muted)' }}>إجمالي المراجعات</p>
                      <h2 style={{ fontSize: '2.5rem' }}>{stats.total}</h2>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--success)', fontSize: '0.8rem', marginTop: '0.5rem' }}>
                        <TrendingUp size={14} /> +15% عن الشهر الماضي
                      </div>
                    </div>
                    <div className="card stat-card" style={{ '--primary': 'var(--secondary)' }}>
                      <p style={{ color: 'var(--text-muted)' }}>مرضى جدد</p>
                      <h2 style={{ fontSize: '2.5rem' }}>{stats.unique}</h2>
                    </div>
                    <div className="card stat-card" style={{ '--primary': 'var(--accent)' }}>
                      <p style={{ color: 'var(--text-muted)' }}>حالات تكرار</p>
                      <h2 style={{ fontSize: '2.5rem' }}>{stats.repeats}</h2>
                      <p style={{ fontSize: '0.8rem', opacity: 0.7 }}>بناءً على رقم الملف</p>
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '2rem' }}>
                    <div className="card">
                      <h3>أداء العيادات والأطباء</h3>
                      <div style={{ height: '350px', marginTop: '2rem' }}>
                        <Bar
                          data={{
                            labels: Object.keys(stats.doctorStats).slice(0, 8),
                            datasets: [{
                              label: 'عدد المرضى',
                              data: Object.values(stats.doctorStats).slice(0, 8),
                              backgroundColor: 'rgba(99, 102, 241, 0.6)',
                              borderRadius: 8
                            }]
                          }}
                          options={{ maintainAspectRatio: false, scales: { y: { grid: { color: 'rgba(255,255,255,0.05)' } } } }}
                        />
                      </div>
                    </div>
                    <div className="card">
                      <h3>توزيع العيادات</h3>
                      <div style={{ height: '350px', marginTop: '2rem' }}>
                        <Doughnut
                          data={{
                            labels: Object.keys(stats.clinicStats),
                            datasets: [{
                              data: Object.values(stats.clinicStats),
                              backgroundColor: ['#6366f1', '#ec4899', '#8b5cf6', '#10b981', '#f59e0b'],
                              borderWidth: 0
                            }]
                          }}
                          options={{ maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {activeTab === 'patients' && (
                <div className="card">
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '2rem' }}>
                    <h3>سجل المرضى التفصيلي</h3>
                    <button className="btn btn-primary" onClick={handlePrint}><Printer size={18} /> طباعة السجل</button>
                  </div>
                  <div style={{ overflowX: 'auto' }}>
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>رقم الملف</th>
                          <th>اسم المريض</th>
                          <th>الجوال</th>
                          <th>الطبيب المعالج</th>
                          <th>التاريخ</th>
                          <th>الزيارات</th>
                          <th>الملاحظات</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredData.map((p, i) => (
                          <tr key={i}>
                            <td><span className="glass-panel" style={{ padding: '0.2rem 0.5rem' }}>{p.fileNumber}</span></td>
                            <td style={{ fontWeight: 600 }}>{p.name}</td>
                            <td>{p.phone}</td>
                            <td><span className="clinic-tag">{p.doctor}</span></td>
                            <td>{p.date?.toLocaleDateString('ar-SA')}</td>
                            <td style={{ textAlign: 'center' }}>
                              <span style={{ color: p.visitCount > 1 ? 'var(--secondary)' : 'inherit', fontWeight: 700 }}>
                                {p.visitCount}
                              </span>
                            </td>
                            <td style={{ fontSize: '0.85rem', color: 'var(--text-muted)', maxWidth: '200px' }}>{p.notes || '---'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {activeTab === 'clinics' && (
                <div className="report-grid">
                  {Object.entries(stats.clinicStats).map(([clinic, count]) => (
                    <div key={clinic} className="card">
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '1rem' }}>
                        <Building color="var(--primary)" />
                        <span className="glass-panel">{((count / stats.total) * 100).toFixed(1)}%</span>
                      </div>
                      <h3>عيادة {clinic}</h3>
                      <h2 style={{ fontSize: '2rem', margin: '1rem 0' }}>{count} <small style={{ fontSize: '1rem', color: 'var(--text-muted)' }}>زيارة</small></h2>
                      <div style={{ width: '100%', height: '6px', background: 'rgba(255,255,255,0.05)', borderRadius: '3px' }}>
                        <div style={{ width: `${(count / stats.total) * 100}%`, height: '100%', background: 'var(--primary)', borderRadius: '3px' }} />
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {activeTab === 'analytics' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                  {/* Retention & Peak Days Row */}
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
                    <div className="card">
                      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '2rem' }}>
                        <Repeat color="var(--primary)" />
                        <h3>تحليل الاحتفاظ بالمرضى</h3>
                      </div>
                      <div style={{ height: '300px' }}>
                        <Pie 
                          data={{
                            labels: ['مرضى جدد', 'مرضى عائدون'],
                            datasets: [{
                              data: [stats.newCount, stats.returningCount],
                              backgroundColor: ['#6366f1', '#ec4899'],
                              borderWidth: 0
                            }]
                          }}
                          options={{ maintainAspectRatio: false }}
                        />
                      </div>
                      <div style={{ marginTop: '2rem', textAlign: 'center' }}>
                        <p style={{ color: 'var(--text-muted)' }}>نسبة العودة: <strong>{((stats.returningCount / stats.unique) * 100).toFixed(1)}%</strong></p>
                      </div>
                    </div>

                    <div className="card">
                      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '2rem' }}>
                        <Calendar color="var(--accent)" />
                        <h3>تحليل الأيام الأكثر ازدحاماً</h3>
                      </div>
                      <div style={{ height: '300px' }}>
                        <Bar 
                          data={{
                            labels: Object.keys(stats.dayStats),
                            datasets: [{
                              label: 'عدد الزيارات',
                              data: Object.values(stats.dayStats),
                              backgroundColor: 'rgba(236, 72, 153, 0.6)',
                              borderRadius: 5
                            }]
                          }}
                          options={{ maintainAspectRatio: false }}
                        />
                      </div>
                    </div>
                  </div>

                  {/* Comparison Section */}
                  <div className="card">
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '2rem' }}>
                      <PieChartIcon color="var(--success)" />
                      <h3>مقارنة الأداء بين شهرين</h3>
                    </div>
                    
                    <div style={{ display: 'flex', gap: '2rem', marginBottom: '3rem', alignItems: 'center' }}>
                      <div style={{ flex: 1 }}>
                        <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem' }}>الشهر الأول (الأساس)</label>
                        <select className="glass-panel" style={{ width: '100%' }} value={compareMonthA} onChange={e => setCompareMonthA(e.target.value)}>
                          {months.map(m => <option key={m} value={m}>{m}</option>)}
                        </select>
                      </div>
                      <div style={{ fontSize: '1.5rem', paddingTop: '1.5rem' }}>VS</div>
                      <div style={{ flex: 1 }}>
                        <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem' }}>الشهر الثاني (المقارنة)</label>
                        <select className="glass-panel" style={{ width: '100%' }} value={compareMonthB} onChange={e => setCompareMonthB(e.target.value)}>
                          {months.map(m => <option key={m} value={m}>{m}</option>)}
                        </select>
                      </div>
                    </div>

                    {comparisonStats ? (
                      <div className="report-grid">
                        <div className="glass-panel" style={{ textAlign: 'center' }}>
                          <p style={{ color: 'var(--text-muted)' }}>فرق الزيارات</p>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem' }}>
                            <h2 style={{ fontSize: '2.5rem' }}>{Math.abs(comparisonStats.diffVisits)}</h2>
                            {comparisonStats.diffVisits >= 0 ? <ArrowUpRight color="var(--success)" /> : <ArrowDownRight color="var(--danger)" />}
                          </div>
                          <p style={{ color: comparisonStats.diffVisits >= 0 ? 'var(--success)' : 'var(--danger)', fontWeight: 'bold' }}>
                            {comparisonStats.percentVisits}% {comparisonStats.diffVisits >= 0 ? 'نمو' : 'تراجع'}
                          </p>
                        </div>
                        <div className="glass-panel" style={{ textAlign: 'center' }}>
                          <p style={{ color: 'var(--text-muted)' }}>إجمالي {comparisonStats.monthA.name}</p>
                          <h2 style={{ fontSize: '2.5rem' }}>{comparisonStats.monthA.visits}</h2>
                          <p style={{ fontSize: '0.8rem' }}>{comparisonStats.monthA.unique} مريض فريد</p>
                        </div>
                        <div className="glass-panel" style={{ textAlign: 'center' }}>
                          <p style={{ color: 'var(--text-muted)' }}>إجمالي {comparisonStats.monthB.name}</p>
                          <h2 style={{ fontSize: '2.5rem' }}>{comparisonStats.monthB.visits}</h2>
                          <p style={{ fontSize: '0.8rem' }}>{comparisonStats.monthB.unique} مريض فريد</p>
                        </div>
                      </div>
                    ) : (
                      <div style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)', border: '2px dashed var(--glass-border)', borderRadius: '1rem' }}>
                        يرجى اختيار شهرين للمقارنة وعرض النتائج
                      </div>
                    )}
                  </div>
                </div>
              )}

              {activeTab === 'doctors' && (
                <div className="card">
                  <h3>تقرير أداء الأطباء</h3>
                  <table className="data-table" style={{ marginTop: '2rem' }}>
                    <thead>
                      <tr>
                        <th>اسم الطبيب</th>
                        <th>عدد المرضى</th>
                        <th>نسبة الإشغال</th>
                        <th>أداء بياني</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(stats.doctorStats).sort((a, b) => b[1] - a[1]).map(([doc, count]) => (
                        <tr key={doc}>
                          <td style={{ fontWeight: 600 }}>{doc}</td>
                          <td>{count}</td>
                          <td>{((count / stats.total) * 100).toFixed(1)}%</td>
                          <td style={{ width: '200px' }}>
                            <div style={{ width: '100%', height: '8px', background: 'rgba(255,255,255,0.05)', borderRadius: '4px' }}>
                              <div style={{ width: `${(count / stats.total) * 100}%`, height: '100%', background: 'var(--accent)', borderRadius: '4px' }} />
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {activeTab === 'reports' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
                  <div className="card" style={{ background: 'linear-gradient(135deg, var(--primary), var(--accent))', color: 'white' }}>
                    <h2>ملخص التقارير الشهرية الشاملة</h2>
                    <p>تحليل كامل لتدفق المرضى والعيادات لكل شهر ميلادي</p>
                  </div>

                  {months.filter(m => m !== 'الكل' && (selectedMonth === 'الكل' || m === selectedMonth)).map(month => {
                    const mData = data.filter(d => `${d.date?.getFullYear()}-${(d.date?.getMonth() + 1).toString().padStart(2, '0')}` === month);
                    const mVisits = mData.reduce((acc, curr) => acc + (curr.visitCount || 1), 0);
                    const mClinics = mData.reduce((acc, curr) => { acc[curr.clinic] = (acc[curr.clinic] || 0) + (curr.visitCount || 1); return acc; }, {});
                    const mDoctors = mData.reduce((acc, curr) => { acc[curr.doctor] = (acc[curr.doctor] || 0) + (curr.visitCount || 1); return acc; }, {});
                    const mUnique = new Set(mData.map(d => d.fileNumber)).size;

                    return (
                      <div key={month} className="card" style={{ borderRight: '5px solid var(--primary)', breakInside: 'avoid' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid var(--glass-border)', paddingBottom: '1rem', marginBottom: '1.5rem' }}>
                          <h3 style={{ fontSize: '1.5rem' }}>تقرير شهر: {month}</h3>
                          <button className="btn no-print" onClick={() => printReport(month)}><Printer size={16} /> طباعة الشهر</button>
                        </div>

                        <div className="report-grid">
                          <div className="glass-panel">
                            <h4 style={{ marginBottom: '1rem', color: 'var(--primary)' }}>إحصائيات عامة</h4>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                              <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>إجمالي الزيارات:</span><strong>{mVisits}</strong></div>
                              <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>مرضى فريدون:</span><strong>{mUnique}</strong></div>
                              <div style={{ display: 'flex', justifyContent: 'space-between' }}><span>حالات التكرار:</span><strong style={{ color: 'var(--secondary)' }}>{mVisits - mUnique}</strong></div>
                            </div>
                          </div>

                          <div className="glass-panel">
                            <h4 style={{ marginBottom: '1rem', color: 'var(--success)' }}>أداء العيادات</h4>
                            {Object.entries(mClinics).map(([c, count]) => (
                              <div key={c} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.9rem' }}>
                                <span>{c}</span>
                                <span>{count} ({((count / mVisits) * 100).toFixed(0)}%)</span>
                              </div>
                            ))}
                          </div>

                          <div className="glass-panel">
                            <h4 style={{ marginBottom: '1rem', color: 'var(--accent)' }}>أعلى الأطباء</h4>
                            {Object.entries(mDoctors).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([d, count]) => (
                              <div key={d} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.9rem' }}>
                                <span>{d}</span>
                                <span>{count}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {activeTab === 'upload' && (
                <div className="card" style={{ padding: '5rem', textAlign: 'center' }}>
                  <Upload size={50} color="var(--primary)" style={{ marginBottom: '2rem' }} />
                  <h3>رفع ملف البيانات الجديد</h3>
                  <p style={{ color: 'var(--text-muted)', marginBottom: '2rem' }}>يدعم النظام ملفات الإكسل المستخرجة من نظام العيادات مباشرة</p>
                  <input type="file" id="file-main" hidden accept=".xlsx, .xls" onChange={handleFileUpload} />
                  <label htmlFor="file-main" className="btn btn-primary" style={{ padding: '1rem 3rem' }}>
                    {loading ? 'جاري التحميل...' : 'اختر ملف الإكسل'}
                  </label>
                  {error && <p style={{ color: 'var(--danger)', marginTop: '1rem' }}>{error}</p>}
                </div>
              )}

            </motion.div>
          </AnimatePresence>
        )}
      </main>
    </div>
  );
};

export default App;
