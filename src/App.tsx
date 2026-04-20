import React, { useState, useMemo, useEffect } from 'react';
import { 
  School, LogOut, Users, CheckCircle2, Clock,
  TrendingUp, AlertTriangle, LayoutDashboard, FileText,
  Search, Settings, Calendar as CalendarIcon,
  ListOrdered, Loader2, UploadCloud, Printer 
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwZHhZ-RPWUpBGIlw0qTFPUmOPmq9WpcvW4WLklcjb_A9U3MW0luIXYPnHznI29ThpbMA/exec"; 

type Section = 'dashboard' | 'reports' | 'search' | 'settings';

const ARABIC_DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
const getArabicDayName = (date: Date) => ARABIC_DAYS[date.getDay()];

const isSameDate = (dateString: string, compareDate: Date) => {
  if (!dateString) return false;
  const d = new Date(dateString);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() === compareDate.getFullYear() && 
           d.getMonth() === compareDate.getMonth() && 
           d.getDate() === compareDate.getDate();
  }
  const day = compareDate.getDate().toString();
  const month = (compareDate.getMonth() + 1).toString();
  const arDay = day.replace(/\d/g, (x:any) => '٠١٢٣٤٥٦٧٨٩'[x]);
  const arMonth = month.replace(/\d/g, (x:any) => '٠١٢٣٤٥٦٧٨٩'[x]);
  const str = dateString.toString();
  return (str.includes(day) || str.includes(arDay)) && (str.includes(month) || str.includes(arMonth));
};

// 💉 تحديث مُولد الوثائق لدعم تقارير الطلاب والمعلمين معاً
const generateOfficialReport = (title: string, dataList: any[], schoolName: string, dateStr: string, type: 'students' | 'teachers' = 'students') => {
  const printWindow = window.open('', '_blank');
  if (!printWindow) {
    alert('الرجاء السماح بالنوافذ المنبثقة (Pop-ups) لطباعة التقرير.');
    return;
  }

  let tableHeader = type === 'students' 
    ? `<tr><th style="width: 5%;">م</th><th style="width: 30%;">اسم الطالب</th><th style="width: 15%;">نوع المخالفة</th><th style="width: 10%;">الفصل</th><th style="width: 25%;">المعلم الراصد</th><th style="width: 15%;">الوقت</th></tr>`
    : `<tr><th style="width: 10%;">م</th><th style="width: 45%;">اسم المعلم </th><th style="width: 25%;">الفصل المسند</th><th style="width: 20%;">حالة الرصد</th></tr>`;

  let tableContent = '';
  if (dataList.length === 0) {
      tableContent = `<tr><td colspan="${type === 'students' ? 6 : 4}" style="text-align:center; padding: 20px; font-weight: bold;">لا يوجد سجلات في هذا التقرير.</td></tr>`;
  } else {
      dataList.forEach((item, index) => {
          tableContent += type === 'students' 
            ? `<tr>
                 <td>${index + 1}</td>
                 <td style="font-weight: bold;">${item.name}</td>
                 <td>${item.type || 'غياب / تأخير'}</td>
                 <td>${item.className}</td>
                 <td>${item.teacher}</td>
                 <td dir="ltr">${item.time}</td>
               </tr>`
            : `<tr>
                 <td>${index + 1}</td>
                 <td style="font-weight: bold; text-align: right; padding-right: 20px;">${item.name}</td>
                 <td>${item.className}</td>
                 <td style="color: red; font-weight: bold;">لم يتم الرصد</td>
               </tr>`;
      });
  }

  const htmlContent = `
    <html dir="rtl" lang="ar">
    <head>
      <title>تقرير نظام راصد - ${title}</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@500;700;900&display=swap');
        @page { size: A4; margin: 15mm; }
        body { font-family: 'Tajawal', Tahoma, sans-serif; color: #000; margin: 0; padding: 0; -webkit-print-color-adjust: exact; }
        .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 3px double #000; padding-bottom: 15px; margin-bottom: 30px; }
        .right-header { text-align: right; font-weight: 900; line-height: 1.6; font-size: 13pt; }
        .left-header { text-align: left; font-weight: 700; line-height: 1.6; font-size: 12pt; }
        .title { text-align: center; font-size: 18pt; font-weight: 900; margin: 30px 0; text-decoration: underline; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 40px; font-size: 12pt; }
        th, td { border: 1px solid #000; padding: 10px 8px; text-align: center; }
        th { background-color: #f1f5f9; font-weight: 900; font-size: 13pt; }
        .signatures { display: flex; justify-content: space-between; margin-top: 70px; font-weight: 900; font-size: 14pt; padding: 0 20px; }
        .watermark { position: fixed; top: 45%; left: 50%; transform: translate(-50%, -50%) rotate(-45deg); font-size: 130px; font-weight: 900; color: rgba(0,0,0,0.03); z-index: -1; white-space: nowrap; pointer-events: none;}
        .footer-note { text-align: center; font-size: 9pt; font-weight: bold; color: #64748b; margin-top: 50px; border-top: 1px solid #cbd5e1; padding-top: 10px; }
      </style>
    </head>
    <body>
       <div class="header">
          <div class="right-header">
            سلطنة عُمان<br>
            وزارة التعليم<br>
            المديرية العامة للتعليم بمحافظة شمال الباطنة<br>
          </div>
          <div class="left-header">
            المدرسة: ${schoolName || '____________________'}<br>
            التاريخ: ${dateStr}<br>
            النظام: راصد الإدارة
          </div>
       </div>
       
       <div class="title">${title}</div>
       
       <table>
         <thead>
           ${tableHeader}
         </thead>
         <tbody>
           ${tableContent}
         </tbody>
       </table>

       <div class="signatures">
         ${type === 'students' ? '<div>مُعد التقرير: ........................</div>' : '<div>الختم الرسمي:</div>'}
         <div>يعتمد، مدير المدرسة: ........................</div>
       </div>

       <div class="footer-note">
          تم استخراج هذا التقرير إلكترونياً من نظام راصد - برمجة وتطوير: محمد الزعابي
       </div>

       <script>
         window.onload = function() {
           setTimeout(function() { window.print(); }, 500);
         };
       </script>
    </body>
    </html>
  `;

  printWindow.document.open();
  printWindow.document.write(htmlContent);
  printWindow.document.close();
};

export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  // 💉 استرجاع الكود المحفوظ تلقائياً
  const [schoolCode, setSchoolCode] = useState(() => localStorage.getItem('rased_admin_code') || '');
  const [schoolName, setSchoolName] = useState(() => localStorage.getItem('rased_school_name') || '');

  useEffect(() => {
    localStorage.setItem('rased_school_name', schoolName);
  }, [schoolName]);

  const [activeSection, setActiveSection] = useState<Section>('dashboard');
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [dashboardData, setDashboardData] = useState({ teachers: [], logs: [] });

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (schoolCode.length < 2) return;
    setIsLoading(true);
    setErrorMsg('');
    try {
      const cacheBuster = new Date().getTime();
      const response = await fetch(`${SCRIPT_URL}?schoolCode=${schoolCode}&t=${cacheBuster}`);
      const result = await response.json();
      if (result.status === "success") {
        setDashboardData(result.data);
        setIsLoggedIn(true);
        // 💉 حفظ كود المدرسة بنجاح في الذاكرة لتسهيل الدخول القادم
        localStorage.setItem('rased_admin_code', schoolCode);
      } else {
        setErrorMsg('حدث خطأ من السيرفر: ' + result.message);
      }
    } catch (error) {
      console.error("تفاصيل الخطأ:", error); 
      setErrorMsg('تأكد من اتصالك بالإنترنت ومن صحة الرابط');
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    // 💉 عند الخروج لا نمسح الكود ليبقى جاهزاً، فقط نمسح الجلسة
    setActiveSection('dashboard');
    setDashboardData({ teachers: [], logs: [] });
  };

  if (!isLoggedIn) return <LoginScreen schoolCode={schoolCode} setSchoolCode={setSchoolCode} onLogin={handleLogin} isLoading={isLoading} errorMsg={errorMsg} />;

  return (
    <div className="flex h-screen overflow-hidden bg-[#f8fafc] text-slate-800 font-sans" dir="rtl">
      <div className="fixed inset-0 z-0 pointer-events-none overflow-hidden">
        <div className="absolute top-[-10%] right-[-5%] w-[40vw] h-[40vw] rounded-full bg-indigo-100/50 blur-[100px]" />
        <div className="absolute bottom-[-10%] left-[-5%] w-[30vw] h-[30vw] rounded-full bg-amber-100/40 blur-[80px]" />
      </div>

      <Sidebar activeSection={activeSection} setActiveSection={setActiveSection} />

      <div className="flex-1 flex flex-col h-full relative z-10 overflow-hidden">
        <TopHeader onLogout={handleLogout} schoolCode={schoolCode} schoolName={schoolName} />
        
        <main className="flex-1 overflow-y-auto p-4 sm:p-8">
          <AnimatePresence mode="wait">
            <motion.div key={activeSection} initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.3 }} className="max-w-7xl mx-auto h-full">
              {activeSection === 'dashboard' && <DashboardHome data={dashboardData} schoolName={schoolName} />}
              {activeSection === 'reports' && <ReportsPage data={dashboardData} schoolName={schoolName} />}
              {activeSection === 'search' && <SearchPage data={dashboardData} />}
              {activeSection === 'settings' && <SettingsPage schoolCode={schoolCode} schoolName={schoolName} setSchoolName={setSchoolName} />}
            </motion.div>
          </AnimatePresence>
        </main>
      </div>
    </div>
  );
}

// =========================================================
// 1️⃣ اللوحة الرئيسية
// =========================================================
function DashboardHome({ data, schoolName }: { data: any, schoolName: string }) {
  const [activeTab, setActiveTab] = useState<'absent' | 'late' | 'truant'>('absent');
  const todayDate = new Date();
  const todayArabicName = getArabicDayName(todayDate);

  const expectedTeachersToday = useMemo(() => data.teachers.filter((t: any) => t.day === todayArabicName), [data.teachers, todayArabicName]);
  const todayLogs = useMemo(() => data.logs.filter((log: any) => isSameDate(log.time, todayDate)), [data.logs, todayDate]);

  const uniqueTodayLogs = useMemo(() => {
    const map = new Map();
    todayLogs.forEach((log: any) => {
      if (!map.has(log.teacherName)) map.set(log.teacherName, log);
    });
    return Array.from(map.values());
  }, [todayLogs]);

  const lateTeachers = useMemo(() => {
    const loggedTeacherNames = new Set(uniqueTodayLogs.map((l: any) => l.teacherName));
    return expectedTeachersToday.filter((t: any) => !loggedTeacherNames.has(t.name));
  }, [expectedTeachersToday, uniqueTodayLogs]);

  const extractList = (key: string, label: string) => {
    let list: any[] = [];
    const seen = new Set(); 
    const chronologicalLogs = [...todayLogs].reverse();

    chronologicalLogs.forEach((log: any) => {
      if (log[key] && log[key] !== "لا يوجد" && log[key] !== "حضور كامل") {
        log[key].split("، ").forEach((name: string) => {
          const studentName = name.trim();
          if(studentName) {
            const uniqueKey = `${studentName}-${log.teacherName}`;
            if (!seen.has(uniqueKey)) {
              seen.add(uniqueKey);
              list.push({ 
                name: studentName, 
                type: label,
                teacher: log.teacherName, 
                className: log.className, 
                time: new Date(log.time).toLocaleTimeString('ar-SA', {hour: '2-digit', minute:'2-digit'}) 
              });
            }
          }
        });
      }
    });
    return list;
  };

  const allAbsent = useMemo(() => extractList('absentStudents', 'غياب'), [todayLogs]);
  const allLate = useMemo(() => extractList('lateStudents', 'تأخير'), [todayLogs]);
  const allTruant = useMemo(() => extractList('truantStudents', 'تسرب'), [todayLogs]);

  const completionRate = expectedTeachersToday.length === 0 ? 0 : Math.round((todayLogs.length / expectedTeachersToday.length) * 100);
  const activeListData = activeTab === 'absent' ? allAbsent : activeTab === 'late' ? allLate : allTruant;

  const handlePrintDaily = () => {
    const combinedData = [...allAbsent, ...allLate, ...allTruant];
    const dateStr = todayDate.toLocaleDateString('ar-OM');
    generateOfficialReport('التقرير اليومي لرصد المخالفات', combinedData, schoolName, dateStr, 'students');
  };

  // 💉 دالة طباعة تقرير المعلمين المتأخرين
  const handlePrintLateTeachers = () => {
    const dateStr = todayDate.toLocaleDateString('ar-OM');
    generateOfficialReport('كشف المعلمين المتأخرين عن رصد غياب الحصة الأولى', lateTeachers, schoolName, dateStr, 'teachers');
  };

  return (
    <div className="space-y-8 pb-10">
      <div className="bg-indigo-900 text-white p-4 rounded-2xl shadow-lg flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="font-black text-xl">إحصائيات اليوم المباشرة</h2>
          <p className="text-indigo-200 text-sm">مخصصة لمعلمي الحصة الأولى ليوم ({todayArabicName})</p>
        </div>
        <div className="flex items-center gap-3">
          <div className="bg-amber-400 text-indigo-950 font-black px-4 py-2 rounded-xl">
            {expectedTeachersToday.length} معلم مستهدف
          </div>
          <button onClick={handlePrintDaily} className="flex items-center gap-2 bg-white text-indigo-900 px-4 py-2 rounded-xl font-bold hover:bg-slate-100 transition shadow-sm active:scale-95">
            <Printer size={18} /> طباعة التقرير الرسمي
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-right">
        <StatCard title="نسبة إنجاز الرصد" value={`${completionRate}%`} icon={TrendingUp} color="indigo" />
        <StatCard title="إجمالي المخالفات اليوم" value={allAbsent.length + allLate.length + allTruant.length} subtitle="حالة" icon={Users} color="amber" />
        <StatCard title="معلمون متأخرون" value={lateTeachers.length} subtitle="معلم" icon={Clock} color="rose" />
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-12 gap-8">
        <div className="xl:col-span-4 flex flex-col gap-4">
          <div className="flex items-center justify-between px-2">
            <div className="flex items-center gap-3">
              <AlertTriangle className="text-rose-500" size={24} />
              <h2 className="text-xl font-black text-slate-800">تأخر رصد اليوم</h2>
            </div>
            
            {/* 💉 زر طباعة تقرير المعلمين بجانب العنوان */}
            {lateTeachers.length > 0 && (
              <button 
                onClick={handlePrintLateTeachers}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-lg text-xs font-bold transition-all shadow-sm active:scale-95"
                title="طباعة كشف المتقاعسين"
              >
                <Printer size={14} /> طباعة الكشف
              </button>
            )}
          </div>
          <div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 border border-white/80 shadow-md flex-1 overflow-y-auto max-h-[500px] custom-scrollbar">
            {lateTeachers.length === 0 ? (
              <div className="py-10 text-center text-emerald-600 font-bold"><CheckCircle2 size={40} className="mx-auto mb-3" />اكتمل الرصد!</div>
            ) : (
              <div className="space-y-3">
                {lateTeachers.map((teacher: any, idx: number) => (
                  <div key={idx} className="flex items-center gap-3 p-3 rounded-2xl bg-rose-50 border border-rose-100">
                    <div className="w-10 h-10 rounded-xl bg-rose-200 text-rose-600 flex items-center justify-center"><Clock size={18} /></div>
                    <div>
                      <h3 className="font-bold text-slate-800 text-sm">{teacher.name}</h3>
                      <p className="text-xs text-rose-600 font-bold">{teacher.className}</p>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        <div className="xl:col-span-8 flex flex-col gap-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between px-2 gap-3">
            <div className="flex items-center gap-3">
              <ListOrdered className="text-amber-500" size={24} />
              <h2 className="text-xl font-black text-slate-800">مخالفات الطلاب (اليوم)</h2>
            </div>
            <div className="flex bg-white rounded-xl p-1 shadow-sm border border-slate-200">
              <button onClick={() => setActiveTab('absent')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition ${activeTab === 'absent' ? 'bg-rose-100 text-rose-700' : 'text-slate-500 hover:bg-slate-50'}`}>🔴 غياب ({allAbsent.length})</button>
              <button onClick={() => setActiveTab('late')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition ${activeTab === 'late' ? 'bg-amber-100 text-amber-700' : 'text-slate-500 hover:bg-slate-50'}`}>🟠 تأخير ({allLate.length})</button>
              <button onClick={() => setActiveTab('truant')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition ${activeTab === 'truant' ? 'bg-purple-100 text-purple-700' : 'text-slate-500 hover:bg-slate-50'}`}>🟣 تسرب ({allTruant.length})</button>
            </div>
          </div>
          
          <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] border border-white p-6 shadow-md overflow-hidden max-h-[500px] overflow-y-auto custom-scrollbar">
            {activeListData.length === 0 ? (
              <div className="py-10 text-center text-slate-500 font-bold">لا يوجد حالات مسجلة.</div>
            ) : (
              <table className="w-full text-right text-sm">
                <thead>
                  <tr className="bg-slate-100 text-slate-600">
                    <th className="p-3 rounded-r-xl">الطالب</th>
                    <th className="p-3">المعلم</th>
                    <th className="p-3">الفصل</th>
                    <th className="p-3 rounded-l-xl">الوقت</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {activeListData.map((item, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="p-3 font-black text-indigo-900">{item.name}</td>
                      <td className="p-3 font-bold text-slate-600">{item.teacher}</td>
                      <td className="p-3"><span className="bg-indigo-50 text-indigo-600 px-2 py-1 rounded text-xs font-bold">{item.className}</span></td>
                      <td className="p-3 text-slate-400 font-mono">{item.time}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// =========================================================
// 2️⃣ صفحة أرشيف التقارير
// =========================================================
function ReportsPage({ data, schoolName }: { data: any, schoolName: string }) {
  const [selectedDateStr, setSelectedDateStr] = useState(new Date().toISOString().split('T')[0]);
  
  const reportData = useMemo(() => {
    const targetDate = new Date(selectedDateStr);
    const targetArabicDay = getArabicDayName(targetDate);
    const expected = data.teachers.filter((t: any) => t.day === targetArabicDay);
    const logsForDate = data.logs.filter((log: any) => isSameDate(log.time, targetDate));

    const uniqueLogsForDate = new Map();
    const chronologicalLogs = [...logsForDate].reverse();

    chronologicalLogs.forEach((log: any) => {
      if (!uniqueLogsForDate.has(log.teacherName)) {
        uniqueLogsForDate.set(log.teacherName, { ...log });
      } else {
        const existing = uniqueLogsForDate.get(log.teacherName);
        const mergeStrings = (s1: string, s2: string) => {
             const arr1 = s1 && s1 !== "لا يوجد" && s1 !== "حضور كامل" ? s1.split("، ") : [];
             const arr2 = s2 && s2 !== "لا يوجد" && s2 !== "حضور كامل" ? s2.split("، ") : [];
             const combined = Array.from(new Set([...arr1, ...arr2]));
             return combined.length > 0 ? combined.join("، ") : "لا يوجد";
        };
        existing.absentStudents = mergeStrings(existing.absentStudents, log.absentStudents);
        existing.lateStudents = mergeStrings(existing.lateStudents, log.lateStudents);
        existing.truantStudents = mergeStrings(existing.truantStudents, log.truantStudents);
        existing.time = log.time;
      }
    });

    return expected.map((teacher: any) => {
      const teacherLog = uniqueLogsForDate.get(teacher.name);
      return {
        name: teacher.name,
        className: teacher.className,
        status: teacherLog ? 'مكتمل' : 'متأخر',
        absents: teacherLog?.absentStudents || 'لا يوجد',
        lates: teacherLog?.lateStudents || 'لا يوجد',
        truants: teacherLog?.truantStudents || 'لا يوجد',
        time: teacherLog ? new Date(teacherLog.time).toLocaleTimeString('ar-SA', {hour: '2-digit', minute:'2-digit'}) : '-'
      };
    });
  }, [data, selectedDateStr]);

  // 💉 زر الطباعة للأرشيف
  const handlePrintArchive = () => {
    let combinedData: any[] = [];
    reportData.forEach((row: any) => {
      if (row.status === 'مكتمل') {
        const extract = (str: string, type: string) => {
          if (str !== 'لا يوجد' && str !== 'حضور كامل') {
            str.split('، ').forEach(name => {
              if(name.trim()) combinedData.push({ name: name.trim(), type: type, teacher: row.name, className: row.className, time: row.time });
            });
          }
        };
        extract(row.absents, 'غياب');
        extract(row.lates, 'تأخير');
        extract(row.truants, 'تسرب');
      }
    });
    
    const dateStr = new Date(selectedDateStr).toLocaleDateString('ar-OM');
    generateOfficialReport('التقرير الأرشيفي الشامل لغياب ومخالفات الطلاب', combinedData, schoolName, dateStr, 'students');
  };

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-2">
        <div className="flex items-center gap-3">
          <FileText className="text-indigo-500" size={28} />
          <h2 className="text-2xl font-black text-slate-800">أرشيف التقارير</h2>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 bg-white px-4 py-2 rounded-xl shadow-sm border border-slate-200">
            <CalendarIcon size={20} className="text-indigo-500" />
            <input type="date" value={selectedDateStr} onChange={(e) => setSelectedDateStr(e.target.value)} className="bg-transparent font-bold text-slate-700 outline-none" />
          </div>
          <button onClick={handlePrintArchive} className="flex items-center gap-2 bg-indigo-600 text-white px-4 py-2 rounded-xl font-bold hover:bg-indigo-700 transition shadow-sm active:scale-95">
            <Printer size={18} /> استخراج PDF رسمي
          </button>
        </div>
      </div>

      <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] p-6 border border-white shadow-lg flex-1 overflow-hidden flex flex-col">
        <div className="overflow-x-auto flex-1 custom-scrollbar">
          <table className="w-full text-right border-collapse min-w-[800px]">
            <thead className="bg-slate-100/80 border-b border-slate-200 sticky top-0 z-10">
              <tr>
                <th className="p-4 font-black text-slate-600">المعلم</th>
                <th className="p-4 font-black text-slate-600">الفصل</th>
                <th className="p-4 font-black text-slate-600 text-center">الرصد</th>
                <th className="p-4 font-black text-slate-600">تفاصيل المخالفات</th>
                <th className="p-4 font-black text-slate-600">الوقت</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {reportData.length === 0 ? (
                <tr><td colSpan={5} className="py-10 text-center font-bold text-slate-400">لا يوجد معلمين مسندين لهذا اليوم.</td></tr>
              ) : (
                reportData.map((row: any, idx: number) => (
                  <tr key={idx} className="hover:bg-slate-50 transition-colors">
                    <td className="p-4 font-black text-indigo-900">{row.name}</td>
                    <td className="p-4 font-bold text-slate-600">{row.className}</td>
                    <td className="p-4 text-center">
                      <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-xs font-bold ${row.status === 'مكتمل' ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>
                        {row.status === 'مكتمل' ? <CheckCircle2 size={14}/> : <Clock size={14}/>} {row.status}
                      </span>
                    </td>
                    <td className="p-4 text-sm font-bold text-slate-600">
                      <div className="flex flex-col gap-1 items-start">
                        {row.absents !== 'لا يوجد' && row.absents !== 'حضور كامل' && <span className="bg-rose-50 text-rose-700 px-2 py-1 rounded-md text-xs border border-rose-100 w-fit">🔴 غياب: {row.absents}</span>}
                        {row.lates !== 'لا يوجد' && <span className="bg-amber-50 text-amber-700 px-2 py-1 rounded-md text-xs border border-amber-100 w-fit">🟠 تأخير: {row.lates}</span>}
                        {row.truants !== 'لا يوجد' && <span className="bg-purple-50 text-purple-700 px-2 py-1 rounded-md text-xs border border-purple-100 w-fit">🟣 تسرب: {row.truants}</span>}
                        {((row.absents === 'لا يوجد' || row.absents === 'حضور كامل') && row.lates === 'لا يوجد' && row.truants === 'لا يوجد' && row.status === 'مكتمل') && <span className="text-emerald-500 font-bold bg-emerald-50 px-2 py-1 rounded-md text-xs">✔️ لا توجد مخالفات</span>}
                      </div>
                    </td>
                    <td className="p-4 text-slate-500 text-sm font-mono">{row.time}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// === المكونات المساعدة للواجهة ===
function SearchPage({ data }: { data: any }) {
  const [searchQuery, setSearchQuery] = useState('');

  const results = useMemo(() => {
    if (!searchQuery.trim()) return [];
    const term = searchQuery.trim();
    let found: any[] = [];

    data.logs.forEach((log: any) => {
      const check = (listStr: string, label: string, color: string) => {
        if (listStr && listStr !== 'لا يوجد' && listStr !== 'حضور كامل') {
          listStr.split('، ').forEach(name => {
            if (name.includes(term)) {
              found.push({ name, label, color, teacher: log.teacherName, className: log.className, time: new Date(log.time).toLocaleTimeString('ar-OM', { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute:'2-digit' }) });
            }
          });
        }
      };
      check(log.absentStudents, '🔴 غياب', 'bg-rose-50 text-rose-700 border-rose-200');
      check(log.lateStudents, '🟠 تأخير', 'bg-amber-50 text-amber-700 border-amber-200');
      check(log.truantStudents, '🟣 تسرب', 'bg-purple-50 text-purple-700 border-purple-200');
    });
    return found.reverse();
  }, [searchQuery, data.logs]);

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col">
      <div className="flex items-center gap-3 px-2">
        <Search className="text-indigo-500" size={28} />
        <h2 className="text-2xl font-black text-slate-800">البحث المتقدم عن الطلاب</h2>
      </div>
      <div className="bg-white p-4 sm:p-6 rounded-[2rem] shadow-sm border border-slate-200 flex items-center gap-4 transition-all focus-within:ring-4 ring-indigo-500/10 focus-within:border-indigo-300">
        <Search className="text-indigo-400" size={28} />
        <input type="text" value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="اكتب اسم الطالب للبحث عن المخالفات (غياب، تأخير، تسرب)..." className="flex-1 bg-transparent text-lg font-bold text-slate-800 outline-none placeholder:text-slate-400" />
      </div>
      <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] p-6 border border-white shadow-lg flex-1 overflow-hidden flex flex-col">
        <div className="overflow-x-auto flex-1 custom-scrollbar">
          {searchQuery.trim() === '' ? (
            <div className="h-full flex flex-col items-center justify-center text-slate-400 opacity-60">
              <div className="w-24 h-24 bg-slate-100 rounded-full flex items-center justify-center mb-4"><Search size={40} className="text-slate-400" /></div>
              <p className="text-xl font-bold">ابدأ بكتابة اسم الطالب لتتبع سجله</p>
            </div>
          ) : results.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center text-emerald-500">
              <div className="w-24 h-24 bg-emerald-50 rounded-full flex items-center justify-center mb-4"><CheckCircle2 size={48} className="text-emerald-500" /></div>
              <p className="text-2xl font-black mb-2">سجل الطالب نظيف!</p>
              <p className="font-bold text-emerald-600/80">لم يتم رصد أي غياب أو تأخير أو تسرب بهذا الاسم.</p>
            </div>
          ) : (
            <table className="w-full text-right border-collapse min-w-[800px]">
              <thead className="bg-slate-100/80 border-b border-slate-200 sticky top-0 z-10">
                <tr>
                  <th className="p-4 font-black text-slate-600 rounded-tr-xl">اسم الطالب</th>
                  <th className="p-4 font-black text-slate-600">نوع المخالفة</th>
                  <th className="p-4 font-black text-slate-600">المعلم / الحصة</th>
                  <th className="p-4 font-black text-slate-600">الفصل</th>
                  <th className="p-4 font-black text-slate-600 rounded-tl-xl">التاريخ والوقت</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {results.map((row, idx) => (
                  <tr key={idx} className="hover:bg-slate-50 transition-colors">
                    <td className="p-4 font-black text-indigo-900">{row.name}</td>
                    <td className="p-4"><span className={`px-3 py-1.5 rounded-lg text-xs font-bold border ${row.color}`}>{row.label}</span></td>
                    <td className="p-4 font-bold text-slate-600">{row.teacher}</td>
                    <td className="p-4"><span className="bg-indigo-50 text-indigo-600 px-3 py-1 rounded-md text-xs font-bold">{row.className}</span></td>
                    <td className="p-4 text-slate-500 font-mono text-sm">{row.time}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

function SettingsPage({ schoolCode, schoolName, setSchoolName }: any) {
  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{ type: 'idle' | 'success' | 'error', msg: string }>({ type: 'idle', msg: '' });

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    setUploadStatus({ type: 'idle', msg: 'جاري قراءة الملف...' });
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);
      const teachersList = jsonData.map(row => ({
        name: row['اسم المعلم'] || row['الاسم'] || row['المعلم'] || '',
        className: row['الفصل'] || row['الفصل المسند'] || row['الصفوف'] || '',
        day: row['اليوم'] || ''
      })).filter(t => t.name && t.day); 
      if (teachersList.length === 0) throw new Error("لم يتم العثور على بيانات صحيحة.");
      setUploadStatus({ type: 'idle', msg: `تم استخراج ${teachersList.length} معلم، جاري الإرسال للسحابة...` });
      const response = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: "bulk_upload_teachers", schoolCode: schoolCode, teachers: teachersList }) });
      const result = await response.json();
      if (result.status === "success") setUploadStatus({ type: 'success', msg: result.message });
      else throw new Error(result.message);
    } catch (error: any) {
      setUploadStatus({ type: 'error', msg: error.message || "حدث خطأ أثناء رفع الملف." });
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col">
      <div className="flex items-center gap-3 px-2">
        <Settings className="text-indigo-500" size={28} />
        <h2 className="text-2xl font-black text-slate-800">الإعدادات</h2>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 sm:p-8 border border-white/80 shadow-[0_8px_30px_rgb(0,0,0,0.04)]">
          <h3 className="text-xl font-bold text-slate-800 mb-6 flex items-center gap-2"><School className="text-amber-500" size={24} /> معلومات المدرسة</h3>
          <div className="space-y-5">
            <div className="space-y-2">
              <label className="block text-sm font-bold text-slate-700 ml-2">اسم المدرسة</label>
              <input type="text" value={schoolName} onChange={(e) => setSchoolName(e.target.value)} placeholder="مثال: مدرسة الإمام الشافعي (9-12)" className="w-full px-4 py-3 rounded-2xl bg-white border border-indigo-100 text-indigo-900 font-bold outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-400/20 transition-all" />
            </div>
            <div className="space-y-2">
              <label className="block text-sm font-bold text-slate-700 ml-2">كود المدرسة المسجل حالياً</label>
              <div className="w-full px-4 py-3 rounded-2xl bg-indigo-50 border border-indigo-100 text-indigo-900 font-black text-xl text-center tracking-[0.3em] font-mono" dir="ltr">{schoolCode}</div>
            </div>
          </div>
        </div>
        <div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 sm:p-8 border border-white/80 shadow-[0_8px_30px_rgb(0,0,0,0.04)]">
          <h3 className="text-xl font-bold text-slate-800 mb-4 flex items-center gap-2"><Users className="text-blue-500" size={24} /> إدارة المعلمين المرجعية</h3>
          <p className="text-slate-500 font-bold text-sm mb-6 leading-relaxed">
            قم برفع ملف Excel (.xlsx) يحتوي على 3 أعمدة رئيسية: <br/>
            <span className="text-indigo-600 bg-indigo-50 px-2 py-1 rounded mx-1">اسم المعلم</span><span className="text-indigo-600 bg-indigo-50 px-2 py-1 rounded mx-1">الفصل</span><span className="text-indigo-600 bg-indigo-50 px-2 py-1 rounded mx-1">اليوم</span>
          </p>
          <div className="relative border-2 border-dashed border-indigo-200 rounded-2xl p-8 flex flex-col items-center justify-center text-center bg-indigo-50/30 hover:bg-indigo-50/80 transition-colors group">
            {isUploading ? (
              <div className="flex flex-col items-center"><Loader2 size={40} className="text-indigo-500 animate-spin mb-4" /><h4 className="font-bold text-slate-800">{uploadStatus.msg}</h4></div>
            ) : (
              <>
                <div className="bg-white p-4 rounded-full shadow-sm mb-4 group-hover:scale-110 transition-transform duration-300"><UploadCloud size={32} className="text-indigo-500" /></div>
                <h4 className="font-bold text-slate-800 mb-1">اضغط هنا لاختيار ملف الإكسل</h4>
                <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" disabled={isUploading} />
              </>
            )}
          </div>
          {uploadStatus.type === 'success' && <div className="mt-4 p-4 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-xl font-bold flex items-center gap-2"><CheckCircle2 size={20} />{uploadStatus.msg}</div>}
          {uploadStatus.type === 'error' && <div className="mt-4 p-4 bg-rose-50 border border-rose-200 text-rose-700 rounded-xl font-bold flex items-center gap-2"><AlertTriangle size={20} />{uploadStatus.msg}</div>}
        </div>
      </div>
    </div>
  );
}

function StatCard({ title, value, subtitle, icon: Icon, color }: any) {
  const colors:any = { indigo: 'from-indigo-500 to-blue-600 bg-indigo-50 text-indigo-600', amber: 'from-amber-400 to-orange-500 bg-amber-50 text-amber-600', rose: 'from-rose-400 to-red-500 bg-rose-50 text-rose-600' };
  return (
    <div className="bg-white rounded-[2rem] p-6 shadow-xl border border-white flex justify-between items-center group overflow-hidden relative">
      <div className={`absolute inset-y-0 right-0 w-2 bg-gradient-to-b ${colors[color].split(' ').slice(0,2).join(' ')}`} />
      <div>
        <p className="text-sm font-bold text-slate-400 mb-1">{title}</p>
        <h3 className="text-4xl font-black text-slate-800">{value} <span className="text-sm text-slate-300">{subtitle}</span></h3>
      </div>
      <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${colors[color].split(' ').slice(2).join(' ')} group-hover:scale-110 transition-transform`}><Icon size={28} /></div>
    </div>
  );
}

function Sidebar({ activeSection, setActiveSection }: any) {
  const navItems = [
    { id: 'dashboard', label: 'اللوحة الرئيسية', icon: LayoutDashboard },
    { id: 'reports', label: 'أرشيف التقارير', icon: FileText },
    { id: 'search', label: 'بحث الطلاب', icon: Search },
    { id: 'settings', label: 'الإعدادات', icon: Settings },
  ];
  return (
    <aside className="w-20 lg:w-72 bg-indigo-950 text-indigo-100 flex flex-col z-20 shadow-2xl transition-all duration-300">
      <div className="h-24 flex items-center px-8 border-b border-white/10"><School className="text-amber-400 ml-3" /><span className="hidden lg:block font-black text-xl text-white">راصد الإدارة</span></div>
      <nav className="flex-1 py-8 px-4 space-y-2">
        {navItems.map((item) => (
          <button key={item.id} onClick={() => setActiveSection(item.id)} className={`w-full flex items-center gap-4 px-4 py-3.5 rounded-2xl transition-all ${activeSection === item.id ? 'bg-indigo-600 text-white shadow-lg' : 'hover:bg-indigo-800/50 text-indigo-300'}`}>
            <item.icon size={22} className={activeSection === item.id ? 'text-amber-400' : ''} />
            <span className="hidden lg:block font-bold">{item.label}</span>
          </button>
        ))}
      </nav>
    </aside>
  );
}

function TopHeader({ onLogout, schoolCode, schoolName }: any) {
  return (
    <header className="h-24 px-10 flex items-center justify-between bg-white/40 backdrop-blur-xl border-b border-white/60 sticky top-0 z-30">
      <div>
        <h2 className="text-2xl font-black text-indigo-900">{schoolName ? schoolName : 'مرحباً بك، مدير المدرسة'}</h2>
        <p className="text-sm font-medium text-slate-500">كود المدرسة: {schoolCode}</p>
      </div>
      <button onClick={onLogout} className="flex items-center gap-2 px-5 py-3 rounded-2xl bg-rose-50 text-rose-600 font-bold hover:bg-rose-100 transition-all border border-rose-100"><LogOut size={18} /><span>خروج</span></button>
    </header>
  );
}

function LoginScreen({ schoolCode, setSchoolCode, onLogin, isLoading, errorMsg }: any) {
  return (
    <div className="min-h-screen flex items-center justify-center relative bg-slate-50" dir="rtl">
      <div className="absolute inset-0 bg-gradient-to-br from-indigo-900 via-slate-900 to-indigo-950" />
      <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} className="relative z-10 w-full max-w-lg p-12 bg-white/10 backdrop-blur-2xl rounded-[2.5rem] border border-white/20 shadow-2xl">
        <div className="flex flex-col items-center mb-10">
          <div className="bg-gradient-to-br from-amber-400 to-amber-600 p-4 rounded-3xl mb-6"><School size={48} className="text-white" /></div>
          <h1 className="text-4xl font-black text-white mb-2">راصد الإدارة</h1>
          <p className="text-indigo-200">نظام المتابعة الذكي للقيادة المدرسية</p>
        </div>
        <form onSubmit={onLogin} className="space-y-6">
          <input type="text" value={schoolCode} onChange={(e) => setSchoolCode(e.target.value)} className="w-full px-6 py-4 rounded-2xl bg-white/10 border border-white/20 text-white text-center text-3xl font-mono outline-none focus:border-amber-400" placeholder="••••" dir="ltr" />
          {errorMsg && <p className="text-rose-400 text-center font-bold">{errorMsg}</p>}
          <button disabled={isLoading} className="w-full bg-amber-500 hover:bg-amber-400 text-indigo-950 font-black py-4 rounded-2xl transition-all">{isLoading ? 'جاري الاتصال...' : 'تسجيل الدخول'}</button>
        </form>
      </motion.div>
    </div>
  );
}