import React, { useEffect, useMemo, useState } from 'react';
import {
  ArrowLeft,
  School,
  LogOut,
  Users,
  CheckCircle2,
  Clock,
  TrendingUp,
  AlertTriangle,
  LayoutDashboard,
  FileText,
  Search,
  Settings,
  Calendar as CalendarIcon,
  ListOrdered,
  Loader2,
  UploadCloud,
  Printer,
  ShieldCheck,
  RefreshCw,
  Wifi,
  Database,
  Menu,
  UserCheck,
  FileSpreadsheet,
  Code
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { Capacitor } from '@capacitor/core';
import { Filesystem, Directory } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';

import { AdminProvider, useAdmin } from '../context/AdminContext';

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwZHhZ-RPWUpBGIlw0qTFPUmOPmq9WpcvW4WLklcjb_A9U3MW0luIXYPnHznI29ThpbMA/exec';

type Section = 'dashboard' | 'substitutions' | 'reports' | 'search' | 'settings';
type TabKey = 'absent' | 'late' | 'truant';

type TeacherRow = {
  name: string;
  className: string;
  day: string;
  period: string;
};

type LogRow = {
  teacherName: string;
  className: string;
  absentStudents: string;
  time: string;
  lateStudents: string;
  truantStudents: string;
  period: string;
};

type DashboardData = {
  teachers: TeacherRow[];
  logs: LogRow[];
};

const ARABIC_DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
const getArabicDayName = (date: Date) => ARABIC_DAYS[date.getDay()];

const normalizeText = (text: string) => {
  if (!text) return '';
  return String(text).replace(/[أإآ]/g, 'ا').replace(/\s+/g, ' ').trim();
};

const formatDate = (date: Date) => date.toLocaleDateString('ar-OM', { year: 'numeric', month: 'long', day: 'numeric' });

const isSameDate = (dateString: string, compareDate: Date) => {
  if (!dateString) return false;
  const d = new Date(dateString);
  if (!Number.isNaN(d.getTime())) {
    return d.getFullYear() === compareDate.getFullYear()
      && d.getMonth() === compareDate.getMonth()
      && d.getDate() === compareDate.getDate();
  }

  const day = compareDate.getDate().toString();
  const month = (compareDate.getMonth() + 1).toString();
  const arDay = day.replace(/\d/g, (x: any) => '٠١٢٣٤٥٦٧٨٩'[x]);
  const arMonth = month.replace(/\d/g, (x: any) => '٠١٢٣٤٥٦٧٨٩'[x]);
  const str = String(dateString);
  return (str.includes(day) || str.includes(arDay)) && (str.includes(month) || str.includes(arMonth));
};

const splitList = (value: string) => {
  if (!value || value === 'لا يوجد' || value === 'حضور كامل') return [];
  return String(value).split('،').map(item => item.trim()).filter(Boolean);
};

const uniqueBy = <T,>(items: T[], keyGetter: (item: T) => string) => {
  const map = new Map<string, T>();
  items.forEach(item => {
    const key = keyGetter(item);
    if (!map.has(key)) map.set(key, item);
  });
  return Array.from(map.values());
};

const sanitizeFileName = (name: string) =>
  String(name || 'report')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '_')
    .slice(0, 80);

const dataUrlToBase64 = (dataUrl: string) => {
  const parts = dataUrl.split(',');
  return parts.length > 1 ? parts[1] : dataUrl;
};

const extractReportBody = (htmlContent: string) => {
  const bodyMatch = htmlContent.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return bodyMatch ? bodyMatch[1] : htmlContent;
};

const extractReportStyles = (htmlContent: string) => {
  const styleMatches = Array.from(htmlContent.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi));
  return styleMatches.map(match => match[1]).join('\n');
};

const createPdfFromHtml = async (
  htmlContent: string,
  title: string,
  orientation: 'portrait' | 'landscape' = 'portrait'
) => {
  const container = document.createElement('div');
  container.setAttribute('dir', 'rtl');
  container.style.position = 'fixed';
  container.style.left = '-10000px';
  container.style.top = '0';
  container.style.width = orientation === 'landscape' ? '1123px' : '794px';
  container.style.background = '#ffffff';
  container.style.zIndex = '-1';
  container.style.padding = orientation === 'landscape' ? '38px' : '56px';
  container.style.boxSizing = 'border-box';
  container.innerHTML = `<style>${extractReportStyles(htmlContent)}</style>${extractReportBody(htmlContent)}`;
  document.body.appendChild(container);

  try {
    await new Promise(resolve => setTimeout(resolve, 350));

    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const pdf = new jsPDF({ orientation, unit: 'mm', format: 'a4', compress: true });
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imgWidth = pageWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;
    const imgData = canvas.toDataURL('image/jpeg', 0.95);

    let heightLeft = imgHeight;
    let position = 0;

    pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
    heightLeft -= pageHeight;

    while (heightLeft > 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;
    }

    return pdf;
  } finally {
    document.body.removeChild(container);
  }
};

const outputOfficialPdf = async (
  htmlContent: string,
  title: string,
  orientation: 'portrait' | 'landscape' = 'portrait'
) => {
  const pdf = await createPdfFromHtml(htmlContent, title, orientation);
  const fileName = `${sanitizeFileName(title)}_${Date.now()}.pdf`;

  // Android / Capacitor: حفظ PDF داخل مستندات التطبيق ثم فتح نافذة المشاركة
  if (Capacitor.isNativePlatform()) {
    const dataUri = pdf.output('datauristring');
    const base64Data = dataUrlToBase64(dataUri);

    // حفظ نسخة دائمة في Documents
    const savedFile = await Filesystem.writeFile({
      path: `RasedAdminReports/${fileName}`,
      data: base64Data,
      directory: Directory.Documents,
      recursive: true
    });

    alert(`تم حفظ التقرير بنجاح في الجهاز:\n${fileName}`);

    // فتح نافذة المشاركة بعد الحفظ
    await Share.share({
      title,
      text: 'تقرير رسمي صادر من راصد الإدارة بصيغة PDF.',
      url: savedFile.uri,
      dialogTitle: 'فتح / مشاركة / طباعة تقرير راصد الإدارة PDF'
    });

    return;
  }

  // الويب: تحميل مباشر
  pdf.save(fileName);
};

const generateOfficialReport = async (
  title: string,
  dataList: any[],
  schoolName: string,
  dateStr: string,
  type: 'students' | 'teachers' = 'students'
) => {
  const tableHeader = type === 'students'
    ? `<tr><th style="width: 5%;">م</th><th style="width: 30%;">اسم الطالب</th><th style="width: 15%;">نوع المخالفة</th><th style="width: 10%;">الفصل</th><th style="width: 25%;">المعلم الراصد</th><th style="width: 15%;">الوقت</th></tr>`
    : `<tr><th style="width: 10%;">م</th><th style="width: 45%;">اسم المعلم</th><th style="width: 25%;">الفصل المسند</th><th style="width: 20%;">حالة الرصد</th></tr>`;

  let tableContent = '';
  if (dataList.length === 0) {
    tableContent = `<tr><td colspan="${type === 'students' ? 6 : 4}" style="text-align:center; padding: 20px; font-weight: bold;">لا يوجد سجلات في هذا التقرير.</td></tr>`;
  } else {
    dataList.forEach((item, index) => {
      tableContent += type === 'students'
        ? `<tr>
             <td>${index + 1}</td>
             <td style="font-weight: bold;">${item.name || ''}</td>
             <td>${item.type || 'غياب / تأخير'}</td>
             <td>${item.className || ''}</td>
             <td>${item.teacher || ''}</td>
             <td dir="ltr">${item.time || ''}</td>
           </tr>`
        : `<tr>
             <td>${index + 1}</td>
             <td style="font-weight: bold; text-align: right; padding-right: 20px;">${item.name || ''}</td>
             <td>${item.className || ''}</td>
             <td style="color: red; font-weight: bold;">لم يتم الرصد</td>
           </tr>`;
    });
  }

  const htmlContent = `
    <html dir="rtl" lang="ar">
    <head>
      <title>تقرير نظام راصد - ${title}</title>
      <style>
        body { font-family: Tahoma, Arial, sans-serif; color: #000; margin: 0; padding: 0; -webkit-print-color-adjust: exact; background: #fff; }
        .report-page { width: 100%; box-sizing: border-box; }
        .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 3px double #000; padding-bottom: 15px; margin-bottom: 30px; }
        .right-header { text-align: right; font-weight: 900; line-height: 1.6; font-size: 13pt; }
        .left-header { text-align: left; font-weight: 700; line-height: 1.6; font-size: 12pt; }
        .title { text-align: center; font-size: 18pt; font-weight: 900; margin: 30px 0; text-decoration: underline; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 40px; font-size: 12pt; }
        th, td { border: 1px solid #000; padding: 10px 8px; text-align: center; vertical-align: middle; }
        th { background-color: #f1f5f9; font-weight: 900; font-size: 13pt; }
        .signatures { display: flex; justify-content: space-between; margin-top: 70px; font-weight: 900; font-size: 14pt; padding: 0 20px; }
        .footer-note { text-align: center; font-size: 9pt; font-weight: bold; color: #64748b; margin-top: 50px; border-top: 1px solid #cbd5e1; padding-top: 10px; }
      </style>
    </head>
    <body>
      <div class="report-page">
        <div class="header">
            <div class="right-header">سلطنة عُمان<br>وزارة التعليم<br>المديرية العامة للتعليم بمحافظة شمال الباطنة</div>
            <div class="left-header">المدرسة: ${schoolName || '____________________'}<br>التاريخ: ${dateStr}<br>النظام: راصد الإدارة</div>
        </div>
        <div class="title">${title}</div>
        <table><thead>${tableHeader}</thead><tbody>${tableContent}</tbody></table>
        <div class="signatures">
          ${type === 'students' ? '<div>مُعد التقرير: ........................</div>' : '<div>الختم الرسمي:</div>'}
          <div>يعتمد، مدير المدرسة: ........................</div>
        </div>
        <div class="footer-note">تم استخراج هذا التقرير إلكترونياً من نظام راصد - برمجة وتطوير: ALZAABI MOHAMMAD</div>
      </div>
    </body>
    </html>`;

  await outputOfficialPdf(htmlContent, title, 'portrait');
};

export default function App() {
  return (
    <AdminProvider>
      <AdminDashboardCore />
    </AdminProvider>
  );
}

function AdminDashboardCore() {
  const { dashboardData, setDashboardData, isDataLoaded, lastSavedAt } = useAdmin() as any;
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [schoolCode, setSchoolCode] = useState(() => localStorage.getItem('rased_admin_code') || '');
  const [schoolName, setSchoolName] = useState(() => localStorage.getItem('rased_school_name') || '');
  const [activeSection, setActiveSection] = useState<Section>('dashboard');
  const [isLoading, setIsLoading] = useState(false);
  const [isRefreshingMain, setIsRefreshingMain] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [lastCloudSyncAt, setLastCloudSyncAt] = useState('');

  const [subsData, setSubsData] = useState<any[]>([]);
  const [isSubsLoaded, setIsSubsLoaded] = useState(false);
  const [isSubsLoading, setIsSubsLoading] = useState(false);

  useEffect(() => {
    localStorage.setItem('rased_school_name', schoolName);
  }, [schoolName]);

  const fetchDashboardData = async (code = schoolCode, silent = false) => {
    if (!code.trim()) return;
    if (silent) setIsRefreshingMain(true);
    else setIsLoading(true);
    setErrorMsg('');

    try {
      const cacheBuster = Date.now();
      const response = await fetch(`${SCRIPT_URL}?schoolCode=${encodeURIComponent(code.trim())}&t=${cacheBuster}`);
      const result = await response.json();
      if (result.status === 'success') {
        setDashboardData({ teachers: result.data?.teachers || [], logs: result.data?.logs || [] });
        setIsLoggedIn(true);
        setLastCloudSyncAt(new Date().toISOString());
        localStorage.setItem('rased_admin_code', code.trim());
      } else {
        setErrorMsg('حدث خطأ من السيرفر: ' + (result.message || 'غير محدد'));
      }
    } catch (error) {
      if (dashboardData?.logs?.length > 0 || dashboardData?.teachers?.length > 0) {
        setIsLoggedIn(true);
        alert('⚠️ تعذر الاتصال بالسحابة. يتم عرض النسخة المؤرشفة محليًا.');
      } else {
        setErrorMsg('تأكد من اتصالك بالإنترنت ومن صحة رابط السحابة');
      }
    } finally {
      setIsLoading(false);
      setIsRefreshingMain(false);
    }
  };

  const fetchSubstitutions = async (forceRefresh = false) => {
    if (isSubsLoaded && !forceRefresh) return;
    setIsSubsLoading(true);
    try {
      const res = await fetch(`${SCRIPT_URL}?action=getSubstitutions&role=admin&schoolCode=${encodeURIComponent(schoolCode)}`);
      const result = await res.json();
      if (result.status === 'success') {
        setSubsData(result.data || []);
        setIsSubsLoaded(true);
      } else {
        alert(result.message || 'تعذر جلب بيانات الاحتياط.');
      }
    } catch (e) {
      alert('خطأ في الاتصال بالسحابة لاستدعاء الاحتياط.');
    } finally {
      setIsSubsLoading(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (schoolCode.trim().length < 2) return;
    await fetchDashboardData(schoolCode, false);
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    setActiveSection('dashboard');
    setIsSubsLoaded(false);
    setIsSidebarOpen(false);
  };

  if (!isDataLoaded) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50" dir="rtl">
        <div className="flex flex-col items-center gap-4">
          <Loader2 className="w-10 h-10 animate-spin text-indigo-600" />
          <p className="text-sm font-black text-slate-500">جاري تحميل أرشيف راصد الإدارة...</p>
        </div>
      </div>
    );
  }

  if (!isLoggedIn) {
    return <LoginScreen schoolCode={schoolCode} setSchoolCode={setSchoolCode} onLogin={handleLogin} isLoading={isLoading} errorMsg={errorMsg} />;
  }

  return (
    <div className="flex h-screen overflow-hidden bg-[#f8fafc] text-slate-800 font-sans" dir="rtl">
      <div className="fixed inset-0 z-0 pointer-events-none overflow-hidden">
        <div className="absolute top-[-10%] right-[-5%] w-[40vw] h-[40vw] rounded-full bg-indigo-100/50 blur-[100px]" />
        <div className="absolute bottom-[-10%] left-[-5%] w-[30vw] h-[30vw] rounded-full bg-amber-100/40 blur-[80px]" />
      </div>

      <MobileSidebarOverlay isOpen={isSidebarOpen} onClose={() => setIsSidebarOpen(false)} />
      <Sidebar activeSection={activeSection} setActiveSection={(section: Section) => { setActiveSection(section); setIsSidebarOpen(false); }} isOpen={isSidebarOpen} />

      <div className="flex-1 flex flex-col h-full relative z-10 overflow-hidden">
        <TopHeader
          onLogout={handleLogout}
          onMenu={() => setIsSidebarOpen(true)}
          onRefresh={() => fetchDashboardData(schoolCode, true)}
          isRefreshing={isRefreshingMain}
          schoolCode={schoolCode}
          schoolName={schoolName}
          lastCloudSyncAt={lastCloudSyncAt}
          lastSavedAt={lastSavedAt}
        />

        <main className="flex-1 overflow-y-auto p-4 sm:p-8">
          <AnimatePresence mode="wait">
            <motion.div key={activeSection} initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.3 }} className="max-w-7xl mx-auto h-full">
              {activeSection === 'dashboard' && <DashboardHome data={dashboardData} schoolName={schoolName} />}
              {activeSection === 'substitutions' && <SubstitutionsRadar schoolName={schoolName} data={subsData} isLoading={isSubsLoading} onFetch={fetchSubstitutions} />}
              {activeSection === 'reports' && <ReportsPage data={dashboardData} schoolName={schoolName} />}
              {activeSection === 'search' && <SearchPage data={dashboardData} />}
              {activeSection === 'settings' && <SettingsPage schoolCode={schoolCode} schoolName={schoolName} setSchoolName={setSchoolName} onUploadComplete={() => fetchDashboardData(schoolCode, true)} />}
            </motion.div>
          </AnimatePresence>
        </main>
      </div>
    </div>
  );
}

function MobileSidebarOverlay({ isOpen, onClose }: { isOpen: boolean; onClose: () => void }) {
  if (!isOpen) return null;
  return <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-30 lg:hidden" onClick={onClose} />;
}

function Sidebar({ activeSection, setActiveSection, isOpen }: any) {
  const navItems = [
    { id: 'dashboard', label: 'اللوحة الرئيسية', icon: LayoutDashboard },
    { id: 'substitutions', label: 'رادار الاحتياط', icon: ShieldCheck },
    { id: 'reports', label: 'أرشيف التقارير', icon: FileText },
    { id: 'search', label: 'بحث الطلاب', icon: Search },
    { id: 'settings', label: 'الإعدادات', icon: Settings }
  ];

  return (
    <aside className={`fixed lg:static top-0 bottom-0 right-0 w-72 lg:w-72 bg-indigo-950 text-indigo-100 flex flex-col z-40 shadow-2xl transition-transform duration-300 ${isOpen ? 'translate-x-0' : 'translate-x-full lg:translate-x-0'}`}>
      <div className="h-24 flex items-center px-8 border-b border-white/10">
        <School className="text-amber-400 ml-3" />
        <span className="font-black text-xl text-white">راصد الإدارة</span>
      </div>
      <nav className="flex-1 py-8 px-4 space-y-2">
        {navItems.map((item) => (
          <button key={item.id} onClick={() => setActiveSection(item.id)} className={`w-full flex items-center gap-4 px-4 py-3.5 rounded-2xl transition-all ${activeSection === item.id ? 'bg-indigo-600 text-white shadow-lg' : 'hover:bg-indigo-800/50 text-indigo-300'}`}>
            <item.icon size={22} className={activeSection === item.id ? 'text-amber-400' : ''} />
            <span className="font-bold">{item.label}</span>
          </button>
        ))}
      </nav>
    </aside>
  );
}

function TopHeader({ onLogout, onMenu, onRefresh, isRefreshing, schoolCode, schoolName, lastCloudSyncAt, lastSavedAt }: any) {
  const lastSyncText = lastCloudSyncAt ? new Date(lastCloudSyncAt).toLocaleTimeString('ar-OM', { hour: '2-digit', minute: '2-digit' }) : 'لم تتم مزامنة حديثة';
  const lastSaveText = lastSavedAt ? new Date(lastSavedAt).toLocaleTimeString('ar-OM', { hour: '2-digit', minute: '2-digit' }) : 'غير محدد';

  return (
    <header className="h-24 px-4 sm:px-10 flex items-center justify-between bg-white/40 backdrop-blur-xl border-b border-white/60 sticky top-0 z-30">
      <div className="flex items-center gap-3 min-w-0">
        <button onClick={onMenu} className="lg:hidden p-2.5 rounded-2xl bg-white border border-slate-100 text-indigo-900 shadow-sm">
          <Menu size={22} />
        </button>
        <div className="min-w-0">
          <h2 className="text-lg sm:text-2xl font-black text-indigo-900 truncate">{schoolName ? schoolName : 'مرحباً بك، مدير المدرسة'}</h2>
          <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-[10px] sm:text-xs font-bold text-slate-500 mt-1">
            <span>كود المدرسة: {schoolCode}</span>
            <span className="flex items-center gap-1"><Wifi size={12} /> آخر مزامنة: {lastSyncText}</span>
            <span className="hidden sm:flex items-center gap-1"><Database size={12} /> آخر حفظ: {lastSaveText}</span>
          </div>
        </div>
      </div>
      <div className="flex items-center gap-2">
        <button onClick={onRefresh} className="p-3 rounded-2xl bg-white border border-slate-100 text-indigo-700 shadow-sm hover:bg-indigo-50 active:scale-95 transition-all" title="تحديث البيانات">
          <RefreshCw size={18} className={isRefreshing ? 'animate-spin' : ''} />
        </button>
        <button onClick={onLogout} className="flex items-center gap-2 px-4 sm:px-5 py-3 rounded-2xl bg-rose-50 text-rose-600 font-bold hover:bg-rose-100 transition-all border border-rose-100">
          <LogOut size={18} /><span className="hidden sm:inline">خروج</span>
        </button>
      </div>
    </header>
  );
}

function DashboardHome({ data, schoolName }: { data: DashboardData; schoolName: string }) {
  const [activeTab, setActiveTab] = useState<TabKey>('absent');
  const todayDate = new Date();
  const todayArabicName = getArabicDayName(todayDate);

  const expectedTeachersToday = useMemo(() => {
    const normalizedToday = normalizeText(todayArabicName);
    const todayClasses = (data.teachers || []).filter((t: TeacherRow) => normalizeText(t.day) === normalizedToday);
    const firstPeriodClasses = todayClasses.filter((t: TeacherRow) => {
      const p = String(t.period).trim();
      return p === '1' || p.includes('الأولى') || p === '01';
    });
    return uniqueBy(firstPeriodClasses, (t: TeacherRow) => normalizeText(t.name));
  }, [data.teachers, todayArabicName]);

  const todayLogs = useMemo(() => (data.logs || []).filter((log: LogRow) => isSameDate(log.time, todayDate)), [data.logs, todayDate]);
  const uniqueTodayLogs = useMemo(() => uniqueBy(todayLogs, (log: LogRow) => normalizeText(log.teacherName)), [todayLogs]);

  const lateTeachers = useMemo(() => {
    const loggedTeacherNames = new Set(uniqueTodayLogs.map((l: LogRow) => normalizeText(l.teacherName)));
    return expectedTeachersToday.filter((t: TeacherRow) => !loggedTeacherNames.has(normalizeText(t.name)));
  }, [expectedTeachersToday, uniqueTodayLogs]);

  const extractList = (key: 'absentStudents' | 'lateStudents' | 'truantStudents', label: string) => {
    const list: any[] = [];
    const seen = new Set<string>();
    const chronologicalLogs = [...todayLogs].reverse();
    chronologicalLogs.forEach((log: any) => {
      splitList(log[key]).forEach(studentName => {
        const uniqueKey = `${normalizeText(studentName)}-${normalizeText(log.teacherName)}-${label}`;
        if (!seen.has(uniqueKey)) {
          seen.add(uniqueKey);
          list.push({
            name: studentName,
            type: label,
            teacher: log.teacherName,
            className: log.className,
            time: new Date(log.time).toLocaleTimeString('ar-OM', { hour: '2-digit', minute: '2-digit' })
          });
        }
      });
    });
    return list;
  };

  const allAbsent = useMemo(() => extractList('absentStudents', 'غياب'), [todayLogs]);
  const allLate = useMemo(() => extractList('lateStudents', 'تأخير'), [todayLogs]);
  const allTruant = useMemo(() => extractList('truantStudents', 'تسرب'), [todayLogs]);
  const completionRate = expectedTeachersToday.length === 0 ? 0 : Math.min(100, Math.round((uniqueTodayLogs.length / expectedTeachersToday.length) * 100));
  const activeListData = activeTab === 'absent' ? allAbsent : activeTab === 'late' ? allLate : allTruant;

  const handlePrintDaily = async () => {
    await generateOfficialReport('التقرير اليومي لرصد المخالفات', [...allAbsent, ...allLate, ...allTruant], schoolName, todayDate.toLocaleDateString('ar-OM'), 'students');
  };

  const handlePrintLateTeachers = async () => {
    await generateOfficialReport('كشف المعلمين المتأخرين عن رصد غياب الحصة الأولى', lateTeachers, schoolName, todayDate.toLocaleDateString('ar-OM'), 'teachers');
  };

  return (
    <div className="space-y-8 pb-10">
      <div className="bg-indigo-900 text-white p-4 rounded-2xl shadow-lg flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="font-black text-xl">إحصائيات اليوم المباشرة</h2>
          <p className="text-indigo-200 text-sm">مخصصة لمعلمي الحصة الأولى ليوم ({todayArabicName}) - {formatDate(todayDate)}</p>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          <div className="bg-amber-400 text-indigo-950 font-black px-4 py-2 rounded-xl">{expectedTeachersToday.length} معلم مستهدف</div>
          <button onClick={handlePrintDaily} className="flex items-center gap-2 bg-white text-indigo-900 px-4 py-2 rounded-xl font-bold hover:bg-slate-100 transition shadow-sm active:scale-95">
            <Printer size={18} /> استخراج PDF رسمي
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-6 text-right">
        <StatCard title="نسبة إنجاز الرصد" value={`${completionRate}%`} icon={TrendingUp} color="indigo" />
        <StatCard title="المعلمون الراصدون" value={uniqueTodayLogs.length} subtitle="معلم" icon={UserCheck} color="green" />
        <StatCard title="إجمالي المخالفات" value={allAbsent.length + allLate.length + allTruant.length} subtitle="حالة" icon={Users} color="amber" />
        <StatCard title="معلمون متأخرون" value={lateTeachers.length} subtitle="معلم" icon={Clock} color="rose" />
      </div>

      <div className="grid grid-cols-1 xl:grid-cols-12 gap-8">
        <div className="xl:col-span-4 flex flex-col gap-4">
          <div className="flex items-center justify-between px-2">
            <div className="flex items-center gap-3"><AlertTriangle className="text-rose-500" size={24} /><h2 className="text-xl font-black text-slate-800">تأخر رصد اليوم</h2></div>
            {lateTeachers.length > 0 && <button onClick={handlePrintLateTeachers} className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-lg text-xs font-bold transition-all shadow-sm active:scale-95"><Printer size={14} /> PDF الكشف</button>}
          </div>
          <div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 border border-white/80 shadow-md flex-1 overflow-y-auto max-h-[500px] custom-scrollbar">
            {lateTeachers.length === 0 ? <div className="py-10 text-center text-emerald-600 font-bold"><CheckCircle2 size={40} className="mx-auto mb-3" />اكتمل الرصد!</div> : (
              <div className="space-y-3">{lateTeachers.map((teacher: any, idx: number) => <div key={idx} className="flex items-center gap-3 p-3 rounded-2xl bg-rose-50 border border-rose-100"><div className="w-10 h-10 rounded-xl bg-rose-200 text-rose-600 flex items-center justify-center"><Clock size={18} /></div><div><h3 className="font-bold text-slate-800 text-sm">{teacher.name}</h3><p className="text-xs text-rose-600 font-bold">{teacher.className}</p></div></div>)}</div>
            )}
          </div>
        </div>

        <div className="xl:col-span-8 flex flex-col gap-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between px-2 gap-3">
            <div className="flex items-center gap-3"><ListOrdered className="text-amber-500" size={24} /><h2 className="text-xl font-black text-slate-800">مخالفات الطلاب (اليوم)</h2></div>
            <div className="flex bg-white rounded-xl p-1 shadow-sm border border-slate-200 overflow-x-auto">
              <button onClick={() => setActiveTab('absent')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition whitespace-nowrap ${activeTab === 'absent' ? 'bg-rose-100 text-rose-700' : 'text-slate-500 hover:bg-slate-50'}`}>🔴 غياب ({allAbsent.length})</button>
              <button onClick={() => setActiveTab('late')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition whitespace-nowrap ${activeTab === 'late' ? 'bg-amber-100 text-amber-700' : 'text-slate-500 hover:bg-slate-50'}`}>🟠 تأخير ({allLate.length})</button>
              <button onClick={() => setActiveTab('truant')} className={`px-4 py-1.5 rounded-lg font-bold text-sm transition whitespace-nowrap ${activeTab === 'truant' ? 'bg-purple-100 text-purple-700' : 'text-slate-500 hover:bg-slate-50'}`}>🟣 تسرب ({allTruant.length})</button>
            </div>
          </div>

          <DataTable activeListData={activeListData} />
        </div>
      </div>
    </div>
  );
}

function DataTable({ activeListData }: { activeListData: any[] }) {
  return (
    <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] border border-white p-6 shadow-md overflow-hidden max-h-[500px] overflow-y-auto custom-scrollbar">
      {activeListData.length === 0 ? <div className="py-10 text-center text-slate-500 font-bold">لا يوجد حالات مسجلة.</div> : (
        <table className="w-full text-right text-sm min-w-[650px]">
          <thead><tr className="bg-slate-100 text-slate-600"><th className="p-3 rounded-r-xl">الطالب</th><th className="p-3">المعلم</th><th className="p-3">الفصل</th><th className="p-3 rounded-l-xl">الوقت</th></tr></thead>
          <tbody className="divide-y divide-slate-100">
            {activeListData.map((item, idx) => <tr key={idx} className="hover:bg-slate-50 transition-colors"><td className="p-3 font-black text-indigo-900">{item.name}</td><td className="p-3 font-bold text-slate-600">{item.teacher}</td><td className="p-3"><span className="bg-indigo-50 text-indigo-600 px-2 py-1 rounded text-xs font-bold">{item.className}</span></td><td className="p-3 text-slate-400 font-mono">{item.time}</td></tr>)}
          </tbody>
        </table>
      )}
    </div>
  );
}

function SubstitutionsRadar({ schoolName, data, isLoading, onFetch }: { schoolName: string; data: any[]; isLoading: boolean; onFetch: (force: boolean) => void }) {
  useEffect(() => { onFetch(false); }, []);

  const matrixData = useMemo(() => {
    const todayCA = new Date().toLocaleDateString('en-CA');
    const todayData = (data || []).filter(d => String(d.date || '').split('T')[0] === todayCA);
    const grouped: any = {};
    todayData.forEach(item => {
      if (!grouped[item.absent]) grouped[item.absent] = { department: item.department, periods: {} };
      const match = String(item.period || '').match(/\d+/);
      const pNum = match ? match[0] : item.period;
      grouped[item.absent].periods[pNum] = { sub: item.sub, class: item.class, status: item.status };
    });
    return Object.keys(grouped).map(absentName => ({ absentName, department: grouped[absentName].department, periods: grouped[absentName].periods }));
  }, [data]);

  const printMatrixReport = async () => {
    let tableRows = '';
    if (matrixData.length === 0) {
      tableRows = `<tr><td colspan="9" style="text-align:center; padding: 20px;">لا توجد حصص احتياط مسجلة لهذا اليوم.</td></tr>`;
    } else {
      matrixData.forEach((row, idx) => {
        let pCells = '';
        for (let i = 1; i <= 8; i++) {
          const p = row.periods[i];
          pCells += p
            ? `<td style="font-size: 10pt; font-weight: bold; background-color: ${p.status === 'Executed' ? '#ecfdf5' : '#fffbeb'}; border: 2px solid ${p.status === 'Executed' ? '#10b981' : '#f59e0b'};">${p.sub}<br><span style="font-size:8pt; color:#666;">(${p.class})</span></td>`
            : `<td style="color:#cbd5e1;">-</td>`;
        }
        tableRows += `<tr><td style="font-weight:900; background-color:#f8fafc;">${idx + 1}. ${row.absentName}<br><span style="font-size:8pt; font-weight:normal; color:#64748b;">${row.department}</span></td>${pCells}</tr>`;
      });
    }

    const html = `
      <html dir="rtl" lang="ar">
      <head>
        <title>سجل الاحتياط اليومي</title>
        <style>
          body { font-family: Tahoma, Arial, sans-serif; color: #000; background: #fff; -webkit-print-color-adjust: exact; }
          .report-page { width: 100%; box-sizing: border-box; }
          h2 { text-align: center; text-decoration: underline; margin-bottom: 20px; }
          table { width: 100%; border-collapse: collapse; text-align: center; font-size: 10pt; }
          th, td { border: 1px solid #000; padding: 8px; vertical-align: middle; }
          th { background-color: #cbd5e1; font-weight: 900; }
          .header { display: flex; justify-content: space-between; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; font-weight:bold; }
        </style>
      </head>
      <body>
        <div class="report-page">
          <div class="header"><div>المدرسة: ${schoolName || '____________________'}</div><div>التاريخ: ${new Date().toLocaleDateString('ar-OM')}</div></div>
          <h2>سجل توزيع حصص الاحتياط اليومي (مصفوفة الإدارة)</h2>
          <table><thead><tr><th style="width:20%">المعلم الغائب</th><th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>6</th><th>7</th><th>8</th></tr></thead><tbody>${tableRows}</tbody></table>
          <div style="margin-top:40px; display:flex; justify-content:space-between; font-weight:bold;"><div>إعداد المناوب الإداري: ....................</div><div>يعتمد، مدير المدرسة: ....................</div></div>
        </div>
      </body>
      </html>`;

    await outputOfficialPdf(html, 'سجل توزيع حصص الاحتياط اليومي', 'landscape');
  };

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-2">
        <div className="flex items-center gap-3"><ShieldCheck className="text-emerald-600" size={32} /><div><h2 className="text-2xl font-black text-slate-800">رادار الاحتياط الشامل</h2><p className="text-sm font-bold text-slate-500">متابعة حية لتكليفات المعلمين الأوائل وتنفيذها</p></div></div>
        <div className="flex items-center gap-3"><button onClick={() => onFetch(true)} className="p-3 bg-white border border-slate-200 rounded-xl hover:bg-slate-50 transition-all shadow-sm"><RefreshCw className={`text-slate-600 ${isLoading ? 'animate-spin' : ''}`} size={20} /></button><button onClick={printMatrixReport} className="flex items-center gap-2 bg-emerald-600 text-white px-5 py-3 rounded-xl font-black hover:bg-emerald-700 transition shadow-lg active:scale-95"><Printer size={20} /> PDF مصفوفة الاحتياط</button></div>
      </div>

      <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] border border-white shadow-lg flex-1 overflow-hidden flex flex-col">
        <div className="overflow-x-auto flex-1 custom-scrollbar p-6">
          {matrixData.length === 0 ? <div className="flex flex-col items-center justify-center h-full text-slate-400 opacity-60"><ShieldCheck size={64} className="mb-4" /><p className="text-xl font-bold">لا توجد حصص احتياط مسجلة لهذا اليوم</p></div> : (
            <table className="w-full text-center border-collapse min-w-[900px]"><thead className="bg-slate-100/80 border-b-2 border-slate-300 sticky top-0 z-10"><tr><th className="p-4 font-black text-slate-700 text-right rounded-tr-xl border-l border-slate-200">المعلم الغائب / القسم</th>{[1,2,3,4,5,6,7,8].map(n => <th key={n} className="p-4 font-black text-slate-700 border-l border-slate-200 w-24">الحصة {n}</th>)}</tr></thead><tbody className="divide-y divide-slate-100">{matrixData.map((row, idx) => <tr key={idx} className="hover:bg-slate-50/50 transition-colors"><td className="p-4 text-right border-l border-slate-100"><p className="font-black text-slate-800 text-sm">{row.absentName}</p><p className="font-bold text-slate-500 text-[10px]">{row.department}</p></td>{[1,2,3,4,5,6,7,8].map(n => { const p = row.periods[n]; return <td key={n} className="p-2 border-l border-slate-100 align-middle">{p ? <div className={`flex flex-col items-center justify-center p-2 rounded-xl border-2 transition-all ${p.status === 'Executed' ? 'bg-emerald-50 border-emerald-500 text-emerald-800 shadow-[0_0_10px_rgba(16,185,129,0.2)]' : 'bg-amber-50 border-amber-400 text-amber-800 border-dashed'}`}><span className="font-black text-[11px] leading-tight">{p.sub}</span><span className="font-bold text-[9px] opacity-70">صف {p.class}</span>{p.status === 'Executed' ? <CheckCircle2 size={12} className="text-emerald-500 mt-1" /> : <Clock size={12} className="text-amber-500 mt-1" />}</div> : <span className="text-slate-300">-</span>}</td>; })}</tr>)}</tbody></table>
          )}
        </div>
        <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-center gap-6 text-xs font-bold text-slate-600"><div className="flex items-center gap-2"><div className="w-3 h-3 bg-emerald-50 border-2 border-emerald-500 rounded-full" /> تم الدخول للفصل</div><div className="flex items-center gap-2"><div className="w-3 h-3 bg-amber-50 border-2 border-amber-400 border-dashed rounded-full" /> بانتظار المعلم</div></div>
      </div>
    </div>
  );
}

function ReportsPage({ data, schoolName }: { data: DashboardData; schoolName: string }) {
  const [selectedDateStr, setSelectedDateStr] = useState(new Date().toISOString().split('T')[0]);

  const reportData = useMemo(() => {
    const targetDate = new Date(selectedDateStr);
    const targetArabicDay = getArabicDayName(targetDate);
    const normalizedTargetDay = normalizeText(targetArabicDay);
    const expected = (data.teachers || []).filter((t: TeacherRow) => normalizeText(t.day) === normalizedTargetDay);
    const logsForDate = (data.logs || []).filter((log: LogRow) => isSameDate(log.time, targetDate));
    const uniqueLogsForDate = new Map<string, any>();

    [...logsForDate].reverse().forEach((log: LogRow) => {
      const key = normalizeText(log.teacherName);
      if (!uniqueLogsForDate.has(key)) uniqueLogsForDate.set(key, { ...log });
      else {
        const existing = uniqueLogsForDate.get(key);
        const mergeStrings = (s1: string, s2: string) => {
          const combined = Array.from(new Set([...splitList(s1), ...splitList(s2)]));
          return combined.length ? combined.join('، ') : 'لا يوجد';
        };
        existing.absentStudents = mergeStrings(existing.absentStudents, log.absentStudents);
        existing.lateStudents = mergeStrings(existing.lateStudents, log.lateStudents);
        existing.truantStudents = mergeStrings(existing.truantStudents, log.truantStudents);
        existing.time = log.time;
      }
    });

    return expected.map((teacher: TeacherRow) => {
      const teacherLog = uniqueLogsForDate.get(normalizeText(teacher.name));
      return { name: teacher.name, className: teacher.className, status: teacherLog ? 'مكتمل' : 'متأخر', absents: teacherLog?.absentStudents || 'لا يوجد', lates: teacherLog?.lateStudents || 'لا يوجد', truants: teacherLog?.truantStudents || 'لا يوجد', time: teacherLog ? new Date(teacherLog.time).toLocaleTimeString('ar-OM', { hour: '2-digit', minute: '2-digit' }) : '-' };
    });
  }, [data, selectedDateStr]);

  const handlePrintArchive = async () => {
    const combinedData: any[] = [];
    reportData.forEach((row: any) => {
      if (row.status === 'مكتمل') {
        const extract = (str: string, type: string) => splitList(str).forEach(name => combinedData.push({ name, type, teacher: row.name, className: row.className, time: row.time }));
        extract(row.absents, 'غياب');
        extract(row.lates, 'تأخير');
        extract(row.truants, 'تسرب');
      }
    });
    await generateOfficialReport('التقرير الأرشيفي الشامل لغياب ومخالفات الطلاب', combinedData, schoolName, new Date(selectedDateStr).toLocaleDateString('ar-OM'), 'students');
  };

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-2"><div className="flex items-center gap-3"><FileText className="text-indigo-500" size={28} /><h2 className="text-2xl font-black text-slate-800">أرشيف التقارير</h2></div><div className="flex items-center gap-3"><div className="flex items-center gap-2 bg-white px-4 py-2 rounded-xl shadow-sm border border-slate-200"><CalendarIcon size={20} className="text-indigo-500" /><input type="date" value={selectedDateStr} onChange={(e) => setSelectedDateStr(e.target.value)} className="bg-transparent font-bold text-slate-700 outline-none" /></div><button onClick={handlePrintArchive} className="flex items-center gap-2 bg-indigo-600 text-white px-4 py-2 rounded-xl font-bold hover:bg-indigo-700 transition shadow-sm active:scale-95"><Printer size={18} /> استخراج PDF رسمي</button></div></div>
      <div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] p-6 border border-white shadow-lg flex-1 overflow-hidden flex flex-col"><div className="overflow-x-auto flex-1 custom-scrollbar"><table className="w-full text-right border-collapse min-w-[800px]"><thead className="bg-slate-100/80 border-b border-slate-200 sticky top-0 z-10"><tr><th className="p-4 font-black text-slate-600">المعلم</th><th className="p-4 font-black text-slate-600">الفصل</th><th className="p-4 font-black text-slate-600 text-center">الرصد</th><th className="p-4 font-black text-slate-600">تفاصيل المخالفات</th><th className="p-4 font-black text-slate-600">الوقت</th></tr></thead><tbody className="divide-y divide-slate-100">{reportData.length === 0 ? <tr><td colSpan={5} className="py-10 text-center font-bold text-slate-400">لا يوجد معلمين مسندين لهذا اليوم.</td></tr> : reportData.map((row: any, idx: number) => <tr key={idx} className="hover:bg-slate-50 transition-colors"><td className="p-4 font-black text-indigo-900">{row.name}</td><td className="p-4 font-bold text-slate-600">{row.className}</td><td className="p-4 text-center"><span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-xl text-xs font-bold ${row.status === 'مكتمل' ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>{row.status === 'مكتمل' ? <CheckCircle2 size={14} /> : <Clock size={14} />} {row.status}</span></td><td className="p-4 text-sm font-bold text-slate-600"><div className="flex flex-col gap-1 items-start">{splitList(row.absents).length > 0 && <span className="bg-rose-50 text-rose-700 px-2 py-1 rounded-md text-xs border border-rose-100 w-fit">🔴 غياب: {row.absents}</span>}{splitList(row.lates).length > 0 && <span className="bg-amber-50 text-amber-700 px-2 py-1 rounded-md text-xs border border-amber-100 w-fit">🟠 تأخير: {row.lates}</span>}{splitList(row.truants).length > 0 && <span className="bg-purple-50 text-purple-700 px-2 py-1 rounded-md text-xs border border-purple-100 w-fit">🟣 تسرب: {row.truants}</span>}{splitList(row.absents).length === 0 && splitList(row.lates).length === 0 && splitList(row.truants).length === 0 && row.status === 'مكتمل' && <span className="text-emerald-500 font-bold bg-emerald-50 px-2 py-1 rounded-md text-xs">✔️ لا توجد مخالفات</span>}</div></td><td className="p-4 text-slate-500 text-sm font-mono">{row.time}</td></tr>)}</tbody></table></div></div>
    </div>
  );
}

function SearchPage({ data }: { data: DashboardData }) {
  const [searchQuery, setSearchQuery] = useState('');
  const results = useMemo(() => {
    if (!searchQuery.trim()) return [];
    const term = normalizeText(searchQuery.trim());
    const found: any[] = [];
    (data.logs || []).forEach((log: LogRow) => {
      const check = (listStr: string, label: string, color: string) => splitList(listStr).forEach(name => {
        if (normalizeText(name).includes(term)) found.push({ name, label, color, teacher: log.teacherName, className: log.className, time: new Date(log.time).toLocaleString('ar-OM', { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) });
      });
      check(log.absentStudents, '🔴 غياب', 'bg-rose-50 text-rose-700 border-rose-200');
      check(log.lateStudents, '🟠 تأخير', 'bg-amber-50 text-amber-700 border-amber-200');
      check(log.truantStudents, '🟣 تسرب', 'bg-purple-50 text-purple-700 border-purple-200');
    });
    return found.reverse();
  }, [searchQuery, data.logs]);

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col"><div className="flex items-center gap-3 px-2"><Search className="text-indigo-500" size={28} /><h2 className="text-2xl font-black text-slate-800">البحث المتقدم عن الطلاب</h2></div><div className="bg-white p-4 sm:p-6 rounded-[2rem] shadow-sm border border-slate-200 flex items-center gap-4 transition-all focus-within:ring-4 ring-indigo-500/10 focus-within:border-indigo-300"><Search className="text-indigo-400" size={28} /><input type="text" value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} placeholder="اكتب اسم الطالب للبحث عن المخالفات..." className="flex-1 bg-transparent text-lg font-bold text-slate-800 outline-none placeholder:text-slate-400" /></div><div className="bg-white/70 backdrop-blur-2xl rounded-[2rem] p-6 border border-white shadow-lg flex-1 overflow-hidden flex flex-col"><div className="overflow-x-auto flex-1 custom-scrollbar">{searchQuery.trim() === '' ? <div className="h-full flex flex-col items-center justify-center text-slate-400 opacity-60"><div className="w-24 h-24 bg-slate-100 rounded-full flex items-center justify-center mb-4"><Search size={40} className="text-slate-400" /></div><p className="text-xl font-bold">ابدأ بكتابة اسم الطالب لتتبع سجله</p></div> : results.length === 0 ? <div className="h-full flex flex-col items-center justify-center text-emerald-500"><div className="w-24 h-24 bg-emerald-50 rounded-full flex items-center justify-center mb-4"><CheckCircle2 size={48} className="text-emerald-500" /></div><p className="text-2xl font-black mb-2">سجل الطالب نظيف!</p><p className="font-bold text-emerald-600/80">لم يتم رصد أي غياب أو تأخير أو تسرب بهذا الاسم.</p></div> : <table className="w-full text-right border-collapse min-w-[800px]"><thead className="bg-slate-100/80 border-b border-slate-200 sticky top-0 z-10"><tr><th className="p-4 font-black text-slate-600 rounded-tr-xl">اسم الطالب</th><th className="p-4 font-black text-slate-600">نوع المخالفة</th><th className="p-4 font-black text-slate-600">المعلم</th><th className="p-4 font-black text-slate-600">الفصل</th><th className="p-4 font-black text-slate-600 rounded-tl-xl">التاريخ والوقت</th></tr></thead><tbody className="divide-y divide-slate-100">{results.map((row, idx) => <tr key={idx} className="hover:bg-slate-50 transition-colors"><td className="p-4 font-black text-indigo-900">{row.name}</td><td className="p-4"><span className={`px-3 py-1.5 rounded-lg text-xs font-bold border ${row.color}`}>{row.label}</span></td><td className="p-4 font-bold text-slate-600">{row.teacher}</td><td className="p-4"><span className="bg-indigo-50 text-indigo-600 px-3 py-1 rounded-md text-xs font-bold">{row.className}</span></td><td className="p-4 text-slate-500 font-mono text-sm">{row.time}</td></tr>)}</tbody></table>}</div></div></div>
  );
}

function SettingsPage({ schoolCode, schoolName, setSchoolName, onUploadComplete }: any) {
  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{ type: 'idle' | 'success' | 'error'; msg: string }>({ type: 'idle', msg: '' });

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
      const teachersList = jsonData.map(row => ({ name: row['اسم المعلم'] || row['الاسم'] || row['المعلم'] || '', className: row['الفصل'] || row['الفصل المسند'] || row['الصفوف'] || '', day: row['اليوم'] || '', period: row['الحصة'] || row['رقم الحصة'] || '' })).filter(t => t.name && t.day && t.period);
      if (teachersList.length === 0) throw new Error('تأكد من وجود الأعمدة: اسم المعلم، الفصل، اليوم، الحصة.');
      setUploadStatus({ type: 'idle', msg: `تم استخراج ${teachersList.length} سجل، جاري الإرسال للسحابة...` });
      const response = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'bulk_upload_teachers', schoolCode, teachers: teachersList }) });
      const result = await response.json();
      if (result.status === 'success') { setUploadStatus({ type: 'success', msg: result.message }); onUploadComplete?.(); }
      else throw new Error(result.message);
    } catch (error: any) {
      setUploadStatus({ type: 'error', msg: error.message || 'حدث خطأ أثناء رفع الملف.' });
    } finally {
      setIsUploading(false);
      e.target.value = '';
    }
  };

  return (
    <div className="space-y-6 pb-10 h-full flex flex-col"><div className="flex items-center gap-3 px-2"><Settings className="text-indigo-500" size={28} /><h2 className="text-2xl font-black text-slate-800">الإعدادات</h2></div><div className="grid grid-cols-1 lg:grid-cols-2 gap-6"><div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 sm:p-8 border border-white/80 shadow-[0_8px_30px_rgb(0,0,0,0.04)]"><h3 className="text-xl font-bold text-slate-800 mb-6 flex items-center gap-2"><School className="text-amber-500" size={24} /> معلومات المدرسة</h3><div className="space-y-5"><div className="space-y-2"><label className="block text-sm font-bold text-slate-700 ml-2">اسم المدرسة</label><input type="text" value={schoolName} onChange={(e) => setSchoolName(e.target.value)} placeholder="مثال: مدرسة الإمام الشافعي" className="w-full px-4 py-3 rounded-2xl bg-white border border-indigo-100 text-indigo-900 font-bold outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-400/20 transition-all" /></div><div className="space-y-2"><label className="block text-sm font-bold text-slate-700 ml-2">كود المدرسة</label><div className="w-full px-4 py-3 rounded-2xl bg-indigo-50 border border-indigo-100 text-indigo-900 font-black text-xl text-center tracking-[0.3em] font-mono" dir="ltr">{schoolCode}</div></div></div></div><div className="bg-white/60 backdrop-blur-2xl rounded-[2rem] p-6 sm:p-8 border border-white/80 shadow-[0_8px_30px_rgb(0,0,0,0.04)]"><h3 className="text-xl font-bold text-slate-800 mb-4 flex items-center gap-2"><FileSpreadsheet className="text-blue-500" size={24} /> إدارة المعلمين المرجعية</h3><p className="text-slate-500 font-bold text-sm mb-6 leading-relaxed">ارفع ملف Excel يحتوي على الأعمدة: اسم المعلم، الفصل، اليوم، الحصة. سيتم استبدال بيانات المدرسة الحالية لتجنب التكرار.</p><div className="relative border-2 border-dashed border-indigo-200 rounded-2xl p-8 flex flex-col items-center justify-center text-center bg-indigo-50/30 hover:bg-indigo-50/80 transition-colors group">{isUploading ? <div className="flex flex-col items-center"><Loader2 size={40} className="text-indigo-500 animate-spin mb-4" /><h4 className="font-bold text-slate-800">{uploadStatus.msg}</h4></div> : <><div className="bg-white p-4 rounded-full shadow-sm mb-4 group-hover:scale-110 transition-transform duration-300"><UploadCloud size={32} className="text-indigo-500" /></div><h4 className="font-bold text-slate-800 mb-1">اضغط هنا لاختيار ملف الإكسل</h4><input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" disabled={isUploading} /></>}</div>{uploadStatus.type === 'success' && <div className="mt-4 p-4 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-xl font-bold flex items-center gap-2"><CheckCircle2 size={20} />{uploadStatus.msg}</div>}{uploadStatus.type === 'error' && <div className="mt-4 p-4 bg-rose-50 border border-rose-200 text-rose-700 rounded-xl font-bold flex items-center gap-2"><AlertTriangle size={20} />{uploadStatus.msg}</div>}</div></div></div>
  );
}

function StatCard({ title, value, subtitle, icon: Icon, color }: any) {
  const classes: any = { indigo: 'from-indigo-500 to-blue-600 bg-indigo-50 text-indigo-600', amber: 'from-amber-400 to-orange-500 bg-amber-50 text-amber-600', rose: 'from-rose-400 to-red-500 bg-rose-50 text-rose-600', green: 'from-emerald-400 to-green-500 bg-emerald-50 text-emerald-600' };
  const gradient = classes[color].split(' ').slice(0, 2).join(' ');
  const iconClass = classes[color].split(' ').slice(2).join(' ');
  return <div className="bg-white rounded-[2rem] p-6 shadow-xl border border-white flex justify-between items-center group overflow-hidden relative"><div className={`absolute inset-y-0 right-0 w-2 bg-gradient-to-b ${gradient}`} /><div><p className="text-sm font-bold text-slate-400 mb-1">{title}</p><h3 className="text-4xl font-black text-slate-800">{value} <span className="text-sm text-slate-300">{subtitle}</span></h3></div><div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${iconClass} group-hover:scale-110 transition-transform`}><Icon size={28} /></div></div>;
}

function LoginScreen({ schoolCode, setSchoolCode, onLogin, isLoading, errorMsg }: any) {
  return (
    <div className="min-h-[100dvh] w-full flex flex-col items-center justify-center font-sans overflow-hidden relative px-6 bg-[#f0f4f8]" dir="rtl">
      <div className="absolute top-[-10%] right-[-10%] w-[50%] h-[40%] bg-indigo-300/20 rounded-full blur-[100px] pointer-events-none" />
      <div className="absolute bottom-[10%] left-[-10%] w-[40%] h-[50%] bg-blue-300/20 rounded-full blur-[100px] pointer-events-none" />

      <motion.main initial={{ opacity: 0, y: 18, scale: 0.98 }} animate={{ opacity: 1, y: 0, scale: 1 }} transition={{ duration: 0.55 }} className="w-full max-w-md relative z-10 flex flex-col items-center py-10">
        <div className="text-center mb-8 shrink-0">
          <div className="inline-flex items-center justify-center p-5 rounded-[2rem] bg-white/80 backdrop-blur-xl mb-6 shadow-xl border border-white/70 text-[#002366]"><School className="w-12 h-12" /></div>
          <h1 className="text-5xl font-black text-[#002366] tracking-tight mb-2">راصد</h1>
          <p className="text-slate-500 font-black tracking-wide text-sm">بوابة الإدارة المدرسية</p>
          <p className="text-[10px] font-bold text-slate-400 mt-2">هوية موحدة لعائلة راصد</p>
        </div>

        <div className="w-full bg-white/85 backdrop-blur-xl rounded-[2.5rem] p-8 shadow-xl border border-white/70">
          <div className="text-center mb-6">
            <div className="w-11 h-11 rounded-2xl bg-indigo-50 border border-indigo-100 mx-auto flex items-center justify-center text-[#002366] mb-3"><ShieldCheck size={22} /></div>
            <h2 className="text-xl font-black text-[#002366]">دخول الإدارة</h2>
            <p className="text-[10px] font-bold text-slate-400 mt-1">أدخل كود المدرسة للاتصال بلوحة القيادة</p>
          </div>

          <form onSubmit={onLogin} className="space-y-6">
            <div className="space-y-2">
              <label className="block text-xs font-bold text-slate-600 px-1 text-right">كود المدرسة</label>
              <div className="relative group">
                <div className="absolute inset-y-0 right-0 pr-4 flex items-center pointer-events-none text-[#002366] z-10"><School className="w-6 h-6" /></div>
                <input type="text" value={schoolCode} onChange={(e) => setSchoolCode(e.target.value.trim())} className="block w-full pr-14 pl-4 py-4 bg-slate-50 border border-indigo-100 rounded-2xl focus:ring-4 focus:ring-indigo-100 text-[#002366] font-black text-2xl outline-none text-center placeholder:text-slate-300 tracking-[0.25em]" placeholder="••••" required dir="ltr" />
              </div>
              {errorMsg && <p className="text-rose-500 text-xs font-bold text-center mt-2 animate-in fade-in">{errorMsg}</p>}
            </div>

            <button type="submit" disabled={!schoolCode || isLoading} className="w-full bg-[#002366] text-white py-4 rounded-2xl font-black text-base flex items-center justify-center gap-3 shadow-xl hover:scale-[1.02] active:scale-95 transition-all disabled:opacity-50">
              {isLoading ? <Loader2 className="animate-spin" /> : <><span>دخول آمن</span><ArrowLeft className="w-5 h-5" /></>}
            </button>
          </form>
        </div>

        <div className="mt-8 text-center opacity-60">
          <p className="text-slate-500 text-[10px] font-bold mb-1">برمجة وتطوير</p>
          <div className="flex items-center justify-center gap-1.5"><Code size={12} className="text-[#002366]" /><span className="text-[#002366] text-[11px] font-black tracking-widest uppercase">ALZAABI MOHAMMAD</span></div>
        </div>
      </motion.main>
    </div>
  );
}
