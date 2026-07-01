import React, { createContext, useContext, useState, useEffect, useRef } from 'react';
import { Filesystem, Directory, Encoding } from '@capacitor/filesystem';
import { Capacitor } from '@capacitor/core';

interface DashboardData {
  teachers: any[];
  logs: any[];
}

interface AdminStoredData {
  version: string;
  timestamp: string;
  dashboardData: DashboardData;
}

interface AdminContextType {
  dashboardData: DashboardData;
  setDashboardData: React.Dispatch<React.SetStateAction<DashboardData>>;
  isDataLoaded: boolean;
  lastSavedAt: string;
  clearAdminData: () => Promise<void>;
}

const AdminContext = createContext<AdminContextType | undefined>(undefined);

// اسم ملف قاعدة بيانات راصد الإدارة فقط
const DBFILENAME = 'admin_raseddatabasev2.json';

// مفاتيح محلية معزولة للإدارة
const ADMIN_STORAGE_KEY = 'admin_dashboardData';
const ADMIN_STORAGE_META_KEY = 'admin_dashboardData_meta';

const EMPTY_DASHBOARD_DATA: DashboardData = {
  teachers: [],
  logs: []
};

const isHeavyEnvironment = () => {
  return Capacitor.isNativePlatform() || (window as any).electron !== undefined;
};

const safeParseJson = (value: any) => {
  try {
    if (!value) return null;
    if (typeof value !== 'string') return value;
    return JSON.parse(value);
  } catch (error) {
    console.warn('⚠️ Failed to parse admin JSON data:', error);
    return null;
  }
};

const normalizeStoredData = (rawData: any): AdminStoredData | null => {
  if (!rawData) return null;

  // الشكل الجديد
  if (rawData.dashboardData) {
    return {
      version: rawData.version || '1.0.0',
      timestamp: rawData.timestamp || '',
      dashboardData: {
        teachers: Array.isArray(rawData.dashboardData.teachers) ? rawData.dashboardData.teachers : [],
        logs: Array.isArray(rawData.dashboardData.logs) ? rawData.dashboardData.logs : []
      }
    };
  }

  // شكل قديم محتمل: البيانات نفسها بدون wrapping
  if (Array.isArray(rawData.teachers) || Array.isArray(rawData.logs)) {
    return {
      version: 'legacy',
      timestamp: '',
      dashboardData: {
        teachers: Array.isArray(rawData.teachers) ? rawData.teachers : [],
        logs: Array.isArray(rawData.logs) ? rawData.logs : []
      }
    };
  }

  return null;
};

export const AdminProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [dashboardData, setDashboardData] = useState<DashboardData>(EMPTY_DASHBOARD_DATA);
  const [lastSavedAt, setLastSavedAt] = useState('');

  const isInitialLoad = useRef(true);
  const saveTimeoutRef = useRef<any>(null);

  // استرجاع أرشيف راصد الإدارة عند فتح التطبيق
  useEffect(() => {
    const loadData = async () => {
      try {
        let loadedData: AdminStoredData | null = null;

        // 1) محاولة القراءة من ملف الجهاز في بيئات Capacitor / Electron
        if (isHeavyEnvironment()) {
          try {
            const result = await Filesystem.readFile({
              path: DBFILENAME,
              directory: Directory.Data,
              encoding: Encoding.UTF8
            });

            const parsed = safeParseJson(result.data);
            loadedData = normalizeStoredData(parsed);

            if (loadedData) {
              console.log('✅ Admin data loaded from device file.');
            }
          } catch (fileError) {
            console.log('ℹ️ No local admin file yet or failed to read file.');
          }
        }

        // 2) fallback إلى localStorage
        if (!loadedData) {
          const localData = localStorage.getItem(ADMIN_STORAGE_KEY);
          const parsed = safeParseJson(localData);
          loadedData = normalizeStoredData(parsed);

          if (loadedData) {
            console.log('✅ Admin data loaded from localStorage.');
          }
        }

        if (loadedData) {
          setDashboardData(loadedData.dashboardData);
          setLastSavedAt(loadedData.timestamp || '');
        }
      } catch (error) {
        console.error('❌ Admin data loading error:', error);
      } finally {
        setIsDataLoaded(true);

        // منع الحفظ المباشر أثناء أول تحميل
        setTimeout(() => {
          isInitialLoad.current = false;
        }, 1000);
      }
    };

    loadData();
  }, []);

  // حفظ أرشيف راصد الإدارة تلقائيًا عند أي تغيير
  useEffect(() => {
    if (isInitialLoad.current) return;

    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    saveTimeoutRef.current = setTimeout(async () => {
      const timestamp = new Date().toISOString();

      const dataToSave: AdminStoredData = {
        version: '1.1.0',
        timestamp,
        dashboardData: {
          teachers: Array.isArray(dashboardData.teachers) ? dashboardData.teachers : [],
          logs: Array.isArray(dashboardData.logs) ? dashboardData.logs : []
        }
      };

      const dataString = JSON.stringify(dataToSave);

      // 1) الحفظ في ملف الجهاز إن كان التطبيق يعمل في بيئة جهاز
      if (isHeavyEnvironment()) {
        try {
          await Filesystem.writeFile({
            path: DBFILENAME,
            data: dataString,
            directory: Directory.Data,
            encoding: Encoding.UTF8
          });
        } catch (fileError) {
          console.warn('⚠️ Failed to save admin data to device file:', fileError);
        }
      }

      // 2) الحفظ دائمًا في localStorage كنسخة احتياطية
      try {
        localStorage.setItem(ADMIN_STORAGE_KEY, dataString);
        localStorage.setItem(
          ADMIN_STORAGE_META_KEY,
          JSON.stringify({
            timestamp,
            teachersCount: dataToSave.dashboardData.teachers.length,
            logsCount: dataToSave.dashboardData.logs.length
          })
        );

        setLastSavedAt(timestamp);
      } catch (storageError) {
        console.warn('⚠️ Failed to save admin data to localStorage:', storageError);
      }
    }, 1500);

    return () => {
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current);
      }
    };
  }, [dashboardData]);

  const clearAdminData = async () => {
    setDashboardData(EMPTY_DASHBOARD_DATA);
    setLastSavedAt('');

    try {
      localStorage.removeItem(ADMIN_STORAGE_KEY);
      localStorage.removeItem(ADMIN_STORAGE_META_KEY);
    } catch (error) {
      console.warn('⚠️ Failed to clear admin localStorage data:', error);
    }

    if (isHeavyEnvironment()) {
      try {
        await Filesystem.deleteFile({
          path: DBFILENAME,
          directory: Directory.Data
        });
      } catch (error) {
        console.log('ℹ️ No admin file to delete or deletion failed.');
      }
    }
  };

  return (
    <AdminContext.Provider
      value={{
        dashboardData,
        setDashboardData,
        isDataLoaded,
        lastSavedAt,
        clearAdminData
      }}
    >
      {children}
    </AdminContext.Provider>
  );
};

export const useAdmin = () => {
  const context = useContext(AdminContext);

  if (context === undefined) {
    throw new Error('useAdmin must be used within an AdminProvider');
  }

  return context;
};
