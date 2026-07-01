import React, { createContext, useContext, useState, useEffect, useRef } from 'react';
import { Filesystem, Directory, Encoding } from '@capacitor/filesystem';
import { Capacitor } from '@capacitor/core';

interface DashboardData {
  teachers: any[];
  logs: any[];
}

interface AdminContextType {
  dashboardData: DashboardData;
  setDashboardData: React.Dispatch<React.SetStateAction<DashboardData>>;
  isDataLoaded: boolean;
}

const AdminContext = createContext<AdminContextType | undefined>(undefined);

// 💉 1. اسم الملف الخاص ببوابة الإدارة فقط
const DBFILENAME = 'admin_raseddatabasev2.json';

export const AdminProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [dashboardData, setDashboardData] = useState<DashboardData>({ teachers: [], logs: [] });
  
  const isInitialLoad = useRef(true);
  const saveTimeoutRef = useRef<any>(null);

  const isHeavyEnvironment = () => Capacitor.isNativePlatform() || (window as any).electron !== undefined;

  // 📥 دالة استرجاع الأرشيف عند فتح التطبيق
  useEffect(() => {
    const loadData = async () => {
      try {
        let data: any = null;

        if (isHeavyEnvironment()) {
          try {
            const result = await Filesystem.readFile({ path: DBFILENAME, directory: Directory.Data, encoding: Encoding.UTF8 });
            if (result.data) data = JSON.parse(result.data as string);
          } catch (e) { console.log('ℹ️ No local admin file yet.'); }
        }

        if (!data) {
          // 💉 2. جلب البيانات باستخدام المفاتيح المعزولة للإدارة "admin_"
          const lsData = localStorage.getItem('admin_dashboardData');
          if (lsData) data = JSON.parse(lsData);
        }

        if (data) {
          if (data.dashboardData) setDashboardData(data.dashboardData);
        }
      } catch (error) {
        console.error('❌ Admin Data loading error', error);
      } finally {
        setIsDataLoaded(true);
        setTimeout(() => { isInitialLoad.current = false; }, 1000);
      }
    };
    loadData();
  }, []);

  // 📤 دالة حفظ الأرشيف تلقائياً عند أي تغيير (من السحابة)
  useEffect(() => {
    if (isInitialLoad.current) return;

    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(async () => {
      const dataToSave = {
        version: '1.0.0',
        timestamp: new Date().toISOString(),
        dashboardData
      };

      const isHeavy = isHeavyEnvironment();

      if (isHeavy) {
        try {
          await Filesystem.writeFile({ path: DBFILENAME, data: JSON.stringify(dataToSave), directory: Directory.Data, encoding: Encoding.UTF8 });
        } catch (e) {}
      }

      try {
        // 💉 3. حفظ البيانات باستخدام المفاتيح المعزولة للإدارة "admin_"
        if (!isHeavy) {
            localStorage.setItem('admin_dashboardData', JSON.stringify(dashboardData));
        }
      } catch (e) {}
    }, 2000); 

    return () => { if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current); };
  }, [dashboardData]);

  return (
    <AdminContext.Provider value={{ dashboardData, setDashboardData, isDataLoaded }}>
      {children}
    </AdminContext.Provider>
  );
};

export const useAdmin = () => {
  const context = useContext(AdminContext);
  if (context === undefined) throw new Error('useAdmin must be used within an AdminProvider');
  return context;
};
