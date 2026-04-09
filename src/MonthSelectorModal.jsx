import React from 'react';
import { Calendar, Loader2 } from 'lucide-react';

export const MonthSelectorModal = ({
  allAvailableMonths,
  selectedStartupMonths,
  setSelectedStartupMonths,
  setSelectedMonths,
  setShowMonthSelector,
  loading
}) => {
  return (
    <div className="fixed inset-0 z-[150] bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-3xl p-8 w-full max-w-sm shadow-2xl text-center">
        <div className="w-14 h-14 bg-[#842327]/10 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <Calendar size={24} className="text-[#842327]"/>
        </div>
        <h3 className="text-lg font-black text-slate-800 mb-2">เลือกเดือนที่ต้องการทำงาน</h3>
        <p className="text-sm text-slate-500 mb-6 leading-relaxed">
            เพื่อเพิ่มความเร็วในการโหลดข้อมูล กรุณาเลือกเดือนที่ต้องการดูข้อมูลเป็นหลัก
        </p>
        {allAvailableMonths.length > 0 ? (
          <div className="space-y-4">
            <div className="flex flex-wrap gap-2 justify-center max-h-48 overflow-y-auto p-1">
              {allAvailableMonths.map(m => (
                <button key={m} onClick={() => setSelectedStartupMonths(prev => prev.includes(m) ? prev.filter(x => x !== m) : [...prev, m])}
                  className={`px-4 py-2 rounded-xl text-sm font-bold transition-all border ${selectedStartupMonths.includes(m) ? 'bg-[#842327] text-white border-[#842327] shadow-md' : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'}`}>
                  {m}
                </button>
              ))}
            </div>
            <p className="text-[10px] text-slate-400">กรุณาเลือกเดือนเพื่อดึงข้อมูล (เพื่อประหยัดโควต้าการอ่านข้อมูล)</p>
            <button onClick={() => { setSelectedMonths(selectedStartupMonths); setShowMonthSelector(false); }} className="w-full py-3.5 bg-gradient-to-r from-[#842327] via-[#D32F2F] to-[#842327] animate-gradient-x shimmer text-white rounded-xl font-black text-sm tracking-wider transition-all shadow-lg shadow-[#842327]/30 hover:shadow-[#D32F2F]/40 hover:-translate-y-0.5 active:scale-[0.98]">
              ยืนยันและเริ่มใช้งาน
            </button>
          </div>
        ) : loading ? (
          <div className="flex flex-col items-center gap-3 py-6">
            <Loader2 className="animate-spin text-[#842327]" size={24} />
            <div className="text-sm font-bold text-slate-400">กำลังตรวจสอบข้อมูล...</div>
          </div>
        ) : (
          <div className="space-y-4">
            <div className="text-sm text-slate-500 py-6 font-semibold border border-slate-100 bg-slate-50 rounded-xl">ยังไม่มีข้อมูลเดือนในโปรเจกต์นี้</div>
            <button onClick={() => { setSelectedMonths([]); setShowMonthSelector(false); }} className="w-full py-3.5 bg-slate-800 hover:bg-slate-700 text-white rounded-xl font-black text-sm tracking-wider transition-all shadow-lg">
              เข้าสู่ Dashboard
            </button>
          </div>
        )}
      </div>
    </div>
  );
};