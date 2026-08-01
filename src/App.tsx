import React, { useState, useRef } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { UploadCloud, FileSpreadsheet, CheckCircle2, AlertCircle, Loader2, Download, X } from 'lucide-react';
import { Student, SUBJECTS } from './types';

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [logs, setLogs] = useState<string[]>([]);
  const [isDone, setIsDone] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).filter(f => f.name.endsWith('.xlsx'));
      setFiles(prev => [...prev, ...newFiles]);
      setIsDone(false);
    }
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const addLog = (msg: string) => {
    setLogs(prev => [...prev, msg]);
  };

  const processFiles = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);
    setLogs([]);
    setIsDone(false);

    try {
      const studentsData: Record<string, Student> = {};
      let validFileCount = 0;

      for (const file of files) {
        if (file.name === 'Ket_Qua_Tong_Hop.xlsx') continue;

        const nameParts = file.name.replace(/\.xlsx$/i, '').split('_');
        if (nameParts.length < 2) {
          addLog(`Bỏ qua file ${file.name}: Tên file không đúng định dạng (Môn_GiaiĐoạn.xlsx)`);
          continue;
        }

        const subject = nameParts[0];
        if (!SUBJECTS.includes(subject)) {
          addLog(`Bỏ qua file ${file.name}: Môn học '${subject}' không nằm trong danh sách hỗ trợ.`);
          continue;
        }

        addLog(`Đang xử lý: ${file.name} (Môn: ${subject})`);
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
          addLog(`Lỗi: Không tìm thấy sheet nào trong file ${file.name}`);
          continue;
        }

        let rowCount = 0;
        worksheet.eachRow((row, rowNumber) => {
          // Bỏ qua 6 dòng đầu (dòng 6 là tiêu đề)
          if (rowNumber < 7) return;

          // Cột 7: Số BD, Cột 2: Họ và tên, Cột 3: Ngày sinh, Cột 4: Giới, Cột 9: Điểm, Cột 10: Ghi chú
          const sbdCell = row.getCell(7);
          let sbd = sbdCell.text?.trim();
          if (!sbd && sbdCell.value !== null && sbdCell.value !== undefined) {
            sbd = String(sbdCell.value).trim();
          }
          
          if (!sbd) return; // Không có số báo danh

          if (sbd.endsWith('.0')) {
            sbd = sbd.slice(0, -2);
          }

          const name = row.getCell(2).text || String(row.getCell(2).value || '');
          const rowText = row.values ? JSON.stringify(row.values).toUpperCase() : '';
          if (!name.trim() || name.includes('Ngày') || name.includes('tháng') || rowText.includes('TỔ TRƯỞNG') || rowText.includes('GIÁO VIÊN')) return; // Bỏ qua dòng rác/footer
          let dob = row.getCell(3).value;
          if (dob instanceof Date) {
            // format date slightly or just keep date object, ExcelJS can write dates back
          } else {
            dob = row.getCell(3).text || String(dob || '');
          }
          const gender = row.getCell(4).text || String(row.getCell(4).value || '');
          const scoreValue = row.getCell(9).value;
          const note = row.getCell(10).text || String(row.getCell(10).value || ''); // Lớp

          if (!studentsData[sbd]) {
            studentsData[sbd] = {
              sbd,
              name,
              dob,
              gender,
              lop: note,
              scores: SUBJECTS.reduce((acc, sub) => {
                acc[sub] = [];
                return acc;
              }, {} as Record<string, number[]>)
            };
          }

          if (scoreValue !== null && scoreValue !== undefined && typeof scoreValue === 'number') {
            studentsData[sbd].scores[subject].push(scoreValue);
          } else {
            // Cố gắng parse float nếu là chuỗi
            const parsedScore = parseFloat(String(scoreValue));
            if (!isNaN(parsedScore)) {
              studentsData[sbd].scores[subject].push(parsedScore);
            }
          }
          rowCount++;
        });

        validFileCount++;
        addLog(`Đã đọc ${rowCount} dòng dữ liệu từ ${file.name}`);
      }

      if (validFileCount === 0 || Object.keys(studentsData).length === 0) {
        addLog("Không có dữ liệu hợp lệ nào được tìm thấy.");
        setIsProcessing(false);
        return;
      }

      addLog("Đang tổng hợp điểm và tạo file kết quả...");

      // Tạo file kết quả
      const resultWorkbook = new ExcelJS.Workbook();
      
      const studentList = Object.values(studentsData);
      studentList.sort((a, b) => {
        if (a.lop !== b.lop) return a.lop.localeCompare(b.lop);
        return a.sbd.localeCompare(b.sbd);
      });

      const lops = Array.from(new Set(studentList.map(s => s.lop))).filter(Boolean).sort();
      
      const sheetsToCreate = [
        { name: 'Kết quả', data: studentList },
        ...lops.map(lop => ({ name: lop, data: studentList.filter(s => s.lop === lop) }))
      ];

      sheetsToCreate.forEach(sheetInfo => {
        const resultSheet = resultWorkbook.addWorksheet(sheetInfo.name);

        const activeSubjects = SUBJECTS.filter(sub => {
          return sheetInfo.data.some(student => student.scores[sub].length > 0);
        });

        // Tạo header
        const headers = [
          'STT', 'Họ và tên', 'Ngày sinh', 'Giới', 'Số BD',
          ...activeSubjects.map(sub => `TB_${sub}`),
          'TB_All', 'Lớp'
        ];
        resultSheet.addRow(headers);
        
        // Định dạng header
        const headerRow = resultSheet.getRow(1);
        headerRow.font = { name: 'Times New Roman', size: 12, bold: true };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

        sheetInfo.data.forEach((student, index) => {
          let totalAvg = 0;
          let countSubjects = 0;
          
          const rowValues: any[] = [
            index + 1,
            student.name,
            student.dob,
            student.gender,
            student.sbd
          ];

          activeSubjects.forEach(sub => {
            const scores = student.scores[sub];
            if (scores.length > 0) {
              const avg = scores.reduce((a, b) => a + b, 0) / scores.length;
              rowValues.push(Math.round(avg * 100) / 100);
              totalAvg += avg;
              countSubjects++;
            } else {
              rowValues.push(null);
            }
          });

          let tbAll = null;
          if (countSubjects === 6) {
            tbAll = Math.round((totalAvg / 6) * 100) / 100;
          } else if (countSubjects > 0) {
            tbAll = Math.round((totalAvg / countSubjects) * 100) / 100;
          }
          
          rowValues.push(tbAll);
          rowValues.push(student.lop);

          resultSheet.addRow(rowValues);
        });

        // Tự động điều chỉnh độ rộng cột và thêm viền, font
        const tbAllIndex = 5 + activeSubjects.length;
        const lopIndex = tbAllIndex + 1;

        resultSheet.columns.forEach((column, i) => {
          let maxLength = 0;
          column.eachCell?.({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? cell.value.toString().length : 10;
            if (columnLength > maxLength) {
              maxLength = columnLength;
            }
            
            const isBoldColumn = i === tbAllIndex || i === lopIndex; // TB_All and Lớp
            cell.font = { name: 'Times New Roman', size: 12, bold: cell.row === 1 || isBoldColumn };
            
            if (isBoldColumn && cell.row > 1) {
              cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
            
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            };
          });
          column.width = Math.max(10, maxLength + 2); // Minimum width of 10
        });
      });

      const buffer = await resultWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'Ket_Qua_Tong_Hop.xlsx');

      addLog("HOÀN THÀNH! File 'Ket_Qua_Tong_Hop.xlsx' đã được tải xuống.");
      setIsDone(true);
    } catch (error: any) {
      console.error(error);
      addLog(`CÓ LỖI XẢY RA: ${error.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="h-screen bg-slate-50 text-slate-800 font-sans flex flex-col overflow-hidden">
      <header className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-white shrink-0">
        <div>
          <h1 className="text-xl font-semibold tracking-tight text-slate-900">THPT LÊ QUÝ ĐÔN - NB</h1>
          <p className="text-xs text-slate-500 mt-1">Hệ thống tổng hợp và phân tích điểm thi 4 GD</p>          
        </div>
        <div className="flex items-center gap-3">
          <div className="text-xs font-medium text-slate-500 bg-slate-100 px-3 py-1.5 rounded-full">
            {isProcessing ? 'Đang xử lý...' : isDone ? 'Hoàn thành' : 'Sẵn sàng'}
          </div>
          <div className={`w-2.5 h-2.5 rounded-full ${isProcessing ? 'bg-amber-400 animate-pulse' : isDone ? 'bg-emerald-500' : 'bg-blue-500'}`}></div>
        </div>
      </header>

      <main className="flex-1 flex overflow-hidden">
        <aside className="w-72 border-r border-slate-200 flex flex-col h-full bg-white shadow-[1px_0_10px_rgba(0,0,0,0.02)] z-10">
          <div className="p-5 border-b border-slate-100">
            <h2 className="font-semibold text-sm text-slate-800">Danh sách tệp</h2>
            <p className="text-xs text-slate-500 mt-1">{files.length} tệp đã tải lên</p>
          </div>
          <div className="flex-1 overflow-y-auto p-3 space-y-2 text-sm">
            {files.map((f, i) => (
              <div key={i} className="flex justify-between items-center p-2.5 bg-slate-50 border border-slate-100 rounded-lg hover:border-slate-200 transition-colors group">
                <div className="flex items-center gap-2 overflow-hidden">
                  <FileSpreadsheet className="w-4 h-4 text-blue-500 shrink-0" />
                  <span className="truncate text-slate-700">{f.name}</span>
                </div>
                <button onClick={(e) => { e.stopPropagation(); removeFile(i); }} disabled={isProcessing} className="text-slate-400 hover:text-red-500 shrink-0 opacity-0 group-hover:opacity-100 transition-opacity p-1">
                  <X className="w-4 h-4" />
                </button>
              </div>
            ))}
            {files.length === 0 && (
              <div className="p-4 text-center text-slate-400 text-sm">
                Chưa có tệp nào
              </div>
            )}
          </div>
          <div className="p-5 bg-slate-50/50 border-t border-slate-100 space-y-3">
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="w-full flex items-center justify-center gap-2 py-2.5 bg-white border border-slate-200 text-slate-700 text-sm font-medium rounded-lg hover:bg-slate-50 hover:border-slate-300 transition-all shadow-sm"
            >
              <UploadCloud className="w-4 h-4 text-slate-500" />
              Tải lên tệp Excel
            </button>
            <input 
              type="file" 
              multiple 
              accept=".xlsx" 
              className="hidden" 
              ref={fileInputRef}
              onChange={handleFileChange}
            />
            {files.length > 0 && (
              <button 
                onClick={() => setFiles([])}
                disabled={isProcessing}
                className="w-full flex items-center justify-center py-2.5 text-red-600 text-sm font-medium hover:bg-red-50 rounded-lg transition-colors"
              >
                Xóa tất cả
              </button>
            )}
          </div>
        </aside>
        
        <section className="flex-1 flex flex-col overflow-hidden bg-slate-50/50">
          <div className="flex-1 flex flex-col items-center justify-center p-8 relative overflow-hidden">
            <div className="z-10 text-center space-y-6 max-w-lg w-full">
               <div className="w-20 h-20 bg-blue-50 border border-blue-100 text-blue-600 rounded-2xl flex items-center justify-center mx-auto shadow-sm">
                 <FileSpreadsheet className="w-10 h-10" />
               </div>
               <div>
                 <h2 className="text-2xl font-semibold tracking-tight text-slate-900 mb-2">Tổng hợp điểm số</h2>
                 <p className="text-slate-500 text-sm leading-relaxed">Tải lên các tệp Excel (VD: Anh_GHKI.xlsx) và hệ thống sẽ tự động tính điểm trung bình 4 GĐ thi, làm căn cứ để sắp xếp lại các lớp.</p>
               </div>
               
               <button 
                  onClick={processFiles}
                  disabled={isProcessing || files.length === 0}
                  className="w-full sm:w-auto px-8 py-3 bg-blue-600 text-white font-medium rounded-xl shadow-sm hover:bg-blue-700 hover:shadow transition-all disabled:opacity-60 disabled:cursor-not-allowed flex items-center justify-center gap-2 mx-auto"
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin" />
                      Đang xử lý...
                    </>
                  ) : (
                    <>
                      <CheckCircle2 className="w-5 h-5" />
                      Bắt đầu tổng hợp
                    </>
                  )}
               </button>

               {isDone && (
                 <div className="mt-8 p-5 bg-emerald-50 border border-emerald-100 rounded-xl flex items-center gap-4 text-left">
                   <div className="w-10 h-10 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center shrink-0">
                     <Download className="w-5 h-5" />
                   </div>
                   <div>
                     <h3 className="font-semibold text-emerald-900">Hoàn thành</h3>
                     <p className="text-emerald-700 text-sm mt-0.5">File kết quả đã được tải xuống máy tính.</p>
                   </div>
                 </div>
               )}
            </div>
          </div>

          {/* Logs */}
          <div className="h-64 border-t border-slate-200 bg-white flex flex-col">
            <div className="px-4 py-2 border-b border-slate-100 bg-slate-50 flex items-center gap-2 text-xs font-medium text-slate-500 uppercase tracking-wider">
              <AlertCircle className="w-3.5 h-3.5" />
              Nhật ký hệ thống
            </div>
            <div className="flex-1 p-4 overflow-y-auto font-mono text-xs text-slate-600 space-y-1.5 leading-relaxed">
              {logs.length === 0 && <p className="text-slate-400 italic">Hệ thống đang chờ dữ liệu...</p>}
              {logs.map((log, i) => (
                <p key={i} className={`
                    ${log.includes('LỖI') || log.includes('Bỏ qua') ? 'text-amber-600' : ''}
                    ${log.includes('HOÀN THÀNH') ? 'text-emerald-600 font-medium' : ''}
                `}>
                  <span className="text-slate-400 mr-2">[{new Date().toLocaleTimeString()}]</span>
                  {log}
                </p>
              ))}
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}
