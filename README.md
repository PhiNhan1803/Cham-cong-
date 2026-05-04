# Cham-cong-
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="Tính Công Lighting">
    <meta name="mobile-web-app-capable" content="yes">
    <title>Bảng Tính Công Nhân Sự Phim</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');
        body { font-family: 'Inter', sans-serif; -webkit-tap-highlight-color: transparent; }
        .no-scrollbar::-webkit-scrollbar { display: none; }
        input, select, textarea { font-size: 16px !important; }
        .staff-checkbox:checked + label { background-color: #2563eb; color: white; border-color: #2563eb; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .animate-fade-in { animation: fadeIn 0.2s ease-out forwards; }
        @media print { .no-print { display: none !important; } }
        .safe-area-bottom { padding-bottom: env(safe-area-inset-bottom); }
    </style>
</head>
<body class="bg-slate-50 min-h-screen text-slate-900 pb-20">
    <div class="max-w-[1600px] mx-auto p-3 md:p-6">
        <!-- Header -->
        <div class="no-print flex flex-col md:flex-row justify-between items-center mb-4 gap-4 bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
            <div class="text-center md:text-left">
                <h1 class="text-lg font-black text-blue-600 uppercase tracking-tight">Lighting Crew Manager</h1>
                <p class="text-[8px] text-green-500 font-bold uppercase mt-0.5 italic">CHẾ ĐỘ OFFLINE - ĐÃ SẴN SÀNG</p>
            </div>
            <div class="flex space-x-1 bg-slate-100 p-1 rounded-xl w-full md:w-auto overflow-x-auto no-scrollbar">
                <button onclick="switchTab('calc')" id="btn-tab-calc" class="flex-1 md:flex-none px-4 py-2 rounded-lg font-bold transition-all bg-white text-blue-600 shadow-sm text-[10px] uppercase">Máy tính</button>
                <button onclick="switchTab('staff')" id="btn-tab-staff" class="flex-1 md:flex-none px-4 py-2 rounded-lg font-bold transition-all text-slate-600 hover:bg-slate-200 text-[10px] uppercase">Nhân sự</button>
            </div>
        </div>

        <!-- TAB 1: MÁY TÍNH CÔNG -->
        <div id="tab-calc" class="block space-y-4">
            <div class="no-print bg-white rounded-2xl shadow-sm p-4 border border-slate-200">
                <div class="flex justify-between items-center mb-4">
                    <h2 class="text-[10px] font-black uppercase tracking-widest text-slate-400 flex items-center">
                        <span class="w-1.5 h-3 bg-orange-500 rounded-full mr-2"></span>Thông tin show
                    </h2>
                    <button onclick="location.reload()" class="text-[9px] font-black bg-blue-50 text-blue-600 px-3 py-1.5 rounded-lg border border-blue-100 uppercase">Làm mới</button>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                        <label class="text-[9px] font-bold text-slate-400 uppercase mb-1 block ml-1">Tên Dự Án</label>
                        <input type="text" id="job-name" placeholder="Nhập tên dự án..." class="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold outline-none focus:border-blue-500">
                    </div>
                    <div class="bg-blue-50/50 p-3 rounded-2xl border border-blue-100">
                        <span class="text-[9px] font-black text-blue-600 uppercase block mb-2">Onset</span>
                        <input type="date" id="onset-date" class="w-full p-2 bg-white border border-slate-200 rounded-lg text-xs font-bold mb-2">
                        <div class="flex items-center gap-2">
                            <input type="number" id="onset-hour" placeholder="Giờ" class="w-1/2 p-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-center">
                            <input type="number" id="onset-min" placeholder="Phút" class="w-1/2 p-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-center">
                        </div>
                    </div>
                    <div class="bg-orange-50/50 p-3 rounded-2xl border border-orange-100">
                        <span class="text-[9px] font-black text-orange-600 uppercase block mb-2">Wrap</span>
                        <input type="date" id="wrap-date" class="w-full p-2 bg-white border border-slate-200 rounded-lg text-xs font-bold mb-2">
                        <div class="flex items-center gap-2">
                            <input type="number" id="wrap-hour" placeholder="Giờ" class="w-1/2 p-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-center">
                            <input type="number" id="wrap-min" placeholder="Phút" class="w-1/2 p-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-center">
                        </div>
                    </div>
                </div>
            </div>

            <div class="no-print bg-white rounded-2xl shadow-sm p-4 border border-slate-200">
                <h2 class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-4">Thành viên đi làm</h2>
                <div id="staff-selection-grid" class="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-2 mb-6"></div>
                <button onclick="calculate()" class="w-full bg-slate-900 text-white font-black py-4 rounded-xl text-[10px] uppercase tracking-widest active:scale-95 transition-all">Tính công ngay</button>
            </div>

            <div id="result-container" class="space-y-4">
                <div id="result-card" class="bg-white rounded-2xl shadow-sm border border-slate-200 hidden flex flex-col overflow-hidden animate-fade-in">
                    <div class="p-4 border-b border-slate-100">
                        <h2 id="res-job-display" class="text-lg font-black text-slate-800 uppercase text-center">BẢNG KÊ NHÂN SỰ</h2>
                        <div id="res-time-info" class="text-center text-[10px] font-bold text-slate-500 mt-1"></div>
                    </div>
                    <div class="overflow-x-auto p-2">
                        <table class="w-full text-left">
                            <thead class="text-[8px] font-black text-slate-400 uppercase">
                                <tr>
                                    <th class="px-2 py-2">Tên</th>
                                    <th class="px-2 py-2 text-center">Hệ số</th>
                                    <th class="px-2 py-2 text-right">Lương</th>
                                </tr>
                            </thead>
                            <tbody id="salary-table-body" class="text-[10px]"></tbody>
                        </table>
                    </div>
                    <div class="p-4 bg-slate-900 flex justify-between items-center">
                        <div>
                            <p class="text-[8px] font-black text-slate-500 uppercase">Tổng quỹ lương</p>
                            <p id="total-group-salary" class="text-xl font-black text-blue-400">0 đ</p>
                        </div>
                        <button onclick="window.print()" class="bg-white/10 text-white px-6 py-3 rounded-xl font-black text-[9px] uppercase">In / Xuất PDF</button>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 2: NHÂN SỰ -->
        <div id="tab-staff" class="hidden animate-fade-in space-y-4">
            <div class="bg-white rounded-2xl shadow-sm p-4 border border-slate-200">
                <h2 class="text-xs font-black text-slate-800 uppercase mb-3">Nhập danh sách từ Excel/Note</h2>
                <textarea id="excel-paste" placeholder="Dán danh sách tên tại đây (mỗi dòng 1 tên)..." class="w-full h-32 p-3 border border-slate-200 rounded-xl text-sm outline-none bg-slate-50 mb-3"></textarea>
                <button onclick="importNames()" class="w-full bg-blue-600 text-white font-black py-3 rounded-xl text-[10px] uppercase">XÁC NHẬN THÊM VÀO MÁY</button>
                
                <div class="mt-6">
                    <div class="flex justify-between items-center mb-3">
                        <h2 class="text-xs font-black text-slate-800 uppercase">Danh sách trong máy</h2>
                        <button onclick="clearAllStaff()" class="text-[8px] text-red-500 font-bold uppercase">Xóa hết</button>
                    </div>
                    <div id="staff-preview" class="space-y-2"></div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const DEFAULT_SALARY = 900000;
        let staffList = JSON.parse(localStorage.getItem('staff_v7')) || [];
        let selectedStaff = new Set();

        function switchTab(tab) {
            document.getElementById('tab-calc').style.display = (tab === 'calc') ? 'block' : 'none';
            document.getElementById('tab-staff').style.display = (tab === 'staff') ? 'block' : 'none';
            if (tab === 'calc') renderSelectionGrid();
            if (tab === 'staff') renderStaffSettings();
        }

        function importNames() {
            const input = document.getElementById('excel-paste').value;
            const names = input.split('\n').map(n => n.trim()).filter(n => n.length > 0);
            if(names.length === 0) { alert("Vui lòng dán danh sách!"); return; }
            
            names.forEach(name => {
                if (!staffList.some(s => s.name === name)) {
                    staffList.push({ name: name, baseSalary: DEFAULT_SALARY });
                }
            });
            localStorage.setItem('staff_v7', JSON.stringify(staffList));
            document.getElementById('excel-paste').value = "";
            renderStaffSettings();
            alert("Đã thêm thành công!");
        }

        function renderStaffSettings() {
            const container = document.getElementById('staff-preview');
            container.innerHTML = staffList.map((s, i) => `
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-100 flex items-center justify-between">
                    <span class="text-[10px] font-black uppercase">${s.name}</span>
                    <div class="flex items-center gap-2">
                        <input type="number" value="${s.baseSalary}" onchange="updateSalary(${i}, this.value)" class="w-20 bg-white border border-slate-200 rounded px-2 py-1 text-[10px] font-bold">
                        <button onclick="deleteStaff(${i})" class="text-red-500 font-bold px-2">×</button>
                    </div>
                </div>
            `).join('');
        }

        function updateSalary(idx, val) { staffList[idx].baseSalary = parseInt(val); localStorage.setItem('staff_v7', JSON.stringify(staffList)); }
        function deleteStaff(idx) { staffList.splice(idx, 1); localStorage.setItem('staff_v7', JSON.stringify(staffList)); renderStaffSettings(); }
        function clearAllStaff() { if(confirm("Xóa toàn bộ?")) { staffList = []; localStorage.removeItem('staff_v7'); renderStaffSettings(); } }

        function renderSelectionGrid() {
            const container = document.getElementById('staff-selection-grid');
            if(staffList.length === 0) { container.innerHTML = "<p class='col-span-full text-center text-slate-400 text-[10px] py-4'>Chưa có nhân sự. Qua tab NHÂN SỰ để thêm.</p>"; return; }
            container.innerHTML = staffList.map((s, i) => `
                <div>
                    <input type="checkbox" id="check-${i}" class="hidden staff-checkbox" onchange="toggleSelection('${s.name}')" ${selectedStaff.has(s.name)?'checked':''}>
                    <label for="check-${i}" class="flex flex-col p-2 border border-slate-200 rounded-xl cursor-pointer text-center bg-white h-12 justify-center">
                        <span class="text-[9px] font-black truncate uppercase">${s.name}</span>
                    </label>
                </div>
            `).join('');
        }

        function toggleSelection(name) { if(selectedStaff.has(name)) selectedStaff.delete(name); else selectedStaff.add(name); }

        function calculate() {
            const job = document.getElementById('job-name').value || "Show mới";
            const onD = document.getElementById('onset-date').value, wrD = document.getElementById('wrap-date').value;
            const onH = parseInt(document.getElementById('onset-hour').value)||0, onM = parseInt(document.getElementById('onset-min').value)||0;
            const wrH = parseInt(document.getElementById('wrap-hour').value)||0, wrM = parseInt(document.getElementById('wrap-min').value)||0;

            if(!onD || !wrD) { alert("Vui lòng nhập ngày!"); return; }
            const start = new Date(`${onD}T${onH.toString().padStart(2,'0')}:${onM.toString().padStart(2,'0')}`);
            const end = new Date(`${wrD}T${wrH.toString().padStart(2,'0')}:${wrM.toString().padStart(2,'0')}`);
            
            const diffMin = (end - start) / 60000;
            let factor = 1.0;
            if (diffMin > 1090) factor = 2.0; // Over 18h 10p
            else if (diffMin > 930) factor = 1.5; // Over 15h 30p
            else if (diffMin > 910) factor = 1.25; // Over 15h 10p

            const tbody = document.getElementById('salary-table-body');
            let totalMoney = 0;
            
            const results = Array.from(selectedStaff).map(name => {
                const s = staffList.find(x => x.name === name);
                const salary = s ? s.baseSalary * factor : DEFAULT_SALARY * factor;
                totalMoney += salary;
                return `<tr class="border-b border-slate-50"><td class="p-2 font-bold uppercase">${name}</td><td class="text-center font-black">${factor}</td><td class="text-right font-black text-blue-600">${new Intl.NumberFormat('vi-VN').format(salary)}</td></tr>`;
            });

            if(results.length === 0) { alert("Hãy chọn ít nhất 1 người!"); return; }

            document.getElementById('result-card').classList.remove('hidden');
            document.getElementById('res-job-display').innerText = job;
            document.getElementById('res-time-info').innerText = `Giờ quay: ${Math.floor(diffMin/60)}h ${Math.floor(diffMin%60)}p | Hệ số: ${factor}`;
            tbody.innerHTML = results.join('');
            document.getElementById('total-group-salary').innerText = new Intl.NumberFormat('vi-VN').format(totalMoney) + ' đ';
            window.scrollTo({top: document.getElementById('result-card').offsetTop, behavior: 'smooth'});
        }

        window.onload = () => { 
            const now = new Date().toISOString().split('T')[0];
            document.getElementById('onset-date').value = now;
            document.getElementById('wrap-date').value = now;
            renderSelectionGrid(); 
        };
    </script>
</body>
</html>
