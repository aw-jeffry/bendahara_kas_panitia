<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BENDAHARA - Sistem Pencatatan Keuangan</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <script>
        tailwind.config = { theme: { extend: { colors: { primary: '#0ea5e9', success: '#10b981', danger: '#ef4444' } } } }
    </script>
    <style>
        .view-section { display: none; }
        .view-section.active { display: block; }
        .custom-scrollbar::-webkit-scrollbar { height: 8px; width: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .nav-btn.active { background-color: #eff6ff; color: #0ea5e9; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans h-screen flex overflow-hidden">

    <aside class="w-64 bg-white shadow-xl hidden md:flex flex-col z-10">
        <div class="p-6 border-b border-slate-100"><h1 class="text-xl font-bold text-primary">BENDAHARA</h1></div>
        <nav class="flex-1 p-4 space-y-2">
            <button onclick="switchView('dashboard')" id="nav-dashboard" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl font-medium text-slate-600 hover:bg-blue-50 hover:text-primary"><i class="fa-solid fa-chart-pie"></i> Dashboard</button>
            <button onclick="switchView('transactions')" id="nav-transactions" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl font-medium text-slate-600 hover:bg-blue-50 hover:text-primary"><i class="fa-solid fa-money-bill-transfer"></i> Transaksi</button>
            <button onclick="switchView('budgets')" id="nav-budgets" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl font-medium text-slate-600 hover:bg-blue-50 hover:text-primary"><i class="fa-solid fa-bullseye"></i> Anggaran</button>
            <button onclick="switchView('reports')" id="nav-reports" class="nav-btn w-full flex items-center gap-3 px-4 py-3 rounded-xl font-medium text-slate-600 hover:bg-blue-50 hover:text-primary"><i class="fa-solid fa-chart-line"></i> Laporan</button>
        </nav>
    </aside>

    <main class="flex-1 flex flex-col h-screen overflow-hidden">
        <div class="flex-1 overflow-y-auto p-4 md:p-8 custom-scrollbar">
            
            <section id="view-dashboard" class="view-section active">
                <h2 class="text-2xl font-bold mb-6">Dashboard</h2>
                
                <div id="dashboard-image-container" class="mb-6 rounded-2xl overflow-hidden shadow-sm border border-slate-100 hidden h-48 md:h-64">
                    <img id="dashboard-image" src="" alt="Banner Dashboard" class="w-full h-full object-cover">
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                        <p class="text-sm text-slate-500">Sisa Dana</p>
                        <h3 id="dash-balance" class="text-2xl font-bold text-primary">Rp 0</h3>
                    </div>
                    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                        <p class="text-sm text-slate-500">Total Pemasukan</p>
                        <h3 id="dash-income" class="text-2xl font-bold text-success">Rp 0</h3>
                    </div>
                    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                        <p class="text-sm text-slate-500">Total Pengeluaran</p>
                        <h3 id="dash-expense" class="text-2xl font-bold text-danger">Rp 0</h3>
                    </div>
                </div>
            </section>

            <section id="view-transactions" class="view-section">
                <div class="flex justify-between mb-6">
                    <h2 class="text-2xl font-bold">Transaksi</h2>
                    <button onclick="openModal()" class="bg-primary text-white px-4 py-2 rounded-lg font-bold"><i class="fa-solid fa-plus mr-2"></i> Tambah</button>
                </div>
                <div class="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
                    <table class="w-full text-left">
                        <thead class="bg-slate-50">
                            <tr>
                                <th class="p-4">Tanggal</th>
                                <th class="p-4">Tipe</th>
                                <th class="p-4">Kategori</th>
                                <th class="p-4">Rincian</th>
                                <th class="p-4">Jumlah</th>
                                <th class="p-4">Aksi</th>
                            </tr>
                        </thead>
                        <tbody id="tx-body">
                            <tr><td colspan="6" class="p-4 text-center text-slate-400">Memuat data...</td></tr>
                        </tbody>
                    </table>
                </div>
            </section>

            <section id="view-budgets" class="view-section">
                <h2 class="text-2xl font-bold mb-6">Anggaran per Kategori</h2>
                <div id="budget-body" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    <p class="text-slate-400 col-span-full">Memuat data anggaran...</p>
                </div>
            </section>

            <section id="view-reports" class="view-section">
                <div class="flex justify-between mb-6">
                    <h2 class="text-2xl font-bold">Laporan Terperinci</h2>
                    <button onclick="exportToExcel()" class="bg-emerald-600 text-white px-4 py-2 rounded-lg font-bold"><i class="fa-solid fa-file-excel mr-2"></i> Export Excel</button>
                </div>
                <div id="report-summary" class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
                    <div class="bg-white p-4 rounded-xl border border-emerald-100 shadow-sm">
                        <p class="text-xs text-slate-500">Total Pemasukan</p>
                        <p class="font-bold text-lg text-emerald-600" id="report-income">Rp 0</p>
                    </div>
                    <div class="bg-white p-4 rounded-xl border border-red-100 shadow-sm">
                        <p class="text-xs text-slate-500">Total Pengeluaran</p>
                        <p class="font-bold text-lg text-red-600" id="report-expense">Rp 0</p>
                    </div>
                    <div class="bg-white p-4 rounded-xl border border-blue-100 shadow-sm">
                        <p class="text-xs text-slate-500">Sisa Saldo</p>
                        <p class="font-bold text-lg text-blue-600" id="report-balance">Rp 0</p>
                    </div>
                </div>
                <div id="report-body" class="space-y-6">
                    <p class="text-slate-400">Memuat laporan...</p>
                </div>
            </section>
        </div>
    </main>

    <div id="modal" class="fixed inset-0 bg-black/50 hidden flex items-center justify-center z-50">
        <div class="bg-white p-8 rounded-2xl w-96 max-w-[90%]">
            <h3 id="modal-title" class="text-lg font-bold mb-4">Tambah Transaksi</h3>
            <form id="form-tx" onsubmit="saveTransaction(event)">
                <input type="hidden" id="edit-id">
                <label class="block text-sm font-medium mb-1">Tipe Transaksi</label>
                <select id="type" onchange="updateCategories()" class="w-full mb-4 p-2 border rounded-lg">
                    <option value="pemasukan">Pemasukan</option>
                    <option value="pengeluaran">Pengeluaran</option>
                </select>
                
                <label class="block text-sm font-medium mb-1">Tanggal</label>
                <input type="date" id="date" required class="w-full mb-4 p-2 border rounded-lg">
                
                <label class="block text-sm font-medium mb-1">Kategori</label>
                <select id="category" class="w-full mb-4 p-2 border rounded-lg"></select>
                
                <label class="block text-sm font-medium mb-1">Rincian</label>
                <input type="text" id="note" placeholder="Rincian transaksi" required class="w-full mb-4 p-2 border rounded-lg">
                
                <label class="block text-sm font-medium mb-1">Jumlah (Rp)</label>
                <input type="text" id="amount" placeholder="0" required class="w-full mb-4 p-2 border rounded-lg" onkeyup="formatInput(this)">
                
                <div class="flex gap-2">
                    <button type="button" onclick="closeModal()" class="w-full bg-gray-200 p-2 rounded-lg hover:bg-gray-300">Batal</button>
                    <button type="submit" class="w-full bg-primary text-white p-2 rounded-lg hover:bg-sky-600">Simpan</button>
                </div>
            </form>
        </div>
    </div>

    <div id="budget-modal" class="fixed inset-0 bg-black/50 hidden flex items-center justify-center z-50">
        <div class="bg-white p-8 rounded-2xl w-80 max-w-[90%]">
            <h3 class="text-lg font-bold mb-4" id="budget-title-modal">Edit Anggaran</h3>
            <input type="hidden" id="budget-cat-edit">
            <label class="block text-sm font-medium mb-1">Jumlah Anggaran</label>
            <input type="text" id="budget-val-edit" placeholder="0" class="w-full mb-4 p-2 border rounded-lg" onkeyup="formatInput(this)">
            <div class="flex gap-2">
                <button onclick="closeBudgetModal()" class="w-full bg-gray-200 p-2 rounded-lg hover:bg-gray-300">Batal</button>
                <button onclick="saveBudgetFromModal()" class="w-full bg-primary text-white p-2 rounded-lg hover:bg-sky-600">Simpan</button>
            </div>
        </div>
    </div>

    <div id="loading" class="fixed inset-0 bg-black/30 hidden flex items-center justify-center z-50">
        <div class="bg-white p-6 rounded-xl shadow-lg">
            <i class="fa-solid fa-spinner fa-spin text-3xl text-primary"></i>
            <p class="mt-2 text-sm">Memproses...</p>
        </div>
    </div>

    <script>
        let appSettings = {};
        let transactions = [];
        let budgets = {};
        let categories = { pemasukan: [], pengeluaran: [] };
        
        // Format Rupiah
        function formatRupiah(num) { 
            return 'Rp ' + num.toLocaleString('id-ID'); 
        }
        
        // Format tanggal untuk tampilan DD/MM/YY
        function formatDateDisplay(dateStr) { 
            const date = new Date(dateStr);
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = String(date.getFullYear()).slice(-2);
            return `${day}/${month}/${year}`;
        }

        // Format input angka
        function formatInput(el) { 
            let v = el.value.replace(/\D/g, ''); 
            el.value = v ? parseInt(v).toLocaleString('id-ID') : ''; 
        }

        // Update dropdown kategori berdasarkan tipe
        function updateCategories() {
            const type = document.getElementById('type').value;
            const cats = categories[type] || [];
            document.getElementById('category').innerHTML = cats.map(c => `<option value="${c}">${c}</option>`).join('');
        }

        // Show & Hide loading
        function showLoading() { document.getElementById('loading').classList.remove('hidden'); }
        function hideLoading() { document.getElementById('loading').classList.add('hidden'); }

        // Switch view
        function switchView(id) { 
            document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active')); 
            document.getElementById('view-'+id).classList.add('active');
            document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
            document.getElementById('nav-'+id).classList.add('active');
        }

        // Load semua data dari server
        function loadAllData() {
            showLoading();
            google.script.run
                .withSuccessHandler(function(data) {
                    appSettings = data.settings || {}; 
                    
                    transactions = data.transaksi.map(t => ({
                        id: t.id,
                        date: t.tanggal,
                        type: t.tipe.toLowerCase(),
                        category: t.kategori,
                        note: t.deskripsi,
                        amount: t.jumlah
                    }));
                    
                    budgets = {};
                    if (data.anggaran) {
                        data.anggaran.forEach(b => budgets[b.kategori] = b.anggaran);
                    }
                    
                    categories = { 
                        pemasukan: data.kategori['Pemasukan'] || [], 
                        pengeluaran: data.kategori['Pengeluaran'] || [] 
                    };
                    
                    renderAll();
                    hideLoading();
                })
                .withFailureHandler(function(err) {
                    alert('Gagal memuat data: ' + err.message);
                    hideLoading();
                })
                .getData();
        }

        // Render semua tampilan
        function renderAll() {
            renderDashboard();
            renderTransactions();
            renderBudgets();
            renderReports();
        }

        // Render Dashboard
        function renderDashboard() {
            let inc = transactions.filter(t => t.type == 'pemasukan').reduce((a,b) => a + b.amount, 0);
            let exp = transactions.filter(t => t.type == 'pengeluaran').reduce((a,b) => a + b.amount, 0);
            document.getElementById('dash-balance').innerText = formatRupiah(inc - exp);
            document.getElementById('dash-income').innerText = formatRupiah(inc);
            document.getElementById('dash-expense').innerText = formatRupiah(exp);

            const imgContainer = document.getElementById('dashboard-image-container');
            const imgElement = document.getElementById('dashboard-image');
            
            if (appSettings.logoUrl && appSettings.logoUrl.trim() !== '') {
                imgElement.src = appSettings.logoUrl;
                imgContainer.classList.remove('hidden');
            } else {
                imgContainer.classList.add('hidden');
            }
        }

        // Render Tabel Transaksi
        function renderTransactions() {
            const tbody = document.getElementById('tx-body');
            if (transactions.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" class="p-4 text-center text-slate-400">Belum ada transaksi</td></tr>';
                return;
            }
            
            tbody.innerHTML = transactions.map(t => `
                <tr class="border-b hover:bg-slate-50">
                    <td class="p-4">${formatDateDisplay(t.date)}</td>
                    <td class="p-4">
                        <span class="px-2 py-1 rounded text-xs font-bold ${t.type=='pemasukan'?'bg-emerald-100 text-emerald-800':'bg-red-100 text-red-800'}">
                            ${t.type.toUpperCase()}
                        </span>
                    </td>
                    <td class="p-4">${t.category}</td>
                    <td class="p-4">${t.note}</td>
                    <td class="p-4 font-bold">${formatRupiah(t.amount)}</td>
                    <td class="p-4">
                        <button onclick="editTransaction(${t.id})" class="text-blue-500 mr-2 hover:underline"><i class="fa-solid fa-edit"></i></button>
                        <button onclick="deleteTx(${t.id})" class="text-red-500 hover:underline"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `).join('');
        }

        // Render Anggaran
        function renderBudgets() {
            const container = document.getElementById('budget-body');
            const cats = Object.keys(budgets);
            
            if (cats.length === 0) {
                container.innerHTML = '<p class="text-slate-400 col-span-full">Belum ada data anggaran</p>';
                return;
            }
            
            container.innerHTML = cats.map(c => {
                let spent = transactions.filter(t => t.category == c && t.type == 'pengeluaran').reduce((a,b) => a + b.amount, 0);
                let sisa = budgets[c] - spent;
                let sisaClass = sisa >= 0 ? 'text-emerald-600' : 'text-red-600';
                
                return `
                    <div class="p-4 bg-white rounded-xl shadow-sm border border-slate-100">
                        <h4 class="font-bold text-slate-700 mb-2">${c}</h4>
                        <p class="text-sm">Anggaran: <span class="font-bold">${formatRupiah(budgets[c])}</span></p>
                        <p class="text-sm text-red-500">Terpakai: ${formatRupiah(spent)}</p>
                        <p class="text-sm ${sisaClass} font-medium">Sisa: ${formatRupiah(sisa)}</p>
                        <button onclick="editBudget('${c}')" class="mt-3 text-xs bg-gray-100 px-3 py-1.5 rounded-lg hover:bg-gray-200">
                            <i class="fa-solid fa-edit mr-1"></i> Edit Anggaran
                        </button>
                    </div>
                `;
            }).join('');
        }

        // Render Laporan
        function renderReports() {
            let inc = transactions.filter(t => t.type == 'pemasukan').reduce((a,b) => a + b.amount, 0);
            let exp = transactions.filter(t => t.type == 'pengeluaran').reduce((a,b) => a + b.amount, 0);
            
            document.getElementById('report-income').innerText = formatRupiah(inc);
            document.getElementById('report-expense').innerText = formatRupiah(exp);
            document.getElementById('report-balance').innerText = formatRupiah(inc - exp);

            const container = document.getElementById('report-body');
            
            if (transactions.length === 0) {
                container.innerHTML = '<p class="text-slate-400">Belum ada transaksi</p>';
                return;
            }
            
            const sorted = [...transactions].sort((a,b) => new Date(a.date) - new Date(b.date));
            const cats = { pemasukan: {}, pengeluaran: {} };
            
            sorted.forEach(t => {
                if(!cats[t.type][t.category]) cats[t.type][t.category] = [];
                cats[t.type][t.category].push(t);
            });

            container.innerHTML = ['pemasukan', 'pengeluaran'].map(type => {
                const typeCats = Object.keys(cats[type]);
                if (typeCats.length === 0) return '';
                
                return `
                    <h3 class="font-bold text-lg capitalize mt-8 mb-4 border-b pb-2">${type}</h3>
                    ${typeCats.map(cat => {
                        const subTotal = cats[type][cat].reduce((sum, t) => sum + t.amount, 0);
                        return `
                            <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-100 mb-4">
                                <h4 class="font-bold text-primary mb-2">${cat}</h4>
                                <table class="w-full text-sm">
                                    ${cats[type][cat].map(t => `
                                        <tr class="border-t">
                                            <td class="p-2 w-24">${formatDateDisplay(t.date)}</td>
                                            <td class="p-2">${t.note}</td>
                                            <td class="p-2 text-right font-medium">${formatRupiah(t.amount)}</td>
                                        </tr>
                                    `).join('')}
                                    <tr class="border-t-2 font-bold bg-slate-50">
                                        <td colspan="2" class="p-2 text-right">Total per Kategori:</td>
                                        <td class="p-2 text-right">${formatRupiah(subTotal)}</td>
                                    </tr>
                                </table>
                            </div>
                        `;
                    }).join('')}
                `;
            }).join('');
        }

        // Modal Transaksi
        function openModal() {
            document.getElementById('edit-id').value = '';
            document.getElementById('modal-title').innerText = 'Tambah Transaksi';
            document.getElementById('date').valueAsDate = new Date();
            document.getElementById('amount').value = '';
            document.getElementById('note').value = '';
            updateCategories();
            document.getElementById('modal').classList.remove('hidden');
        }

        function closeModal() {
            document.getElementById('modal').classList.add('hidden');
        }

        function editTransaction(id) {
            const t = transactions.find(x => x.id == id);
            if (!t) return;
            
            document.getElementById('edit-id').value = t.id;
            document.getElementById('modal-title').innerText = 'Edit Transaksi';
            document.getElementById('type').value = t.type;
            updateCategories();
            document.getElementById('date').value = t.date;
            document.getElementById('category').value = t.category;
            document.getElementById('note').value = t.note;
            document.getElementById('amount').value = t.amount.toLocaleString('id-ID');
            document.getElementById('modal').classList.remove('hidden');
        }

        // Simpan Transaksi
        function saveTransaction(e) {
            e.preventDefault();
            showLoading();
            
            const id = document.getElementById('edit-id').value;
            const amount = parseInt(document.getElementById('amount').value.replace(/\D/g, ''));
            const typeVal = document.getElementById('type').value;
            
            const dataToSend = {
                id: id ? parseInt(id) : Date.now(),
                tipe: typeVal === 'pemasukan' ? 'Pemasukan' : 'Pengeluaran',
                tanggal: document.getElementById('date').value,
                kategori: document.getElementById('category').value,
                deskripsi: document.getElementById('note').value,
                jumlah: amount
            };
            
            if (id) {
                google.script.run
                    .withSuccessHandler(function(res) {
                        hideLoading();
                        closeModal();
                        if (res.success) loadAllData(); else alert(res.message);
                    })
                    .withFailureHandler(function(err) { hideLoading(); alert('Gagal update: ' + err.message); })
                    .updateTransaksi(dataToSend);
            } else {
                google.script.run
                    .withSuccessHandler(function(res) {
                        hideLoading();
                        closeModal();
                        if (res.success) loadAllData(); else alert(res.message);
                    })
                    .withFailureHandler(function(err) { hideLoading(); alert('Gagal simpan: ' + err.message); })
                    .addTransaksi(dataToSend);
            }
        }

        // Hapus Transaksi
        function deleteTx(id) {
            if (!confirm('Yakin ingin menghapus transaksi ini?')) return;
            showLoading();
            google.script.run
                .withSuccessHandler(function(res) {
                    hideLoading();
                    if (res.success) loadAllData(); else alert(res.message);
                })
                .withFailureHandler(function(err) { hideLoading(); alert('Gagal hapus: ' + err.message); })
                .deleteTransaksi(id);
        }

        // Modal Anggaran
        function editBudget(c) {
            document.getElementById('budget-cat-edit').value = c;
            document.getElementById('budget-title-modal').innerText = 'Edit Anggaran: ' + c;
            document.getElementById('budget-val-edit').value = budgets[c].toLocaleString('id-ID');
            document.getElementById('budget-modal').classList.remove('hidden');
        }

        function closeBudgetModal() {
            document.getElementById('budget-modal').classList.add('hidden');
        }

        // Simpan Anggaran
        function saveBudgetFromModal() {
            const cat = document.getElementById('budget-cat-edit').value;
            const val = parseInt(document.getElementById('budget-val-edit').value.replace(/\D/g, '')); 
            
            showLoading();
            google.script.run
                .withSuccessHandler(function(res) {
                    hideLoading();
                    closeBudgetModal();
                    if (res.success) loadAllData(); else alert(res.message);
                })
                .withFailureHandler(function(err) { hideLoading(); alert('Gagal update anggaran: ' + err.message); })
                .updateAnggaran(cat, val);
        }

        // Export Excel
        function exportToExcel() {
            let inc = transactions.filter(t => t.type == 'pemasukan').reduce((a,b) => a + b.amount, 0);
            let exp = transactions.filter(t => t.type == 'pengeluaran').reduce((a,b) => a + b.amount, 0);

            const wb = XLSX.utils.book_new();
            
            const summary = [
                { Item: "Total Pemasukan", Jumlah: inc },
                { Item: "Total Pengeluaran", Jumlah: exp },
                { Item: "Sisa Saldo", Jumlah: (inc - exp) }
            ];
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Ringkasan");
            
            ['pemasukan', 'pengeluaran'].forEach(type => {
                let exportData = [];
                const typeCats = [...new Set(transactions.filter(t => t.type == type).map(t => t.category))];
                let grandTotal = 0;

                typeCats.forEach(cat => {
                    const catTxs = transactions.filter(t => t.type == type && t.category == cat);
                    let catTotal = 0;
                    catTxs.forEach(t => {
                        exportData.push({
                            "Kategori": cat,
                            "Tanggal": formatDateDisplay(t.date),
                            "Rincian": t.note,
                            "Jumlah": t.amount
                        });
                        catTotal += t.amount;
                    });
                    exportData.push({ "Kategori": "TOTAL " + cat.toUpperCase(), "Tanggal": "", "Rincian": "", "Jumlah": catTotal });
                    grandTotal += catTotal;
                });
                
                exportData.push({ "Kategori": "TOTAL KESELURUHAN", "Tanggal": "", "Rincian": "", "Jumlah": grandTotal });
                XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportData), type.charAt(0).toUpperCase() + type.slice(1));
            });
            
            XLSX.writeFile(wb, "Laporan_Keuangan.xlsx");
        }

        // Load data saat halaman dimuat
        window.onload = loadAllData;
    </script>
</body>
</html>
