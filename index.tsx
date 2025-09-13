// NOVO: Aviso para o TypeScript sobre a variável global do Firebase
declare const firebase: any;

// --- Type Definitions for external libraries ---
declare const XLSX: any;
declare const Chart: any;
declare const ChartDataLabels: any;
declare const jspdf: any;
declare const html2canvas: any;

// NOVO: Bloco de configuração e inicialização do Firebase
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();


// --- DOM Elements ---
const fileUpload = document.getElementById('file-upload') as HTMLInputElement;
const dashboardGrid = document.getElementById('dashboard-grid') as HTMLElement;
const lastUpdate = document.getElementById('last-update') as HTMLElement;
const placeholder = document.getElementById('placeholder') as HTMLElement;
const filterContainer = document.getElementById('filter-container') as HTMLElement;
const chartsContainer = document.getElementById('charts-container') as HTMLElement;
const applyFiltersBtn = document.getElementById('apply-filters-btn') as HTMLButtonElement;
const resetFiltersBtn = document.getElementById('reset-filters-btn') as HTMLButtonElement;
const totalFclDisplay = document.getElementById('total-fcl-display') as HTMLElement;
const totalFclCount = document.getElementById('total-fcl-count') as HTMLElement;

// View Tabs
const viewTabsContainer = document.getElementById('view-tabs-container') as HTMLElement;
const viewVesselBtn = document.getElementById('view-vessel-btn') as HTMLButtonElement;
const viewPoBtn = document.getElementById('view-po-btn') as HTMLButtonElement;
const viewWarehouseBtn = document.getElementById('view-warehouse-btn') as HTMLButtonElement;

// Filter Inputs
const arrivalStartDate = document.getElementById('arrival-start-date') as HTMLInputElement;
const arrivalEndDate = document.getElementById('arrival-end-date') as HTMLInputElement;
const deadlineStartDate = document.getElementById('deadline-start-date') as HTMLInputElement;
const deadlineEndDate = document.getElementById('deadline-end-date') as HTMLInputElement;
const statusFilter = document.getElementById('status-filter') as HTMLSelectElement;
const shipmentTypeFilter = document.getElementById('shipment-type-filter') as HTMLSelectElement;
const cargoTypeFilter = document.getElementById('cargo-type-filter') as HTMLSelectElement;
const poSearchInput = document.getElementById('po-search-input') as HTMLInputElement;
const vesselSearchInput = document.getElementById('vessel-search-input') as HTMLInputElement;
const poFilter = document.getElementById('po-filter') as HTMLSelectElement;
const vesselFilter = document.getElementById('vessel-filter') as HTMLSelectElement;


// Export Buttons
const exportCsvBtn = document.getElementById('export-csv-btn') as HTMLButtonElement;
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;
const exportExcelBtn = document.getElementById('export-excel-btn') as HTMLButtonElement;


// Loading Overlay
const loadingOverlay = document.getElementById('loading-overlay') as HTMLElement;

// Column Visibility
const columnToggleBtn = document.getElementById('column-toggle-btn') as HTMLButtonElement;
const columnToggleDropdown = document.getElementById('column-toggle-dropdown') as HTMLElement;

// Theme Toggle Buttons
const darkModeBtn = document.getElementById('dark-mode-btn') as HTMLButtonElement;
const lightModeBtn = document.getElementById('light-mode-btn') as HTMLButtonElement;

// Modal Elements
const detailsModal = document.getElementById('details-modal') as HTMLElement;
const modalContent = document.getElementById('modal-content') as HTMLElement;
const modalHeaderContent = document.getElementById('modal-header-content') as HTMLElement;
const modalBody = document.getElementById('modal-body') as HTMLElement;
const modalCloseBtn = document.getElementById('modal-close-btn') as HTMLButtonElement;

// Logo Elements
const companyLogo = document.getElementById('company-logo') as HTMLImageElement;
const logoUpload = document.getElementById('logo-upload') as HTMLInputElement;
const removeLogoBtn = document.getElementById('remove-logo-btn') as HTMLButtonElement;


// --- Global State ---
let originalData: any[] = [];
let filteredDataCache: any[] = [];
let chartDataSource: {
    bar: any[];
    pie: { label: string; value: number }[];
    deadline: { label: string; count: number }[];
} = { bar: [], pie: [], deadline: [] };
let mainBarChart: any = null;
let statusPieChart: any = null;
let deadlineDistributionChart: any = null;
let currentView: 'vessel' | 'po' | 'warehouse' = 'vessel';
const TODAY = new Date(); // Use current date
let currentSortKey: string = 'Dias Restantes';
let currentSortOrder: 'asc' | 'desc' = 'asc';
let activeModalItem: any = null;
let isUpdatingFromFirebase = false; // NOVO: Flag para evitar loops

// Column definitions for each view
const viewColumns = {
    vessel: ['SAPCargoPO', 'BL/AWB', 'Carrier', 'Status', 'ShipmentType', 'CargoType', 'Arrival', 'FreeTimeDeadline', 'Dias Restantes', 'Warehouse'],
    po: ['Vessel', 'BL/AWB', 'Carrier', 'Status', 'ShipmentType', 'CargoType', 'Arrival', 'FreeTimeDeadline', 'Dias Restantes', 'Warehouse'],
    warehouse: ['Vessel', 'SAPCargoPO', 'BL/AWB', 'Carrier', 'Status', 'ShipmentType', 'CargoType', 'Arrival', 'FreeTimeDeadline', 'Dias Restantes']
};
let columnVisibility: Record<string, boolean> = {};


// --- Main App Initialization ---
window.addEventListener('load', () => {
    initializeApp();
});

function initializeApp() {
    // Defensive check for external libraries to prevent race conditions
    if (typeof Chart === 'undefined' || typeof ChartDataLabels === 'undefined' || typeof jspdf === 'undefined' || typeof html2canvas === 'undefined' || typeof XLSX === 'undefined') {
        console.warn("External libraries not yet loaded, retrying in 100ms...");
        setTimeout(initializeApp, 100);
        return;
    }

    // Initialize Theme
    const savedTheme = localStorage.getItem('theme');
    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const initialTheme = savedTheme || (prefersDark ? 'dark' : 'light');
    setTheme(initialTheme as 'dark' | 'light');


    // Event Listeners
    fileUpload.addEventListener('change', handleFileUpload);
    applyFiltersBtn.addEventListener('click', applyFiltersAndRender);
    resetFiltersBtn.addEventListener('click', resetFiltersAndRender);
    viewVesselBtn.addEventListener('click', () => setView('vessel'));
    viewPoBtn.addEventListener('click', () => setView('po'));
    viewWarehouseBtn.addEventListener('click', () => setView('warehouse'));
    exportCsvBtn.addEventListener('click', handleGlobalCsvExport);
    exportPdfBtn.addEventListener('click', handlePdfExport);
    exportExcelBtn.addEventListener('click', handleExcelExport);
    darkModeBtn.addEventListener('click', () => setTheme('dark'));
    lightModeBtn.addEventListener('click', () => setTheme('light'));
    poSearchInput.addEventListener('input', applyFiltersAndRender);
    vesselSearchInput.addEventListener('input', applyFiltersAndRender);
    logoUpload.addEventListener('change', handleLogoUpload);
    removeLogoBtn.addEventListener('click', handleRemoveLogo);


    dashboardGrid.addEventListener('click', handleSortClick);
    setupColumnToggles();
    loadSavedLogo();

    // Modal Listeners
    modalCloseBtn.addEventListener('click', closeModal);
    detailsModal.addEventListener('click', (e) => {
        if (e.target === detailsModal) { // Close if clicking on the overlay
            closeModal();
        }
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && !detailsModal.classList.contains('hidden')) {
            closeModal();
        }
    });

    // NOVO: Inicia a escuta por dados do Firebase
    escutarMudancasEmTempoReal();
}

// NOVO: --- Firebase Integration ---
async function salvarDados(dataToSave: any[]) {
    if (isUpdatingFromFirebase) return; // Evita salvar dados que acabaram de chegar do Firebase
    console.log("Enviando dados para o Firebase...");
    try {
        await db.collection("dashboardImportacao").doc("dadosAtuais").set({
            dados: dataToSave,
            ultimaAtualizacao: new Date()
        });
        showToast("Planilha enviada com sucesso! Atualizando para todos...", "success");
    } catch (error) {
        console.error("Erro ao salvar dados no Firebase: ", error);
        showToast("Falha ao sincronizar com o servidor.", "error");
    }
}

function escutarMudancasEmTempoReal() {
    console.log("Iniciando ouvinte de dados do Firebase...");
    db.collection("dashboardImportacao").doc("dadosAtuais").onSnapshot(doc => {
        isUpdatingFromFirebase = true;
        console.log("Dados recebidos do Firebase!");

        if (doc.exists) {
            const data = doc.data();
            if (data && data.dados) {
                originalData = data.dados; // ATUALIZA A VARIÁVEL GLOBAL
                const ultimaAtualizacao = data.ultimaAtualizacao?.toDate();

                // Popula os filtros e renderiza o dashboard com os novos dados
                populateStatusFilter(originalData);
                populateShipmentTypeFilter(originalData);
                populateCargoTypeFilter(originalData);
                populatePoFilter(originalData);
                populateVesselFilter(originalData);
                applyFiltersAndRender();
                
                // Mostra os containers do dashboard
                filterContainer.classList.remove('hidden');
                chartsContainer.classList.remove('hidden');
                viewTabsContainer.classList.remove('hidden');
                exportCsvBtn.classList.remove('hidden');
                exportPdfBtn.classList.remove('hidden');
                exportExcelBtn.classList.remove('hidden');
                totalFclDisplay.classList.remove('hidden');

                if (ultimaAtualizacao) {
                    lastUpdate.textContent = `Dados sincronizados | Atualizado em: ${ultimaAtualizacao.toLocaleString('pt-BR')}`;
                }
                showToast('Dados atualizados em tempo real!', 'success');
            }
        } else {
            console.log("Nenhum dado no Firebase. Aguardando upload.");
            resetUI(); // Garante que a interface esteja limpa
        }
        setTimeout(() => { isUpdatingFromFirebase = false; }, 500);
    }, error => {
        console.error("Erro no ouvinte do Firebase: ", error);
        showToast("Conexão com o servidor perdida.", "error");
    });
}

// --- Theme Management ---
function setTheme(theme: 'dark' | 'light') {
    const isDark = theme === 'dark';
    
    // 1. Toggle class on body
    document.body.classList.toggle('dark', isDark);

    // 2. Toggle button visibility
    darkModeBtn.classList.toggle('hidden', isDark);
    lightModeBtn.classList.toggle('hidden', !isDark);

    // 3. Save to localStorage
    localStorage.setItem('theme', theme);

    // 4. Update Chart.js defaults for new charts
    const textColor = isDark ? '#d1d5db' : '#4b5563';
    const gridColor = isDark ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';
    Chart.defaults.color = textColor;
    Chart.defaults.borderColor = gridColor;

    // 5. If data is loaded, re-render to update existing charts and UI
    if (originalData.length > 0) {
        applyFiltersAndRender(); 
    }
}


// --- Loading Indicator ---
function showLoading() {
    loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    loadingOverlay.classList.add('hidden');
}


// --- Toast Notifications ---
function showToast(message: string, type: 'success' | 'error' | 'warning' = 'success') {
    const toastContainer = document.getElementById('toast-container');
    if (!toastContainer) return;

    const toast = document.createElement('div');
    const icons = { success: 'fa-check-circle', error: 'fa-times-circle', warning: 'fa-exclamation-triangle' };
    const colors = { success: 'bg-green-500', error: 'bg-red-500', warning: 'bg-yellow-500' };
    toast.className = `toast ${colors[type]} text-white py-3 px-5 rounded-lg shadow-xl flex items-center mb-2`;
    toast.innerHTML = `<i class="fas ${icons[type]} mr-3"></i> <p>${message}</p>`;
    toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 5000);
}

// --- File Handling ---
// MODIFICADO: A função agora é 'async' para esperar o salvamento no Firebase
async function handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    const uploadLabel = document.querySelector('label[for="file-upload"]');
    if (!file || !uploadLabel) return;

    uploadLabel.classList.add('opacity-50', 'cursor-not-allowed');
    uploadLabel.innerHTML = `<i class="fas fa-spinner fa-spin mr-2"></i> Processando...`;

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const workbook = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), { type: 'array' });
            const sheetName = "FUP - International Trade - D11";
            if (!workbook.Sheets[sheetName]) throw new Error(`Planilha "${sheetName}" não encontrada.`);
            
            const dataFromSheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: '' });
            
            if (dataFromSheet.length === 0) throw new Error("A planilha está vazia.");

            // MODIFICADO: Em vez de processar localmente, apenas salva no Firebase.
            // O ouvinte 'escutarMudancasEmTempoReal' vai cuidar de atualizar a tela para todos.
            await salvarDados(dataFromSheet);
            
        } catch (err: any) {
            showToast(err.message || 'Erro ao processar arquivo.', 'error');
            resetUI();
        } finally {
            uploadLabel.classList.remove('opacity-50', 'cursor-not-allowed');
            uploadLabel.innerHTML = `<i class="fas fa-upload mr-2"></i> Carregar XLSX`;
            fileUpload.value = '';
        }
    };
    reader.onerror = () => {
        showToast('Não foi possível ler o arquivo.', 'error');
        resetUI();
    };
    reader.readAsArrayBuffer(file);
}

// --- Data Processing & Filtering ---
function excelDateToJSDate(serial: any): Date | null {
    if (!serial) return null;
    if (typeof serial === 'string') {
        if (serial.match(/^\d{5}$/)) { // Looks like an excel serial number as a string
             serial = parseInt(serial, 10);
        } else if (serial.includes('/') || serial.includes('-') || serial.includes('.')) {
             // Attempt to parse common date formats, this might need refinement
             const date = new Date(serial.replace(/(\d{2})\.(\d{2})\.(\d{4})/, '$2/$1/$3')); // Handle DD.MM.YYYY
             return isNaN(date.getTime()) ? null : date;
        } else {
            return null;
        }
    }
    if (typeof serial !== 'number' || serial < 1) return null;
    const utc_days = Math.floor(serial - 25569);
    const date_info = new Date(utc_days * 86400 * 1000);
    return new Date(date_info.getTime() + (date_info.getTimezoneOffset() * 60 * 1000));
}

function applyFiltersAndRender() {
    showLoading();
    setTimeout(() => {
        let filteredData = [...originalData];
        
        // Date Filters
        const arrivalStart = arrivalStartDate.valueAsDate;
        const arrivalEnd = arrivalEndDate.valueAsDate;
        const deadlineStart = deadlineStartDate.valueAsDate;
        const deadlineEnd = deadlineEndDate.valueAsDate;

        if (arrivalStart) filteredData = filteredData.filter(row => {
            const arrivalDate = excelDateToJSDate(row.Arrival);
            return arrivalDate && arrivalDate >= arrivalStart;
        });
        if (arrivalEnd) filteredData = filteredData.filter(row => {
            const arrivalDate = excelDateToJSDate(row.Arrival);
            return arrivalDate && arrivalDate <= arrivalEnd;
        });
        if (deadlineStart) filteredData = filteredData.filter(row => {
            const deadlineDate = excelDateToJSDate(row.FreeTimeDeadline);
            return deadlineDate && deadlineDate >= deadlineStart;
        });
        if (deadlineEnd) filteredData = filteredData.filter(row => {
            const deadlineDate = excelDateToJSDate(row.FreeTimeDeadline);
            return deadlineDate && deadlineDate <= deadlineEnd;
        });

        // PO Search Filter
        const poSearchTerm = poSearchInput.value.trim().toLowerCase();
        if (poSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row.SAPCargoPO || '').toLowerCase().includes(poSearchTerm);
            });
        }
        
        // PO Filter (multi-select)
        const selectedPOs = Array.from(poFilter.selectedOptions).map(opt => opt.value);
        if (selectedPOs.length > 0 && !selectedPOs.includes('')) {
            filteredData = filteredData.filter(row => selectedPOs.includes(row.SAPCargoPO || 'Sem PO'));
        }

        // Vessel Search Filter
        const vesselSearchTerm = vesselSearchInput.value.trim().toLowerCase();
        if (vesselSearchTerm) {
            filteredData = filteredData.filter(row => {
                return String(row.Vessel || '').toLowerCase().includes(vesselSearchTerm);
            });
        }
        
        // Vessel Filter (multi-select)
        const selectedVessels = Array.from(vesselFilter.selectedOptions).map(opt => opt.value);
        if (selectedVessels.length > 0 && !selectedVessels.includes('')) {
            filteredData = filteredData.filter(row => selectedVessels.includes(row.Vessel || 'Sem Navio'));
        }

        // Status Filter (multi-select)
        const selectedStatuses = Array.from(statusFilter.selectedOptions).map(opt => opt.value);
        if (selectedStatuses.length > 0 && !selectedStatuses.includes('')) {
            filteredData = filteredData.filter(row => selectedStatuses.includes(row.Status || 'Sem Status'));
        }

        // ShipmentType Filter (multi-select)
        const selectedShipmentTypes = Array.from(shipmentTypeFilter.selectedOptions).map(opt => opt.value);
        if (selectedShipmentTypes.length > 0 && !selectedShipmentTypes.includes('')) {
            filteredData = filteredData.filter(row => selectedShipmentTypes.includes(row.ShipmentType || 'Sem Tipo'));
        }

        // CargoType Filter (multi-select)
        const selectedCargoTypes = Array.from(cargoTypeFilter.selectedOptions).map(opt => opt.value);
        if (selectedCargoTypes.length > 0 && !selectedCargoTypes.includes('')) {
            filteredData = filteredData.filter(row => selectedCargoTypes.includes(row.CargoType || 'Sem Tipo de Mercadoria'));
        }

        filteredDataCache = filteredData;
        
        let processedData;
        if (currentView === 'vessel') {
            processedData = processDataByVessel(filteredData);
            renderVesselDashboard(processedData);
        } else if (currentView === 'po') {
            processedData = processDataByPO(filteredData);
            renderPODashboard(processedData);
        } else { // 'warehouse' view
            processedData = processDataByWarehouse(filteredData);
            renderWarehouseDashboard(processedData);
        }
        
        renderCharts(processedData, filteredData);
        updateTotalContainers(filteredData);
        hideLoading();
    }, 50); // Timeout allows UI to show spinner before processing
}

function resetFiltersAndRender() {
    showLoading();
    setTimeout(() => {
        arrivalStartDate.value = '';
        arrivalEndDate.value = '';
        deadlineStartDate.value = '';
        deadlineEndDate.value = '';
        poSearchInput.value = '';
        vesselSearchInput.value = '';
        Array.from(statusFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(shipmentTypeFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(cargoTypeFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(poFilter.options).forEach((opt, i) => opt.selected = i === 0);
        Array.from(vesselFilter.options).forEach((opt, i) => opt.selected = i === 0);
        applyFiltersAndRender();
        hideLoading();
    }, 50);
}


// --- View Switching ---
function setView(view: 'vessel' | 'po' | 'warehouse') {
    if (currentView === view) return;
    currentView = view;
    
    const buttons = { vessel: viewVesselBtn, po: viewPoBtn, warehouse: viewWarehouseBtn };
    Object.values(buttons).forEach(btn => {
        btn.classList.add('text-gray-500', 'border-transparent');
        btn.classList.remove('border-blue-600', 'text-blue-600');
    });

    buttons[view].classList.add('border-blue-600', 'text-blue-600');
    buttons[view].classList.remove('text-gray-500', 'border-transparent');
    
    if (originalData.length > 0) {
        populateColumnToggles(); // Repopulate columns for the new view
        applyFiltersAndRender();
    }
}

// --- Data Grouping and Processing ---
function calculateRisk(shipment: any) {
    const deadline = excelDateToJSDate(shipment.FreeTimeDeadline);
    const isDelivered = (shipment.Status || '').toLowerCase().includes('delivered');
    let daysToDeadline: number | null = null;
    let risk = 'low';

    if (isDelivered) {
        risk = 'none';
    } else if (deadline) {
        const timeDiff = deadline.getTime() - TODAY.getTime();
        daysToDeadline = Math.ceil(timeDiff / (1000 * 3600 * 24));
        if (daysToDeadline < 0) risk = 'high';
        else if (daysToDeadline <= 7) risk = 'medium';
    }
    return { ...shipment, daysToDeadline, risk };
}

function processDataByVessel(data: any[]) {
    return processDataGeneric(data, 'Vessel', 'Sem Navio');
}

function processDataByPO(data: any[]) {
    return processDataGeneric(data, 'SAPCargoPO', 'Sem PO');
}

function processDataByWarehouse(data: any[]) {
    return processDataGeneric(data, 'Warehouse', 'Sem Armazém');
}

function processDataGeneric(data: any[], groupKey: string, defaultName: string) {
    const grouped = data.reduce((acc, row) => {
        const name = row[groupKey] ? String(row[groupKey]).trim().toUpperCase() : defaultName;
        if (!acc[name]) acc[name] = [];
        acc[name].push(row);
        return acc;
    }, {});

    return Object.entries(grouped).map(([name, shipments]) => {
        const processedShipments = (shipments as any[]).map(calculateRisk);
        const riskCounts = { high: 0, medium: 0, low: 0, none: 0 };
        processedShipments.forEach(s => (riskCounts as any)[s.risk]++);
        
        let overallRisk = 'low';
        if (riskCounts.high > 0) overallRisk = 'high';
        else if (riskCounts.medium > 0) overallRisk = 'medium';
        else if (riskCounts.none === processedShipments.length) overallRisk = 'none';

        const totalContainers = processedShipments.reduce((sum, s) => {
            const fclCount = parseInt(s.FCL, 10) || 0;
            const lclCount = parseInt(s.LCL, 10) || 0;
            return sum + fclCount + lclCount;
        }, 0);
        
        return { name, shipments: processedShipments, totalFCL: totalContainers, overallRisk, riskCounts };
    }).sort((a, b) => {
        const riskOrder = { high: 0, medium: 1, low: 2, none: 3 };
        return (riskOrder as any)[a.overallRisk] - (riskOrder as any)[b.overallRisk];
    });
}


// --- UI Rendering ---
function renderVesselDashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderText: "Nenhum navio encontrado para os filtros aplicados.",
        cardTitlePrefix: '',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderPODashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderText: "Nenhuma PO encontrada para os filtros aplicados.",
        cardTitlePrefix: 'PO: ',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderWarehouseDashboard(data: any[]) {
    const headers = getVisibleColumns();
    const renderConfig = {
        placeholderText: "Nenhum armazém encontrado para os filtros aplicados.",
        cardTitlePrefix: '',
        columns: headers,
    };
    renderDashboard(data, renderConfig);
}

function renderDashboard(data: any[], config: any) {
    dashboardGrid.innerHTML = '';
    if (data.length === 0) {
        placeholder.classList.remove('hidden');
        placeholder.querySelector('h2')!.textContent = config.placeholderText;
        placeholder.querySelector('p')!.textContent = "Tente limpar os filtros ou carregar um novo arquivo.";
        return;
    }
    placeholder.classList.add('hidden');

    data.forEach(item => {
        const card = createDashboardCard(
            `${config.cardTitlePrefix}${item.name}`,
            item.totalFCL,
            config.columns,
            item.shipments,
            item.overallRisk
        );
        card.addEventListener('click', () => openModal(item));
        dashboardGrid.appendChild(card);
    });
}

function createDashboardCard(title: string, totalFCL: number, headers: string[], shipments: any[], risk: string) {
    const card = document.createElement('div');
    card.className = `card risk-${risk} cursor-pointer`;

    // Sort shipments based on global state
    const sortedShipments = sortData(shipments);

    const shipmentsHtml = sortedShipments.map((s: any) => {
        const daysText = s.risk === 'none' ? `<span class="font-semibold text-green-700">Entregue</span>` : s.daysToDeadline !== null ? `<span class="font-bold ${s.daysToDeadline < 0 ? 'text-red-600' : 'text-gray-800'}">${s.daysToDeadline}</span>` : 'N/A';

        const rowData = headers.map(header => {
            let cellContent = s[header] || '';
            if (header === 'Dias Restantes') {
                return `<td class="px-3 py-2 text-center text-xs">${daysText}</td>`;
            }
             if (header === 'FreeTimeDeadline') {
                return `<td class="px-3 py-2 whitespace-nowrap text-xs font-semibold">${cellContent}</td>`;
            }
            return `<td class="px-3 py-2 whitespace-nowrap text-xs">${cellContent}</td>`;
        }).join('');

        return `<tr class="row-risk-${s.risk} hover:bg-opacity-50">${rowData}</tr>`;
    }).join('');
    
    const getSortIndicator = (key: string) => {
        if (key === currentSortKey) {
            return currentSortOrder === 'asc' ? '<i class="fas fa-arrow-up ml-1"></i>' : '<i class="fas fa-arrow-down ml-1"></i>';
        }
        return '';
    };

    card.innerHTML = `<div class="p-4 border-b border-gray-200">
            <div class="flex justify-between items-center">
                <h3 class="font-extrabold text-lg text-gray-800">${title}</h3>
                <div class="text-right">
                    <span class="block text-2xl font-bold text-blue-600">${totalFCL}</span>
                    <span class="text-sm font-medium text-gray-500">Containers</span>
                </div>
            </div>
        </div>
        <div class="flex-grow table-responsive">
            <table class="min-w-full text-sm">
                <thead class="bg-gray-50"><tr class="border-b">
                    ${headers.map(h => `<th class="px-3 py-2 text-left font-semibold text-gray-500 text-xs uppercase tracking-wider sortable-header" data-sort-key="${h}">${h} ${getSortIndicator(h)}</th>`).join('')}
                </tr></thead>
                <tbody class="bg-white divide-y divide-gray-200">${shipmentsHtml}</tbody>
            </table>
        </div>`;
    return card;
}

// --- Chart Rendering ---
function renderCharts(processedData: any[], filteredData: any[]) {
    // Save data for export
    chartDataSource.bar = processedData;

    const isDarkMode = document.body.classList.contains('dark');
    const textColor = isDarkMode ? '#d1d5db' : '#4b5563';
    
    // --- Chart Color Helpers ---
    const getRiskColor = (risk: string) => {
        const colors: Record<string, string> = {
            high: 'rgba(239, 68, 68, 0.7)',    // red
            medium: 'rgba(249, 115, 22, 0.7)', // orange
            low: 'rgba(34, 197, 94, 0.7)',     // green
            none: 'rgba(107, 114, 128, 0.7)'   // gray
        };
        return colors[risk] || colors.none;
    };

    const getStatusColor = (status: string) => {
        const s = status.toLowerCase();
        if (s.includes('delivered')) return 'rgba(34, 197, 94, 0.7)';     // green
        if (s.includes('cleared')) return 'rgba(59, 130, 246, 0.7)';      // blue
        if (s.includes('sem status')) return 'rgba(239, 68, 68, 0.7)';     // red
        if (s.includes('presence') || s.includes('unloaded')) return 'rgba(249, 115, 22, 0.7)'; // orange
        if (s.includes('on board') || s.includes('transshipment')) return 'rgba(168, 85, 247, 0.7)'; // purple
        // Default color for other statuses
        return `rgba(${Math.floor(Math.random() * 155) + 100}, ${Math.floor(Math.random() * 155) + 100}, ${Math.floor(Math.random() * 155) + 100}, 0.7)`;
    };


    let barChartTitle: string;
    if (currentView === 'vessel') {
        barChartTitle = 'Total de Containers por Navio';
    } else if (currentView === 'po') {
        barChartTitle = 'Total de Containers por PO';
    } else { // Warehouse view
        barChartTitle = 'Total de Containers por Armazém';
    }
    document.getElementById('bar-chart-title')!.textContent = barChartTitle;
    
    const labels = processedData.map(item => item.name);
    const data = processedData.map(item => item.totalFCL);
    
    // Bar Chart
    if (mainBarChart) mainBarChart.destroy();
    Chart.register(ChartDataLabels);
    mainBarChart = new Chart(document.getElementById('main-bar-chart') as HTMLCanvasElement, {
        type: 'bar',
        data: { 
            labels, 
            datasets: [{ 
                label: 'Total Containers', 
                data, 
                backgroundColor: processedData.map(item => getRiskColor(item.overallRisk)),
                // Custom property to hold risk data for tooltips
                riskCounts: processedData.map(item => item.riskCounts)
            }] 
        },
        options: { 
            indexAxis: 'y', 
            responsive: true, 
            scales: {
                x: {
                    beginAtZero: true,
                    grace: '5%' // Add extra space at the top
                }
            },
            plugins: { 
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: textColor,
                    font: {
                        weight: 'bold'
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.x !== null) {
                                label += context.parsed.x;
                            }
                            return label;
                        },
                        afterLabel: function(context) {
                            const riskData = (context.dataset as any).riskCounts[context.dataIndex];
                            return [
                                `High Risk: ${riskData.high}`,
                                `Medium Risk: ${riskData.medium}`,
                                `Low Risk: ${riskData.low}`,
                                `Delivered: ${riskData.none}`
                            ];
                        }
                    }
                }
            } 
        }
    });

    // Pie Chart
    const statusCounts = filteredData.reduce((acc, row) => {
        const status = row.Status || 'Sem Status';
        const fclCount = parseInt(row.FCL, 10) || 0;
        const lclCount = parseInt(row.LCL, 10) || 0;
        const totalContainers = fclCount + lclCount;

        if (!acc[status]) {
            acc[status] = { count: 0, fcl: 0 };
        }
        acc[status].count += 1;
        acc[status].fcl += totalContainers;
        return acc;
    }, {} as Record<string, {count: number, fcl: number}>);
    
    // Save data for export
    chartDataSource.pie = Object.entries(statusCounts).map(([label, data]) => ({ label, value: (data as { fcl: number }).fcl }));

    const totalContainersInView = Object.values(statusCounts).reduce((sum: number, s) => sum + (s as { fcl: number }).fcl, 0);

    if (statusPieChart) statusPieChart.destroy();
    statusPieChart = new Chart(document.getElementById('status-pie-chart') as HTMLCanvasElement, {
        type: 'pie',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts).map(s => (s as { fcl: number }).fcl),
                backgroundColor: Object.keys(statusCounts).map(status => getStatusColor(status))
            }]
        },
        options: { 
            responsive: true, 
            plugins: { 
                legend: { position: 'right' },
                tooltip: {
                       callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = Number(context.parsed as any) || 0;
                                const total = totalContainersInView as number;
                                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                                return `${label}: ${value} Containers (${percentage}%)`;
                            }
                       }
                }
            },
            onClick: (event, elements, chart) => {
                if (elements.length > 0) {
                    const clickedIndex = elements[0].index;
                    const clickedStatus = chart.data.labels![clickedIndex] as string;
                    
                    // Reset selections and select only the clicked one
                    Array.from(statusFilter.options).forEach(opt => {
                        opt.selected = (opt.value === clickedStatus);
                    });

                    applyFiltersAndRender();
                }
            }
        }
    });

    // Deadline Distribution Chart
    const deadlineCanvas = document.getElementById('deadline-distribution-chart') as HTMLCanvasElement;
    const bins = {
        overdue: { label: 'Atrasado (<0)', count: 0, color: 'rgba(239, 68, 68, 0.7)' }, // red
        '0-7': { label: '0-7 Dias', count: 0, color: 'rgba(249, 115, 22, 0.7)' }, // orange
        '8-15': { label: '8-15 Dias', count: 0, color: 'rgba(245, 158, 11, 0.7)' }, // amber
        '16-30': { label: '16-30 Dias', count: 0, color: 'rgba(132, 204, 22, 0.7)' }, // lime
        '31+': { label: '31+ Dias', count: 0, color: 'rgba(34, 197, 94, 0.7)' }, // green
        delivered: { label: 'Entregue', count: 0, color: 'rgba(107, 114, 128, 0.7)' } // gray
    };

    const enrichedFilteredData = filteredData.map(calculateRisk);

    for (const shipment of enrichedFilteredData) {
        const { daysToDeadline, risk } = shipment;
        if (risk === 'none') {
            (bins.delivered.count)++;
        } else if (daysToDeadline !== null) {
            if (daysToDeadline < 0) (bins.overdue.count)++;
            else if (daysToDeadline <= 7) (bins['0-7'].count)++;
            else if (daysToDeadline <= 15) (bins['8-15'].count)++;
            else if (daysToDeadline <= 30) (bins['16-30'].count)++;
            else (bins['31+'].count)++;
        }
    }
    
    // Save data for export
    chartDataSource.deadline = Object.values(bins).map(b => ({ label: b.label, count: b.count }));


    if (deadlineDistributionChart) deadlineDistributionChart.destroy();
    deadlineDistributionChart = new Chart(deadlineCanvas, {
        type: 'bar',
        data: {
            labels: Object.values(bins).map(b => b.label),
            datasets: [{
                label: 'Nº de Cargas',
                data: Object.values(bins).map(b => b.count),
                backgroundColor: Object.values(bins).map(b => b.color)
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: textColor,
                    font: {
                        weight: 'bold'
                    }
                }
            }
        }
    });
}


// --- Helper Functions ---
function populateStatusFilter(data: any[]) {
    const statuses = [...new Set(data.map(row => row.Status || 'Sem Status'))].sort();
    statusFilter.innerHTML = '<option value="" selected>-- Todos --</option>';
    statuses.forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        statusFilter.appendChild(option);
    });
}

function populatePoFilter(data: any[]) {
    const pos = [...new Set(data.map(row => row.SAPCargoPO || 'Sem PO'))].sort();
    poFilter.innerHTML = '<option value="" selected>-- Todos --</option>';
    pos.forEach(po => {
        const option = document.createElement('option');
        option.value = po;
        option.textContent = po;
        poFilter.appendChild(option);
    });
}

function populateVesselFilter(data: any[]) {
    const vessels = [...new Set(data.map(row => row.Vessel || 'Sem Navio'))].sort();
    vesselFilter.innerHTML = '<option value="" selected>-- Todos --</option>';
    vessels.forEach(vessel => {
        const option = document.createElement('option');
        option.value = vessel;
        option.textContent = vessel;
        vesselFilter.appendChild(option);
    });
}

function populateShipmentTypeFilter(data: any[]) {
    const shipmentTypes = [...new Set(data.map(row => row.ShipmentType || 'Sem Tipo'))].sort();
    shipmentTypeFilter.innerHTML = '<option value="" selected>-- Todos --</option>';
    shipmentTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        shipmentTypeFilter.appendChild(option);
    });
}

function populateCargoTypeFilter(data: any[]) {
    const cargoTypes = [...new Set(data.map(row => row.CargoType || 'Sem Tipo de Mercadoria'))].sort();
    cargoTypeFilter.innerHTML = '<option value="" selected>-- Todos --</option>';
    cargoTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        cargoTypeFilter.appendChild(option);
    });
}

function updateTotalContainers(data: any[]) {
    const total = data.reduce((sum, row) => {
        const fclCount = parseInt(row.FCL, 10) || 0;
        const lclCount = parseInt(row.LCL, 10) || 0;
        return sum + fclCount + lclCount;
    }, 0);
    totalFclCount.textContent = total.toString();
}

function resetUI() {
    dashboardGrid.innerHTML = '';
    placeholder.classList.remove('hidden');
    filterContainer.classList.add('hidden');
    chartsContainer.classList.add('hidden');
    viewTabsContainer.classList.add('hidden');
    exportCsvBtn.classList.add('hidden');
    exportPdfBtn.classList.add('hidden');
    exportExcelBtn.classList.add('hidden');
    totalFclDisplay.classList.add('hidden');
    originalData = [];
    if (mainBarChart) mainBarChart.destroy();
    if (statusPieChart) statusPieChart.destroy();
    if (deadlineDistributionChart) deadlineDistributionChart.destroy();
    lastUpdate.textContent = 'Carregue um arquivo .xlsx para começar';
    setView('vessel');
}

// --- Sorting ---
function handleSortClick(event: MouseEvent) {
    const target = event.target as HTMLElement;
    const header = target.closest('.sortable-header');
    if (!header) return;

    const sortKey = header.getAttribute('data-sort-key');
    if (!sortKey) return;

    if (sortKey === currentSortKey) {
        currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortKey = sortKey;
        currentSortOrder = 'asc';
    }
    applyFiltersAndRender();
}

function sortData(data: any[]) {
    return [...data].sort((a, b) => {
        const valA = a[currentSortKey];
        const valB = b[currentSortKey];
        
        let comparison = 0;
        
        if (currentSortKey === 'Arrival' || currentSortKey === 'FreeTimeDeadline') {
            const dateA = excelDateToJSDate(valA);
            const dateB = excelDateToJSDate(valB);
            if (dateA && dateB) {
                comparison = dateA.getTime() - dateB.getTime();
            } else if (dateA) {
                comparison = -1;
            } else if (dateB) {
                comparison = 1;
            }
        } else if (currentSortKey === 'Dias Restantes') {
             // Handle nulls for delivered items
            const daysA = a.daysToDeadline ?? Infinity;
            const daysB = b.daysToDeadline ?? Infinity;
            comparison = daysA - daysB;
        } else {
            // Default string/number sort
            const strA = String(valA || '').toLowerCase();
            const strB = String(valB || '').toLowerCase();
            if (strA < strB) comparison = -1;
            if (strA > strB) comparison = 1;
        }
        
        return currentSortOrder === 'asc' ? comparison : -comparison;
    });
}

// --- Column Visibility ---
function setupColumnToggles() {
    loadColumnVisibility();
    populateColumnToggles();

    columnToggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        columnToggleDropdown.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
        if (!columnToggleDropdown.classList.contains('hidden') && !columnToggleBtn.contains(e.target as Node) && !columnToggleDropdown.contains(e.target as Node)) {
            columnToggleDropdown.classList.add('hidden');
        }
    });

    columnToggleDropdown.addEventListener('change', (e) => {
        const checkbox = e.target as HTMLInputElement;
        columnVisibility[checkbox.value] = checkbox.checked;
        saveColumnVisibility();
        applyFiltersAndRender();
    });
}

function populateColumnToggles() {
    const columns = viewColumns[currentView];
    // Re-check defaults if a column has no saved setting
    columns.forEach(col => {
        if (columnVisibility[col] === undefined) {
            columnVisibility[col] = true;
        }
    });
    
    columnToggleDropdown.innerHTML = columns.map(col => `
        <label class="flex items-center px-4 py-2 hover:bg-gray-100 cursor-pointer">
            <input type="checkbox" class="form-checkbox h-4 w-4 text-blue-600 border-gray-300 rounded" value="${col}" ${columnVisibility[col] ? 'checked' : ''}>
            <span class="ml-3 text-sm text-gray-700">${col}</span>
        </label>
    `).join('');
}

function getVisibleColumns(): string[] {
    return viewColumns[currentView].filter(col => columnVisibility[col]);
}

function loadColumnVisibility() {
    try {
        const saved = localStorage.getItem('columnVisibility');
        if (saved) {
            columnVisibility = JSON.parse(saved);
        } else {
            // Default all to visible
            columnVisibility = {};
             Object.values(viewColumns).flat().forEach(col => columnVisibility[col] = true);
        }
    } catch(e) {
        console.error("Could not load column visibility from localStorage", e);
        columnVisibility = {};
        Object.values(viewColumns).flat().forEach(col => columnVisibility[col] = true);
    }
}

function saveColumnVisibility() {
    localStorage.setItem('columnVisibility', JSON.stringify(columnVisibility));
}

// --- Exporting ---
function handleExcelExport() {
    if (filteredDataCache.length === 0) {
        showToast("Não há dados para exportar.", "warning");
        return;
    }
    
    const workbook = XLSX.utils.book_new();

    // 1. Filtered Data Sheet
    const visibleHeaders = getVisibleColumns();
    const dataToExport = filteredDataCache.map(row => {
        const newRow: Record<string, any> = {};
        visibleHeaders.forEach(header => {
            newRow[header] = row[header] ?? '';
        });
        return newRow;
    });
    const filteredDataSheet = XLSX.utils.json_to_sheet(dataToExport, { header: visibleHeaders });
    XLSX.utils.book_append_sheet(workbook, filteredDataSheet, 'Filtered Data');

    // 2. Main Bar Chart Data Sheet
    if (chartDataSource.bar.length > 0) {
        const barDataForSheet = chartDataSource.bar.map(item => ({
            'Item': item.name,
            'Total Containers': item.totalFCL,
            'High Risk': item.riskCounts.high,
            'Medium Risk': item.riskCounts.medium,
            'Low Risk': item.riskCounts.low,
            'Delivered': item.riskCounts.none
        }));
        const barChartSheet = XLSX.utils.json_to_sheet(barDataForSheet);
        XLSX.utils.book_append_sheet(workbook, barChartSheet, 'Chart Data (Main View)');
    }

    // 3. Status Pie Chart Data Sheet
    if (chartDataSource.pie.length > 0) {
        const pieDataForSheet = chartDataSource.pie.map(item => ({
            'Status': item.label,
            'Total Containers': item.value
        }));
        const pieChartSheet = XLSX.utils.json_to_sheet(pieDataForSheet);
        XLSX.utils.book_append_sheet(workbook, pieChartSheet, 'Chart Data (Status)');
    }

    // 4. Deadline Distribution Chart Data Sheet
    if (chartDataSource.deadline.length > 0) {
        const deadlineDataForSheet = chartDataSource.deadline.map(item => ({
            'Prazo (Dias)': item.label,
            'Nº de Cargas': item.count
        }));
        const deadlineChartSheet = XLSX.utils.json_to_sheet(deadlineDataForSheet);
        XLSX.utils.book_append_sheet(workbook, deadlineChartSheet, 'Chart Data (Deadlines)');
    }

    // Create a filename and trigger the download
    const fileName = `dashboard_export_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
    
    showToast("Exportação para Excel iniciada.", "success");
}

function handleGlobalCsvExport() {
    if (filteredDataCache.length === 0) {
        showToast("Não há dados para exportar.", "warning");
        return;
    }
    const csvContent = convertShipmentsToCSV(filteredDataCache);
    downloadCSV(csvContent, `dashboard_export_${currentView}_${new Date().toISOString().split('T')[0]}.csv`);
}

function convertShipmentsToCSV(data: any[]): string {
    if (data.length === 0) return "";

    const allHeaders = Object.keys(data[0] || {});
    const visibleHeaders = getVisibleColumns();

    // Use all headers but prioritize the order and visibility from the user's selection
    const headers = [...visibleHeaders, ...allHeaders.filter(h => !visibleHeaders.includes(h) && !['daysToDeadline', 'risk', 'riskCounts'].includes(h))];

    const rows = data.map(shipment =>
        headers.map(header => {
            let value = shipment[header] ?? '';
            // Ensure values with commas are wrapped in quotes
            return `"${String(value).replace(/"/g, '""')}"`;
        }).join(',')
    );
    return [headers.join(','), ...rows].join('\n');
}

function downloadCSV(csvContent: string, fileName: string) {
    const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' }); // Add BOM for Excel
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function handlePdfExport() {
    showLoading();
    const reportContent = document.getElementById('report-content');
    if (!reportContent) {
        hideLoading();
        return;
    }

    try {
        const { jsPDF } = jspdf;
        const canvas = await html2canvas(reportContent, {
            scale: 2, // Higher scale for better quality
            useCORS: true,
            logging: false,
            windowWidth: reportContent.scrollWidth,
            windowHeight: reportContent.scrollHeight
        });
        
        const imgData = canvas.toDataURL('image/png');
        const pdf = new jsPDF({
            orientation: 'landscape',
            unit: 'px',
            format: [canvas.width, canvas.height]
        });
        
        pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
        pdf.save(`dashboard_export_${currentView}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (error) {
        console.error("Error generating PDF:", error);
        showToast("Ocorreu um erro ao gerar o PDF.", "error");
    } finally {
        hideLoading();
    }
}


// --- Modal ---
function openModal(item: any) {
    activeModalItem = item;
    renderModalContent();
    detailsModal.classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
    // For animation
    setTimeout(() => {
        detailsModal.classList.add('modal-open');
    }, 10);
}

function closeModal() {
    detailsModal.classList.remove('modal-open');
     setTimeout(() => {
        detailsModal.classList.add('hidden');
        document.body.classList.remove('overflow-hidden');
        modalBody.innerHTML = ''; // Clear content
        modalHeaderContent.innerHTML = '';
        activeModalItem = null;
    }, 300); // Match CSS transition duration
}

function renderModalContent() {
    if (!activeModalItem) return;

    const { name, totalFCL, shipments } = activeModalItem;
    
    // 1. Render Header
    const prefix = (currentView === 'po') ? 'PO: ' : '';
    const title = `${prefix}${name}`;
    
    const uniqueWarehouses = [...new Set(shipments
        .map((s: any) => s.Warehouse)
        .filter((wh: string | undefined) => wh && wh.trim() !== ''))];

    const warehouseDisplayHtml = uniqueWarehouses.length > 0 
        ? `<div class="mt-1 flex items-center">
             <i class="fas fa-warehouse text-gray-500 mr-2"></i>
             <p class="text-sm text-gray-600 font-medium">${uniqueWarehouses.join(', ')}</p>
           </div>`
        : '';

    modalHeaderContent.innerHTML = `
        <div class="flex justify-between items-start">
            <div>
                <h3 class="font-extrabold text-xl text-gray-800">${title}</h3>
                ${warehouseDisplayHtml}
            </div>
            <div class="text-right ml-8 flex-shrink-0">
                <span class="block text-3xl font-bold text-blue-600">${totalFCL}</span>
                <span class="text-sm font-medium text-gray-500">Containers</span>
            </div>
        </div>
    `;

    // 2. Render Body (Table)
    const blContainerCounts = shipments.reduce((acc: Record<string, number>, shipment: any) => {
        const blAwb = shipment['BL/AWB'];
        if (blAwb) {
            const fcl = parseInt(shipment.FCL, 10) || 0;
            const lcl = parseInt(shipment.LCL, 10) || 0;
            acc[blAwb] = (acc[blAwb] || 0) + fcl + lcl;
        }
        return acc;
    }, {});

    // Define a fixed, comprehensive set of columns for the detail modal to ensure consistency.
    // This bypasses the main view's column visibility settings for a complete detailed view.
    let headers: string[] = [];
    const baseHeaders = [
        'BL/AWB',
        'Qtd. Contêineres (BL)', // This is a virtual column
        'Carrier',
        'Warehouse',
        'Status',
        'ShipmentType',
        'CargoType',
        'Arrival',
        'FreeTimeDeadline',
        'Dias Restantes'
    ];
    
    // Add view-specific columns at the beginning
    if (currentView === 'vessel') {
        headers = ['SAPCargoPO', ...baseHeaders];
    } else if (currentView === 'po') {
        headers = ['Vessel', ...baseHeaders];
    } else { // warehouse view
        headers = ['Vessel', 'SAPCargoPO', ...baseHeaders];
    }

    // Map internal data keys to user-friendly Portuguese headers for the modal.
    const headerDisplayMap: Record<string, string> = {
        'SAPCargoPO': 'PO',
        'Vessel': 'Navio',
        'ShipmentType': 'Tipo de Carga',
        'CargoType': 'Tipo de Mercadoria'
    };

    const sortedShipments = sortData(shipments);
    
    const shipmentsHtml = sortedShipments.map((shipment: any) => {
        const daysText = shipment.risk === 'none' 
            ? `<span class="font-semibold text-green-700">Entregue</span>` 
            : shipment.daysToDeadline !== null 
            ? `<span class="font-bold ${shipment.daysToDeadline < 0 ? 'text-red-600' : 'text-gray-800'}">${shipment.daysToDeadline}</span>` 
            : 'N/A';

        const rowData = headers.map(header => {
            if (header === 'Qtd. Contêineres (BL)') {
                const blAwb = shipment['BL/AWB'] || '';
                const count = blAwb ? blContainerCounts[blAwb] : 0;
                return `<td class="px-2 py-2 text-center font-bold text-blue-600">${count}</td>`;
            }
            
            let cellContent = shipment[header] || '';
            if (header === 'Dias Restantes') {
                return `<td class="px-2 py-2 text-center">${daysText}</td>`;
            }
            if (header === 'FreeTimeDeadline' || header === 'Arrival') {
                return `<td class="px-2 py-2 whitespace-nowrap">${cellContent}</td>`;
            }
            if (header === 'BL/AWB' || header === 'SAPCargoPO') {
                return `<td class="px-2 py-2 whitespace-nowrap">${cellContent}</td>`;
            }
            if (header === 'Vessel') {
                return `<td class="px-2 py-2 break-all">${cellContent}</td>`;
            }
            return `<td class="px-2 py-2 break-words">${cellContent}</td>`;
        }).join('');

        return `<tr class="row-risk-${shipment.risk}">${rowData}</tr>`;
    }).join('');

    modalBody.innerHTML = `
        <div class="table-responsive">
            <table class="min-w-full text-sm">
                <thead class="bg-gray-50"><tr class="border-b">
                    ${headers.map(h => `<th class="px-2 py-2 text-left font-semibold text-gray-500 uppercase tracking-wider text-xs">${headerDisplayMap[h] || h}</th>`).join('')}
                </tr></thead>
                <tbody class="bg-white divide-y divide-gray-200">${shipmentsHtml}</tbody>
            </table>
        </div>
    `;
}

// --- Logo Management ---
function loadSavedLogo() {
    const savedLogo = localStorage.getItem('companyLogo');
    if (savedLogo) {
        companyLogo.src = savedLogo;
        companyLogo.classList.remove('hidden');
        removeLogoBtn.classList.remove('hidden');
    }
}

function handleLogoUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const logoDataUrl = e.target?.result as string;
        if (logoDataUrl) {
            localStorage.setItem('companyLogo', logoDataUrl);
            companyLogo.src = logoDataUrl;
            companyLogo.classList.remove('hidden');
            removeLogoBtn.classList.remove('hidden');
            showToast('Logo atualizado com sucesso!', 'success');
        }
    };
    reader.onerror = () => {
         showToast('Erro ao carregar o logo.', 'error');
    };
    reader.readAsDataURL(file);
    target.value = ''; // Reset input so the same file can be chosen again
}

function handleRemoveLogo() {
    localStorage.removeItem('companyLogo');
    companyLogo.src = '';
    companyLogo.classList.add('hidden');
    removeLogoBtn.classList.add('hidden');
    showToast('Logo removido.', 'success');
}
