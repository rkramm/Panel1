// ============================================
// CONFIGURACIÓN SUPABASE (Reemplaza con tus credenciales)
// ============================================
const { createClient } = supabase;
const SUPABASE_URL = 'https://jwmyhrldrqxhwletbtyy.supabase.co/rest/v1/';
const SUPABASE_KEY = 'sb_publishable_3qvX38tpEJ76PjGv3mmYYg_Hv-WAoMO';

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// ============================================
// ESTADO GLOBAL
// ============================================
let currentUser = null;
let gastos = [];
let cuentasAhorro = [];
let prestamos = [];

// ============================================
// INICIALIZACIÓN
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    initAuth();
    initTabs();
    initFileUpload();
    initModals();
    checkSession();
});

// ============================================
// AUTENTICACIÓN (Magic Link por Email)
// ============================================
function initAuth() {
    const loginBtn = document.getElementById('google-login');
    loginBtn.addEventListener('click', async () => {
        const email = prompt('Ingresa tu email:');
        if (!email) return;
        
        const { error } = await supabase.auth.signInWithOtp({
            email: email,
            options: {
                emailRedirectTo: window.location.origin
            }
        });
        
        if (error) {
            showToast('Error: ' + error.message);
        } else {
            showToast('Revisa tu email para el link de acceso');
        }
    });

    document.getElementById('logout-btn').addEventListener('click', async () => {
        await supabase.auth.signOut();
        location.reload();
    });
}

async function checkSession() {
    const { data: { session } } = await supabase.auth.getSession();
    if (session) {
        currentUser = session.user;
        showApp();
    } else {
        showAuth();
    }
}

function showAuth() {
    document.getElementById('auth-screen').classList.remove('hidden');
    document.getElementById('app-screen').classList.add('hidden');
}

function showApp() {
    document.getElementById('auth-screen').classList.add('hidden');
    document.getElementById('app-screen').classList.remove('hidden');
    document.getElementById('user-name').textContent = currentUser.email;
    loadData();
}

// ============================================
// NAVEGACIÓN POR TABS
// ============================================
function initTabs() {
    document.querySelectorAll('.nav-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.nav-tab').forEach(t => {
                t.classList.remove('text-emerald-400', 'border-emerald-400');
                t.classList.add('text-slate-400', 'border-transparent');
            });
            tab.classList.remove('text-slate-400', 'border-transparent');
            tab.classList.add('text-emerald-400', 'border-emerald-400');

            document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
            document.getElementById(`tab-${tab.dataset.tab}`).classList.remove('hidden');

            if (tab.dataset.tab === 'dashboard') renderDashboard();
            if (tab.dataset.tab === 'gastos') renderGastos();
            if (tab.dataset.tab === 'ahorro') renderAhorro();
        });
    });
}

// ============================================
// CARGA DE DATOS
// ============================================
async function loadData() {
    await Promise.all([
        loadGastos(),
        loadCuentas(),
        loadPrestamos()
    ]);
    renderDashboard();
}

async function loadGastos() {
    const { data, error } = await supabase
        .from('gastos')
        .select('*')
        .eq('user_id', currentUser.id)
        .order('fecha', { ascending: false });
    
    if (error) {
        showToast('Error cargando gastos');
        return;
    }
    gastos = data || [];
    updateMonthFilter();
}

async function loadCuentas() {
    const { data, error } = await supabase
        .from('cuentas_ahorro')
        .select('*')
        .eq('user_id', currentUser.id);
    
    if (error) {
        showToast('Error cargando cuentas');
        return;
    }
    cuentasAhorro = data || [];
}

async function loadPrestamos() {
    const { data, error } = await supabase
        .from('prestamos')
        .select('*, cuentas_ahorro(nombre)')
        .eq('user_id', currentUser.id)
        .eq('estado', 'activo');
    
    if (error) {
        showToast('Error cargando préstamos');
        return;
    }
    prestamos = data || [];
}

// ============================================
// UPLOAD DE EXCEL
// ============================================
function initFileUpload() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('border-emerald-500', 'bg-slate-700');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('border-emerald-500', 'bg-slate-700');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-emerald-500', 'bg-slate-700');
        const files = e.dataTransfer.files;
        if (files.length) processExcel(files[0]);
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) processExcel(e.target.files[0]);
    });
}

function processExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            const nuevosGastos = [];
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row[0]) continue;

                const fecha = parseDate(row[0]);
                const descripcion = String(row[1] || '').trim();
                const titular = String(row[2] || '').trim();
                const montoRaw = row[3];
                const cuotasPendientes = row[4] ? parseInt(row[4]) : null;
                const valorCuota = row[5] ? parseFloat(row[5]) : null;

                let monto = 0;
                if (typeof montoRaw === 'number') {
                    monto = montoRaw;
                } else if (typeof montoRaw === 'string') {
                    monto = parseFloat(montoRaw.replace(/\./g, '').replace(',', '.'));
                }

                if (!fecha || isNaN(monto) || monto === 0) continue;

                const categoria = detectarCategoria(descripcion);

                nuevosGastos.push({
                    user_id: currentUser.id,
                    fecha: fecha,
                    descripcion: descripcion,
                    titular: titular,
                    monto: Math.abs(monto),
                    cuotas_pendientes: cuotasPendientes,
                    valor_cuota: valorCuota,
                    categoria: categoria,
                    porcentaje_usuario: 50,
                    es_manual: false,
                    created_at: new Date().toISOString()
                });
            }

            if (nuevosGastos.length === 0) {
                showToast('No se encontraron datos válidos en el archivo');
                return;
            }

            saveGastosBatch(nuevosGastos);
        } catch (err) {
            console.error(err);
            showToast('Error procesando el archivo Excel');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseDate(dateValue) {
    if (typeof dateValue === 'number') {
        return new Date((dateValue - 25569) * 86400 * 1000).toISOString().split('T')[0];
    }
    if (typeof dateValue === 'string') {
        const parts = dateValue.split(/[\/\-]/);
        if (parts.length === 3) {
            if (parts[2].length === 4) {
                return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }
            return dateValue;
        }
    }
    return null;
}

function detectarCategoria(descripcion) {
    const desc = descripcion.toLowerCase();
    const keywords = {
        supermercado: ['supermercado', 'jumbo', 'lider', 'unimarc', 'santa isabel', 'alvi', 'acuenta', 'masxmenos'],
        transporte: ['uber', 'didi', 'beat', 'cabify', 'combustible', 'shell', 'copec', 'bencina', 'metro', 'bus'],
        servicios: ['agua', 'luz', 'electricidad', 'gas', 'internet', 'telefonia', 'movistar', 'entel', 'vtr', 'mundo'],
        salud: ['farmacia', 'salco', 'cruz verde', 'ahumada', 'medico', 'hospital', 'clinica', 'dentista'],
        entretenimiento: ['netflix', 'spotify', 'youtube', 'cine', 'hbo', 'disney', 'amazon prime', 'steam'],
        restaurantes: ['restaurant', 'mcdonalds', 'burger', 'pizza', 'sushi', 'rappi', 'pedidosya', 'ubereats'],
        hogar: ['homecenter', 'sodimac', 'easy', 'ikea', 'ferreteria', 'muebles'],
        educacion: ['universidad', 'colegio', 'curso', 'udemy', 'coursera', 'libro', 'libreria']
    };

    for (const [cat, words] of Object.entries(keywords)) {
        if (words.some(w => desc.includes(w))) return cat;
    }
    return 'otros';
}

async function saveGastosBatch(gastosArray) {
    showToast(`Procesando ${gastosArray.length} gastos...`);
    
    const { data, error } = await supabase
        .from('gastos')
        .insert(gastosArray)
        .select();

    if (error) {
        showToast('Error guardando gastos: ' + error.message);
        return;
    }

    showToast(`${gastosArray.length} gastos importados exitosamente`);
    await loadGastos();
    renderDashboard();
}

// ============================================
// RENDERIZADO: DASHBOARD
// ============================================
function renderDashboard() {
    const mesActual = new Date().toISOString().slice(0, 7);
    const gastosMes = gastos.filter(g => g.fecha.startsWith(mesActual));
    
    const totalMes = gastosMes.reduce((sum, g) => sum + g.monto, 0);
    const miParte = gastosMes.reduce((sum, g) => sum + (g.monto * (g.porcentaje_usuario / 100)), 0);

    document.getElementById('total-mes').textContent = formatCurrency(totalMes);
    document.getElementById('mi-parte').textContent = formatCurrency(miParte);

    const porCategoria = {};
    gastosMes.forEach(g => {
        porCategoria[g.categoria] = (porCategoria[g.categoria] || 0) + g.monto;
    });

    const chartContainer = document.getElementById('chart-container');
    chartContainer.innerHTML = '';
    
    const categoriasOrdenadas = Object.entries(porCategoria).sort((a, b) => b[1] - a[1]);
    const maxVal = Math.max(...categoriasOrdenadas.map(c => c[1]));

    categoriasOrdenadas.forEach(([cat, monto]) => {
        const porcentaje = (monto / totalMes * 100).toFixed(1);
        const barWidth = (monto / maxVal * 100).toFixed(0);
        
        chartContainer.innerHTML += `
            <div class="flex items-center gap-3">
                <div class="w-24 text-xs text-slate-400 capitalize">${getCatEmoji(cat)} ${cat}</div>
                <div class="flex-1 bg-slate-700 rounded-full h-6 overflow-hidden">
                    <div class="bg-emerald-500 h-full rounded-full flex items-center justify-end pr-2" style="width: ${barWidth}%">
                        <span class="text-xs text-white font-medium">${porcentaje}%</span>
                    </div>
                </div>
                <div class="w-20 text-right text-sm font-medium">${formatCurrency(monto)}</div>
            </div>
        `;
    });

    const recentContainer = document.getElementById('recent-activity');
    recentContainer.innerHTML = '';
    gastos.slice(0, 5).forEach(g => {
        recentContainer.innerHTML += `
            <div class="flex items-center justify-between py-2 border-b border-slate-700 last:border-0">
                <div>
                    <p class="text-sm font-medium">${g.descripcion}</p>
                    <p class="text-xs text-slate-500">${formatDate(g.fecha)} • ${g.categoria}</p>
                </div>
                <div class="text-right">
                    <p class="text-sm font-medium text-white">${formatCurrency(g.monto)}</p>
                    <p class="text-xs text-emerald-400">Tu parte: ${g.porcentaje_usuario}%</p>
                </div>
            </div>
        `;
    });
}

// ============================================
// RENDERIZADO: GASTOS
// ============================================
function renderGastos() {
    const filter = document.getElementById('month-filter').value;
    let filtered = gastos;
    
    if (filter !== 'all') {
        filtered = gastos.filter(g => g.fecha.startsWith(filter));
    }

    const container = document.getElementById('gastos-list');
    container.innerHTML = '';

    if (filtered.length === 0) {
        container.innerHTML = '<p class="text-center text-slate-500 py-8">No hay gastos registrados</p>';
        return;
    }

    filtered.forEach(g => {
        const miParte = g.monto * (g.porcentaje_usuario / 100);
        const otraParte = g.monto - miParte;
        
        container.innerHTML += `
            <div class="bg-slate-800 rounded-xl p-4 border border-slate-700 cursor-pointer hover:border-slate-500 transition" onclick="editGasto('${g.id}')">
                <div class="flex justify-between items-start mb-2">
                    <div>
                        <p class="font-medium text-sm">${g.descripcion}</p>
                        <p class="text-xs text-slate-500">${formatDate(g.fecha)} • ${getCatEmoji(g.categoria)} ${g.categoria}</p>
                    </div>
                    <div class="text-right">
                        <p class="font-bold text-white">${formatCurrency(g.monto)}</p>
                        ${g.cuotas_pendientes ? `<p class="text-xs text-amber-400">${g.cuotas_pendientes} cuotas de ${formatCurrency(g.valor_cuota)}</p>` : ''}
                    </div>
                </div>
                <div class="flex gap-2 mt-2">
                    <div class="flex-1 bg-emerald-900/30 rounded-lg px-2 py-1 text-center">
                        <p class="text-xs text-slate-400">Tú (${g.porcentaje_usuario}%)</p>
                        <p class="text-sm font-medium text-emerald-400">${formatCurrency(miParte)}</p>
                    </div>
                    <div class="flex-1 bg-slate-700/30 rounded-lg px-2 py-1 text-center">
                        <p class="text-xs text-slate-400">Otra persona (${100 - g.porcentaje_usuario}%)</p>
                        <p class="text-sm font-medium text-slate-300">${formatCurrency(otraParte)}</p>
                    </div>
                </div>
            </div>
        `;
    });
}

function updateMonthFilter() {
    const select = document.getElementById('month-filter');
    const meses = [...new Set(gastos.map(g => g.fecha.slice(0, 7)))].sort().reverse();
    
    select.innerHTML = '<option value="all">Todos los meses</option>';
    meses.forEach(m => {
        const [year, month] = m.split('-');
        const nombreMes = new Date(year, month - 1).toLocaleDateString('es-ES', { month: 'long', year: 'numeric' });
        select.innerHTML += `<option value="${m}">${nombreMes}</option>`;
    });
    
    select.addEventListener('change', renderGastos);
}

// ============================================
// RENDERIZADO: AHORRO
// ============================================
function renderAhorro() {
    const container = document.getElementById('cuentas-ahorro');
    container.innerHTML = '';

    cuentasAhorro.forEach(c => {
        const prestamosCuenta = prestamos.filter(p => p.cuenta_id === c.id);
        const totalPrestado = prestamosCuenta.reduce((sum, p) => sum + p.monto, 0);
        const saldoReal = c.saldo - totalPrestado;

        container.innerHTML += `
            <div class="bg-slate-800 rounded-xl p-4 border border-slate-700">
                <div class="flex justify-between items-start mb-3">
                    <div>
                        <h4 class="font-bold">${c.nombre}</h4>
                        <p class="text-xs text-slate-500">Creada el ${formatDate(c.created_at)}</p>
                    </div>
                    <div class="text-right">
                        <p class="text-2xl font-bold text-emerald-400">${formatCurrency(saldoReal)}</p>
                        <p class="text-xs text-slate-500">Saldo disponible</p>
                    </div>
                </div>
                ${totalPrestado > 0 ? `<p class="text-xs text-amber-400 mb-2">-${formatCurrency(totalPrestado)} prestado</p>` : ''}
                <button onclick="retirarCuenta('${c.id}')" class="w-full bg-slate-700 hover:bg-slate-600 py-2 rounded-lg text-sm font-medium transition">
                    Retirar / Prestar
                </button>
            </div>
        `;
    });

    const prestamosContainer = document.getElementById('prestamos-activos');
    prestamosContainer.innerHTML = '';

    if (prestamos.length === 0) {
        prestamosContainer.innerHTML = '<p class="text-center text-slate-500 py-4">No hay préstamos activos</p>';
        return;
    }

    prestamos.forEach(p => {
        const montoCuota = p.monto / p.cuotas_total;
        const cuotasRestantes = p.cuotas_total - p.cuotas_pagadas;
        const saldoPendiente = montoCuota * cuotasRestantes;

        prestamosContainer.innerHTML += `
            <div class="bg-slate-900 rounded-lg p-3 border border-slate-700">
                <div class="flex justify-between items-start">
                    <div>
                        <p class="font-medium text-sm">${p.motivo}</p>
                        <p class="text-xs text-slate-500">De: ${p.cuentas_ahorro?.nombre || 'Cuenta eliminada'}</p>
                    </div>
                    <div class="text-right">
                        <p class="text-sm font-bold">${formatCurrency(p.monto)}</p>
                        <p class="text-xs ${cuotasRestantes === 0 ? 'text-emerald-400' : 'text-amber-400'}">
                            ${p.cuotas_pagadas}/${p.cuotas_total} cuotas pagadas
                        </p>
                    </div>
                </div>
                <div class="mt-2 flex items-center justify-between">
                    <div class="text-xs text-slate-400">
                        Cuota mensual: ${formatCurrency(montoCuota)}<br>
                        Saldo pendiente: ${formatCurrency(saldoPendiente)}
                    </div>
                    ${cuotasRestantes > 0 ? `
                        <button onclick="pagarCuota('${p.id}')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1 rounded text-xs font-medium">
                            Pagar Cuota
                        </button>
                    ` : `
                        <span class="text-emerald-400 text-xs font-medium">✓ Pagado</span>
                    `}
                </div>
                <div class="mt-2 bg-slate-800 rounded-full h-2 overflow-hidden">
                    <div class="bg-emerald-500 h-full rounded-full transition-all" style="width: ${(p.cuotas_pagadas / p.cuotas_total * 100)}%"></div>
                </div>
            </div>
        `;
    });
}

// ============================================
// MODALES Y FORMULARIOS
// ============================================
function initModals() {
    document.getElementById('cancel-gasto').addEventListener('click', () => toggleModal('modal-gasto', false));
    document.getElementById('form-gasto').addEventListener('submit', async (e) => {
        e.preventDefault();
        const id = document.getElementById('gasto-id').value;
        const porcentaje = parseInt(document.getElementById('gasto-porcentaje').value);
        const categoria = document.getElementById('gasto-categoria').value;

        const { error } = await supabase
            .from('gastos')
            .update({ porcentaje_usuario: porcentaje, categoria: categoria })
            .eq('id', id);

        if (error) {
            showToast('Error actualizando gasto');
            return;
        }

        toggleModal('modal-gasto', false);
        await loadGastos();
        renderGastos();
        renderDashboard();
        showToast('Gasto actualizado');
    });

    document.getElementById('add-cuenta-btn').addEventListener('click', () => toggleModal('modal-cuenta', true));
    document.getElementById('cancel-cuenta').addEventListener('click', () => toggleModal('modal-cuenta', false));
    document.getElementById('form-cuenta').addEventListener('submit', async (e) => {
        e.preventDefault();
        const nombre = document.getElementById('cuenta-nombre').value;
        const saldo = parseFloat(document.getElementById('cuenta-saldo').value);

        const { error } = await supabase
            .from('cuentas_ahorro')
            .insert([{ user_id: currentUser.id, nombre, saldo }]);

        if (error) {
            showToast('Error creando cuenta');
            return;
        }

        toggleModal('modal-cuenta', false);
        document.getElementById('form-cuenta').reset();
        await loadCuentas();
        renderAhorro();
        showToast('Cuenta creada exitosamente');
    });

    document.getElementById('cancel-prestamo').addEventListener('click', () => toggleModal('modal-prestamo', false));
    document.getElementById('form-prestamo').addEventListener('submit', async (e) => {
        e.preventDefault();
        const cuentaId = document.getElementById('prestamo-cuenta-id').value;
        const monto = parseFloat(document.getElementById('prestamo-monto').value);
        const motivo = document.getElementById('prestamo-motivo').value;
        const cuotas = parseInt(document.getElementById('prestamo-cuotas').value);

        const { error } = await supabase
            .from('prestamos')
            .insert([{
                user_id: currentUser.id,
                cuenta_id: cuentaId,
                monto: monto,
                motivo: motivo,
                cuotas_total: cuotas,
                cuotas_pagadas: 0,
                estado: 'activo'
            }]);

        if (error) {
            showToast('Error creando préstamo');
            return;
        }

        toggleModal('modal-prestamo', false);
        document.getElementById('form-prestamo').reset();
        await Promise.all([loadCuentas(), loadPrestamos()]);
        renderAhorro();
        showToast('Préstamo registrado');
    });

    document.getElementById('add-manual-btn').addEventListener('click', () => {
        document.getElementById('manual-fecha').value = new Date().toISOString().split('T')[0];
        toggleModal('modal-manual', true);
    });
    document.getElementById('cancel-manual').addEventListener('click', () => toggleModal('modal-manual', false));
    document.getElementById('form-manual').addEventListener('submit', async (e) => {
        e.preventDefault();
        const gasto = {
            user_id: currentUser.id,
            fecha: document.getElementById('manual-fecha').value,
            descripcion: document.getElementById('manual-desc').value,
            monto: parseFloat(document.getElementById('manual-monto').value),
            categoria: document.getElementById('manual-categoria').value,
            porcentaje_usuario: 50,
            es_manual: true
        };

        const { error } = await supabase.from('gastos').insert([gasto]);
        if (error) {
            showToast('Error guardando gasto');
            return;
        }

        toggleModal('modal-manual', false);
        document.getElementById('form-manual').reset();
        await loadGastos();
        renderGastos();
        showToast('Gasto agregado');
    });

    document.getElementById('prestamo-monto').addEventListener('input', calcularCuota);
    document.getElementById('prestamo-cuotas').addEventListener('input', calcularCuota);
}

function toggleModal(id, show) {
    document.getElementById(id).classList.toggle('hidden', !show);
}

function calcularCuota() {
    const monto = parseFloat(document.getElementById('prestamo-monto').value) || 0;
    const cuotas = parseInt(document.getElementById('prestamo-cuotas').value) || 1;
    document.getElementById('monto-cuota').textContent = formatCurrency(monto / cuotas);
}

// ============================================
// ACCIONES
// ============================================
window.editGasto = async function(id) {
    const gasto = gastos.find(g => g.id === id);
    if (!gasto) return;

    document.getElementById('gasto-id').value = gasto.id;
    document.getElementById('gasto-desc').value = gasto.descripcion;
    document.getElementById('gasto-monto').value = formatCurrency(gasto.monto);
    document.getElementById('gasto-porcentaje').value = gasto.porcentaje_usuario;
    document.getElementById('gasto-categoria').value = gasto.categoria;
    
    toggleModal('modal-gasto', true);
};

window.retirarCuenta = function(id) {
    const cuenta = cuentasAhorro.find(c => c.id === id);
    if (!cuenta) return;

    document.getElementById('prestamo-cuenta-id').value = cuenta.id;
    document.getElementById('prestamo-cuenta-nombre').value = cuenta.nombre;
    toggleModal('modal-prestamo', true);
};

window.pagarCuota = async function(prestamoId) {
    const prestamo = prestamos.find(p => p.id === prestamoId);
    if (!prestamo) return;

    const nuevasCuotas = prestamo.cuotas_pagadas + 1;
    const estado = nuevasCuotas >= prestamo.cuotas_total ? 'pagado' : 'activo';

    const { error } = await supabase
        .from('prestamos')
        .update({ cuotas_pagadas: nuevasCuotas, estado: estado })
        .eq('id', prestamoId);

    if (error) {
        showToast('Error registrando pago');
        return;
    }

    await loadPrestamos();
    renderAhorro();
    showToast('Cuota pagada exitosamente');
};

// ============================================
// UTILIDADES
// ============================================
function formatCurrency(value) {
    return new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', maximumFractionDigits: 0 }).format(value);
}

function formatDate(dateStr) {
    return new Date(dateStr).toLocaleDateString('es-ES', { day: 'numeric', month: 'short', year: 'numeric' });
}

function getCatEmoji(cat) {
    const emojis = {
        supermercado: '🛒', transporte: '🚗', servicios: '💡', salud: '🏥',
        entretenimiento: '🎬', restaurantes: '🍽️', hogar: '🏠', educacion: '📚', otros: '📦'
    };
    return emojis[cat] || '📦';
}

function showToast(message) {
    const toast = document.createElement('div');
    toast.className = 'fixed bottom-4 left-4 right-4 bg-slate-800 border border-slate-600 text-white px-4 py-3 rounded-xl shadow-2xl z-50 text-sm text-center';
    toast.textContent = message;
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(20px)';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ============================================
// SERVICE WORKER REGISTRATION (PWA)
// ============================================
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(err => console.log('SW registration failed'));
}