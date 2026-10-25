// Librerías pesadas: se cargan solo cuando hacen falta (no bloquean el inicio)
const LIBS = {
    xlsx: 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js',
    chart: 'https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.js'
};
const _libPromises = {};
function cargarLib(url) {
    if (!_libPromises[url]) {
        _libPromises[url] = new Promise((ok, fail) => {
            const sc = document.createElement('script');
            sc.src = url;
            sc.onload = ok;
            sc.onerror = () => { delete _libPromises[url]; fail(new Error(url)); };
            document.head.appendChild(sc);
        });
    }
    return _libPromises[url];
}
const enReposo = window.requestIdleCallback ? (f) => requestIdleCallback(f) : (f) => setTimeout(f, 1500);

// ===== Semana (Viernes -> Jueves), límites y fechas =====
const MAX_SESIONES_DIA = 6;
const DIA_INICIO_SEMANA = 5; // getDay(): 0=Dom ... 5=Vie ... 6=Sáb
const NOMBRES_DIA = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];

// Fecha local en formato YYYY-MM-DD (evita el desfase UTC de toISOString/valueAsDate)
function ymd(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function fmtDM(d) {
    return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}`;
}
// Medianoche local del viernes que inicia la semana de "fecha"
function inicioSemanaDe(fecha) {
    const d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
    d.setDate(d.getDate() - ((d.getDay() - DIA_INICIO_SEMANA + 7) % 7));
    return d;
}
function fechasSemana() {
    return Array.from({ length: 7 }, (_, i) => {
        const d = new Date(semanaGrilla);
        d.setDate(d.getDate() + i);
        return d;
    });
}

let modoRegistro = 'fecha', grillaSucia = false, semanaGrilla = inicioSemanaDe(new Date());

const DB_NAME = 'SesionesAppDB';
const STORE_NAME = 'sesiones_registradas';
let db, miGrafico;

const request = indexedDB.open(DB_NAME, 1);
request.onupgradeneeded = (e) => {
    db = e.target.result;
    if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
    }
};
request.onsuccess = (e) => {
    db = e.target.result;
    cargarTemaPreferido();
    setModoRegistro(localStorage.getItem('sesiones-modo') || 'fecha', true);
    actualizarTodo();
    // Precarga en segundo plano para que Excel/Gráfico respondan rápido
    enReposo(() => { cargarLib(LIBS.chart).catch(() => {}); cargarLib(LIBS.xlsx).catch(() => {}); });
};

document.getElementById('fechaRegistro').value = ymd(new Date());

function actualizarTodo() {
    const fDesde = document.getElementById('filtroDesde').value;
    const fHasta = document.getElementById('filtroHasta').value;

    db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => {
        const todos = e.target.result;
        let filtrados = todos;
        if (fDesde) filtrados = filtrados.filter(r => r.fecha >= fDesde);
        if (fHasta) filtrados = filtrados.filter(r => r.fecha <= fHasta);

        calcularDashboard(todos, filtrados, fDesde, fHasta);
        renderizarTabla(filtrados);
        if (modoRegistro === 'grilla' && !grillaSucia) renderizarGrilla(todos);
        document.getElementById('badgeFiltro').style.display = (fDesde || fHasta) ? 'inline' : 'none';
    };
}

function calcularDashboard(totalDB, filtrados, fDesde, fHasta) {
    const ahora = new Date();
    const inicioSemana = inicioSemanaDe(ahora); // semana Vie -> Jue
    const inicioSemanaPasada = new Date(inicioSemana);
    inicioSemanaPasada.setDate(inicioSemana.getDate() - 7);
    const finSemanaPasada = new Date(inicioSemana);
    finSemanaPasada.setMilliseconds(-1);
    const inicioMes = new Date(ahora.getFullYear(), ahora.getMonth(), 1);

    let sMin = 0, sCnt = 0, sPasadaCnt = 0, mMin = 0, mCnt = 0;

    totalDB.forEach(r => {
        const f = new Date(r.fecha + "T12:00:00");
        if (f >= inicioSemana) { sMin += r.minutos; sCnt++; }
        if (f >= inicioSemanaPasada && f <= finSemanaPasada) { sPasadaCnt++; }
        if (f >= inicioMes) { mMin += r.minutos; mCnt++; }
    });

    document.getElementById('statSemana').innerText = fmtHM(sMin);
    document.getElementById('countSemana').innerText = `${sCnt} sesiones`;
    document.getElementById('statMes').innerText = fmtHM(mMin);
    document.getElementById('countMes').innerText = `${mCnt} sesiones`;
    document.getElementById('statTotal').innerText = fmtHM(totalDB.reduce((s, r) => s + r.minutos, 0));
    document.getElementById('countTotal').innerText = `${totalDB.length} sesiones total`;

    document.getElementById('avgSemanaActual').innerText = (sCnt / 7).toFixed(1);
    document.getElementById('avgSemanaPasada').innerText = (sPasadaCnt / 7).toFixed(1);

    if (fDesde && fHasta) {
        const d1 = new Date(fDesde), d2 = new Date(fHasta);
        const dias = Math.max(1, Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1);
        document.getElementById('avgFiltro').innerText = (filtrados.length / dias).toFixed(1);
    } else {
        document.getElementById('avgFiltro').innerText = "0.0";
    }
}

// ORDEN CRONOLÓGICO: Antiguo -> Actual
function renderizarTabla(registros) {
    const cronologico = [...registros].sort((a,b) => a.fecha.localeCompare(b.fecha) || a.ts - b.ts);
    
    let suma = 0;
    const tbody = document.querySelector('#tablaRegistros tbody');
    tbody.innerHTML = '';
    
    if(cronologico.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center">Sin datos</td></tr>';
        return;
    }

    const filas = cronologico.map(item => {
        suma += item.minutos;
        return `<tr>
            <td>${item.fecha.split('-').reverse().join('/')}</td>
            <td>${item.minutos} min</td>
            <td>${fmtHM(suma)}</td>
            <td><button class="btn-delete" onclick="eliminarSesion(${item.id})">Borrar</button></td>
        </tr>`;
    });
    tbody.innerHTML = filas.join('');
}

async function exportarExcelFiltrado() {
    try { await cargarLib(LIBS.xlsx); }
    catch { return alert("No se pudo cargar la librería de Excel. Revisá tu conexión."); }
    const d = document.getElementById('filtroDesde').value;
    const h = document.getElementById('filtroHasta').value;
    db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => {
        let res = e.target.result;
        if (d) res = res.filter(r => r.fecha >= d);
        if (h) res = res.filter(r => r.fecha <= h);
        
        // Orden cronológico para Excel
        res.sort((a,b) => a.fecha.localeCompare(b.fecha) || a.ts - a.ts);
        
        let suma = 0;
        const data = res.map(r => {
            suma += r.minutos;
            return {
                "Fecha": new Date(r.fecha + "T12:00:00"),
                "Sesión (min)": r.minutos,
                "Suma Progresiva (min)": suma,
                "Suma Progresiva (HH:MM)": fmtHM(suma)
            };
        });

        const ws = XLSX.utils.json_to_sheet(data, { cellDates: true });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Reporte Cronológico");
        XLSX.writeFile(wb, `Sesiones_Reporte_Progresivo.xlsx`);
    };
}

async function ejecutarReporteGrafico() {
    try { await cargarLib(LIBS.chart); }
    catch { return alert("No se pudo cargar la librería de gráficos. Revisá tu conexión."); }
    const d = document.getElementById('filtroDesde').value;
    const h = document.getElementById('filtroHasta').value;
    db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => {
        let res = e.target.result;
        if (d) res = res.filter(r => r.fecha >= d);
        if (h) res = res.filter(r => r.fecha <= h);
        
        if (res.length === 0 && !(d && h)) return alert("Sin datos");
        
        res.sort((a,b) => a.fecha.localeCompare(b.fecha) || a.ts - a.ts);
        
        document.getElementById('chartCard').style.display = 'block';
        const totalMin = res.reduce((s, r) => s + r.minutos, 0);
        const out = document.getElementById('reporteOutput');
        out.style.display = 'block';
        out.innerHTML = `ANÁLISIS CRONOLÓGICO: ${res.length} sesiones | Total: ${fmtHM(totalMin)}`;
        
        dibujarGrafico(res, d, h);
    };
}

function dibujarGrafico(registros, desde, hasta) {
    const agrupado = {};
    registros.forEach(r => {
        if (!agrupado[r.fecha]) agrupado[r.fecha] = { s: 0, m: 0 };
        agrupado[r.fecha].s++; agrupado[r.fecha].m += r.minutos;
    });
    
    // Eje X continuo: todos los días del rango, también los sin actividad (valor 0)
    const conActividad = Object.keys(agrupado).sort();
    const inicio = desde || conActividad[0];
    const fin = hasta || conActividad[conActividad.length - 1];
    const labels = [];
    const cur = new Date(inicio + 'T12:00:00'), finD = new Date(fin + 'T12:00:00');
    for (; cur <= finD; cur.setDate(cur.getDate() + 1)) labels.push(ymd(cur));
    const accent = getComputedStyle(document.body).getPropertyValue('--accent').trim();
    if (miGrafico) miGrafico.destroy();
    
    const ctx = document.getElementById('graficoSesiones').getContext('2d');
    miGrafico = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels.map(l => fmtDM(new Date(l + 'T12:00:00'))),
            datasets: [
                { label: 'Sesiones/Día', data: labels.map(l => agrupado[l] ? agrupado[l].s : 0), borderColor: accent, tension: 0.3, yAxisID: 'y' },
                { label: 'Horas Totales/Día', data: labels.map(l => agrupado[l] ? Number((agrupado[l].m / 60).toFixed(2)) : 0), borderColor: '#27ae60', tension: 0.3, yAxisID: 'y1' }
            ]
        },
        options: { scales: { y: { beginAtZero: true, ticks: { precision: 0 } }, y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false } } } }
    });
}

function guardarSesion() {
    const fecha = document.getElementById('fechaRegistro').value;
    const minutos = parseInt(document.getElementById('minutosSesion').value);
    if (!fecha || !minutos || minutos < 1) return;
    const tx = db.transaction([STORE_NAME], 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    store.getAll().onsuccess = (e) => {
        const delDia = e.target.result.filter(r => r.fecha === fecha).length;
        if (delDia >= MAX_SESIONES_DIA) {
            tx.abort();
            return alert(`Ese día ya tiene ${MAX_SESIONES_DIA} sesiones (máximo diario).`);
        }
        store.add({ fecha, minutos, ts: Date.now() });
    };
    tx.oncomplete = () => { document.getElementById('minutosSesion').value = ''; actualizarTodo(); };
}

function eliminarSesion(id) {
    if(confirm("¿Eliminar?")) db.transaction([STORE_NAME], 'readwrite').objectStore(STORE_NAME).delete(id).onsuccess = actualizarTodo;
}

function fmtHM(m) {
    const h = Math.floor(m/60);
    const min = Math.round(m%60);
    return h > 0 ? `${h}h ${min}m` : `${min}m`;
}

function cambiarTema(t) { document.body.className = t; localStorage.setItem('sesiones-tema', t); actualizarTodo(); }
function cargarTemaPreferido() { const t = localStorage.getItem('sesiones-tema') || 'theme-light'; document.body.className = t; document.getElementById('selectTema').value = t; }
function descargarRespaldo() { db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => { const blob = new Blob([JSON.stringify(e.target.result, null, 2)], {type: "application/json"}); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `backup.json`; a.click(); }; }
function importarBackup(event) { const lector = new FileReader(); lector.onload = (e) => { const datos = JSON.parse(e.target.result); const tx = db.transaction([STORE_NAME], 'readwrite'); datos.forEach(d => { delete d.id; tx.objectStore(STORE_NAME).add(d); }); tx.oncomplete = () => { alert("Importado"); actualizarTodo(); }; }; lector.readAsText(event.target.files[0]); }

// ===================== MODO DE REGISTRO + GRILLA =====================
function confirmarDescarte() {
    return !grillaSucia || confirm('Hay cambios sin guardar en la grilla. ¿Descartarlos?');
}

function setModoRegistro(modo, inicial) {
    if (!inicial) {
        if (modo === modoRegistro) return;
        if (modoRegistro === 'grilla' && !confirmarDescarte()) return;
    }
    modoRegistro = modo;
    grillaSucia = false;
    localStorage.setItem('sesiones-modo', modo);
    const g = modo === 'grilla';
    document.getElementById('btnModoFecha').classList.toggle('activo', !g);
    document.getElementById('btnModoGrilla').classList.toggle('activo', g);
    document.getElementById('formFecha').style.display = g ? 'none' : '';
    document.getElementById('hintGrilla').style.display = g ? 'block' : 'none';
    const card = document.getElementById('grillaCard');
    card.style.display = g ? 'block' : 'none';
    if (g && !inicial) {
        renderizarGrilla();
        card.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function cambiarSemanaGrilla(delta) {
    if (!confirmarDescarte()) return;
    semanaGrilla = new Date(semanaGrilla);
    semanaGrilla.setDate(semanaGrilla.getDate() + 7 * delta);
    grillaSucia = false;
    setGrillaMsg('');
    renderizarGrilla();
}
function irASemanaDeFecha(valor) {
    if (!valor || !confirmarDescarte()) return;
    semanaGrilla = inicioSemanaDe(new Date(valor + 'T12:00:00'));
    grillaSucia = false;
    setGrillaMsg('');
    renderizarGrilla();
}
function irASemanaActual() {
    if (!confirmarDescarte()) return;
    semanaGrilla = inicioSemanaDe(new Date());
    grillaSucia = false;
    setGrillaMsg('');
    renderizarGrilla();
}

function setGrillaMsg(texto, ok) {
    const el = document.getElementById('grillaMsg');
    el.textContent = texto;
    el.className = ok ? 'ok' : '';
}

function renderizarGrilla(todos) {
    if (!todos) {
        db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => renderizarGrilla(e.target.result);
        return;
    }
    const fechas = fechasSemana();
    const claves = fechas.map(ymd);
    const hoy = ymd(new Date());
    const porDia = claves.map(k => todos.filter(r => r.fecha === k).sort((a, b) => (a.ts || 0) - (b.ts || 0) || a.id - b.id));
    const columnas = Math.max(MAX_SESIONES_DIA, ...porDia.map(a => a.length)); // sesiones por día

    document.getElementById('grillaRango').innerText =
        `${NOMBRES_DIA[fechas[0].getDay()]} ${fmtDM(fechas[0])} → ${NOMBRES_DIA[fechas[6].getDay()]} ${fmtDM(fechas[6])}`;

    // Filas = días (Vie -> Jue), columnas = sesiones 1..N, última columna = total del día
    let html = '<thead><tr><th class="col-dia">Día</th>' +
        Array.from({ length: columnas }, (_, r) => `<th>${r + 1}</th>`).join('') +
        '<th>Total</th></tr></thead><tbody>';

    for (let c = 0; c < 7; c++) {
        html += `<tr><th class="col-dia ${claves[c] === hoy ? 'hoy' : ''}">${NOMBRES_DIA[fechas[c].getDay()]}<small>${fmtDM(fechas[c])}</small></th>`;
        for (let r = 0; r < columnas; r++) {
            const s = porDia[c][r];
            const bloqueada = !s && r >= MAX_SESIONES_DIA;
            const datos = s ? `data-id="${s.id}" data-orig="${s.minutos}" value="${s.minutos}"` : '';
            html += `<td class="${claves[c] === hoy ? 'hoy' : ''}"><input type="number" min="1" max="1440" step="1" inputmode="numeric" `
                  + `data-dia="${c}" data-fila="${r}" ${datos} ${bloqueada ? 'disabled' : ''} aria-label="${NOMBRES_DIA[fechas[c].getDay()]} sesión ${r + 1}"></td>`;
        }
        html += `<td class="tot" id="totDia${c}"></td></tr>`;
    }
    html += '</tbody>';

    document.getElementById('tablaGrilla').innerHTML = html;
    actualizarTotalesGrilla();
}

function actualizarTotalesGrilla() {
    const sumas = Array(7).fill(0), cnts = Array(7).fill(0);
    document.querySelectorAll('#tablaGrilla input').forEach(inp => {
        const n = parseInt(inp.value, 10);
        if (n > 0) { sumas[+inp.dataset.dia] += n; cnts[+inp.dataset.dia]++; }
    });
    for (let c = 0; c < 7; c++) {
        const el = document.getElementById('totDia' + c);
        if (el) el.innerHTML = cnts[c] ? `<strong>${fmtHM(sumas[c])}</strong><small>${cnts[c]}/${MAX_SESIONES_DIA}</small>` : '<small>–</small>';
    }
    const tc = cnts.reduce((a, b) => a + b, 0), tm = sumas.reduce((a, b) => a + b, 0);
    document.getElementById('grillaTotalSemana').innerText = `${tc} sesiones · ${fmtHM(tm)}`;
}

// Guarda solo las diferencias: celda nueva = alta, celda cambiada = modificación, celda vaciada = baja
function guardarGrilla() {
    const tabla = document.getElementById('tablaGrilla');
    const fechas = fechasSemana().map(ymd);
    const base = Date.now();
    const altas = [], cambios = [], bajas = [];

    for (const inp of tabla.querySelectorAll('input')) {
        const txt = inp.value.trim();
        const dia = +inp.dataset.dia, fila = +inp.dataset.fila;
        const id = inp.dataset.id ? +inp.dataset.id : null;
        const orig = inp.dataset.orig ? +inp.dataset.orig : null;
        const n = txt === '' ? null : Number(txt);
        if (n !== null && (!Number.isInteger(n) || n < 1 || n > 1440)) {
            inp.focus();
            return alert('Los minutos deben ser un número entero entre 1 y 1440.');
        }
        if (id === null) { if (n !== null) altas.push({ fecha: fechas[dia], minutos: n, ts: base + fila }); }
        else if (n === null) bajas.push(id);
        else if (n !== orig) cambios.push({ id, fecha: fechas[dia], minutos: n });
    }

    if (!altas.length && !cambios.length && !bajas.length) return setGrillaMsg('No hay cambios para guardar.');
    if (bajas.length && !confirm(`Se eliminarán ${bajas.length} sesión(es) cuyas celdas dejaste vacías. ¿Continuar?`)) return;

    const tx = db.transaction([STORE_NAME], 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    altas.forEach(a => store.add(a));
    bajas.forEach(id => store.delete(id));
    cambios.forEach(c => {
        store.get(c.id).onsuccess = (e) => {
            const rec = e.target.result;
            if (rec) { rec.minutos = c.minutos; store.put(rec); }
            else store.add({ fecha: c.fecha, minutos: c.minutos, ts: base });
        };
    });
    tx.oncomplete = () => {
        grillaSucia = false;
        setGrillaMsg(`✓ Guardado: ${altas.length} nueva(s), ${cambios.length} modificada(s), ${bajas.length} eliminada(s).`, true);
        actualizarTodo();
    };
    tx.onerror = () => alert('No se pudo guardar. Intentá de nuevo.');
}

// Eventos de la grilla (delegados)
(function () {
    const tabla = document.getElementById('tablaGrilla');
    tabla.addEventListener('input', (e) => {
        const inp = e.target;
        if (inp.tagName !== 'INPUT') return;
        inp.classList.toggle('cambiado', inp.value.trim() !== (inp.dataset.orig ?? ''));
        grillaSucia = !!tabla.querySelector('input.cambiado');
        setGrillaMsg(grillaSucia ? 'Cambios sin guardar' : '');
        actualizarTotalesGrilla();
    });
    // Enter: avanza de izquierda a derecha y, al terminar el día, pasa al siguiente
    tabla.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter' || e.target.tagName !== 'INPUT') return;
        e.preventDefault();
        const inputs = [...tabla.querySelectorAll('input:not([disabled])')]
            .sort((a, b) => a.dataset.dia - b.dataset.dia || a.dataset.fila - b.dataset.fila);
        const sig = inputs[inputs.indexOf(e.target) + 1];
        if (sig) { sig.focus(); sig.select(); }
    });
})();
