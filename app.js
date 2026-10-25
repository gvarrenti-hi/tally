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
    actualizarTodo();
};

document.getElementById('fechaRegistro').valueAsDate = new Date();

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
        document.getElementById('badgeFiltro').style.display = (fDesde || fHasta) ? 'inline' : 'none';
    };
}

function calcularDashboard(totalDB, filtrados, fDesde, fHasta) {
    const ahora = new Date();
    const inicioSemana = new Date(ahora);
    inicioSemana.setDate(ahora.getDate() - ahora.getDay());
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

    cronologico.forEach(item => {
        suma += item.minutos;
        tbody.innerHTML += `<tr>
            <td>${item.fecha.split('-').reverse().join('/')}</td>
            <td>${item.minutos} min</td>
            <td>${fmtHM(suma)}</td>
            <td><button class="btn-delete" onclick="eliminarSesion(${item.id})">Borrar</button></td>
        </tr>`;
    });
}

function exportarExcelFiltrado() {
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

function ejecutarReporteGrafico() {
    const d = document.getElementById('filtroDesde').value;
    const h = document.getElementById('filtroHasta').value;
    db.transaction([STORE_NAME], 'readonly').objectStore(STORE_NAME).getAll().onsuccess = (e) => {
        let res = e.target.result;
        if (d) res = res.filter(r => r.fecha >= d);
        if (h) res = res.filter(r => r.fecha <= h);
        
        if (res.length === 0) return alert("Sin datos");
        
        res.sort((a,b) => a.fecha.localeCompare(b.fecha) || a.ts - a.ts);
        
        document.getElementById('chartCard').style.display = 'block';
        const totalMin = res.reduce((s, r) => s + r.minutos, 0);
        const out = document.getElementById('reporteOutput');
        out.style.display = 'block';
        out.innerHTML = `ANÁLISIS CRONOLÓGICO: ${res.length} sesiones | Total: ${fmtHM(totalMin)}`;
        
        dibujarGrafico(res);
    };
}

function dibujarGrafico(registros) {
    const agrupado = {};
    registros.forEach(r => {
        if (!agrupado[r.fecha]) agrupado[r.fecha] = { s: 0, m: 0 };
        agrupado[r.fecha].s++; agrupado[r.fecha].m += r.minutos;
    });
    
    const labels = Object.keys(agrupado).sort();
    const accent = getComputedStyle(document.body).getPropertyValue('--accent').trim();
    if (miGrafico) miGrafico.destroy();
    
    const ctx = document.getElementById('graficoSesiones').getContext('2d');
    miGrafico = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels.map(l => l.split('-').reverse().slice(0,2).join('/')),
            datasets: [
                { label: 'Sesiones/Día', data: labels.map(l => agrupado[l].s), borderColor: accent, tension: 0.3, yAxisID: 'y' },
                { label: 'Horas Totales/Día', data: labels.map(l => (agrupado[l].m / 60).toFixed(2)), borderColor: '#27ae60', tension: 0.3, yAxisID: 'y1' }
            ]
        },
        options: { scales: { y1: { position: 'right', grid: { drawOnChartArea: false } } } }
    });
}

async function guardarSesion() {
    const fecha = document.getElementById('fechaRegistro').value;
    const minutos = parseInt(document.getElementById('minutosSesion').value);
    if (!fecha || !minutos) return;
    const tx = db.transaction([STORE_NAME], 'readwrite');
    tx.objectStore(STORE_NAME).add({ fecha, minutos, ts: Date.now() });
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