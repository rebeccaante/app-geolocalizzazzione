// Popola automaticamente il menu degli anni dal 2024 al 2060
const selectAnno = document.getElementById('selectAnno');
if (selectAnno) {
    selectAnno.innerHTML = ''; // Svuota opzioni esistenti
    const annoCorrente = new Date().getFullYear();
    
    for (let anno = 2024; anno <= 2060; anno++) {
        const option = document.createElement('option');
        option.value = anno;
        option.textContent = anno;
        if (anno === annoCorrente) option.selected = true; // Seleziona l'anno corrente di default
        selectAnno.appendChild(option);
    }
}
const SUPABASE_URL = 'https://fasjhqmnzxhtnpiuvzdk.supabase.co';
const SUPABASE_KEY = 'sb_publishable_EqZTCd8mOinAbEXk2kYoIw_PTAOvGOv';
const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// Regolamento QTECH S.r.l.
function calcolaOreTarget(km, pernottamento) {
    if (pernottamento || km <= 40) return 8.0;   // 0-40 km -> 8h[cite: 1]
    if (km <= 100) return 7.5;                   // 41-100 km -> 7.5h[cite: 1]
    if (km <= 150) return 6.5;                   // 101-150 km -> 6.5h[cite: 1]
    return 8.0;                                  // Oltre 150 km dal 2° gg -> 8h[cite: 1]
}

// Visualizzazione dello Storico da Supabase
document.getElementById('btnCaricaStorico').addEventListener('click', async () => {
    const mese = parseInt(document.getElementById('selectMese').value);
    const anno = parseInt(document.getElementById('selectAnno').value);
    
    // Cerca il report salvato
    const { data: report } = await _supabase
        .from('report_mensili')
        .select('*')
        .eq('anno', anno)
        .eq('mese', mese)
        .maybeSingle();

    if (!report) {
        alert('Nessun dato registrato per il mese e anno selezionati.');
        document.getElementById('tableBody').innerHTML = '';
        document.getElementById('tableFooter').innerHTML = '';
        return;
    }

    // Carica i dati delle timbrature
    const { data: timbrature } = await _supabase
        .from('timbrature_elaborate')
        .select('*')
        .eq('report_id', report.id)
        .order('giorno_entrata', { ascending: true });

    renderTabella(timbrature);
});

// Funzione di Disegno della Tabella e Calcolo delle Medie Generali
function renderTabella(dati) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    let sommaTarget = 0;
    let sommaEffettuate = 0;
    let conteggioTarget = 0;
    let conteggioEffettuate = 0;

    dati.forEach(riga => {
        const tr = document.createElement('tr');
        if (riga.errore_timbratura) tr.classList.add('bg-red-50');

        tr.innerHTML = `
            <td class="p-3 font-medium">${riga.dipendente}</td>
            <td class="p-3">${riga.giorno_entrata}</td>
            <td class="p-3">${riga.ora_entrata}</td>
            <td class="p-3">${riga.giorno_uscita || '-'}</td>
            <td class="p-3">${riga.ora_uscita || '-'}</td>
            <td class="p-3">${riga.luogo_partenza}</td>
            <td class="p-3">${riga.luogo_cantiere}</td>
            <td class="p-3 text-center">${riga.km_percorsi} km</td>
            <td class="p-3 text-center font-semibold">${riga.ore_target} h</td>
            <td class="p-3 text-center font-semibold">${riga.errore_timbratura ? '-' : riga.ore_effettuate + ' h'}</td>
            <td class="p-3 text-xs ${riga.errore_timbratura ? 'text-red-600 font-bold' : ''}">
                ${riga.errore_timbratura ? 'Errore nella timbratura' : ''}
            </td>
        `;
        tbody.appendChild(tr);

        // Accumulo per Media Generale Totale
        if (riga.ore_target > 0) {
            sommaTarget += parseFloat(riga.ore_target);
            conteggioTarget++;
        }
        if (!riga.errore_timbratura && riga.ore_effettuate > 0) {
            sommaEffettuate += parseFloat(riga.ore_effettuate);
            conteggioEffettuate++;
        }
    });

    // Riga delle Medie Totali in Fondo
    const mediaTarget = conteggioTarget > 0 ? (sommaTarget / conteggioTarget).toFixed(2) : 0;
    const mediaEffettuate = conteggioEffettuate > 0 ? (sommaEffettuate / conteggioEffettuate).toFixed(2) : 0;

    const tfoot = document.getElementById('tableFooter');
    tfoot.innerHTML = `
        <tr>
            <td colspan="8" class="p-3 text-right uppercase tracking-wider">Media Totale Tutti i Dipendenti:</td>
            <td class="p-3 text-center text-yellow-300 text-base font-bold">${mediaTarget} h</td>
            <td class="p-3 text-center text-green-300 text-base font-bold">${mediaEffettuate} h</td>
            <td></td>
        </tr>
    `;
}
