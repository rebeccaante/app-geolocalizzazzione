// 1. INIZIALIZZAZIONE SUPABASE
const SUPABASE_URL = 'https://fasjhqmnzxhtnpiuvzdk.supabase.co';
const SUPABASE_KEY = 'sb_publishable_EqZTCd8mOinAbEXk2kYoIw_PTAOvGOv';
const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// 2. POPOLAMENTO ANNI (2026 - 2060)
document.addEventListener('DOMContentLoaded', () => {
    const selectAnno = document.getElementById('selectAnno');
    if (selectAnno) {
        selectAnno.innerHTML = ''; 
        const annoCorrente = new Date().getFullYear();
        
        for (let anno = 2026; anno <= 2060; anno++) {
            const option = document.createElement('option');
            option.value = anno;
            option.textContent = anno;
            if (anno === annoCorrente) option.selected = true; 
            selectAnno.appendChild(option);
        }
    }
});

// 3. REGOLAMENTO QTECH S.R.L.
function calcolaOreTarget(km, pernottamento) {
    if (pernottamento || km <= 40) return 8.0;   // 0-40 km -> 8h
    if (km <= 100) return 7.5;                   // 41-100 km -> 7.5h
    if (km <= 150) return 6.5;                   // 101-150 km -> 6.5h
    return 8.0;                                  // Oltre 150 km -> 8h dal 2° gg
}

// 4. CONSULTAZIONE STORICO MESE / ANNO DA SUPABASE
document.getElementById('btnCaricaStorico').addEventListener('click', async () => {
    const mese = parseInt(document.getElementById('selectMese').value);
    const anno = parseInt(document.getElementById('selectAnno').value);
    
    const { data: report, error } = await _supabase
        .from('report_mensili')
        .select('*')
        .eq('anno', anno)
        .eq('mese', mese)
        .maybeSingle();

    if (error || !report) {
        alert('Nessun dato registrato per il mese e anno selezionati.');
        document.getElementById('tableBody').innerHTML = '';
        document.getElementById('tableFooter').innerHTML = '';
        return;
    }

    const { data: timbrature } = await _supabase
        .from('timbrature_elaborate')
        .select('*')
        .eq('report_id', report.id)
        .order('giorno_entrata', { ascending: true });

    renderTabella(timbrature || []);
});

// 5. CARICAMENTO, LETTURA EXCEL ED ELABORAZIONE
document.getElementById('btnElabora').addEventListener('click', async () => {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Seleziona prima un file Excel!');
        return;
    }

    const btn = document.getElementById('btnElabora');
    btn.disabled = true;
    btn.innerText = 'Elaborazione in corso...';

    const reader = new FileReader();

    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const righeGrezze = XLSX.utils.sheet_to_json(firstSheet);

            if (!righeGrezze || righeGrezze.length === 0) {
                alert('Il file caricato è vuoto o non è stato letto correttamente.');
                return;
            }

            // Anagrafica Sedi Dipendenti
            const { data: dipendentiAnagrafica } = await _supabase.from('dipendenti').select('*');
            const mappaSedi = {};
            if (dipendentiAnagrafica) {
                dipendentiAnagrafica.forEach(d => {
                    mappaSedi[d.nome_cognome.trim().toLowerCase()] = d.sede_appartenenza;
                });
            }

            const mese = parseInt(document.getElementById('selectMese').value);
            const anno = parseInt(document.getElementById('selectAnno').value);

            // Inserisci o Recupera Report Mensile
            let { data: report } = await _supabase
                .from('report_mensili')
                .select('id')
                .eq('anno', anno)
                .eq('mese', mese)
                .maybeSingle();

            if (!report) {
                const { data: newReport, error: errIns } = await _supabase
                    .from('report_mensili')
                    .insert([{ anno, mese }])
                    .select()
                    .single();
                if (errIns) throw errIns;
                report = newReport;
            }

            // Elaborazione Turni e Anomalie
            const risultati = [];
            let cantiereGiornoPrecedente = {};

            for (let i = 0; i < righeGrezze.length; i++) {
                const r = righeGrezze[i];
                const nomeDip = (r.Dipendente || r.Nome || r.dipendente || '').toString().trim();
                const verso = (r.Verso || r.verso || '').toString().toUpperCase().trim();

                if (verso === 'ENTRATA') {
                    const dtIn = new Date(r.DataOra || r.Data || r.ora);
                    let dtOut = null;
                    let erroreTimbratura = false;

                    // Cerca l'uscita corrispondente
                    for (let j = i + 1; j < righeGrezze.length; j++) {
                        const prox = righeGrezze[j];
                        const proxNome = (prox.Dipendente || prox.Nome || prox.dipendente || '').toString().trim();
                        if (proxNome !== nomeDip) break;

                        const proxVerso = (prox.Verso || prox.verso || '').toString().toUpperCase().trim();
                        if (proxVerso === 'ENTRATA') {
                            erroreTimbratura = true;
                            break;
                        }
                        if (proxVerso === 'USCITA') {
                            dtOut = new Date(prox.DataOra || prox.Data || prox.ora);
                            break;
                        }
                    }

                    if (!dtOut) erroreTimbratura = true;

                    const cantiereAttuale = (r.Cantiere || r.Luogo || r.Citta || 'Cantiere').toString().trim();
                    const dipKey = nomeDip.toLowerCase();
                    const sedePura = mappaSedi[dipKey] || 'Milano (MI)';

                    let isPernottamento = false;
                    if (cantiereGiornoPrecedente[dipKey] && cantiereGiornoPrecedente[dipKey] === cantiereAttuale) {
                        isPernottamento = true;
                    }
                    cantiereGiornoPrecedente[dipKey] = cantiereAttuale;

                    const km = isPernottamento ? 0 : parseFloat(r.Km || r.km || 0);
                    const oreTarget = calcolaOreTarget(km, isPernottamento);

                    let oreEffettuate = 0;
                    if (dtOut && !erroreTimbratura) {
                        const diffMs = dtOut - dtIn;
                        oreEffettuate = Math.round((diffMs / (1000 * 60 * 60)) * 100) / 100;
                    }

                    risultati.push({
                        report_id: report.id,
                        dipendente: nomeDip,
                        giorno_entrata: dtIn.toISOString().split('T')[0],
                        ora_entrata: dtIn.toTimeString().split(' ')[0].substring(0, 5),
                        giorno_uscita: dtOut ? dtOut.toISOString().split('T')[0] : null,
                        ora_uscita: dtOut ? dtOut.toTimeString().split(' ')[0].substring(0, 5) : null,
                        luogo_partenza: isPernottamento ? 'Pernottamento Cantiere' : sedePura,
                        luogo_cantiere: cantiereAttuale,
                        km_percorsi: km,
                        ore_target: oreTarget,
                        ore_effettuate: erroreTimbratura ? 0 : oreEffettuate,
                        errore_timbratura: erroreTimbratura
                    });
                }
            }

            if (risultati.length > 0) {
                await _supabase.from('timbrature_elaborate').delete().eq('report_id', report.id);
                const { error: errDB } = await _supabase.from('timbrature_elaborate').insert(risultati);
                if (errDB) throw errDB;
            }

            renderTabella(risultati);
            alert('Elaborazione completata con successo!');

        } catch (err) {
            console.error('Errore:', err);
            alert('Si è verificato un errore durante l\'elaborazione: ' + (err.message || err));
        } finally {
            btn.disabled = false;
            btn.innerText = 'Elabora e Salva';
        }
    };

    reader.readAsArrayBuffer(file);
});

// 6. DISEGNO TABELLA E CALCOLO MEDIE
function renderTabella(dati) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    let sommaTarget = 0;
    let sommaEffettuate = 0;
    let cntTarget = 0;
    let cntEffettuate = 0;

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

        if (riga.ore_target > 0) {
            sommaTarget += parseFloat(riga.ore_target);
            cntTarget++;
        }
        if (!riga.errore_timbratura && riga.ore_effettuate > 0) {
            sommaEffettuate += parseFloat(riga.ore_effettuate);
            cntEffettuate++;
        }
    });

    const mediaTarget = cntTarget > 0 ? (sommaTarget / cntTarget).toFixed(2) : 0;
    const mediaEffettuate = cntEffettuate > 0 ? (sommaEffettuate / cntEffettuate).toFixed(2) : 0;

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
