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
    if (pernottamento || km <= 40) return 8.0;
    if (km <= 100) return 7.5;
    if (km <= 150) return 6.5;
    return 8.0;
}

// Helper per trovare valori nell'Excel indipendentemente dal nome esatto della colonna
function getValoreCoppia(riga, paroleChiave) {
    const chiavi = Object.keys(riga);
    for (let key of chiavi) {
        const keyClean = key.toLowerCase().trim();
        for (let p of paroleChiave) {
            if (keyClean.includes(p)) {
                return riga[key];
            }
        }
    }
    return '';
}

// 4. CONSULTAZIONE STORICO
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

// 5. CARICAMENTO, LETTURA ED ELABORAZIONE FLUIDA
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
                alert('Il file caricato è vuoto.');
                return;
            }

            // Anagrafica Sedi Dipendenti da Supabase
            const { data: dipendentiAnagrafica } = await _supabase.from('dipendenti').select('*');
            const mappaSedi = {};
            if (dipendentiAnagrafica) {
                dipendentiAnagrafica.forEach(d => {
                    mappaSedi[d.nome_cognome.trim().toLowerCase()] = d.sede_appartenenza;
                });
            }

            const mese = parseInt(document.getElementById('selectMese').value);
            const anno = parseInt(document.getElementById('selectAnno').value);

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

            const risultati = [];
            let cantiereGiornoPrecedente = {};

            // Mappatura dinamica delle righe dell'Excel
            const righeNormalizzate = righeGrezze.map(r => ({
                dipendente: String(getValoreCoppia(r, ['dipendente', 'nome', 'cognome', 'user', 'persona'])).trim(),
                verso: String(getValoreCoppia(r, ['verso', 'tipo', 'causale', 'in/out', 'operazione'])).toUpperCase().trim(),
                dataOra: getValoreCoppia(r, ['data', 'ora', 'timestamp', 'time', 'giorno']),
                cantiere: String(getValoreCoppia(r, ['cantiere', 'luogo', 'location', 'indirizzo', 'posto', 'città', 'citta'])).trim(),
                km: parseFloat(getValoreCoppia(r, ['km', 'chilometri', 'distanza']) || 0)
            })).filter(r => r.dipendente && r.dataOra);

            // Ordina cronologicamente per dipendente e orario
            righeNormalizzate.sort((a, b) => new Date(a.dataOra) - new Date(b.dataOra));

            for (let i = 0; i < righeNormalizzate.length; i++) {
                const r = righeNormalizzate[i];
                const verso = r.verso;

                // Accetta sia 'ENTRATA', 'E', 'IN', 'INGRESSO'
                if (verso.includes('ENTR') || verso === 'E' || verso === 'IN' || verso.includes('INGR')) {
                    const dtIn = new Date(r.dataOra);
                    let dtOut = null;
                    let erroreTimbratura = false;

                    // Cerca la corrispondente uscita
                    for (let j = i + 1; j < righeNormalizzate.length; j++) {
                        const prox = righeNormalizzate[j];
                        if (prox.dipendente.toLowerCase() !== r.dipendente.toLowerCase()) break;

                        const proxVerso = prox.verso;
                        if (proxVerso.includes('ENTR') || proxVerso === 'E' || proxVerso === 'IN' || proxVerso.includes('INGR')) {
                            erroreTimbratura = true; // Doppia entrata
                            break;
                        }
                        if (proxVerso.includes('USC') || proxVerso === 'U' || proxVerso === 'OUT' || proxVerso.includes('USCITA')) {
                            dtOut = new Date(prox.dataOra);
                            break;
                        }
                    }

                    if (!dtOut) erroreTimbratura = true;

                    const dipKey = r.dipendente.toLowerCase();
                    const cantiereAttuale = r.cantiere || 'Cantiere';
                    const sedePura = mappaSedi[dipKey] || 'Milano (MI)';

                    let isPernottamento = false;
                    if (cantiereGiornoPrecedente[dipKey] && cantiereGiornoPrecedente[dipKey] === cantiereAttuale) {
                        isPernottamento = true;
                    }
                    cantiereGiornoPrecedente[dipKey] = cantiereAttuale;

                    const km = isPernottamento ? 0 : r.km;
                    const oreTarget = calcolaOreTarget(km, isPernottamento);

                    let oreEffettuate = 0;
                    if (dtOut && !erroreTimbratura) {
                        const diffMs = dtOut - dtIn;
                        oreEffettuate = Math.round((diffMs / (1000 * 60 * 60)) * 100) / 100;
                    }

                    risultati.push({
                        report_id: report.id,
                        dipendente: r.dipendente,
                        giorno_entrata: isNaN(dtIn) ? '-' : dtIn.toISOString().split('T')[0],
                        ora_entrata: isNaN(dtIn) ? '-' : dtIn.toTimeString().split(' ')[0].substring(0, 5),
                        giorno_uscita: (dtOut && !isNaN(dtOut)) ? dtOut.toISOString().split('T')[0] : null,
                        ora_uscita: (dtOut && !isNaN(dtOut)) ? dtOut.toTimeString().split(' ')[0].substring(0, 5) : null,
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
                if (errDB) console.error('Errore Supabase:', errDB);
            }

            renderTabella(risultati);
            alert('Elaborazione completata!');

        } catch (err) {
            console.error(err);
            alert('Errore durante l\'elaborazione: ' + (err.message || err));
        } finally {
            btn.disabled = false;
            btn.innerText = 'Elabora e Salva';
        }
    };

    reader.readAsArrayBuffer(file);
});

// 6. DISEGNO TABELLA E MEDIE
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
