// 1. INIZIALIZZAZIONE SUPABASE
const SUPABASE_URL = 'https://fasjhqmnzxhtnpiuvzdk.supabase.co';
const SUPABASE_KEY = 'sb_publishable_EqZTCd8mOinAbEXk2kYoIw_PTAOvGOv';
const _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

const COORDINATE_SEDI = {
    'milano': { lat: 45.5085, lon: 9.2231, nome: 'Milano (Via Gentile 7)' },
    'san giovanni lupatoto': { lat: 45.3855, lon: 11.0362, nome: 'San Giovanni Lupatoto (VR)' },
    'verona': { lat: 45.3855, lon: 11.0362, nome: 'San Giovanni Lupatoto (VR)' }
};

function calcolaDistanzaKm(lat1, lon1, lat2, lon2) {
    if (!lat1 || !lon1 || !lat2 || !lon2) return 0;
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(R * c * 10) / 10;
}

function calcolaOreTarget(km, pernottamento) {
    if (pernottamento || km <= 40) return 8.0;
    if (km <= 100) return 7.5;
    if (km <= 150) return 6.5;
    return 8.0;
}

// 2. NAVIGAZIONE SCHEDE (TABS)
const tabElaborazione = document.getElementById('tabElaborazione');
const tabAnagrafica = document.getElementById('tabAnagrafica');
const sezioneElaborazione = document.getElementById('sezioneElaborazione');
const sezioneAnagrafica = document.getElementById('sezioneAnagrafica');

tabElaborazione.addEventListener('click', () => {
    sezioneElaborazione.classList.remove('hidden');
    sezioneAnagrafica.classList.add('hidden');
    tabElaborazione.className = 'px-4 py-2 text-sm font-semibold rounded-lg bg-blue-600 text-white shadow-sm';
    tabAnagrafica.className = 'px-4 py-2 text-sm font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200';
});

tabAnagrafica.addEventListener('click', () => {
    sezioneAnagrafica.classList.remove('hidden');
    sezioneElaborazione.classList.add('hidden');
    tabAnagrafica.className = 'px-4 py-2 text-sm font-semibold rounded-lg bg-blue-600 text-white shadow-sm';
    tabElaborazione.className = 'px-4 py-2 text-sm font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200';
    caricaAnagrafica();
});

// 3. INIZIALIZZAZIONE ANNI
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

// 4. GESTIONE ANAGRAFICA DIPENDENTI
async function caricaAnagrafica() {
    const { data: dipendenti } = await _supabase.from('dipendenti').select('*').order('nome_cognome');
    const tbody = document.getElementById('tableAnagraficaBody');
    tbody.innerHTML = '';

    if (dipendenti) {
        dipendenti.forEach(d => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="p-3 font-medium text-slate-800">${d.nome_cognome}</td>
                <td class="p-3 text-slate-600">${d.sede_appartenenza}</td>
                <td class="p-3 text-right">
                    <button onclick="eliminaDipendente('${d.id}')" class="text-red-600 hover:underline text-xs font-bold">Elimina</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    }
}

document.getElementById('btnSalvaDipendente').addEventListener('click', async () => {
    const nome = document.getElementById('inputNomeDipendente').value.trim().toUpperCase();
    const sede = document.getElementById('selectSedeAppartenenza').value;

    if (!nome) {
        alert('Inserisci il nome e cognome del dipendente!');
        return;
    }

    await _supabase.from('dipendenti').upsert([{ nome_cognome: nome, sede_appartenenza: sede }], { onConflict: 'nome_cognome' });
    document.getElementById('inputNomeDipendente').value = '';
    caricaAnagrafica();
    alert('Dipendente salvato con successo!');
});

async function eliminaDipendente(id) {
    if (confirm('Vuoi davvero eliminare questo dipendente dall\'anagrafica?')) {
        await _supabase.from('dipendenti').delete().eq('id', id);
        caricaAnagrafica();
    }
}

// 5. CARICAMENTO ED ELABORAZIONE TIMBRATURE
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
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const righeGrezze = XLSX.utils.sheet_to_json(firstSheet);

            // Carica anagrafica sedi aggiornata da Supabase
            const { data: dipendentiAnagrafica } = await _supabase.from('dipendenti').select('*');
            const mappaSedi = {};
            if (dipendentiAnagrafica) {
                dipendentiAnagrafica.forEach(d => {
                    mappaSedi[d.nome_cognome.trim().toLowerCase()] = d.sede_appartenenza;
                });
            }

            const mese = parseInt(document.getElementById('selectMese').value);
            const anno = parseInt(document.getElementById('selectAnno').value);

            let { data: report } = await _supabase.from('report_mensili').select('id').eq('anno', anno).eq('mese', mese).maybeSingle();
            if (!report) {
                const { data: newReport, error: errIns } = await _supabase.from('report_mensili').insert([{ anno, mese }]).select().single();
                if (errIns) throw errIns;
                report = newReport;
            }

            const righe = righeGrezze.map(r => {
                const cognome = (r['Cognome'] || '').toString().trim();
                const nome = (r['Nome'] || '').toString().trim();
                const dipendente = `${cognome} ${nome}`.trim();
                let dataOra = r['Data Ora'];
                if (dataOra && !(dataOra instanceof Date)) dataOra = new Date(dataOra);
                const verso = (r['Verso'] || '').toString().toUpperCase().trim();
                const lat = parseFloat(r['Latitudine'] || 0);
                const lon = parseFloat(r['Longitudine'] || 0);
                const posizione = (r['Posizione'] || '').toString().trim();
                return { dipendente, verso, dataOra, lat, lon, posizione };
            }).filter(r => r.dipendente && r.dataOra && !isNaN(r.dataOra.getTime()));

            const risultati = [];
            const pad = num => String(num).padStart(2, '0');
            const formatData = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
            const formatOra = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;

            let i = 0;
            let cantiereGiornoPrecedente = {};

            while (i < righe.length) {
                const r = righe[i];

                if (r.verso === 'E' || r.verso.includes('ENTR')) {
                    const dtIn = r.dataOra;
                    const dipKey = r.dipendente.toLowerCase();
                    const nomeSede = mappaSedi[dipKey] || 'Milano (Via Gentile 7)';
                    
                    let coordSede = COORDINATE_SEDI['milano'];
                    if (nomeSede.toLowerCase().includes('lupatoto') || nomeSede.toLowerCase().includes('verona')) {
                        coordSede = COORDINATE_SEDI['san giovanni lupatoto'];
                    }

                    const kmReali = calcolaDistanzaKm(coordSede.lat, coordSede.lon, r.lat, r.lon);

                    let dtOut = null;
                    let coppiaperta = false;

                    if (i + 1 < righe.length) {
                        const prox = righe[i + 1];
                        if (prox.dipendente.toLowerCase() === dipKey && (prox.verso === 'U' || prox.verso.includes('USC'))) {
                            dtOut = prox.dataOra;
                            coppiaperta = true;
                        }
                    }

                    const cantiereAttuale = `${r.lat.toFixed(3)},${r.lon.toFixed(3)}`;
                    let isPernottamento = false;
                    if (cantiereGiornoPrecedente[dipKey] && cantiereGiornoPrecedente[dipKey] === cantiereAttuale) {
                        isPernottamento = true;
                    }
                    cantiereGiornoPrecedente[dipKey] = cantiereAttuale;

                    const kmFinali = isPernottamento ? 0 : kmReali;
                    const oreTarget = calcolaOreTarget(kmFinali, isPernottamento);

                    if (coppiaperta && dtOut) {
                        const diffMs = dtOut - dtIn;
                        const oreEffettuate = Math.round((diffMs / (1000 * 60 * 60)) * 100) / 100;

                        risultati.push({
                            report_id: report.id,
                            dipendente: r.dipendente,
                            giorno_entrata: formatData(dtIn),
                            ora_entrata: formatOra(dtIn),
                            giorno_uscita: formatData(dtOut),
                            ora_uscita: formatOra(dtOut),
                            luogo_partenza: isPernottamento ? 'Pernottamento Cantiere' : coordSede.nome,
                            luogo_cantiere: r.posizione || `${r.lat}, ${r.lon}`,
                            km_percorsi: kmFinali,
                            ore_target: oreTarget,
                            ore_effettuate: oreEffettuate,
                            errore_timbratura: false
                        });
                        i += 2;
                    } else {
                        risultati.push({
                            report_id: report.id,
                            dipendente: r.dipendente,
                            giorno_entrata: formatData(dtIn),
                            ora_entrata: formatOra(dtIn),
                            giorno_uscita: null,
                            ora_uscita: null,
                            luogo_partenza: isPernottamento ? 'Pernottamento Cantiere' : coordSede.nome,
                            luogo_cantiere: r.posizione || `${r.lat}, ${r.lon}`,
                            km_percorsi: kmFinali,
                            ore_target: oreTarget,
                            ore_effettuate: 0,
                            errore_timbratura: true
                        });
                        i += 1;
                    }
                } else {
                    const dipKey = r.dipendente.toLowerCase();
                    const nomeSede = mappaSedi[dipKey] || 'Milano (Via Gentile 7)';
                    risultati.push({
                        report_id: report.id,
                        dipendente: r.dipendente,
                        giorno_entrata: null,
                        ora_entrata: null,
                        giorno_uscita: formatData(r.dataOra),
                        ora_uscita: formatOra(r.dataOra),
                        luogo_partenza: nomeSede,
                        luogo_cantiere: r.posizione || `${r.lat}, ${r.lon}`,
                        km_percorsi: 0,
                        ore_target: 8.0,
                        ore_effettuate: 0,
                        errore_timbratura: true
                    });
                    i += 1;
                }
            }

            if (risultati.length > 0) {
                await _supabase.from('timbrature_elaborate').delete().eq('report_id', report.id);
                await _supabase.from('timbrature_elaborate').insert(risultati);
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

function renderTabella(dati) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    let sommaTarget = 0, sommaEffettuate = 0, cntTarget = 0, cntEffettuate = 0;

    dati.forEach(riga => {
        const tr = document.createElement('tr');
        if (riga.errore_timbratura) tr.classList.add('bg-red-50');

        let cantiereHtml = riga.luogo_cantiere;
        if (riga.luogo_cantiere && riga.luogo_cantiere.startsWith('http')) {
            cantiereHtml = `<a href="${riga.luogo_cantiere}" target="_blank" class="text-blue-600 underline font-medium">Apri Mappa 📍</a>`;
        }

        tr.innerHTML = `
            <td class="p-3 font-medium">${riga.dipendente}</td>
            <td class="p-3">${riga.giorno_entrata || '-'}</td>
            <td class="p-3">${riga.ora_entrata || '-'}</td>
            <td class="p-3">${riga.giorno_uscita || '-'}</td>
            <td class="p-3">${riga.ora_uscita || '-'}</td>
            <td class="p-3 text-xs">${riga.luogo_partenza}</td>
            <td class="p-3 text-xs max-w-xs truncate">${cantiereHtml}</td>
            <td class="p-3 text-center font-medium text-blue-700">${riga.km_percorsi} km</td>
            <td class="p-3 text-center font-semibold">${riga.ore_target} h</td>
            <td class="p-3 text-center font-semibold">${riga.errore_timbratura ? '-' : riga.ore_effettuate + ' h'}</td>
            <td class="p-3 text-xs ${riga.errore_timbratura ? 'text-red-600 font-bold' : ''}">
                ${riga.errore_timbratura ? 'Errore nella timbratura' : ''}
            </td>
        `;
        tbody.appendChild(tr);

        if (riga.ore_target > 0) { sommaTarget += parseFloat(riga.ore_target); cntTarget++; }
        if (!riga.errore_timbratura && riga.ore_effettuate > 0) { sommaEffettuate += parseFloat(riga.ore_effettuate); cntEffettuate++; }
    });

    const mediaTarget = cntTarget > 0 ? (sommaTarget / cntTarget).toFixed(2) : 0;
    const mediaEffettuate = cntEffettuate > 0 ? (sommaEffettuate / cntEffettuate).toFixed(2) : 0;

    const tfoot = document.getElementById('tableFooter');
    tfoot.innerHTML = `
        <tr>
            <td colspan="8" class="p-3 text-right uppercase tracking-wider font-semibold">Media Totale Tutti i Dipendenti:</td>
            <td class="p-3 text-center text-yellow-300 text-base font-bold">${mediaTarget} h</td>
            <td class="p-3 text-center text-green-300 text-base font-bold">${mediaEffettuate} h</td>
            <td></td>
        </tr>
    `;
}
