'use client';

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
type Mode = 'dedupe' | 'merge' | 'merge-dedupe' | 'join';
type JoinParts = { leftOnly: boolean; inner: boolean; rightOnly: boolean };
type TableFile = { name: string; headers: string[]; rows: Record<string, unknown>[] };
const accepted = '.xlsx,.xls,.csv';

async function readExcel(file: File): Promise<TableFile> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' });
  const headers = (matrix[0] ?? []).map((value, index) => String(value || `Colonne ${index + 1}`).trim());
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '', raw: false });
  return { name: file.name, headers, rows };
}
function saveWorkbook(rows: Record<string, unknown>[], name: string, headers?: string[]) {
  const sheet = XLSX.utils.json_to_sheet(rows, headers ? { header: headers } : undefined);
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, 'Résultat'); XLSX.writeFile(book, name);
}

function createJoin(left: TableFile, right: TableFile, leftKey: string, rightKey: string, parts: JoinParts) {
  const rightNames = new Map(right.headers.map((header) => [header, left.headers.includes(header) ? `${header}_droite` : header]));
  const headers = [...left.headers, ...right.headers.map((header) => rightNames.get(header)!)];
  const index = new Map<string, { row: Record<string, unknown>; index: number }[]>();
  right.rows.forEach((row, rowIndex) => { const value = String(row[rightKey] ?? '').trim().toLocaleLowerCase('fr'); const group = index.get(value) ?? []; group.push({ row, index: rowIndex }); index.set(value, group); });
  const matchedRight = new Set<number>(); const rows: Record<string, unknown>[] = []; let matched = 0; let leftOnly = 0;
  const combine = (leftRow?: Record<string, unknown>, rightRow?: Record<string, unknown>) => { const result: Record<string, unknown> = {}; left.headers.forEach((header) => { result[header] = leftRow?.[header] ?? ''; }); right.headers.forEach((header) => { result[rightNames.get(header)!] = rightRow?.[header] ?? ''; }); return result; };
  left.rows.forEach((leftRow) => { const value = String(leftRow[leftKey] ?? '').trim().toLocaleLowerCase('fr'); const matches = index.get(value) ?? []; if (matches.length) { matches.forEach((item) => { if (parts.inner) rows.push(combine(leftRow, item.row)); matchedRight.add(item.index); matched++; }); } else { leftOnly++; if (parts.leftOnly) rows.push(combine(leftRow)); } });
  const rightOnly = right.rows.length - matchedRight.size;
  if (parts.rightOnly) right.rows.forEach((rightRow, index) => { if (!matchedRight.has(index)) rows.push(combine(undefined, rightRow)); });
  return { rows, headers, matched, leftOnly, rightOnly };
}

export default function Home() {
  const [mode, setMode] = useState<Mode>('dedupe'); const [files, setFiles] = useState<TableFile[]>([]);
  const [keys, setKeys] = useState<string[]>([]); const [keep, setKeep] = useState<'first' | 'last'>('first');
  const [leftKey, setLeftKey] = useState(''); const [rightKey, setRightKey] = useState(''); const [joinParts, setJoinParts] = useState<JoinParts>({ leftOnly: false, inner: true, rightOnly: false });
  const [showJoinHelp, setShowJoinHelp] = useState(false);
  const [busy, setBusy] = useState(false); const [dragging, setDragging] = useState(false); const [message, setMessage] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const reset = (next: Mode) => { setMode(next); setFiles([]); setKeys([]); setLeftKey(''); setRightKey(''); setJoinParts({ leftOnly: false, inner: true, rightOnly: false }); setShowJoinHelp(false); setMessage(''); };
  const allCompatible = useMemo(() => { if (files.length < 2) return true; const base = [...files[0].headers].sort().join('\u0000'); return files.every((file) => [...file.headers].sort().join('\u0000') === base); }, [files]);
  const totalRows = files.reduce((sum, file) => sum + file.rows.length, 0);
  const previewStats = useMemo(() => {
    if (!files.length || !keys.length || mode === 'merge' || (mode === 'merge-dedupe' && !allCompatible)) return null;
    const rows = mode === 'dedupe' ? files[0].rows : files.flatMap((file) => file.rows);
    const signatures = new Set<string>();
    for (const row of rows) signatures.add(keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'));
    return { total: rows.length, duplicates: rows.length - signatures.size, remaining: signatures.size };
  }, [allCompatible, files, keys, mode]);
  useEffect(() => { if (mode === 'join' && files[0] && !leftKey) setLeftKey(files[0].headers[0] ?? ''); if (mode === 'join' && files[1] && !rightKey) { const common = files[1].headers.find((header) => header === leftKey); setRightKey(common ?? files[1].headers[0] ?? ''); } }, [files, leftKey, mode, rightKey]);
  const joinPreview = useMemo(() => mode === 'join' && files.length === 2 && leftKey && rightKey ? createJoin(files[0], files[1], leftKey, rightKey, joinParts) : null, [files, joinParts, leftKey, mode, rightKey]);
  const joinLabel = joinParts.leftOnly && joinParts.inner && joinParts.rightOnly ? 'Tout' : joinParts.leftOnly && joinParts.inner ? 'Gauche + intersection' : joinParts.inner && joinParts.rightOnly ? 'Droite + intersection' : joinParts.leftOnly && joinParts.rightOnly ? 'Externes A + B' : joinParts.leftOnly ? 'Externe gauche' : joinParts.rightOnly ? 'Externe droite' : 'Interne';
  const joinPresetKey = `${Number(joinParts.leftOnly)}${Number(joinParts.inner)}${Number(joinParts.rightOnly)}`;
  const toggleJoinPart = (part: keyof JoinParts) => setJoinParts((current) => { const next = { ...current, [part]: !current[part] }; return next.leftOnly || next.inner || next.rightOnly ? next : current; });
  async function addFiles(list: FileList | File[]) {
    const selected = Array.from(list).filter((file) => /\.(xlsx?|csv)$/i.test(file.name));
    if (!selected.length) { setMessage('Choisissez un fichier Excel ou CSV valide.'); return; }
    setBusy(true); setMessage('');
    try { const parsed = await Promise.all(selected.map(readExcel)); setFiles(mode === 'dedupe' ? [parsed[0]] : (current) => mode === 'join' ? [...current, ...parsed].slice(0, 2) : [...current, ...parsed]); if (mode !== 'merge' && mode !== 'join' && !keys.length) setKeys(parsed[0].headers); }
    catch { setMessage('Impossible de lire ce fichier. Vérifiez son format.'); } finally { setBusy(false); }
  }
  function onDrop(event: DragEvent<HTMLDivElement>) { event.preventDefault(); setDragging(false); addFiles(event.dataTransfer.files); }
  function process() {
    if (!files.length) return;
    if (mode === 'join') {
      if (files.length !== 2 || !leftKey || !rightKey) { setMessage('Ajoutez deux fichiers et choisissez les deux clés de jointure.'); return; }
      const result = createJoin(files[0], files[1], leftKey, rightKey, joinParts); saveWorkbook(result.rows, 'jointure_excel.xlsx', result.headers); setMessage(`${result.rows.length.toLocaleString('fr-FR')} lignes générées par la jointure.`);
    } else if (mode === 'dedupe') {
      if (!keys.length) { setMessage('Sélectionnez au moins une colonne de comparaison.'); return; }
      const source = files[0].rows; const seen = new Set<string>(); const ordered = keep === 'last' ? [...source].reverse() : source;
      const clean = ordered.filter((row) => { const signature = keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'); if (seen.has(signature)) return false; seen.add(signature); return true; });
      if (keep === 'last') clean.reverse(); saveWorkbook(clean, `sans_doublons_${files[0].name.replace(/\.[^.]+$/, '')}.xlsx`, files[0].headers); setMessage(`${source.length - clean.length} doublon(s) supprimé(s). Le résultat a été téléchargé.`);
    } else if (mode === 'merge') {
      if (!allCompatible) { setMessage('Les colonnes ne correspondent pas. Corrigez les fichiers signalés.'); return; }
      saveWorkbook(files.flatMap((file) => file.rows), 'fichiers_fusionnes.xlsx', files[0].headers); setMessage(`${files.length} fichiers et ${totalRows.toLocaleString('fr-FR')} lignes ont été fusionnés.`);
    } else {
      if (!allCompatible) { setMessage('Les colonnes ne correspondent pas. Corrigez les fichiers signalés.'); return; }
      if (!keys.length) { setMessage('Sélectionnez au moins une colonne pour identifier les doublons.'); return; }
      const merged = files.flatMap((file) => file.rows); const seen = new Set<string>(); const ordered = keep === 'last' ? [...merged].reverse() : merged;
      const clean = ordered.filter((row) => { const signature = keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'); if (seen.has(signature)) return false; seen.add(signature); return true; });
      if (keep === 'last') clean.reverse(); saveWorkbook(clean, 'fusion_sans_doublons.xlsx', files[0].headers); setMessage(`${files.length} fichiers fusionnés · ${merged.length - clean.length} doublon(s) supprimé(s) · ${clean.length.toLocaleString('fr-FR')} lignes conservées.`);
    }
  }
  const title = mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner des fichiers' : mode === 'merge-dedupe' ? 'Fusionner et dédoublonner' : 'Créer une jointure';
  const description = mode === 'dedupe' ? 'Détectez les lignes identiques selon les colonnes de votre choix.' : mode === 'merge' ? 'Regroupez les lignes de fichiers qui possèdent les mêmes colonnes.' : mode === 'merge-dedupe' ? 'Réunissez vos fichiers, puis retirez les doublons en une seule opération.' : 'Reliez deux tableaux grâce à une colonne commune.';
  return <main>
    <nav><a className="brand" href="#"><span className="brandmark">X</span><span>Excel<span>Flow</span> <small>by Dhafer</small></span></a><div className="privacy"><span>✓</span> Vos fichiers restent sur votre appareil</div></nav>
    <section className="hero"><div className="eyebrow">OUTILS EXCEL, SANS COMPLICATION</div><h1>Vos fichiers Excel,<br/><em>propres et réunis.</em></h1></section>
    <section className="workspace">
      <div className="tabs" role="tablist"><button className={mode === 'dedupe' ? 'active' : ''} onClick={() => reset('dedupe')}><span className="tabicon">⌁</span> Supprimer les doublons</button><button className={mode === 'merge' ? 'active' : ''} onClick={() => reset('merge')}><span className="tabicon">⊕</span> Fusionner</button><button className={mode === 'merge-dedupe' ? 'active' : ''} onClick={() => reset('merge-dedupe')}><span className="tabicon">◎</span> Fusionner + dédoublonner</button><button className={mode === 'join' ? 'active' : ''} onClick={() => reset('join')}><span className="tabicon">⌘</span> Jointure</button></div>
      <div className="toolcard">
        <div className="toolhead"><div><span className="step">01</span><h2>{title}</h2><p>{description}</p></div><div className="format">XLSX&nbsp;&nbsp; XLS&nbsp;&nbsp; CSV</div></div>
        <input ref={input} hidden type="file" accept={accepted} multiple={mode !== 'dedupe'} onChange={(e: ChangeEvent<HTMLInputElement>) => e.target.files && addFiles(e.target.files)} />
        <div className={`dropzone ${dragging ? 'dragging' : ''} ${busy ? 'isLoading' : ''}`} aria-busy={busy} onClick={() => !busy && input.current?.click()} onPointerMove={(e) => { const box = e.currentTarget.getBoundingClientRect(); e.currentTarget.style.setProperty('--mouse-x', `${e.clientX - box.left}px`); e.currentTarget.style.setProperty('--mouse-y', `${e.clientY - box.top}px`); }} onDragOver={(e) => { e.preventDefault(); if (!busy) setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop}>
          <div className={`uploadIcon ${busy ? 'loading' : ''}`}>{busy ? <span className="spinner" /> : <span className="arrow">⇧</span>}</div>
          <strong>{busy ? 'Analyse de vos données…' : dragging ? 'Relâchez pour importer' : mode !== 'dedupe' ? 'Déposez vos fichiers ici' : 'Déposez votre fichier ici'}</strong>
          <span>{busy ? 'Lecture des colonnes et des lignes' : 'ou cliquez pour parcourir'}</span>
          {busy ? <div className="loadingTrack"><span /></div> : <small>{mode !== 'dedupe' ? 'Plusieurs fichiers autorisés · ' : ''}50 Mo maximum par fichier</small>}
        </div>
        {!!files.length && <div className="settings">
          <div className="filelist">{files.map((file, index) => <div className="file" key={`${file.name}-${index}`}><span className="fileIcon">{mode === 'join' ? index === 0 ? 'A' : 'B' : 'XL'}</span><div><strong>{file.name}</strong><small>{file.rows.length.toLocaleString('fr-FR')} lignes · {file.headers.length} colonnes</small></div><button aria-label={`Retirer ${file.name}`} onClick={() => setFiles(files.filter((_, i) => i !== index))}>×</button></div>)}{mode !== 'dedupe' && (mode !== 'join' || files.length < 2) && <button className="addmore" onClick={() => input.current?.click()}>+ Ajouter {mode === 'join' ? 'le second fichier' : 'd’autres fichiers'}</button>}</div>
          {mode !== 'dedupe' && mode !== 'join' && <div className={`compat ${allCompatible ? '' : 'error'}`}><span>{allCompatible ? '✓' : '!'}</span><div><strong>{allCompatible ? 'Colonnes compatibles' : 'Colonnes incompatibles'}</strong><small>{allCompatible ? 'Les fichiers seront réunis dans l’ordre affiché.' : 'Chaque fichier doit contenir exactement les mêmes colonnes.'}</small></div></div>}
          {mode !== 'merge' && mode !== 'join' && <div className="options"><label>Colonnes utilisées pour identifier un doublon</label><div className="chips">{files[0].headers.map((header) => <button key={header} className={keys.includes(header) ? 'selected' : ''} onClick={() => setKeys(keys.includes(header) ? keys.filter((key) => key !== header) : [...keys, header])}>{keys.includes(header) ? '✓ ' : ''}{header}</button>)}</div><label>Occurrence à conserver</label><div className="radio"><button className={keep === 'first' ? 'selected' : ''} onClick={() => setKeep('first')}>◉ Première ligne</button><button className={keep === 'last' ? 'selected' : ''} onClick={() => setKeep('last')}>◉ Dernière ligne</button></div></div>}
          {mode === 'join' && files.length === 2 && <div className="joinBuilder">
            <div className="joinKeys"><label><span>Table A · clé de jointure</span><select value={leftKey} onChange={(e) => setLeftKey(e.target.value)}>{files[0].headers.map((header) => <option key={header}>{header}</option>)}</select></label><div className="joinLink"><span></span><b>=</b><span></span></div><label><span>Table B · clé de jointure</span><select value={rightKey} onChange={(e) => setRightKey(e.target.value)}>{files[1].headers.map((header) => <option key={header}>{header}</option>)}</select></label></div>
            <div className="joinTypes">
              <div className="joinTypeHeading"><label>Sélection : <strong>{joinLabel}</strong></label><button className="joinHelpButton" aria-label="Comprendre les zones de jointure" aria-expanded={showJoinHelp} onClick={() => setShowJoinHelp(!showJoinHelp)}>?</button></div>
              {showJoinHelp && <div className="joinHelp" role="note">
                <div className="helpTitle"><span>?</span><div><strong>Comprendre les jointures</strong><p>Une jointure compare la clé choisie dans les fichiers A et B. Sélectionnez le résultat que vous souhaitez conserver.</p></div></div>
                <div className="helpOptions">
                  <div><b>A − B</b><p><strong>Externe gauche</strong>Lignes de A sans correspondance dans B.</p></div>
                  <div><b>A ∩ B</b><p><strong>Interne</strong>Lignes ayant une correspondance dans les deux fichiers.</p></div>
                  <div><b>A + ∩</b><p><strong>Gauche + intersection</strong>Toutes les lignes de A, complétées avec B lorsqu’une correspondance existe.</p></div>
                  <div><b>∩ + B</b><p><strong>Droite + intersection</strong>Toutes les lignes de B, complétées avec A lorsqu’une correspondance existe.</p></div>
                  <div><b>B − A</b><p><strong>Externe droite</strong>Lignes de B sans correspondance dans A.</p></div>
                  <div><b>A ∪ B</b><p><strong>Tout</strong>Toutes les lignes des deux fichiers, avec ou sans correspondance.</p></div>
                </div>
              </div>}
              <div className="joinHint">Cliquez sur les zones ou choisissez une combinaison.</div>
              <div className="joinPickerClean"><div className="vennDiagramClean" role="group" aria-label="Choisir les zones de la jointure"><span className="circleOutline circleA"></span><span className="circleOutline circleB"></span><button className={`joinZone zoneLeft ${joinParts.leftOnly ? 'active' : ''}`} aria-pressed={joinParts.leftOnly} aria-label="Lignes présentes uniquement dans A" onClick={() => toggleJoinPart('leftOnly')}><span>A</span></button><button className={`joinZone zoneMiddle ${joinParts.inner ? 'active' : ''}`} aria-pressed={joinParts.inner} aria-label="Correspondances entre A et B" onClick={() => toggleJoinPart('inner')}><span>∩</span></button><button className={`joinZone zoneRight ${joinParts.rightOnly ? 'active' : ''}`} aria-pressed={joinParts.rightOnly} aria-label="Lignes présentes uniquement dans B" onClick={() => toggleJoinPart('rightOnly')}><span>B</span></button></div><div className="zoneLegend"><span>A uniquement</span><span>Commun</span><span>B uniquement</span></div></div>
              <div className="joinPresets">{[{key:'100',label:'Externe gauche',parts:{leftOnly:true,inner:false,rightOnly:false}},{key:'010',label:'Interne',parts:{leftOnly:false,inner:true,rightOnly:false}},{key:'110',label:'Gauche + intersection',parts:{leftOnly:true,inner:true,rightOnly:false}},{key:'011',label:'Droite + intersection',parts:{leftOnly:false,inner:true,rightOnly:true}},{key:'001',label:'Externe droite',parts:{leftOnly:false,inner:false,rightOnly:true}},{key:'111',label:'Tout',parts:{leftOnly:true,inner:true,rightOnly:true}}].map((preset) => <button key={preset.key} className={joinPresetKey === preset.key ? 'active' : ''} onClick={() => setJoinParts(preset.parts)}>{joinPresetKey === preset.key ? '✓ ' : ''}{preset.label}</button>)}</div>
            </div>
            {joinPreview && <div className="joinResult" aria-live="polite"><span className={joinParts.inner ? 'included' : 'excluded'}><strong>{joinPreview.matched.toLocaleString('fr-FR')}</strong> correspondances</span><span className={joinParts.leftOnly ? 'included' : 'excluded'}><strong>{joinPreview.leftOnly.toLocaleString('fr-FR')}</strong> A uniquement</span><span className={joinParts.rightOnly ? 'included' : 'excluded'}><strong>{joinPreview.rightOnly.toLocaleString('fr-FR')}</strong> B uniquement</span><span className="resultTotal"><strong>{joinPreview.rows.length.toLocaleString('fr-FR')}</strong> lignes en sortie</span></div>}
          </div>}
          {previewStats && <div className="liveStats" aria-live="polite"><span><strong>{previewStats.duplicates.toLocaleString('fr-FR')}</strong> doublon(s) détecté(s)</span><i></i><span><strong>{previewStats.remaining.toLocaleString('fr-FR')}</strong> lignes après traitement</span></div>}
          <button className="primary" disabled={busy || (mode !== 'dedupe' && files.length < 2)} onClick={process}>{mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner et télécharger' : mode === 'merge-dedupe' ? 'Fusionner, dédoublonner et télécharger' : 'Créer la jointure et télécharger'} <span>→</span></button>
        </div>}{message && <div className="message" role="status">{message}</div>}
      </div>
    </section>
    <footer><span className="credit">Développé par <strong>Dhafer Harbaoui</strong></span><span>Traitement local & privé</span><span>© 2026 ExcelFlow</span></footer>
  </main>;
}
