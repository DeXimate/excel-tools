'use client';

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
type Mode = 'dedupe' | 'merge' | 'merge-dedupe';
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

export default function Home() {
  const [mode, setMode] = useState<Mode>('dedupe'); const [files, setFiles] = useState<TableFile[]>([]);
  const [keys, setKeys] = useState<string[]>([]); const [keep, setKeep] = useState<'first' | 'last'>('first');
  const [busy, setBusy] = useState(false); const [dragging, setDragging] = useState(false); const [message, setMessage] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const reset = (next: Mode) => { setMode(next); setFiles([]); setKeys([]); setMessage(''); };
  const allCompatible = useMemo(() => { if (files.length < 2) return true; const base = [...files[0].headers].sort().join('\u0000'); return files.every((file) => [...file.headers].sort().join('\u0000') === base); }, [files]);
  const totalRows = files.reduce((sum, file) => sum + file.rows.length, 0);
  const previewStats = useMemo(() => {
    if (!files.length || !keys.length || mode === 'merge' || (mode === 'merge-dedupe' && !allCompatible)) return null;
    const rows = mode === 'dedupe' ? files[0].rows : files.flatMap((file) => file.rows);
    const signatures = new Set<string>();
    for (const row of rows) signatures.add(keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'));
    return { total: rows.length, duplicates: rows.length - signatures.size, remaining: signatures.size };
  }, [allCompatible, files, keys, mode]);
  async function addFiles(list: FileList | File[]) {
    const selected = Array.from(list).filter((file) => /\.(xlsx?|csv)$/i.test(file.name));
    if (!selected.length) { setMessage('Choisissez un fichier Excel ou CSV valide.'); return; }
    setBusy(true); setMessage('');
    try { const parsed = await Promise.all(selected.map(readExcel)); setFiles(mode === 'dedupe' ? [parsed[0]] : (current) => [...current, ...parsed]); if (mode !== 'merge' && !keys.length) setKeys(parsed[0].headers); }
    catch { setMessage('Impossible de lire ce fichier. Vérifiez son format.'); } finally { setBusy(false); }
  }
  function onDrop(event: DragEvent<HTMLDivElement>) { event.preventDefault(); setDragging(false); addFiles(event.dataTransfer.files); }
  function process() {
    if (!files.length) return;
    if (mode === 'dedupe') {
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
  const title = mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner des fichiers' : 'Fusionner et dédoublonner';
  const description = mode === 'dedupe' ? 'Détectez les lignes identiques selon les colonnes de votre choix.' : mode === 'merge' ? 'Regroupez les lignes de fichiers qui possèdent les mêmes colonnes.' : 'Réunissez vos fichiers, puis retirez les doublons en une seule opération.';
  return <main>
    <nav><a className="brand" href="#"><span className="brandmark">X</span><span>Excel<span>Flow</span> <small>by Dhafer</small></span></a><div className="privacy"><span>✓</span> Vos fichiers restent sur votre appareil</div></nav>
    <section className="hero"><div className="eyebrow">OUTILS EXCEL, SANS COMPLICATION</div><h1>Vos fichiers Excel,<br/><em>propres et réunis.</em></h1></section>
    <section className="workspace">
      <div className="tabs" role="tablist"><button className={mode === 'dedupe' ? 'active' : ''} onClick={() => reset('dedupe')}><span className="tabicon">⌁</span> Supprimer les doublons</button><button className={mode === 'merge' ? 'active' : ''} onClick={() => reset('merge')}><span className="tabicon">⊕</span> Fusionner</button><button className={mode === 'merge-dedupe' ? 'active' : ''} onClick={() => reset('merge-dedupe')}><span className="tabicon">◎</span> Fusionner + dédoublonner</button></div>
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
          <div className="filelist">{files.map((file, index) => <div className="file" key={`${file.name}-${index}`}><span className="fileIcon">XL</span><div><strong>{file.name}</strong><small>{file.rows.length.toLocaleString('fr-FR')} lignes · {file.headers.length} colonnes</small></div><button aria-label={`Retirer ${file.name}`} onClick={() => setFiles(files.filter((_, i) => i !== index))}>×</button></div>)}{mode !== 'dedupe' && <button className="addmore" onClick={() => input.current?.click()}>+ Ajouter d’autres fichiers</button>}</div>
          {mode !== 'dedupe' && <div className={`compat ${allCompatible ? '' : 'error'}`}><span>{allCompatible ? '✓' : '!'}</span><div><strong>{allCompatible ? 'Colonnes compatibles' : 'Colonnes incompatibles'}</strong><small>{allCompatible ? 'Les fichiers seront réunis dans l’ordre affiché.' : 'Chaque fichier doit contenir exactement les mêmes colonnes.'}</small></div></div>}
          {mode !== 'merge' && <div className="options"><label>Colonnes utilisées pour identifier un doublon</label><div className="chips">{files[0].headers.map((header) => <button key={header} className={keys.includes(header) ? 'selected' : ''} onClick={() => setKeys(keys.includes(header) ? keys.filter((key) => key !== header) : [...keys, header])}>{keys.includes(header) ? '✓ ' : ''}{header}</button>)}</div><label>Occurrence à conserver</label><div className="radio"><button className={keep === 'first' ? 'selected' : ''} onClick={() => setKeep('first')}>◉ Première ligne</button><button className={keep === 'last' ? 'selected' : ''} onClick={() => setKeep('last')}>◉ Dernière ligne</button></div></div>}
          {previewStats && <div className="liveStats" aria-live="polite"><span><strong>{previewStats.duplicates.toLocaleString('fr-FR')}</strong> doublon(s) détecté(s)</span><i></i><span><strong>{previewStats.remaining.toLocaleString('fr-FR')}</strong> lignes après traitement</span></div>}
          <button className="primary" disabled={busy || (mode !== 'dedupe' && files.length < 2)} onClick={process}>{mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner et télécharger' : 'Fusionner, dédoublonner et télécharger'} <span>→</span></button>
        </div>}{message && <div className="message" role="status">{message}</div>}
      </div>
    </section>
    <footer><span className="credit">Développé par <strong>Dhafer Harbaoui</strong></span><span>Traitement local & privé</span><span>© 2026 ExcelFlow</span></footer>
  </main>;
}
