'use client';

import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
type Mode = 'dedupe' | 'merge' | 'merge-dedupe' | 'join' | 'convert';
type ConversionFormat = 'xlsx' | 'xls' | 'csv';
type JoinParts = { leftOnly: boolean; inner: boolean; rightOnly: boolean };
type TableFile = { name: string; headers: string[]; rows: Record<string, unknown>[]; detectedDelimiter?: string };
type ModeWorkspace = {
  files: TableFile[];
  keys: string[];
  keep: 'first' | 'last';
  leftKey: string;
  rightKey: string;
  joinParts: JoinParts;
  outputColumns: string[] | null;
  conversionFormat: ConversionFormat;
  csvDelimiter: string;
  message: string;
};
const accepted = '.xlsx,.xls,.csv';

async function readExcel(file: File): Promise<TableFile> {
  let detectedDelimiter: string | undefined;
  let workbook: XLSX.WorkBook;
  if (/\.csv$/i.test(file.name)) {
    const text = await file.text();
    detectedDelimiter = detectDelimiter(text);
    workbook = XLSX.read(text, { type: 'string', cellDates: true, FS: detectedDelimiter });
  } else workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' });
  const headers = (matrix[0] ?? []).map((value, index) => String(value || `Colonne ${index + 1}`).trim());
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '', raw: false });
  return { name: file.name, headers, rows, detectedDelimiter };
}
function saveWorkbook(rows: Record<string, unknown>[], name: string, headers?: string[]) {
  const sheet = XLSX.utils.json_to_sheet(rows, headers ? { header: headers } : undefined);
  const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, 'Résultat'); XLSX.writeFile(book, name);
}
function detectDelimiter(text: string) {
  const candidates = [',', ';', '\t', '|'];
  const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter((line) => line.trim()).slice(0, 12);
  let best = { delimiter: ',', score: -1 };
  for (const delimiter of candidates) {
    const counts = lines.map((line) => { let count = 0; let quoted = false; for (let i = 0; i < line.length; i++) { if (line[i] === '"') quoted = !quoted; else if (line[i] === delimiter && !quoted) count++; } return count; });
    const positive = counts.filter(Boolean); const consistency = positive.length ? positive.filter((count) => count === positive[0]).length / lines.length : 0;
    const score = consistency * 100 + (positive[0] ?? 0);
    if (positive.length && score > best.score) best = { delimiter, score };
  }
  return best.delimiter;
}
function saveConversion(rows: Record<string, unknown>[], headers: string[], sourceName: string, format: ConversionFormat, delimiter: string) {
  const sheet = XLSX.utils.json_to_sheet(rows, { header: headers });
  const base = sourceName.replace(/\.[^.]+$/, '');
  if (format === 'csv') {
    const csv = XLSX.utils.sheet_to_csv(sheet, { FS: delimiter });
    const url = URL.createObjectURL(new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a'); link.href = url; link.download = `${base}_converti.csv`; link.click(); URL.revokeObjectURL(url);
  } else {
    const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, sheet, 'Données'); XLSX.writeFile(book, `${base}_converti.${format}`, { bookType: format });
  }
}

const normalize = (value: unknown) => String(value ?? '').trim().toLocaleLowerCase('fr').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
function suggestJoinKeys(left: TableFile, right: TableFile) {
  let best = { left: left.headers[0] ?? '', right: right.headers[0] ?? '', score: -1 };
  for (const leftHeader of left.headers) for (const rightHeader of right.headers) {
    const leftName = normalize(leftHeader); const rightName = normalize(rightHeader);
    const nameScore = leftName === rightName ? 80 : leftName.includes(rightName) || rightName.includes(leftName) ? 35 : 0;
    const rightValues = new Set(right.rows.slice(0, 1000).map((row) => normalize(row[rightHeader])).filter(Boolean));
    const leftValues = left.rows.slice(0, 1000).map((row) => normalize(row[leftHeader])).filter(Boolean);
    const overlap = leftValues.length ? leftValues.filter((value) => rightValues.has(value)).length / leftValues.length : 0;
    const score = nameScore + overlap * 100;
    if (score > best.score) best = { left: leftHeader, right: rightHeader, score };
  }
  return best;
}
function selectColumns(rows: Record<string, unknown>[], columns: string[]) {
  return rows.map((row) => Object.fromEntries(columns.map((column) => [column, row[column] ?? ''])));
}

function createJoin(left: TableFile, right: TableFile, leftKey: string, rightKey: string, parts: JoinParts) {
  const rightNames = new Map(right.headers.map((header) => [header, left.headers.includes(header) ? `${header}_droite` : header]));
  const headers = [...left.headers, ...right.headers.map((header) => rightNames.get(header)!)];
  const index = new Map<string, { row: Record<string, unknown>; index: number }[]>();
  right.rows.forEach((row, rowIndex) => { const value = normalize(row[rightKey]); if (!value) return; const group = index.get(value) ?? []; group.push({ row, index: rowIndex }); index.set(value, group); });
  const matchedRight = new Set<number>(); const rows: Record<string, unknown>[] = []; let matched = 0; let leftOnly = 0;
  const combine = (leftRow?: Record<string, unknown>, rightRow?: Record<string, unknown>) => { const result: Record<string, unknown> = {}; left.headers.forEach((header) => { result[header] = leftRow?.[header] ?? ''; }); right.headers.forEach((header) => { result[rightNames.get(header)!] = rightRow?.[header] ?? ''; }); return result; };
  left.rows.forEach((leftRow) => { const value = normalize(leftRow[leftKey]); const matches = value ? index.get(value) ?? [] : []; if (matches.length) { matches.forEach((item) => { if (parts.inner) rows.push(combine(leftRow, item.row)); matchedRight.add(item.index); matched++; }); } else { leftOnly++; if (parts.leftOnly) rows.push(combine(leftRow)); } });
  const rightOnly = right.rows.length - matchedRight.size;
  if (parts.rightOnly) right.rows.forEach((rightRow, index) => { if (!matchedRight.has(index)) rows.push(combine(undefined, rightRow)); });
  return { rows, headers, matched, leftOnly, rightOnly };
}

function countJoin(left: TableFile, right: TableFile, leftKey: string, rightKey: string, parts: JoinParts) {
  const rightCounts = new Map<string, number>();
  let emptyRight = 0;
  right.rows.forEach((row) => { const value = normalize(row[rightKey]); if (!value) { emptyRight++; return; } rightCounts.set(value, (rightCounts.get(value) ?? 0) + 1); });
  const matchedKeys = new Set<string>();
  let matched = 0; let leftOnly = 0;
  left.rows.forEach((row) => { const value = normalize(row[leftKey]); const count = value ? rightCounts.get(value) ?? 0 : 0; if (count) { matched += count; matchedKeys.add(value); } else leftOnly++; });
  let rightOnly = emptyRight;
  rightCounts.forEach((count, value) => { if (!matchedKeys.has(value)) rightOnly += count; });
  const outputRows = (parts.inner ? matched : 0) + (parts.leftOnly ? leftOnly : 0) + (parts.rightOnly ? rightOnly : 0);
  return { matched, leftOnly, rightOnly, outputRows };
}

function getJoinHeaders(left: TableFile, right: TableFile) {
  return [...left.headers, ...right.headers.map((header) => left.headers.includes(header) ? `${header}_droite` : header)];
}

export default function Home() {
  const [mode, setMode] = useState<Mode>('dedupe'); const [files, setFiles] = useState<TableFile[]>([]);
  const [keys, setKeys] = useState<string[]>([]); const [keep, setKeep] = useState<'first' | 'last'>('first');
  const [leftKey, setLeftKey] = useState(''); const [rightKey, setRightKey] = useState(''); const [joinParts, setJoinParts] = useState<JoinParts>({ leftOnly: false, inner: true, rightOnly: false });
  const [outputColumns, setOutputColumns] = useState<string[] | null>(null);
  const [conversionFormat, setConversionFormat] = useState<ConversionFormat>('xlsx');
  const [csvDelimiter, setCsvDelimiter] = useState(';');
  const [showJoinHelp, setShowJoinHelp] = useState(false);
  const [duplicatePage, setDuplicatePage] = useState(0);
  const [busy, setBusy] = useState(false); const [dragging, setDragging] = useState(false); const [message, setMessage] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const folderInput = useRef<HTMLInputElement>(null);
  const smartAppliedFor = useRef('');
  const workspaces = useRef<Partial<Record<Mode, ModeWorkspace>>>({});
  const switchMode = (next: Mode) => {
    if (next === mode) return;
    workspaces.current[mode] = { files, keys, keep, leftKey, rightKey, joinParts, outputColumns, conversionFormat, csvDelimiter, message };
    const saved = workspaces.current[next];
    setMode(next);
    setFiles(saved?.files ?? []);
    setKeys(saved?.keys ?? []);
    setKeep(saved?.keep ?? 'first');
    setLeftKey(saved?.leftKey ?? '');
    setRightKey(saved?.rightKey ?? '');
    setJoinParts(saved?.joinParts ?? { leftOnly: false, inner: true, rightOnly: false });
    setOutputColumns(saved?.outputColumns ?? null);
    setConversionFormat(saved?.conversionFormat ?? 'xlsx');
    setCsvDelimiter(saved?.csvDelimiter ?? ';');
    setShowJoinHelp(false);
    setMessage(saved?.message ?? '');
    setDragging(false);
    if (input.current) input.current.value = '';
  };
  const allCompatible = useMemo(() => { if (files.length < 2) return true; const base = [...files[0].headers].sort().join('\u0000'); return files.every((file) => [...file.headers].sort().join('\u0000') === base); }, [files]);
  const totalRows = files.reduce((sum, file) => sum + file.rows.length, 0);
  const previewStats = useMemo(() => {
    if (!files.length || !keys.length || mode === 'merge' || (mode === 'merge-dedupe' && !allCompatible)) return null;
    const rows = mode === 'dedupe' ? files[0].rows : files.flatMap((file) => file.rows);
    const signatures = new Set<string>();
    for (const row of rows) signatures.add(keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'));
    return { total: rows.length, duplicates: rows.length - signatures.size, remaining: signatures.size };
  }, [allCompatible, files, keys, mode]);
  const duplicateRows = useMemo(() => {
    if (mode !== 'dedupe' || !files[0] || !keys.length) return [];
    const rows = files[0].rows;
    const signature = (row: Record<string, unknown>) => keys.map((key) => normalize(row[key])).join('\u0001');
    if (keep === 'last') {
      const lastIndexes = new Map<string, number>();
      rows.forEach((row, index) => lastIndexes.set(signature(row), index));
      return rows.flatMap((row, index) => {
        const keptIndex = lastIndexes.get(signature(row));
        return keptIndex !== index && keptIndex !== undefined ? [{ row, rowNumber: index + 2, keptRow: rows[keptIndex], keptRowNumber: keptIndex + 2, signature: signature(row) }] : [];
      });
    }
    const seen = new Map<string, { row: Record<string, unknown>; rowNumber: number }>();
    return rows.flatMap((row, index) => { const value = signature(row); const kept = seen.get(value); if (kept) return [{ row, rowNumber: index + 2, keptRow: kept.row, keptRowNumber: kept.rowNumber, signature: value }]; seen.set(value, { row, rowNumber: index + 2 }); return []; });
  }, [files, keep, keys, mode]);
  const duplicateDisplayRows = useMemo(() => {
    const groups = new Map<string, typeof duplicateRows>();
    duplicateRows.forEach((item) => groups.set(item.signature, [...(groups.get(item.signature) ?? []), item]));
    return Array.from(groups.values()).flatMap((items, groupIndex) => [
      { row: items[0].keptRow, rowNumber: items[0].keptRowNumber, status: 'Conservée', group: groupIndex + 1 },
      ...items.map((item) => ({ row: item.row, rowNumber: item.rowNumber, status: 'Supprimée', group: groupIndex + 1 })),
    ]);
  }, [duplicateRows]);
  const duplicatePageSize = 50;
  const duplicatePageCount = Math.max(1, Math.ceil(duplicateDisplayRows.length / duplicatePageSize));
  const visibleDuplicateRows = duplicateDisplayRows.slice(duplicatePage * duplicatePageSize, (duplicatePage + 1) * duplicatePageSize);
  useEffect(() => { setDuplicatePage(0); }, [files, keep, keys, mode]);
  useEffect(() => { if (mode === 'join' && files[0] && !leftKey) setLeftKey(files[0].headers[0] ?? ''); if (mode === 'join' && files[1] && !rightKey) { const common = files[1].headers.find((header) => header === leftKey); setRightKey(common ?? files[1].headers[0] ?? ''); } }, [files, leftKey, mode, rightKey]);
  const joinPreview = useMemo(() => mode === 'join' && files.length === 2 && leftKey && rightKey ? countJoin(files[0], files[1], leftKey, rightKey, joinParts) : null, [files, joinParts, leftKey, mode, rightKey]);
  const availableOutputColumns = useMemo(() => mode === 'join' ? files.length === 2 ? getJoinHeaders(files[0], files[1]) : [] : files[0]?.headers ?? [], [files, mode]);
  const selectedOutputColumns = outputColumns ?? availableOutputColumns;
  const smartJoin = useMemo(() => mode === 'join' && files.length === 2 ? suggestJoinKeys(files[0], files[1]) : null, [files, mode]);
  useEffect(() => { setOutputColumns((current) => current?.filter((column) => availableOutputColumns.includes(column)) ?? null); }, [availableOutputColumns]);
  useEffect(() => {
    if (!smartJoin || files.length !== 2) return;
    const signature = files.map((file) => `${file.name}:${file.rows.length}:${file.headers.join('|')}`).join('::');
    if (smartAppliedFor.current === signature) return;
    smartAppliedFor.current = signature;
    setLeftKey(smartJoin.left); setRightKey(smartJoin.right);
  }, [files, smartJoin]);
  const joinLabel = joinParts.leftOnly && joinParts.inner && joinParts.rightOnly ? 'Tout' : joinParts.leftOnly && joinParts.inner ? 'Gauche + intersection' : joinParts.inner && joinParts.rightOnly ? 'Droite + intersection' : joinParts.leftOnly && joinParts.rightOnly ? 'Externes A + B' : joinParts.leftOnly ? 'Externe gauche' : joinParts.rightOnly ? 'Externe droite' : 'Interne';
  const joinPresetKey = `${Number(joinParts.leftOnly)}${Number(joinParts.inner)}${Number(joinParts.rightOnly)}`;
  const toggleJoinPart = (part: keyof JoinParts) => setJoinParts((current) => { const next = { ...current, [part]: !current[part] }; return next.leftOnly || next.inner || next.rightOnly ? next : current; });
  async function addFiles(list: FileList | File[]) {
    const selected = Array.from(list).filter((file) => /\.(xlsx?|csv)$/i.test(file.name));
    if (!selected.length) { setMessage('Choisissez un fichier Excel ou CSV valide.'); return; }
    setBusy(true); setMessage('');
    try { const parsed = await Promise.all(selected.map(readExcel)); setFiles(mode === 'dedupe' || mode === 'convert' ? [parsed[0]] : (current) => mode === 'join' ? [...current, ...parsed].slice(0, 2) : [...current, ...parsed]); if (mode !== 'merge' && mode !== 'join' && mode !== 'convert' && !keys.length) setKeys(parsed[0].headers); if (mode === 'convert') { const isCsv = /\.csv$/i.test(parsed[0].name); setConversionFormat(isCsv ? 'xlsx' : 'csv'); if (parsed[0].detectedDelimiter) setCsvDelimiter(parsed[0].detectedDelimiter); } }
    catch { setMessage('Impossible de lire ce fichier. Vérifiez son format.'); } finally { setBusy(false); }
  }
  function onDrop(event: DragEvent<HTMLDivElement>) { event.preventDefault(); setDragging(false); addFiles(event.dataTransfer.files); }
  function process() {
    if (!files.length) return;
    if (mode === 'convert') {
      if (!selectedOutputColumns.length) { setMessage('Sélectionnez au moins une colonne pour le fichier final.'); return; }
      saveConversion(selectColumns(files[0].rows, selectedOutputColumns), selectedOutputColumns, files[0].name, conversionFormat, csvDelimiter);
      setMessage(`${files[0].rows.length.toLocaleString('fr-FR')} lignes converties au format ${conversionFormat.toUpperCase()}.`);
    } else if (mode === 'join') {
      if (files.length !== 2 || !leftKey || !rightKey) { setMessage('Ajoutez deux fichiers et choisissez les deux clés de jointure.'); return; }
      if (!selectedOutputColumns.length) { setMessage('Sélectionnez au moins une colonne pour le fichier final.'); return; }
      const estimate = countJoin(files[0], files[1], leftKey, rightKey, joinParts);
      if (estimate.outputRows > 200000) { setMessage(`Cette jointure générerait ${estimate.outputRows.toLocaleString('fr-FR')} lignes. Choisissez des clés plus uniques ou réduisez les zones sélectionnées pour éviter de saturer la mémoire.`); return; }
      const result = createJoin(files[0], files[1], leftKey, rightKey, joinParts); saveWorkbook(selectColumns(result.rows, selectedOutputColumns), 'jointure_excel.xlsx', selectedOutputColumns); setMessage(`${result.rows.length.toLocaleString('fr-FR')} lignes générées par la jointure.`);
    } else if (mode === 'dedupe') {
      if (!keys.length) { setMessage('Sélectionnez au moins une colonne de comparaison.'); return; }
      const source = files[0].rows; const seen = new Set<string>(); const ordered = keep === 'last' ? [...source].reverse() : source;
      const clean = ordered.filter((row) => { const signature = keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'); if (seen.has(signature)) return false; seen.add(signature); return true; });
      if (!selectedOutputColumns.length) { setMessage('Sélectionnez au moins une colonne pour le fichier final.'); return; }
      if (keep === 'last') clean.reverse(); saveWorkbook(selectColumns(clean, selectedOutputColumns), `sans_doublons_${files[0].name.replace(/\.[^.]+$/, '')}.xlsx`, selectedOutputColumns); setMessage(`${source.length - clean.length} doublon(s) supprimé(s). Le résultat a été téléchargé.`);
    } else if (mode === 'merge') {
      if (!allCompatible) { setMessage('Les colonnes ne correspondent pas. Corrigez les fichiers signalés.'); return; }
      if (!selectedOutputColumns.length) { setMessage('Sélectionnez au moins une colonne pour le fichier final.'); return; }
      saveWorkbook(selectColumns(files.flatMap((file) => file.rows), selectedOutputColumns), 'fichiers_fusionnes.xlsx', selectedOutputColumns); setMessage(`${files.length} fichiers et ${totalRows.toLocaleString('fr-FR')} lignes ont été fusionnés.`);
    } else {
      if (!allCompatible) { setMessage('Les colonnes ne correspondent pas. Corrigez les fichiers signalés.'); return; }
      if (!keys.length) { setMessage('Sélectionnez au moins une colonne pour identifier les doublons.'); return; }
      const merged = files.flatMap((file) => file.rows); const seen = new Set<string>(); const ordered = keep === 'last' ? [...merged].reverse() : merged;
      const clean = ordered.filter((row) => { const signature = keys.map((key) => String(row[key] ?? '').trim().toLocaleLowerCase('fr')).join('\u0001'); if (seen.has(signature)) return false; seen.add(signature); return true; });
      if (!selectedOutputColumns.length) { setMessage('Sélectionnez au moins une colonne pour le fichier final.'); return; }
      if (keep === 'last') clean.reverse(); saveWorkbook(selectColumns(clean, selectedOutputColumns), 'fusion_sans_doublons.xlsx', selectedOutputColumns); setMessage(`${files.length} fichiers fusionnés · ${merged.length - clean.length} doublon(s) supprimé(s) · ${clean.length.toLocaleString('fr-FR')} lignes conservées.`);
    }
  }
  const title = mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner des fichiers' : mode === 'merge-dedupe' ? 'Fusionner et dédoublonner' : mode === 'join' ? 'Créer une jointure' : 'Convertir un fichier';
  const description = mode === 'dedupe' ? 'Détectez les lignes identiques selon les colonnes de votre choix.' : mode === 'merge' ? 'Regroupez les lignes de fichiers qui possèdent les mêmes colonnes.' : mode === 'merge-dedupe' ? 'Réunissez vos fichiers, puis retirez les doublons en une seule opération.' : mode === 'join' ? 'Reliez deux tableaux grâce à une colonne commune.' : 'Convertissez Excel et CSV dans les deux sens, directement sur votre appareil.';
  return <main>
    <nav><a className="brand" href="#"><span className="brandmark">X</span><span>Excel<span>Flow</span> <small>by Dhafer</small></span></a><div className="privacy"><span>✓</span> Vos fichiers restent sur votre appareil</div></nav>
    <section className="hero"><div className="eyebrow">OUTILS EXCEL, SANS COMPLICATION</div><h1>Vos fichiers Excel,<br/><em>propres et réunis.</em></h1></section>
    <section className="workspace">
      <div className="tabs" role="tablist"><button className={mode === 'dedupe' ? 'active' : ''} onClick={() => switchMode('dedupe')}><span className="tabicon">⌁</span> Supprimer les doublons</button><button className={mode === 'merge' ? 'active' : ''} onClick={() => switchMode('merge')}><span className="tabicon">⊕</span> Fusionner</button><button className={mode === 'merge-dedupe' ? 'active' : ''} onClick={() => switchMode('merge-dedupe')}><span className="tabicon">◎</span> Fusionner + dédoublonner</button><button className={mode === 'join' ? 'active' : ''} onClick={() => switchMode('join')}><span className="tabicon">⌘</span> Jointure</button><button className={mode === 'convert' ? 'active' : ''} onClick={() => switchMode('convert')}><span className="tabicon">⇄</span> Convertir</button></div>
      <div className="toolcard">
        <div className="toolhead"><div><span className="step">01</span><h2>{title}</h2><p>{description}</p></div><div className="format">XLSX&nbsp;&nbsp; XLS&nbsp;&nbsp; CSV</div></div>
        <input ref={input} hidden type="file" accept={accepted} multiple={mode !== 'dedupe' && mode !== 'convert'} onChange={(e: ChangeEvent<HTMLInputElement>) => e.target.files && addFiles(e.target.files)} />
        <input ref={folderInput} hidden type="file" accept={accepted} multiple {...({ webkitdirectory: '', directory: '' } as Record<string, string>)} onChange={(e: ChangeEvent<HTMLInputElement>) => e.target.files && addFiles(e.target.files)} />
        <div className={`dropzone ${dragging ? 'dragging' : ''} ${busy ? 'isLoading' : ''}`} aria-busy={busy} onClick={() => !busy && input.current?.click()} onPointerMove={(e) => { const box = e.currentTarget.getBoundingClientRect(); e.currentTarget.style.setProperty('--mouse-x', `${e.clientX - box.left}px`); e.currentTarget.style.setProperty('--mouse-y', `${e.clientY - box.top}px`); }} onDragOver={(e) => { e.preventDefault(); if (!busy) setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop}>
          <div className={`uploadIcon ${busy ? 'loading' : ''}`}>{busy ? <span className="spinner" /> : <span className="arrow">⇧</span>}</div>
          <strong>{busy ? 'Analyse de vos données…' : dragging ? 'Relâchez pour importer' : mode === 'dedupe' || mode === 'convert' ? 'Déposez votre fichier ici' : 'Déposez vos fichiers ici'}</strong>
          <span>{busy ? 'Lecture des colonnes et des lignes' : 'ou cliquez pour parcourir'}</span>
          {busy ? <div className="loadingTrack"><span /></div> : <small>{mode !== 'dedupe' && mode !== 'convert' ? 'Plusieurs fichiers autorisés · ' : ''}50 Mo maximum par fichier</small>}
        </div>
        {!files.length && (mode === 'merge' || mode === 'merge-dedupe') && <button className="folderEntry" onClick={() => folderInput.current?.click()}><span>▣</span><div><strong>Importer tout un dossier</strong><small>Seuls les fichiers Excel et CSV du dossier seront ajoutés.</small></div><b>Choisir un dossier →</b></button>}
        {!!files.length && <div className="settings">
          <div className="filelist">{files.map((file, index) => <div className="file" key={`${file.name}-${index}`}><span className="fileIcon">{mode === 'join' ? index === 0 ? 'A' : 'B' : mode === 'convert' ? file.name.split('.').pop()?.toUpperCase() : 'XL'}</span><div><strong>{file.name}</strong><small>{file.rows.length.toLocaleString('fr-FR')} lignes · {file.headers.length} colonnes</small></div><button aria-label={`Retirer ${file.name}`} onClick={() => setFiles(files.filter((_, i) => i !== index))}>×</button></div>)}{mode !== 'dedupe' && mode !== 'convert' && (mode !== 'join' || files.length < 2) && <div className="addActions"><button className="addmore" onClick={() => input.current?.click()}>+ Ajouter {mode === 'join' ? 'le second fichier' : 'des fichiers'}</button>{mode !== 'join' && <button className="addmore folderButton" onClick={() => folderInput.current?.click()}>▣ Importer un dossier</button>}</div>}</div>
          {(mode === 'merge' || mode === 'merge-dedupe') && <div className={`compat ${allCompatible ? '' : 'error'}`}><span>{allCompatible ? '✓' : '!'}</span><div><strong>{allCompatible ? 'Colonnes compatibles' : 'Colonnes incompatibles'}</strong><small>{allCompatible ? 'Les fichiers seront réunis dans l’ordre affiché.' : 'Chaque fichier doit contenir exactement les mêmes colonnes.'}</small></div></div>}
          {(mode === 'dedupe' || mode === 'merge-dedupe') && <div className="options"><div className="keySelectionHead"><div><label>Colonnes utilisées pour identifier un doublon</label><small>{keys.length} sur {files[0].headers.length} sélectionnée(s)</small></div><div><button onClick={() => setKeys([...files[0].headers])}>Tout sélectionner</button><button onClick={() => setKeys([])}>Tout désélectionner</button></div></div><div className="chips keyChips">{files[0].headers.map((header) => <button key={header} className={keys.includes(header) ? 'selected' : ''} onClick={() => setKeys(keys.includes(header) ? keys.filter((key) => key !== header) : [...keys, header])}>{keys.includes(header) ? '✓ ' : ''}{header}</button>)}</div><label>Occurrence à conserver</label><div className="radio"><button className={keep === 'first' ? 'selected' : ''} onClick={() => setKeep('first')}>◉ Première ligne</button><button className={keep === 'last' ? 'selected' : ''} onClick={() => setKeep('last')}>◉ Dernière ligne</button></div></div>}
          {mode === 'join' && files.length === 2 && <div className="joinBuilder">
            <div className="joinKeys"><label><span>Table A · clé de jointure</span><select value={leftKey} onChange={(e) => setLeftKey(e.target.value)}>{files[0].headers.map((header) => <option key={header}>{header}</option>)}</select></label><div className="joinLink"><span></span><b>=</b><span></span></div><label><span>Table B · clé de jointure</span><select value={rightKey} onChange={(e) => setRightKey(e.target.value)}>{files[1].headers.map((header) => <option key={header}>{header}</option>)}</select></label></div>
            {smartJoin && <div className="smartSuggestion"><span className="smartIcon">✦</span><div><strong>Clés détectées automatiquement</strong><small>{smartJoin.left} ↔ {smartJoin.right}</small></div><button onClick={() => { setLeftKey(smartJoin.left); setRightKey(smartJoin.right); }}>Utiliser cette suggestion</button></div>}
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
            {joinPreview && <div className="joinResult" aria-live="polite"><span className={joinParts.inner ? 'included' : 'excluded'}><strong>{joinPreview.matched.toLocaleString('fr-FR')}</strong> correspondances</span><span className={joinParts.leftOnly ? 'included' : 'excluded'}><strong>{joinPreview.leftOnly.toLocaleString('fr-FR')}</strong> A uniquement</span><span className={joinParts.rightOnly ? 'included' : 'excluded'}><strong>{joinPreview.rightOnly.toLocaleString('fr-FR')}</strong> B uniquement</span><span className="resultTotal"><strong>{joinPreview.outputRows.toLocaleString('fr-FR')}</strong> lignes en sortie</span></div>}
          </div>}
          {mode === 'convert' && <div className="converter">
            {/\.csv$/i.test(files[0].name) && <div className="delimiterDetection"><span>✦</span><div><strong>Séparateur détecté automatiquement</strong><small>{csvDelimiter === ',' ? 'Virgule (,)' : csvDelimiter === ';' ? 'Point-virgule (;)' : csvDelimiter === '\t' ? 'Tabulation' : 'Barre verticale (|)'}</small></div><b>Détection intelligente</b></div>}
            <div className="conversionFlow"><div className="sourceFormat"><span>Format source</span><strong>{files[0].name.split('.').pop()?.toUpperCase()}</strong><small>{files[0].rows.length.toLocaleString('fr-FR')} lignes reconnues</small></div><div className="flowArrow"><i></i><b>→</b><i></i></div><div className="targetFormat"><span>Convertir vers</span><div className="formatChoices">{(['xlsx','xls','csv'] as ConversionFormat[]).map((format) => <button key={format} disabled={files[0].name.toLowerCase().endsWith(`.${format}`)} className={conversionFormat === format ? 'active' : ''} onClick={() => setConversionFormat(format)}><b>{format.toUpperCase()}</b><small>{format === 'csv' ? 'Texte universel' : format === 'xls' ? 'Ancien Excel' : 'Excel moderne'}</small></button>)}</div></div></div>
            {conversionFormat === 'csv' && <div className="delimiterChoice"><div><label>Séparateur du fichier CSV</label><small>Le choix recommandé dépend des paramètres régionaux d’Excel.</small></div><div>{[{value:';',label:'Point-virgule',symbol:';'},{value:',',label:'Virgule',symbol:','},{value:'\t',label:'Tabulation',symbol:'TAB'},{value:'|',label:'Barre verticale',symbol:'|'}].map((item) => <button key={item.value} className={csvDelimiter === item.value ? 'active' : ''} onClick={() => setCsvDelimiter(item.value)}><b>{item.symbol}</b><span>{item.label}</span></button>)}</div></div>}
            <div className="conversionPreview"><div><strong>Aperçu des données</strong><small>5 premières lignes après lecture du fichier</small></div><div className="conversionTableWrap"><table><thead><tr>{files[0].headers.map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{files[0].rows.slice(0,5).map((row,index) => <tr key={index}>{files[0].headers.map((header) => <td key={header}>{String(row[header] ?? '') || <em>vide</em>}</td>)}</tr>)}</tbody></table></div></div>
          </div>}
          {!!availableOutputColumns.length && <div className="outputColumns"><div className="outputHeading"><div><label>Colonnes du fichier final</label><small>{selectedOutputColumns.length} sur {availableOutputColumns.length} sélectionnée(s)</small></div><div><button onClick={() => setOutputColumns([...availableOutputColumns])}>Tout sélectionner</button><button onClick={() => setOutputColumns([])}>Tout retirer</button></div></div><div className="chips">{availableOutputColumns.map((column) => <button key={column} className={selectedOutputColumns.includes(column) ? 'selected' : ''} onClick={() => setOutputColumns(selectedOutputColumns.includes(column) ? selectedOutputColumns.filter((item) => item !== column) : [...selectedOutputColumns, column])}>{selectedOutputColumns.includes(column) ? '✓ ' : ''}{column}</button>)}</div></div>}
          {previewStats && <div className="liveStats" aria-live="polite"><span><strong>{previewStats.duplicates.toLocaleString('fr-FR')}</strong> doublon(s) détecté(s)</span><i></i><span><strong>{previewStats.remaining.toLocaleString('fr-FR')}</strong> lignes après traitement</span></div>}
          {mode === 'dedupe' && duplicateRows.length > 0 && <section className="duplicatePanel" aria-label="Liste complète des doublons">
            <div className="duplicateHead"><div><span className="duplicateBadge">{duplicateRows.length.toLocaleString('fr-FR')}</span><div><h3>Groupes de doublons</h3><p>Chaque doublon est affiché avec la ligne correspondante qui sera conservée.</p></div></div><button onClick={() => saveWorkbook(duplicateDisplayRows.map((item) => ({ Statut: item.status, Groupe: item.group, 'N° ligne': item.rowNumber, ...item.row })), `doublons_${files[0].name.replace(/\.[^.]+$/, '')}.xlsx`, ['Statut', 'Groupe', 'N° ligne', ...files[0].headers])}>Télécharger la liste</button></div>
            <div className="duplicateTableWrap"><table className="duplicateTable"><thead><tr><th>Statut</th><th>Groupe</th><th>N° ligne</th>{files[0].headers.map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{visibleDuplicateRows.map((item) => <tr className={item.status === 'Conservée' ? 'keptRow' : 'removedRow'} key={`${item.group}-${item.status}-${item.rowNumber}`}><td><span className="rowStatus">{item.status}</span></td><td>#{item.group}</td><td>{item.rowNumber}</td>{files[0].headers.map((header) => <td key={header} title={String(item.row[header] ?? '')}>{String(item.row[header] ?? '') || <em>vide</em>}</td>)}</tr>)}</tbody></table></div>
            <div className="duplicatePager"><span>Lignes {(duplicatePage * duplicatePageSize + 1).toLocaleString('fr-FR')}–{Math.min((duplicatePage + 1) * duplicatePageSize, duplicateDisplayRows.length).toLocaleString('fr-FR')} sur {duplicateDisplayRows.length.toLocaleString('fr-FR')}</span><div><button disabled={duplicatePage === 0} onClick={() => setDuplicatePage((page) => Math.max(0, page - 1))}>← Précédent</button><b>{duplicatePage + 1} / {duplicatePageCount}</b><button disabled={duplicatePage >= duplicatePageCount - 1} onClick={() => setDuplicatePage((page) => Math.min(duplicatePageCount - 1, page + 1))}>Suivant →</button></div></div>
          </section>}
          <button className="primary" disabled={busy || ((mode === 'merge' || mode === 'merge-dedupe' || mode === 'join') && files.length < 2)} onClick={process}>{mode === 'dedupe' ? 'Supprimer les doublons' : mode === 'merge' ? 'Fusionner et télécharger' : mode === 'merge-dedupe' ? 'Fusionner, dédoublonner et télécharger' : mode === 'join' ? 'Créer la jointure et télécharger' : `Convertir en ${conversionFormat.toUpperCase()} et télécharger`} <span>→</span></button>
        </div>}{message && <div className="message" role="status">{message}</div>}
      </div>
    </section>
    <footer><span className="credit">Développé par <strong>Dhafer Harbaoui</strong></span><span>Traitement local & privé</span><span>© 2026 ExcelFlow</span></footer>
  </main>;
}
