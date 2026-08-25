'use client';

import { ChangeEvent, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

type Row = Record<string, unknown>;
type Mode = 'outstanding' | 'completed';

const OUTSTANDING_COLUMNS = {
  Customer: ['Customer'],
  'Invoice Date': ['Invoice Date'],
  Total: ['Total', 'Total USD'],
};

const COMPLETED_COLUMNS = {
  Customer: ['Customer'],
  'Sub Customer': ['SubCustomer', 'Sub Customer'],
  'Invoice #': ['Invoice #', 'Invoice No', 'Invoice No.', 'Invoice Number'],
  'Invoice Date': ['Invoice Date'],
  Type: ['Type', 'Invoice Type'],
  Subscription: ['Subscription', 'Subscription Type'],
  Total: ['Total', 'Total USD'],
  Created: ['Created', 'CreatedBy'],
  Sent: ['Sent', 'MarkedAsSentBy'],
};

const COMPLETED_OUTPUT_COLUMNS = [
  'Customer', 'Sub Customer', 'Invoice #', 'Invoice Date', 'Type', 'Subscription', 'Total',
];

function normalizeColumn(name: string) {
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findColumns(rows: Row[], aliases: Record<string, string[]>) {
  const keys = rows.length > 0 ? Object.keys(rows[0]) : [];
  const available = new Map(keys.map((key) => [normalizeColumn(key), key]));
  const found: Record<string, string> = {};

  for (const [displayName, candidates] of Object.entries(aliases)) {
    for (const candidate of candidates) {
      const match = available.get(normalizeColumn(candidate));
      if (match) {
        found[displayName] = match;
        break;
      }
    }
  }

  return { found, missing: Object.keys(aliases).filter((name) => !found[name]) };
}

function cleanText(value: unknown) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function parseDate(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const text = cleanText(value);
  if (!text) return null;
  const dayFirst = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (dayFirst) {
    const date = new Date(Number(dayFirst[3]), Number(dayFirst[2]) - 1, Number(dayFirst[1]));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function isoDate(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function extractName(value: unknown) {
  const text = cleanText(value);
  const separator = text.includes('•') ? '•' : text.includes('â€¢') ? 'â€¢' : null;
  return separator ? text.split(separator).at(-1)?.trim() ?? '' : text;
}

function unique(values: unknown[], reverse = false) {
  const result = [...new Set(values.map(cleanText).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true }),
  );
  return reverse ? result.reverse() : result;
}

function csvCell(value: unknown) {
  const text = cleanText(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function downloadCsv(rows: Row[], fileName: string) {
  if (rows.length === 0) return;
  const columns = Object.keys(rows[0]);
  const csv = [
    columns.map(csvCell).join(','),
    ...rows.map((row) => columns.map((column) => csvCell(row[column])).join(',')),
  ].join('\r\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function FilterPicker({ label, options, selected, onChange, disabled = false }: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
  disabled?: boolean;
}) {
  const isDisabled = disabled || options.length === 0;
  const allSelected = options.length > 0 && options.every((option) => selected.includes(option));

  return (
    <fieldset className="filter" disabled={isDisabled}>
      <div className="filter-heading">
        <legend>{label}</legend>
        <label className="select-all">
          <input type="checkbox" checked={allSelected} onChange={(event) => onChange(event.target.checked ? options : [])} />
          All
        </label>
      </div>
      <select multiple aria-label={label} value={selected} onChange={(event) => onChange(Array.from(event.target.selectedOptions, (option) => option.value))}>
        {options.map((option) => <option key={option} value={option}>{option}</option>)}
      </select>
    </fieldset>
  );
}

function DataTable({ rows, limit = 100 }: { rows: Row[]; limit?: number }) {
  if (rows.length === 0) return null;
  const columns = Object.keys(rows[0]);
  const visibleRows = rows.slice(0, limit);

  return (
    <div className="table-shell">
      <table>
        <thead><tr>{columns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
        <tbody>
          {visibleRows.map((row, rowIndex) => (
            <tr key={rowIndex}>{columns.map((column) => <td key={column}>{cleanText(row[column])}</td>)}</tr>
          ))}
        </tbody>
      </table>
      {rows.length > limit && <p className="table-note">Showing the first {limit} of {rows.length} rows.</p>}
    </div>
  );
}

function OutstandingReport({ rows }: { rows: Row[] }) {
  const columns = useMemo(() => findColumns(rows, OUTSTANDING_COLUMNS), [rows]);
  const prepared = useMemo(() => {
    if (columns.missing.length > 0) return [];
    return rows.flatMap((row) => {
      const customer = cleanText(row[columns.found.Customer]);
      const invoiceDate = parseDate(row[columns.found['Invoice Date']]);
      const total = Number(String(row[columns.found.Total]).replaceAll(',', ''));
      if (!customer || !invoiceDate || Number.isNaN(total)) return [];
      return [{ customer, invoiceDate, total, year: String(invoiceDate.getFullYear()), month: invoiceDate.toLocaleString('en', { month: 'long' }), yearMonth: isoDate(invoiceDate).slice(0, 7) }];
    });
  }, [columns, rows]);

  const [customers, setCustomers] = useState<string[]>([]);
  const [years, setYears] = useState<string[]>([]);
  const [months, setMonths] = useState<string[]>([]);
  const [summary, setSummary] = useState<Row[]>([]);

  const customerOptions = unique(prepared.map((row) => row.customer));
  const customerRows = prepared.filter((row) => customers.includes(row.customer));
  const yearOptions = unique(customerRows.map((row) => row.year), true);
  const yearRows = customerRows.filter((row) => years.includes(row.year));
  const monthOptions = [...new Set(yearRows.map((row) => row.month))].sort(
    (a, b) => new Date(`${a} 1, 2000`).getMonth() - new Date(`${b} 1, 2000`).getMonth(),
  );

  function calculate() {
    const selectedRows = yearRows.filter((row) => months.includes(row.month));
    if (selectedRows.length === 0) {
      setSummary([]);
      return;
    }

    const grouped = new Map<string, typeof selectedRows>();
    for (const row of selectedRows) grouped.set(row.customer, [...(grouped.get(row.customer) ?? []), row]);
    const result: Row[] = [];
    for (const customer of [...grouped.keys()].sort()) {
      const customerRows = grouped.get(customer) ?? [];
      result.push({ 'Row Labels': customer, 'Sum of Total': customerRows.reduce((sum, row) => sum + row.total, 0).toFixed(2) });
      const monthlyTotals = new Map<string, number>();
      for (const row of customerRows) monthlyTotals.set(row.yearMonth, (monthlyTotals.get(row.yearMonth) ?? 0) + row.total);
      for (const [yearMonth, total] of [...monthlyTotals].sort()) result.push({ 'Row Labels': yearMonth, 'Sum of Total': total.toFixed(2) });
    }
    setSummary(result);
  }

  if (columns.missing.length > 0) return <div className="notice warning">Missing required columns: {columns.missing.join(', ')}</div>;
  if (prepared.length === 0) return <div className="notice warning">No valid rows found after reading Customer, Invoice Date, and Total.</div>;

  return (
    <>
      <div className="notice success">Loaded {prepared.length} valid rows.</div>
      <section><h2>Input preview</h2><DataTable rows={rows} /></section>
      <section>
        <h2>Choose what to include</h2>
        <div className="filter-grid">
          <FilterPicker label="Customers" options={customerOptions} selected={customers} onChange={(values) => { setCustomers(values); setYears([]); setMonths([]); setSummary([]); }} />
          <FilterPicker label="Years" options={yearOptions} selected={years} onChange={(values) => { setYears(values); setMonths([]); setSummary([]); }} />
          <FilterPicker label="Months" options={monthOptions} selected={months} onChange={(values) => { setMonths(values); setSummary([]); }} />
        </div>
        <button className="primary-button" onClick={calculate}>Calculate outstanding</button>
      </section>
      {summary.length > 0 && (
        <section>
          <div className="section-heading"><div><span className="eyebrow">Ready to export</span><h2>Outstanding summary</h2></div><button className="download-button" onClick={() => downloadCsv(summary, 'report_output.csv')}>Download CSV</button></div>
          <DataTable rows={summary} />
        </section>
      )}
    </>
  );
}

type CompletedRow = Row & {
  Customer: string; 'Sub Customer': string; 'Invoice #': string; 'Invoice Date': string;
  Type: string; Subscription: string; Total: unknown; 'Created By': string; 'Sent By': string;
  Year: string; Date: string;
};

function CompletedReport({ rows }: { rows: Row[] }) {
  const columns = useMemo(() => findColumns(rows, COMPLETED_COLUMNS), [rows]);
  const prepared = useMemo<CompletedRow[]>(() => rows.map((row) => {
    const value = (name: string) => columns.found[name] ? row[columns.found[name]] : '';
    const date = parseDate(value('Invoice Date'));
    return {
      Customer: cleanText(value('Customer')), 'Sub Customer': cleanText(value('Sub Customer')),
      'Invoice #': cleanText(value('Invoice #')), 'Invoice Date': date ? isoDate(date) : cleanText(value('Invoice Date')),
      Type: cleanText(value('Type')), Subscription: cleanText(value('Subscription')), Total: value('Total'),
      'Created By': extractName(value('Created')), 'Sent By': extractName(value('Sent')),
      Year: date ? String(date.getFullYear()) : '', Date: date ? isoDate(date) : '',
    };
  }), [columns, rows]);

  const [customers, setCustomers] = useState<string[]>([]);
  const [subCustomers, setSubCustomers] = useState<string[]>([]);
  const [years, setYears] = useState<string[]>([]);
  const [dates, setDates] = useState<string[]>([]);
  const [creators, setCreators] = useState<string[]>([]);
  const [subscriptions, setSubscriptions] = useState<string[]>([]);
  const [senders, setSenders] = useState<string[]>([]);

  const apply = (source: CompletedRow[], column: keyof CompletedRow, selected: string[], enabled: boolean) => !enabled ? source : selected.length === 0 ? [] : source.filter((row) => selected.includes(cleanText(row[column])));
  const has = (name: string) => Boolean(columns.found[name]);
  const customerRows = apply(prepared, 'Customer', customers, has('Customer'));
  const subCustomerRows = apply(customerRows, 'Sub Customer', subCustomers, has('Sub Customer'));
  const yearRows = apply(subCustomerRows, 'Year', years, has('Invoice Date'));
  const dateRows = apply(yearRows, 'Date', dates, has('Invoice Date'));
  const creatorRows = apply(dateRows, 'Created By', creators, has('Created'));
  const subscriptionRows = apply(creatorRows, 'Subscription', subscriptions, has('Subscription'));
  const finalRows = apply(subscriptionRows, 'Sent By', senders, has('Sent'));
  const output = finalRows.map((row) => Object.fromEntries(COMPLETED_OUTPUT_COLUMNS.map((column) => [column, row[column]])));
  const resetAfter = (...setters: Array<(value: string[]) => void>) => setters.forEach((setter) => setter([]));

  return (
    <>
      {columns.missing.length > 0 && <div className="notice warning">Missing expected columns: {columns.missing.join(', ')}. Missing filters are disabled and missing output fields will be blank.</div>}
      <div className="notice success">Loaded {prepared.length} rows.</div>
      <section><h2>Input preview</h2><DataTable rows={rows} /></section>
      <section>
        <h2>Choose what to include</h2>
        <div className="filter-grid">
          <FilterPicker label="Customers" options={unique(prepared.map((row) => row.Customer))} selected={customers} disabled={!has('Customer')} onChange={(values) => { setCustomers(values); resetAfter(setSubCustomers, setYears, setDates, setCreators, setSubscriptions, setSenders); }} />
          <FilterPicker label="Sub customers" options={unique(customerRows.map((row) => row['Sub Customer']))} selected={subCustomers} disabled={!has('Sub Customer')} onChange={(values) => { setSubCustomers(values); resetAfter(setYears, setDates, setCreators, setSubscriptions, setSenders); }} />
          <FilterPicker label="Years" options={unique(subCustomerRows.map((row) => row.Year), true)} selected={years} disabled={!has('Invoice Date')} onChange={(values) => { setYears(values); resetAfter(setDates, setCreators, setSubscriptions, setSenders); }} />
          <FilterPicker label="Invoice dates" options={unique(yearRows.map((row) => row.Date))} selected={dates} disabled={!has('Invoice Date')} onChange={(values) => { setDates(values); resetAfter(setCreators, setSubscriptions, setSenders); }} />
          <FilterPicker label="Created by" options={unique(dateRows.map((row) => row['Created By']))} selected={creators} disabled={!has('Created')} onChange={(values) => { setCreators(values); resetAfter(setSubscriptions, setSenders); }} />
          <FilterPicker label="Subscription" options={unique(creatorRows.map((row) => row.Subscription))} selected={subscriptions} disabled={!has('Subscription')} onChange={(values) => { setSubscriptions(values); resetAfter(setSenders); }} />
          <FilterPicker label="Sent by" options={unique(subscriptionRows.map((row) => row['Sent By']))} selected={senders} disabled={!has('Sent')} onChange={setSenders} />
        </div>
      </section>
      <section>
        <div className="section-heading"><div><span className="eyebrow">Live result</span><h2>Completed invoices</h2></div><button className="download-button" disabled={output.length === 0} onClick={() => downloadCsv(output, 'completed_output.csv')}>Download CSV</button></div>
        {output.length > 0 ? <DataTable rows={output} /> : <div className="notice neutral">Select values in each available filter to build your report.</div>}
      </section>
    </>
  );
}

export default function Home() {
  const [mode, setMode] = useState<Mode>('outstanding');
  const [rows, setRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  async function readFile(file: File) {
    setError('');
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error('The file does not contain a worksheet.');
      const nextRows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: '' });
      if (nextRows.length === 0) throw new Error('The uploaded file does not contain any rows.');
      setRows(nextRows);
      setFileName(file.name);
    } catch (readError) {
      setRows([]);
      setFileName('');
      setError(readError instanceof Error ? readError.message : 'Could not read this file.');
    }
  }

  function chooseFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void readFile(file);
  }

  function switchMode(nextMode: Mode) {
    setMode(nextMode);
    setRows([]);
    setFileName('');
    setError('');
    if (inputRef.current) inputRef.current.value = '';
  }

  return (
    <main>
      <header className="hero">
        <div className="hero-inner">
          <img src="/logo.png" alt="Invoice Watch" className="logo" />
          <span className="privacy-pill">Processed in your browser</span>
          <h1>Turn invoice exports into clean, ready-to-share reports.</h1>
          <p>Upload an Excel or CSV file, choose the records you need, and download the result in minutes.</p>
        </div>
      </header>

      <div className="workspace">
        <nav className="mode-tabs" aria-label="Report type">
          <button className={mode === 'outstanding' ? 'active' : ''} onClick={() => switchMode('outstanding')}><span>01</span> Outstanding</button>
          <button className={mode === 'completed' ? 'active' : ''} onClick={() => switchMode('completed')}><span>02</span> Completed</button>
        </nav>

        <section className="upload-section">
          <div className="section-kicker"><span>Step one</span><span>{mode === 'outstanding' ? 'Outstanding invoices' : 'Completed invoices'}</span></div>
          <div className={`drop-zone ${isDragging ? 'dragging' : ''}`} onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }} onDragLeave={() => setIsDragging(false)} onDrop={(event) => { event.preventDefault(); setIsDragging(false); const file = event.dataTransfer.files[0]; if (file) void readFile(file); }}>
            <div className="file-mark">XLS</div>
            <div><h2>{fileName || 'Drop your invoice export here'}</h2><p>{fileName ? `${rows.length} rows ready to filter` : 'Supports .xlsx and .csv files'}</p></div>
            <button onClick={() => inputRef.current?.click()}>{fileName ? 'Replace file' : 'Choose file'}</button>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={chooseFile} hidden />
          </div>
          {error && <div className="notice warning">Could not read this file: {error}</div>}
        </section>

        {rows.length > 0 ? (mode === 'outstanding' ? <OutstandingReport key={fileName} rows={rows} /> : <CompletedReport key={fileName} rows={rows} />) : (
          <div className="empty-state"><span>REPORT</span><h2>Your report workspace is ready</h2><p>Upload a file to reveal the preview, filters, and export controls.</p></div>
        )}
      </div>

      <footer><span>Invoice Watch</span><p>Fast reporting. Private files stay on your device.</p></footer>
    </main>
  );
}
