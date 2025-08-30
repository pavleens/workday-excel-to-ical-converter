import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

/**
 * Workday Excel -> iCal converter
 * - Upload .xlsx or .csv exported from Workday
 * - Map columns to fields
 * - Expands weekly patterns (e.g., MWF, TuTh, Mon Wed Fri) into individual dates between Start Date and End Date
 * - Exports a standards-compliant .ics file
 *
 * Notes
 * - Default timezone hint set to America/Vancouver via X-WR-TIMEZONE (events are floating local times)
 * - If your file fails to parse, try exporting CSV from Excel and uploading that
 */
export default function WorkdayExcelToICS() {
  const [rows, setRows] = useState([]); // array of objects, keys are headers
  const [headers, setHeaders] = useState([]); // detected header names
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [mapped, setMapped] = useState({});
  const [titleTemplate, setTitleTemplate] = useState("{Course} {Component} {Section}");
  const [calendarName, setCalendarName] = useState("Workday Schedule");
  const [timezone, setTimezone] = useState("America/Vancouver");
  const inputRef = useRef(null);

  const guessMap = useMemo(() => guessInitialMapping(headers), [headers]);

  // Apply guesses on first load
  React.useEffect(() => {
    setMapped((m) => ({ ...m, ...guessMap }));
  }, [guessMap]);

  const handleFile = async (file) => {
    setError("");
    setInfo("");
    setRows([]);
    setHeaders([]);
    setFileName(file?.name || "");
    if (!file) return;

    try {
      const ext = (file.name.split(".").pop() || "").toLowerCase();

      if (ext === "xlsx" || ext === "xls") {
        const ab = await file.arrayBuffer();
        const wb = XLSX.read(ab, { type: "array", cellDates: true });
        const firstSheet = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheet];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        if (!json.length) throw new Error("No rows detected");
        const hdrs = Object.keys(json[0]);
        setRows(json);
        setHeaders(hdrs);
        setInfo(`Loaded ${json.length} rows from sheet "${firstSheet}"`);
        return;
      }

      if (ext === "csv") {
        const text = await file.text();
        const { data, hdrs } = parseCSV(text);
        if (!data.length) throw new Error("No rows detected");
        setRows(data);
        setHeaders(hdrs);
        setInfo(`Loaded ${data.length} rows from CSV`);
        return;
      }

      throw new Error("Please upload an .xlsx or .csv file");
    } catch (e) {
      console.error(e);
      setError(`Failed to read file. ${e.message || e.toString()}`);
    }
  };

  const onDrop = (e) => {
    e.preventDefault();
    if (!e.dataTransfer.files?.length) return;
    handleFile(e.dataTransfer.files[0]);
  };

  const onBrowse = (e) => {
    const f = e.target.files?.[0];
    if (f) handleFile(f);
  };

  const mappingControls = (
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-6">
      {renderSelect("Title field (optional)", "titleField", headers, mapped, setMapped)}
      {renderText("Title template", titleTemplate, setTitleTemplate, "Use tokens like {Course}, {Component}, {Section}. Ignored if Title field is selected.")}
      {renderSelect("Course (optional)", "courseField", headers, mapped, setMapped)}
      {renderSelect("Component (LEC, LBL, etc, optional)", "componentField", headers, mapped, setMapped)}
      {renderSelect("Section (optional)", "sectionField", headers, mapped, setMapped)}
      {renderSelect("Start date", "startDateField", headers, mapped, setMapped, true)}
      {renderSelect("End date", "endDateField", headers, mapped, setMapped, true)}
      {renderSelect("Start time", "startTimeField", headers, mapped, setMapped, true)}
      {renderSelect("End time", "endTimeField", headers, mapped, setMapped, true)}
      {renderSelect("Days pattern", "daysField", headers, mapped, setMapped, true, "Examples: MWF, TuTh, Mon Wed Fri, TTh")}
      {renderSelect("Location (optional)", "locationField", headers, mapped, setMapped)}
      {renderSelect("Description (optional)", "descField", headers, mapped, setMapped)}
      <div className="flex flex-col">
        <label className="text-sm text-gray-600 mb-1">Timezone hint</label>
        <input className="border rounded-xl px-3 py-2" value={timezone} onChange={(e)=>setTimezone(e.target.value)} placeholder="America/Vancouver" />
        <p className="text-xs text-gray-500 mt-1">Calendars will treat times as local by default. This tag helps consumers group events under a calendar timezone.</p>
      </div>
      <div className="flex flex-col">
        <label className="text-sm text-gray-600 mb-1">Calendar name</label>
        <input className="border rounded-xl px-3 py-2" value={calendarName} onChange={(e)=>setCalendarName(e.target.value)} />
      </div>
    </div>
  );

  const previewTable = (
    <div className="mt-6 overflow-auto max-h-72 border rounded-xl">
      <table className="min-w-full text-sm">
        <thead className="bg-gray-50 sticky top-0">
          <tr>
            {headers.map((h) => (
              <th key={h} className="text-left px-3 py-2 font-semibold border-b">{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.slice(0, 10).map((r, i) => (
            <tr key={i} className="odd:bg-white even:bg-gray-50">
              {headers.map((h) => (
                <td key={h} className="px-3 py-2 border-b align-top">{String(r[h] ?? "")}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );

  const handleGenerate = () => {
    try {
      const req = ["startDateField", "endDateField", "startTimeField", "endTimeField", "daysField"];
      for (const k of req) {
        if (!mapped[k]) throw new Error(`Missing required mapping: ${labelForKey(k)}`);
      }

      const events = [];
      const errors = [];

      rows.forEach((row, idx) => {
        try {
          const startDate = parseDate(row[mapped.startDateField]);
          const endDate = parseDate(row[mapped.endDateField]);
          const startTime = parseTime(row[mapped.startTimeField]);
          const endTime = parseTime(row[mapped.endTimeField]);
          const days = parseDays(row[mapped.daysField]);

          if (!startDate || !endDate || !startTime || !endTime || days.size === 0) {
            throw new Error("Could not parse one of date, time, or days");
          }

          // Build title
          let summary = "";
          if (mapped.titleField) {
            summary = String(row[mapped.titleField] ?? "").trim();
          } else {
            summary = titleTemplate
              .replaceAll("{Course}", String(row[mapped.courseField] ?? "").trim())
              .replaceAll("{Component}", String(row[mapped.componentField] ?? "").trim())
              .replaceAll("{Section}", String(row[mapped.sectionField] ?? "").trim())
              .replace(/\s+/g, " ")
              .trim();
          }
          if (!summary) summary = "Class";

          const location = mapped.locationField ? String(row[mapped.locationField] ?? "").trim() : "";
          const description = mapped.descField ? String(row[mapped.descField] ?? "").trim() : "";

          const occurrences = expandOccurrences(startDate, endDate, days);
          occurrences.forEach((d) => {
            const dtStart = combineDateTime(d, startTime);
            const dtEnd = combineDateTime(d, endTime);
            events.push({ summary, location, description, dtStart, dtEnd });
          });
        } catch (e) {
          errors.push({ idx: idx + 1, error: e.message || String(e) });
        }
      });

      if (!events.length) throw new Error("No events generated. Check mappings and data.");

      const ics = buildICS(events, calendarName, timezone);
      downloadText(ics, sanitizeFileName((calendarName || "schedule")) + ".ics");
      setInfo(`Generated ${events.length} events. ${errors.length ? "Rows with issues: " + errors.length : "All rows parsed."}`);
      if (errors.length) console.warn("Row errors", errors);
    } catch (e) {
      setError(e.message || String(e));
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-white to-gray-100 text-gray-900">
      <div className="max-w-5xl mx-auto px-4 py-10">
        <h1 className="text-3xl font-bold">Workday Excel → iCal</h1>
        <p className="mt-2 text-gray-600">Upload your Workday schedule export and get a clean .ics you can import into Google Calendar, Apple Calendar, or Outlook.</p>

        <div
          className="mt-6 border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer hover:bg-gray-50"
          onDragOver={(e)=>{e.preventDefault();}}
          onDrop={onDrop}
          onClick={()=>inputRef.current?.click()}
        >
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onBrowse} />
          <div className="text-lg">Drop your .xlsx or .csv here, or click to browse</div>
          <div className="text-sm text-gray-500 mt-1">Recommended: export directly from Workday without editing</div>
          {fileName && <div className="mt-3 text-sm text-gray-700">Selected: <span className="font-medium">{fileName}</span></div>}
        </div>

        {info && <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-xl text-green-800">{info}</div>}
        {error && <div className="mt-4 p-3 bg-red-50 border border-red-200 rounded-xl text-red-800">{error}</div>}

        {!!rows.length && (
          <>
            <h2 className="mt-8 text-xl font-semibold">Step 2 - Map your columns</h2>
            <p className="text-sm text-gray-600">We tried to guess based on header names. You can override below. Required fields are marked.</p>
            {mappingControls}
            <h2 className="mt-8 text-xl font-semibold">Preview (first 10 rows)</h2>
            {previewTable}
            <div className="mt-6 flex gap-3">
              <button onClick={handleGenerate} className="px-4 py-2 rounded-xl bg-black text-white shadow hover:opacity-90">Generate .ics</button>
              <button onClick={()=>{setRows([]); setHeaders([]); setMapped({}); setTitleTemplate("{Course} {Component} {Section}"); setError(""); setInfo(""); setFileName("");}} className="px-4 py-2 rounded-xl border">Reset</button>
            </div>
            <p className="text-xs text-gray-500 mt-3">Tip: If parsing fails for .xlsx, save your sheet as CSV and upload that.</p>
          </>
        )}

        <Footer />
      </div>
    </div>
  );
}

function Footer() {
  return (
    <div className="mt-16 text-xs text-gray-500">
      <p>Privacy: everything runs in your browser. Files never leave your device.</p>
    </div>
  );
}

function renderSelect(label, key, headers, mapped, setMapped, required=false, hint="") {
  return (
    <div className="flex flex-col">
      <label className="text-sm text-gray-600 mb-1">{label} {required && <span className="text-red-500">*</span>}</label>
      <select
        className="border rounded-xl px-3 py-2"
        value={mapped[key] || ""}
        onChange={(e)=>setMapped((m)=>({...m, [key]: e.target.value}))}
      >
        <option value="">-- Select a column --</option>
        {headers.map((h)=> <option key={h} value={h}>{h}</option>)}
      </select>
      {hint ? <p className="text-xs text-gray-500 mt-1">{hint}</p> : null}
    </div>
  );
}

function renderText(label, value, onChange, hint="") {
  return (
    <div className="flex flex-col">
      <label className="text-sm text-gray-600 mb-1">{label}</label>
      <input className="border rounded-xl px-3 py-2" value={value} onChange={(e)=>onChange(e.target.value)} />
      {hint ? <p className="text-xs text-gray-500 mt-1">{hint}</p> : null}
    </div>
  );
}

function labelForKey(k){
  const map = {
    titleField: "Title field",
    courseField: "Course",
    componentField: "Component",
    sectionField: "Section",
    startDateField: "Start date",
    endDateField: "End date",
    startTimeField: "Start time",
    endTimeField: "End time",
    daysField: "Days pattern",
    locationField: "Location",
    descField: "Description",
  };
  return map[k] || k;
}

function guessInitialMapping(headers) {
  const h = headers.map((s)=>s.toLowerCase());
  const pick = (...candidates) => {
    for (const c of candidates) {
      const i = h.findIndex((x)=> x.includes(c));
      if (i !== -1) return headers[i];
    }
    return "";
  };

  const daysHeader = pick("days", "meeting pattern", "meets", "days of week");
  const startDateHeader = pick("start date", "from date", "first day");
  const endDateHeader = pick("end date", "to date", "last day");
  const startTimeHeader = pick("start time", "time start", "from time", "begin time");
  const endTimeHeader = pick("end time", "time end", "to time", "finish time");
  const locationHeader = pick("location", "room", "building");
  const titleHeader = pick("title");
  const courseHeader = pick("course", "subject");
  const sectionHeader = pick("section");
  const componentHeader = pick("component", "type");

  return {
    titleField: titleHeader,
    courseField: courseHeader,
    componentField: componentHeader,
    sectionField: sectionHeader,
    startDateField: startDateHeader,
    endDateField: endDateHeader,
    startTimeField: startTimeHeader,
    endTimeField: endTimeHeader,
    daysField: daysHeader,
    locationField: locationHeader,
  };
}

function parseCSV(text) {
  // Simple CSV parser that handles commas, quotes, and newlines
  const rows = [];
  let i = 0; let cur = []; let field = ""; let inQuotes = false;
  const pushField = () => { cur.push(field); field = ""; };
  const pushRow = () => { rows.push(cur); cur = []; };

  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i+1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else { field += ch; }
    } else {
      if (ch === '"') { inQuotes = true; }
      else if (ch === ',') { pushField(); }
      else if (ch === '\n') { pushField(); pushRow(); }
      else if (ch === '\r') { /* ignore */ }
      else { field += ch; }
    }
    i++;
  }
  // last field
  pushField();
  // if last row not pushed
  if (cur.length) pushRow();

  // First row as header
  const hdrs = rows.shift() || [];
  const data = rows.filter(r => r.length && r.some(x => String(x).trim() !== ""))
                   .map(r => Object.fromEntries(hdrs.map((h, idx) => [h, r[idx] ?? ""])));
  return { data, hdrs };
}

function parseDate(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  const s = String(val).trim();
  if (!s) return null;
  // Try ISO or locale-like
  const try1 = new Date(s);
  if (!isNaN(try1)) return new Date(try1.getFullYear(), try1.getMonth(), try1.getDate());
  // Try numeric like 20250904
  const m = s.match(/^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})$/);
  if (m) {
    const [_, y, mo, d] = m;
    return new Date(Number(y), Number(mo)-1, Number(d));
  }
  return null;
}

function parseTime(val) {
  if (!val && val !== 0) return null;
  let s = String(val).trim();
  if (!s) return null;
  // Excel sometimes stores times as HH:MM:SS, or 8:00 AM, or 8 AM, or 20:00
  s = s.toUpperCase();
  // Handle 8 AM format
  let m = s.match(/^(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*(AM|PM)?$/);
  if (m) {
    let h = Number(m[1]);
    const min = Number(m[2] || 0);
    const sec = Number(m[3] || 0);
    const ampm = m[4];
    if (ampm === "AM") { if (h === 12) h = 0; }
    else if (ampm === "PM") { if (h < 12) h += 12; }
    return { h, min, sec };
  }
  // 20:30 or 20:30:00
  m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    return { h: Number(m[1]), min: Number(m[2]), sec: Number(m[3] || 0) };
  }
  return null;
}

function parseDays(val) {
  const set = new Set();
  if (!val && val !== 0) return set;
  let s = String(val).trim();
  if (!s) return set;
  s = s.toUpperCase();
  // Normalize common words to tokens
  const rep = [
    [/THURSDAYS?/, "R"], [/THURS?\b/, "R"], [/THU\b/, "R"], [/\bTH\b/, "R"],
    [/TUESDAYS?/, "TU"], [/TUES?\b/, "TU"], [/\bTUE\b/, "TU"], [/\bTU\b/, "TU"],
    [/MONDAYS?/, "MO"], [/\bMON\b/, "MO"], [/\bMO\b/, "MO"],
    [/WEDNESDAYS?/, "WE"], [/\bWED\b/, "WE"], [/\bWE\b/, "WE"],
    [/FRIDAYS?/, "FR"], [/\bFRI\b/, "FR"], [/\bFR\b/, "FR"],
    [/SATURDAYS?/, "SA"], [/\bSAT\b/, "SA"], [/\bSA\b/, "SA"],
    [/SUNDAYS?/, "SU"], [/\bSUN\b/, "SU"], [/\bSU\b/, "SU"],
  ];
  for (const [pattern, token] of rep) s = s.replace(pattern, token);
  // Replace separators with space
  s = s.replace(/[,&/\\|]+/g, " ").replace(/\s+/g, " ").trim();
  const tokens = new Set();
  if (!s.includes(" ")) {
    // Possibly concatenated like MWF or MTWRF
    for (let i=0; i<s.length; i++) {
      const c = s[i];
      if (c === 'T') {
        // could be TU or R will be separate
        if (s.slice(i, i+2) === 'TU') { tokens.add('TU'); i++; continue; }
      }
      if (c === 'R') { tokens.add('R'); continue; }
      if (c === 'M') { tokens.add('MO'); continue; }
      if (c === 'W') { tokens.add('WE'); continue; }
      if (c === 'F') { tokens.add('FR'); continue; }
      if (c === 'S') {
        // try SA or SU
        if (s.slice(i, i+2) === 'SA') { tokens.add('SA'); i++; continue; }
        if (s.slice(i, i+2) === 'SU') { tokens.add('SU'); i++; continue; }
      }
    }
  } else {
    s.split(" ").forEach(tok => {
      tok = tok.trim(); if (!tok) return;
      if (tok === 'M') tok = 'MO';
      if (tok === 'T') tok = 'TU';
      if (tok === 'W') tok = 'WE';
      if (tok === 'TH') tok = 'R';
      if (tok === 'F') tok = 'FR';
      if (tok === 'R') tok = 'R';
      tokens.add(tok);
    });
  }
  for (const t of tokens) {
    switch (t) {
      case 'MO': set.add(1); break;
      case 'TU': set.add(2); break;
      case 'WE': set.add(3); break;
      case 'R': set.add(4); break; // Thursday
      case 'FR': set.add(5); break;
      case 'SA': set.add(6); break;
      case 'SU': set.add(0); break;
    }
  }
  return set;
}

function expandOccurrences(startDate, endDate, daySet) {
  const out = [];
  if (!startDate || !endDate) return out;
  const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (d <= end) {
    if (daySet.has(d.getDay())) out.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function combineDateTime(date, t) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), t.h, t.min, t.sec || 0);
  return d;
}

function icsDateTimeLocal(dt) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    dt.getFullYear().toString() +
    pad(dt.getMonth() + 1) +
    pad(dt.getDate()) +
    "T" +
    pad(dt.getHours()) +
    pad(dt.getMinutes()) +
    pad(dt.getSeconds())
  );
}

function icsDateTimeUTC(dt) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    dt.getUTCFullYear().toString() +
    pad(dt.getUTCMonth() + 1) +
    pad(dt.getUTCDate()) +
    "T" +
    pad(dt.getUTCHours()) +
    pad(dt.getUTCMinutes()) +
    pad(dt.getUTCSeconds()) +
    "Z"
  );
}

function icsEscape(s) {
  return (s || "")
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/[,;]/g, (m) => `\\${m}`);
}

function foldLines(text) {
  // 75 octets per RFC 5545. This naive wrapper works for ASCII
  const lines = text.split("\n");
  const out = [];
  for (const ln of lines) {
    if (ln.length <= 75) { out.push(ln); continue; }
    let i = 0; const L = ln.length;
    while (i < L) {
      const chunk = ln.slice(i, i+75);
      out.push(i === 0 ? chunk : " " + chunk);
      i += 75;
    }
  }
  return out.join("\n");
}

function buildICS(events, calName, tzHint) {
  const now = new Date();
  let ics = "BEGIN:VCALENDAR\n" +
            "VERSION:2.0\n" +
            "CALSCALE:GREGORIAN\n" +
            "PRODID:-//GistTools//Workday Excel to iCal//EN\n" +
            (calName ? `X-WR-CALNAME:${icsEscape(calName)}\n` : "") +
            (tzHint ? `X-WR-TIMEZONE:${icsEscape(tzHint)}\n` : "");

  for (const ev of events) {
    const uid = `${Math.random().toString(36).slice(2)}@gisttools.local`;
    ics += foldLines(
      "BEGIN:VEVENT\n" +
      `UID:${uid}\n` +
      `DTSTAMP:${icsDateTimeUTC(now)}\n` +
      `DTSTART:${icsDateTimeLocal(ev.dtStart)}\n` +
      `DTEND:${icsDateTimeLocal(ev.dtEnd)}\n` +
      (ev.summary ? `SUMMARY:${icsEscape(ev.summary)}\n` : "") +
      (ev.location ? `LOCATION:${icsEscape(ev.location)}\n` : "") +
      (ev.description ? `DESCRIPTION:${icsEscape(ev.description)}\n` : "") +
      "END:VEVENT"
    ) + "\n";
  }
  ics += "END:VCALENDAR\n";
  return ics;
}

function downloadText(text, name) {
  const blob = new Blob([text], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function sanitizeFileName(s) {
  return s.replace(/[^a-z0-9-_]+/gi, "-").replace(/-+/g, "-").replace(/^-|-$/g, "").toLowerCase();
}
