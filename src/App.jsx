import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

function formatDate(dateVal) {
  const d = dateVal ? new Date(`${dateVal}T00:00:00`) : new Date();
  const day = d.getDate();
  const suffix =
    [, "st", "nd", "rd"][(day % 100) - 11 < 3 && day % 100 > 10 ? 0 : day % 10] ||
    "th";
  const months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  return {
    day,
    suffix,
    month: months[d.getMonth()],
    year: d.getFullYear(),
  };
}

function drawDateText(doc, dateObj, x, y, fontSize) {
  const dayStr = String(dateObj.day);
  const monthYear = ` ${dateObj.month} ${dateObj.year}`;
  doc.setFont("times", "normal");
  doc.setFontSize(fontSize);
  doc.text(dayStr, x, y);
  const dayW = doc.getTextWidth(dayStr);
  doc.setFontSize(fontSize * 0.65);
  doc.text(dateObj.suffix, x + dayW, y - fontSize * 0.18);
  doc.setFontSize(fontSize);
  const supW = doc.getTextWidth(dateObj.suffix) * 0.65;
  doc.text(monthYear, x + dayW + supW + 0.3, y);
}

function drawTextLogo(doc, text, x, y) {
  doc.setFont("times", "bold");
  doc.setFontSize(10);
  doc.setTextColor(30, 30, 30);
  doc.text(text, x, y + 8);
}

function App() {
  const todayStr = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const fileInputRef = useRef(null);
  const [theme, setTheme] = useState(() => {
    if (typeof window === "undefined") return "dark";
    const storedTheme = window.localStorage.getItem("attendance-theme");
    if (storedTheme === "light" || storedTheme === "dark") return storedTheme;
    return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  });

  const [eventTitle, setEventTitle] = useState("Pre-Placement Talk & Interviews");
  const [company, setCompany] = useState("Capgemini");
  const [eventDate, setEventDate] = useState(todayStr);
  const [venue, setVenue] = useState("Sun Hall, Chitkara University");
  const [rowsPerPage, setRowsPerPage] = useState(25);

  const [attendees, setAttendees] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [nameCol, setNameCol] = useState(0);
  const [rollCol, setRollCol] = useState(0);
  const [fileName, setFileName] = useState("");
  const [dragActive, setDragActive] = useState(false);

  const [logo1Data, setLogo1Data] = useState(null);
  const [logo2Data, setLogo2Data] = useState(null);

  const [status, setStatus] = useState({ msg: "", type: "" });
  const [isGenerating, setIsGenerating] = useState(false);

  const [pdfUrl, setPdfUrl] = useState("");
  const [pdfName, setPdfName] = useState("");
  const [previewLabel, setPreviewLabel] = useState("Preview");

  useEffect(() => {
    return () => {
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    };
  }, [pdfUrl]);

  useEffect(() => {
    document.title = "Explore Labs Attandance Generator";
    document.documentElement.style.colorScheme = theme;
    window.localStorage.setItem("attendance-theme", theme);
  }, [theme]);

  const isDark = theme === "dark";

  const statusColorClass =
    status.type === "error"
      ? isDark
        ? "text-rose-300"
        : "text-rose-600"
      : status.type === "success"
        ? isDark
          ? "text-emerald-300"
          : "text-emerald-700"
        : isDark
          ? "text-slate-300"
          : "text-slate-600";

  const readLogo = (file, which) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const result = e.target?.result;
      if (typeof result !== "string") return;
      if (which === "logo1") setLogo1Data(result);
      else setLogo2Data(result);
    };
    reader.readAsDataURL(file);
  };

  const handleFile = (file) => {
    if (!file) return;
    setStatus({ msg: "Reading file...", type: "" });

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (!data.length || !Array.isArray(data[0])) {
          throw new Error("Invalid file format.");
        }

        const hdrs = data[0].map((h) => String(h || "").trim());
        const rows = data
          .slice(1)
          .filter((r) => Array.isArray(r) && r.some((c) => c !== undefined && c !== ""));

        let detectedName = 0;
        let detectedRoll = 0;
        hdrs.forEach((h, i) => {
          if (/name/i.test(h)) detectedName = i;
          if (/roll/i.test(h)) detectedRoll = i;
        });

        setHeaders(hdrs);
        setAttendees(rows);
        setNameCol(detectedName);
        setRollCol(detectedRoll);
        setFileName(file.name);
        setStatus({ msg: `${rows.length} attendees loaded successfully.`, type: "success" });
      } catch {
        setStatus({
          msg: "Could not read file. Please use a valid .xlsx or .xls file.",
          type: "error",
        });
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const generatePDF = async () => {
    const safeRowsPerPage = Math.max(5, Math.min(50, Number.parseInt(rowsPerPage, 10) || 25));
    if (!attendees.length) {
      setStatus({ msg: "Please upload an Excel file with attendee data first.", type: "error" });
      return;
    }

    setIsGenerating(true);
    setStatus({ msg: "Generating PDF...", type: "" });

    try {
      const dateObj = formatDate(eventDate);
      const formattedDatePlain = `${dateObj.day}${dateObj.suffix} ${dateObj.month} ${dateObj.year}`;

      const title = eventTitle.trim() || "Attendance Sheet";
      const companyName = company.trim();
      const venueName = venue.trim();

      const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pageW = 210;
      const pageH = 297;
      const pageMargin = 14;
      const tableW = 120;
      const tableLeft = (pageW - tableW) / 2;

      const pages = [];
      for (let i = 0; i < attendees.length; i += safeRowsPerPage) {
        pages.push(attendees.slice(i, i + safeRowsPerPage));
      }

      for (let p = 0; p < pages.length; p += 1) {
        if (p > 0) doc.addPage();

        const rows = pages[p];
        const startNum = p * safeRowsPerPage + 1;

        let y = 12;
        const logoH = 16;
        const logoW = 42;
        const logo1H = logoH + 8;

        if (logo1Data) {
          try {
            doc.addImage(logo1Data, undefined, pageMargin, y, logoW, logo1H, undefined, "FAST");
          } catch {
            drawTextLogo(doc, "UNIVERSITY", pageMargin, y);
          }
        } else {
          drawTextLogo(doc, "UNIVERSITY", pageMargin, y);
        }

        doc.setFont("times", "bold");
        doc.setFontSize(15);
        doc.setTextColor(20, 20, 20);
        doc.text(title, pageW / 2, y + 12, { align: "center" });

        if (logo2Data) {
          try {
            doc.addImage(
              logo2Data,
              undefined,
              pageW - pageMargin - logoW,
              y,
              logoW,
              logoH,
              undefined,
              "FAST",
            );
          } catch {
            if (companyName) {
              drawTextLogo(doc, companyName.toUpperCase(), pageW - pageMargin - logoW, y);
            }
          }
        } else if (companyName) {
          drawTextLogo(doc, companyName.toUpperCase(), pageW - pageMargin - logoW, y);
        }

        y += logoH + 6;
        doc.setDrawColor(200, 195, 190);
        doc.setLineWidth(0.4);
        doc.line(pageMargin, y, pageW - pageMargin, y);
        y += 6;

        doc.setTextColor(60, 60, 60);
        doc.text("Date: ", pageMargin, y);
        drawDateText(doc, dateObj, pageMargin + doc.getTextWidth("Date: "), y, 10);

        if (venueName) {
          doc.setFont("times", "normal");
          doc.setFontSize(10);
          doc.setTextColor(60, 60, 60);
          doc.text(`Venue: ${venueName}`, pageW - pageMargin, y, { align: "right" });
        }
        y += 9;

        const colW = [12, 44, 34, 30];
        const headerNames = ["S.No.", "Name", "Roll No", "Signature"];
        const rowH = 10;

        doc.setFillColor(8, 128, 124);
        doc.rect(tableLeft, y, tableW, rowH, "F");
        doc.setDrawColor(0, 100, 100);
        doc.setLineWidth(0.4);
        doc.rect(tableLeft, y, tableW, rowH, "S");

        doc.setFont("times", "bold");
        doc.setFontSize(10);
        doc.setTextColor(255, 255, 255);

        let colX = tableLeft;
        headerNames.forEach((h, i) => {
          doc.text(h, colX + 2, y + 7);
          if (i < headerNames.length - 1) {
            doc.setDrawColor(255, 255, 255);
            doc.line(colX + colW[i], y, colX + colW[i], y + rowH);
          }
          colX += colW[i];
        });
        y += rowH;

        rows.forEach((row, idx) => {
          const sno = String(startNum + idx);
          const name = String(row[nameCol] !== undefined ? row[nameCol] : "").trim();
          const roll = String(row[rollCol] !== undefined ? row[rollCol] : "").trim();

          doc.setDrawColor(160, 160, 160);
          doc.setLineWidth(0.2);
          doc.rect(tableLeft, y, tableW, rowH, "S");

          let dividerX = tableLeft;
          colW.slice(0, -1).forEach((w) => {
            dividerX += w;
            doc.setDrawColor(160, 160, 160);
            doc.line(dividerX, y, dividerX, y + rowH);
          });

          doc.setFont("times", "normal");
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);

          let valX = tableLeft;
          [sno, name, roll, ""].forEach((val, i) => {
            if (val) doc.text(val, valX + 2.5, y + 7);
            valX += colW[i];
          });

          y += rowH;
        });

        doc.setFontSize(8);
        doc.setFont("times", "normal");
        doc.setTextColor(160, 155, 150);
        doc.text(`Page ${p + 1} of ${pages.length}`, pageW / 2, pageH - 9, { align: "center" });
        if (companyName) doc.text(companyName, pageMargin, pageH - 9);
        doc.text(formattedDatePlain, pageW - pageMargin, pageH - 9, { align: "right" });
      }

      const blob = doc.output("blob");
      const newUrl = URL.createObjectURL(blob);
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);

      const safeCompany = (company.trim() || "event").replace(/\s+/g, "_");
      const safeDate = eventDate || new Date().toISOString().slice(0, 10);
      const downloadName = `${safeCompany}_Attendance_${safeDate}.pdf`;

      setPdfUrl(newUrl);
      setPdfName(downloadName);
      setPreviewLabel(`Preview - ${pages.length} page(s), ${attendees.length} students`);
      setStatus({
        msg: `Done! ${pages.length} page(s) generated for ${attendees.length} students.`,
        type: "success",
      });
    } catch (error) {
      setStatus({
        msg: `Error generating PDF: ${error instanceof Error ? error.message : "Unknown error"}`,
        type: "error",
      });
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div
      className={`relative min-h-screen overflow-x-hidden ${isDark ? "theme-dark bg-slate-950 text-slate-100" : "theme-light bg-slate-100 text-slate-900"}`}
    >
      <div className="pointer-events-none absolute inset-0 -z-10">
        <div
          className={`absolute -left-20 -top-20 h-80 w-80 rounded-full blur-3xl ${isDark ? "bg-cyan-400/25" : "bg-sky-400/30"}`}
        />
        <div
          className={`absolute right-0 top-1/4 h-96 w-96 rounded-full blur-3xl ${isDark ? "bg-fuchsia-500/20" : "bg-orange-300/35"}`}
        />
        <div
          className={`absolute bottom-0 left-1/3 h-96 w-96 rounded-full blur-3xl ${isDark ? "bg-emerald-400/15" : "bg-emerald-300/30"}`}
        />
      </div>

      <header className={`border-b backdrop-blur-xl ${isDark ? "border-white/10 bg-black/30" : "border-slate-200/80 bg-white/70"}`}>
        <div className="mx-auto flex max-w-6xl items-center gap-4 px-6 py-5">
          <div className="grid h-11 w-11 place-content-center rounded-xl bg-gradient-to-br from-cyan-300 to-emerald-300 text-lg font-black text-slate-950 shadow-[0_0_40px_rgba(34,211,238,0.45)]">
            EL
          </div>
          <div className="flex-1">
            <h1 className={`font-display text-lg font-bold tracking-wide md:text-2xl ${isDark ? "text-white" : "text-slate-900"}`}>
              Explore Labs Attandance Generator
            </h1>
            <p className={`text-sm ${isDark ? "text-cyan-100/75" : "text-slate-600"}`}>
              Upload Excel, configure fields, then generate polished attendance PDFs
            </p>
          </div>
          <button
            type="button"
            onClick={() => setTheme(isDark ? "light" : "dark")}
            className={`rounded-xl px-3 py-2 text-xs font-semibold transition ${
              isDark
                ? "border border-white/20 bg-white/5 text-cyan-100 hover:bg-white/10"
                : "border border-slate-300 bg-white text-slate-700 hover:bg-slate-100"
            }`}
          >
            <span className="inline-flex items-center gap-2">
              {isDark ? (
                <svg
                  aria-hidden="true"
                  viewBox="0 0 24 24"
                  className="h-4 w-4"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.8"
                >
                  <circle cx="12" cy="12" r="4" />
                  <path d="M12 2v2.5M12 19.5V22M4.9 4.9l1.8 1.8M17.3 17.3l1.8 1.8M2 12h2.5M19.5 12H22M4.9 19.1l1.8-1.8M17.3 6.7l1.8-1.8" />
                </svg>
              ) : (
                <svg
                  aria-hidden="true"
                  viewBox="0 0 24 24"
                  className="h-4 w-4"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.8"
                >
                  <path d="M20 15.5A8.5 8.5 0 1 1 8.5 4a7 7 0 0 0 11.5 11.5Z" />
                </svg>
              )}
              {isDark ? "Switch to Light" : "Switch to Dark"}
            </span>
          </button>
        </div>
      </header>

      <main className="mx-auto grid w-full max-w-6xl gap-6 px-4 py-8 md:px-6 lg:grid-cols-3">
        <section className="space-y-6 lg:col-span-2">
          <div className="glass-card p-5 md:p-6">
            <p className="section-label">Event details</p>
            <div className="mt-4 grid gap-4 md:grid-cols-2">
              <label className="field-wrap md:col-span-1">
                <span className="field-label">Event title</span>
                <input
                  className="field-input"
                  value={eventTitle}
                  onChange={(e) => setEventTitle(e.target.value)}
                  placeholder="Pre-Placement Talk & Interviews"
                />
              </label>

              <label className="field-wrap md:col-span-1">
                <span className="field-label">Company / organisation</span>
                <input
                  className="field-input"
                  value={company}
                  onChange={(e) => setCompany(e.target.value)}
                  placeholder="Capgemini"
                />
              </label>

              <label className="field-wrap">
                <span className="field-label">Date</span>
                <input
                  type="date"
                  className="field-input"
                  value={eventDate}
                  onChange={(e) => setEventDate(e.target.value)}
                />
              </label>

              <label className="field-wrap">
                <span className="field-label">Venue</span>
                <input
                  className="field-input"
                  value={venue}
                  onChange={(e) => setVenue(e.target.value)}
                  placeholder="Sun Hall, Chitkara University"
                />
              </label>

              <label className="field-wrap md:col-span-2">
                <span className="field-label">Rows per page</span>
                <input
                  type="number"
                  min={5}
                  max={50}
                  className="field-input"
                  value={rowsPerPage}
                  onChange={(e) => setRowsPerPage(e.target.value)}
                />
                <span className="mt-1 text-xs text-cyan-100/65">Recommended range: 20-30</span>
              </label>
            </div>
          </div>

          <div className="glass-card p-5 md:p-6">
            <p className="section-label">Attendee list (Excel)</p>

            <button
              type="button"
              className={`mt-4 w-full rounded-2xl border-2 border-dashed px-4 py-10 text-left transition ${
                dragActive
                  ? isDark
                    ? "border-cyan-300 bg-cyan-300/10"
                    : "border-sky-400 bg-sky-200/30"
                  : isDark
                    ? "border-white/25 bg-white/5 hover:border-cyan-200/80 hover:bg-cyan-200/5"
                    : "border-slate-300 bg-white/70 hover:border-sky-500 hover:bg-sky-100/40"
              }`}
              onClick={() => fileInputRef.current?.click()}
              onDragOver={(e) => {
                e.preventDefault();
                setDragActive(true);
              }}
              onDragLeave={() => setDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragActive(false);
                handleFile(e.dataTransfer.files?.[0]);
              }}
            >
              <div className="mx-auto max-w-md text-center">
                <div className="mb-3 text-4xl">📂</div>
                <h3 className={`font-display text-lg font-semibold ${isDark ? "text-white" : "text-slate-900"}`}>
                  Drop your Excel file here
                </h3>
                <p className={`mt-1 text-sm ${isDark ? "text-cyan-100/70" : "text-slate-600"}`}>
                  Supports .xlsx, .xls, .csv. Include Name and Roll No columns.
                </p>
              </div>
            </button>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => handleFile(e.target.files?.[0])}
            />

            {fileName && (
              <div className="mt-4 flex items-center gap-2 rounded-xl border border-emerald-300/40 bg-emerald-300/10 px-4 py-2 text-sm text-emerald-100">
                <span className="inline-block h-2 w-2 rounded-full bg-emerald-300" />
                <span className="font-medium">{fileName}</span>
                <span className="text-emerald-100/70">- {attendees.length} students</span>
              </div>
            )}

            {!!headers.length && (
              <div className="mt-4 grid gap-4 md:grid-cols-2">
                <label className="field-wrap">
                  <span className="field-label">Name column</span>
                  <select
                    className="field-input"
                    value={nameCol}
                    onChange={(e) => setNameCol(Number(e.target.value))}
                  >
                    {headers.map((h, i) => (
                      <option key={`${h}-${i}`} value={i}>
                        {h || `Column ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>

                <label className="field-wrap">
                  <span className="field-label">Roll no column</span>
                  <select
                    className="field-input"
                    value={rollCol}
                    onChange={(e) => setRollCol(Number(e.target.value))}
                  >
                    {headers.map((h, i) => (
                      <option key={`${h}-${i}`} value={i}>
                        {h || `Column ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
            )}
          </div>
        </section>

        <aside className="space-y-6">
          <div className="glass-card p-5 md:p-6">
            <p className="section-label">Branding</p>
            <div className="mt-4 space-y-4">
              <label className="field-wrap">
                <span className="field-label">Left logo (University)</span>
                <input
                  type="file"
                  accept="image/*"
                  className="field-file"
                  onChange={(e) => readLogo(e.target.files?.[0], "logo1")}
                />
                {logo1Data && <img src={logo1Data} alt="University logo" className="logo-preview" />}
              </label>

              <label className="field-wrap">
                <span className="field-label">Right logo (Company)</span>
                <input
                  type="file"
                  accept="image/*"
                  className="field-file"
                  onChange={(e) => readLogo(e.target.files?.[0], "logo2")}
                />
                {logo2Data && <img src={logo2Data} alt="Company logo" className="logo-preview" />}
              </label>
            </div>
          </div>

          <div className="glass-card p-5 md:p-6">
            <button
              type="button"
              onClick={generatePDF}
              disabled={isGenerating}
              className="cta-btn"
            >
              {isGenerating ? "Generating..." : "Generate Attendance Sheet PDF"}
            </button>
            <p className={`mt-3 min-h-5 text-sm ${statusColorClass}`}>{status.msg}</p>
          </div>

          {pdfUrl && (
            <div className="glass-card p-4">
              <div className="mb-3 flex items-center justify-between gap-3">
                <p className={`text-sm font-semibold ${isDark ? "text-white" : "text-slate-900"}`}>{previewLabel}</p>
                <a href={pdfUrl} download={pdfName} className="download-btn">
                  Download PDF
                </a>
              </div>
              <iframe src={pdfUrl} title="Attendance PDF preview" className="h-[420px] w-full rounded-xl border border-white/15 bg-white" />
            </div>
          )}
        </aside>
      </main>

      <footer className="mx-auto mt-6 w-full max-w-6xl px-4 pb-8 md:px-6">
        <div className={`h-px w-full ${isDark ? "bg-white/15" : "bg-slate-300"}`} />
        <div className="flex flex-col items-start justify-between gap-3 pt-5 md:flex-row md:items-center">
          <div>
            <p className={`font-display text-sm font-semibold md:text-base ${isDark ? "text-white" : "text-slate-900"}`}>
              Developed and Operated by Dr Isha Kansal
            </p>
            <p className={`mt-1 text-sm ${isDark ? "text-cyan-100/80" : "text-slate-600"}`}>
              Designed by KalawatiPutra.com
            </p>
          </div>

          <a
            href="https://kalawatiputra.com"
            target="_blank"
            rel="noreferrer"
            className={`text-sm font-semibold underline-offset-4 transition hover:underline ${
              isDark ? "text-cyan-200" : "text-sky-700"
            }`}
          >
            KalawatiPutra.com
          </a>
        </div>
      </footer>
    </div>
  );
}

export default App;