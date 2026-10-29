import {
  auth,
  db,
  GoogleAuthProvider,
  signInWithPopup,
  signOut,
  onAuthStateChanged
} from './firebase.js';

import {
  collection,
  getDocs
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const jsPDF = window.jspdf.jsPDF;
import 'https://cdn.jsdelivr.net/npm/jspdf-autotable@3.5.28/+esm';
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-latest/package/xlsx.mjs';

const loginBtn = document.getElementById("login-btn");
const logoutBtn = document.getElementById("logout-btn");
const userInfo = document.getElementById("user-info");
const content = document.getElementById("content");
const generateMissedDetentionPdfBtn = document.getElementById("generate-missed-detention-pdf");
const generateMissedDetentionReceiptPdfBtn = document.getElementById("generate-missed-detention-receipt-pdf");
const exportOutstandingDetentionsBtn = document.getElementById("export-outstanding-detentions");
const missedSlipsMaxDaysInput = document.getElementById("missed-slips-max-days");
const outstandingExcelMaxDaysInput = document.getElementById("outstanding-excel-max-days");

let allStudents = [];
const selectedStudentIds = new Set();

loginBtn.onclick = async () => {
  const provider = new GoogleAuthProvider();
  await signInWithPopup(auth, provider).catch(err => {
    alert("Login failed");
    console.error(err);
  });
};

logoutBtn.onclick = () => signOut(auth);

onAuthStateChanged(auth, async user => {
  if (user) {
    userInfo.textContent = `Signed in as: ${user.displayName} (${user.email})`;
    loginBtn.style.display = "none";
    logoutBtn.style.display = "inline-block";
    content.style.display = "block";
    await loadStudents();
  } else {
    userInfo.textContent = "";
    loginBtn.style.display = "inline-block";
    logoutBtn.style.display = "none";
    content.style.display = "none";
  }
});

exportOutstandingDetentionsBtn.addEventListener("click", async () => {
  await runWithButtonLoading(exportOutstandingDetentionsBtn, "Exporting...", exportOutstandingDetentionsList);
});

generateMissedDetentionPdfBtn.addEventListener("click", async () => {
  await runWithButtonLoading(generateMissedDetentionPdfBtn, "Generating...", exportMissedDetentionNoticePdf);
});

generateMissedDetentionReceiptPdfBtn.addEventListener("click", async () => {
  await runWithButtonLoading(generateMissedDetentionReceiptPdfBtn, "Generating...", exportMissedDetentionNoticeReceiptPdf);
});

async function runWithButtonLoading(button, loadingText, action) {
  const originalText = button.textContent;
  button.disabled = true;
  button.classList.add("loading-button");
  button.textContent = loadingText;

  try {
    await action();
  } finally {
    button.textContent = originalText;
    button.classList.remove("loading-button");
    button.disabled = false;
  }
}

async function loadStudents() {
  const studentSnapshot = await getDocs(collection(db, "students"));

  allStudents = studentSnapshot.docs.map(docSnap => {
    const data = docSnap.data();
    const detentions = getDetentionLedger(data);
    const detentionServedEvents = getDetentionServedEvents(data);
    const detentionStatus = buildDetentionStatus(detentions, detentionServedEvents);
    return {
      studentId: docSnap.id,
      surname: data.surname || '',
      givenName: data.givenName || '',
      rollClass: data.rollClass || '',
      yearGroup: resolveYearGroup(data),
      lateCount: data.lateCount || 0,
      detentionsServed: data.detentionsServed || 0,
      detentionHistory: data.detentionHistory || [],
      lateArrivals: data.lateArrivals || [],
      activeDetention: data.activeDetention || null,
      detentions,
      detentionServedEvents,
      detentionStatus
    };
  }).sort((a, b) => a.surname.localeCompare(b.surname) || a.givenName.localeCompare(b.givenName));
}

function renderFilters() {
  const rollClasses = [...new Set(allStudents.map(student => student.rollClass).filter(Boolean))].sort();
  const yearGroups = [...new Set(allStudents.map(student => student.yearGroup).filter(Boolean))].sort((a, b) => Number(a) - Number(b));

  rollClassSelect.innerHTML = '<option value="">All Roll Classes</option>' + rollClasses.map(value => `<option value="${value}">${value}</option>`).join('');
  yearGroupSelect.innerHTML = '<option value="">All Year Groups</option>' + yearGroups.map(value => `<option value="${value}">${value}</option>`).join('');
}

function renderStudentPicker() {
  const search = studentSearch.value.trim().toLowerCase();
  const scope = historyScope.value;
  const rollClass = rollClassSelect.value;
  const yearGroup = yearGroupSelect.value;

  const visibleStudents = allStudents.filter(student => {
    if (rollClass && student.rollClass !== rollClass) return false;
    if (yearGroup && student.yearGroup !== yearGroup) return false;

    if (!search) return true;
    return `${student.givenName} ${student.surname} ${student.rollClass}`.toLowerCase().includes(search);
  });

  studentPicker.innerHTML = visibleStudents.map(student => `
    <label class="student-choice-row ${scope === 'single' ? 'single-choice' : ''}">
      <input
        class="student-choice"
        type="${scope === 'single' ? 'radio' : 'checkbox'}"
        name="student-selection"
        value="${student.studentId}"
        ${selectedStudentIds.has(student.studentId) ? 'checked' : ''}
      />
      <span>${student.surname}, ${student.givenName} (${student.rollClass})</span>
    </label>
  `).join('');
}

function exportSummaryReport() {
  const date = getFormattedDate();
  const rows = allStudents.map(student => ({
    surname: student.surname,
    givenName: student.givenName,
    rollClass: student.rollClass,
    lateCount: student.lateCount,
    detentionsServed: student.detentionsServed,
    outstandingDetentions: student.detentionStatus?.outstandingCount || 0,
    oldestOutstandingDetention: student.detentionStatus?.oldestOutstandingDetentionDate || 'None'
  }));

  if (exportFormat.value === "pdf") {
    const doc = new jsPDF();
    doc.text("Attendance Assistant Summary Report", 14, 15);
    doc.autoTable({
      startY: 22,
      head: [["Surname", "Given Name", "Roll Class", "Late Count", "Detentions Served", "Outstanding Detentions", "Oldest Outstanding"]],
      body: rows.map(row => [row.surname, row.givenName, row.rollClass, row.lateCount, row.detentionsServed, row.outstandingDetentions, row.oldestOutstandingDetention]),
      styles: { fontSize: 8 }
    });
    doc.save(`attendance_summary_${date}.pdf`);
    return;
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Summary");
  XLSX.writeFile(workbook, `attendance_summary_${date}.xlsx`);
}

function exportAllStudentsReport() {
  const date = getFormattedDate();
  const rows = allStudents.map(student => ({
    studentId: student.studentId,
    surname: student.surname,
    givenName: student.givenName,
    yearGroup: student.yearGroup || '',
    rollClass: student.rollClass,
    lateCount: student.lateCount,
    unjustifiedLateArrivalDates: getUnjustifiedLateArrivalDates(student).join(', '),
    mostRecentCompletedDetention: getMostRecentCompletedDetentionDate(student) || '',
    outstandingDetentions: student.detentionStatus?.outstandingCount || 0,
    oldestOutstandingDetention: student.detentionStatus?.oldestOutstandingDetentionDate || ''
  }));

  if (rows.length === 0) {
    alert("No student records were found.");
    return;
  }

  if (exportFormat.value === "pdf") {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.text("All Students Register", 14, 15);
    doc.autoTable({
      startY: 22,
      head: [["Student ID", "Surname", "Given Name", "Year", "Roll Class", "Late Count", "Unjustified Late Arrival Dates", "Most Recent Completed Detention", "Outstanding Detentions", "Oldest Outstanding"]],
      body: rows.map(row => [
        row.studentId,
        row.surname,
        row.givenName,
        row.yearGroup,
        row.rollClass,
        row.lateCount,
        row.unjustifiedLateArrivalDates,
        row.mostRecentCompletedDetention,
        row.outstandingDetentions,
        row.oldestOutstandingDetention
      ]),
      styles: { fontSize: 7 },
      columnStyles: {
        6: { cellWidth: 58 },
        7: { cellWidth: 28 }
      }
    });
    doc.save(`all_students_register_${date}.pdf`);
    return;
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "All Students");
  XLSX.writeFile(workbook, `all_students_register_${date}.xlsx`);
}

async function exportOutstandingDetentionsList() {
  const maximumDays = getMaximumDays(outstandingExcelMaxDaysInput, 0);
  if (maximumDays === null) return;

  const rows = buildOutstandingDetentionRows(maximumDays);
  if (rows.length === 0) {
    alert(`No outstanding detentions within ${maximumDays} day(s) were found.`);
    return;
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Outstanding Detentions");
  XLSX.writeFile(workbook, `outstanding_detentions_${getFormattedDate()}.xlsx`);
}

async function exportMissedDetentionNoticePdf() {
  const maximumDays = getMaximumDays(missedSlipsMaxDaysInput, 1);
  if (maximumDays === null) return;

  const rows = buildMissedDetentionNoticeRows(maximumDays);

  if (rows.length === 0) {
    alert(`No overdue outstanding detentions within ${maximumDays} day(s) were found.`);
    return;
  }

  const date = getLocalDateString();
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const slipHeight = pageHeight / 2;
  const marginX = 20;
  const maxTextWidth = pageWidth - (marginX * 2);

  const drawSlip = (student, top) => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.text(student.fullName, marginX, top + 14);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text(`Year ${student.yearGroup || "-"}    Roll Class: ${student.rollClass || "-"}`, marginX, top + 22);

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Unexplained late arrival date(s):", marginX, top + 31);

    doc.setFont("helvetica", "normal");
    const lateArrivalDateText = student.unexplainedLateArrivalDates.length > 0
      ? student.unexplainedLateArrivalDates.map(formatDateForDisplay).join(", ")
      : "Date unavailable";
    const lateArrivalDateLines = doc.splitTextToSize(lateArrivalDateText, maxTextWidth);
    doc.text(lateArrivalDateLines, marginX, top + 37, { lineHeightFactor: 1.25 });

    const parts = getMissedDetentionNoticeParts(student.yearGroup);
    const lh = size => (size * 1.3) / doc.internal.scaleFactor;
    let y = top + 37 + (lateArrivalDateLines.length * 5) + 6;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    const introLines = doc.splitTextToSize(parts.intro, maxTextWidth);
    doc.text(introLines, marginX, y, { lineHeightFactor: 1.3 });
    y += introLines.length * lh(10) + 2;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    const roomLines = doc.splitTextToSize(parts.roomLine, maxTextWidth);
    doc.text(roomLines, marginX, y, { lineHeightFactor: 1.3 });
    y += roomLines.length * lh(13) + 2;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    const outroLines = doc.splitTextToSize(parts.outro, maxTextWidth);
    doc.text(outroLines, marginX, y, { lineHeightFactor: 1.3 });
  };

  rows.forEach((student, index) => {
    const slot = index % 2;                       // 0 = top half, 1 = bottom half
    if (index > 0 && slot === 0) {
      doc.addPage();
    }

    drawSlip(student, slot * slipHeight);

    // Dashed cut line under the top slip (only if a second slip follows)
    if (slot === 0 && index < rows.length - 1) {
      doc.setDrawColor(150);
      doc.setLineDashPattern([2, 2], 0);
      doc.line(10, slipHeight, pageWidth - 10, slipHeight);
      doc.setLineDashPattern([], 0);
      doc.setDrawColor(0);
    }
  });

  doc.save(`missed-detention-notices-${date}.pdf`);
}

async function exportMissedDetentionNoticeReceiptPdf() {
  const maximumDays = getMaximumDays(missedSlipsMaxDaysInput, 1);
  if (maximumDays === null) return;

  const rows = buildMissedDetentionNoticeRows(maximumDays);

  if (rows.length === 0) {
    alert(`No overdue outstanding detentions within ${maximumDays} day(s) were found.`);
    return;
  }

  const date = getLocalDateString();

  const RECEIPT_WIDTH = 80;                 // mm, paper width
  const MARGIN_X = 6;                       // printable area is usually ~72mm on 80mm paper
  const MAX_TEXT_WIDTH = RECEIPT_WIDTH - MARGIN_X * 2;
  const BOTTOM_PADDING = 12;                // space so the cutter doesn't slice through text
  const LINE_FACTOR = 1.3;

  // Draws one slip and returns the y position where it ended.
  // Used on a scratch doc to measure, then again on the real doc to print.
  const drawSlip = (d, student) => {
    const lineHeight = size => (size * LINE_FACTOR) / d.internal.scaleFactor;
    let y = 8;

    d.setFont("helvetica", "bold");
    d.setFontSize(14);
    const nameLines = d.splitTextToSize(student.fullName, MAX_TEXT_WIDTH);
    d.text(nameLines, MARGIN_X, y, { lineHeightFactor: LINE_FACTOR });
    y += nameLines.length * lineHeight(14) + 1;

    d.setFont("helvetica", "normal");
    d.setFontSize(10);
    d.text(`Year ${student.yearGroup || "-"}  Roll Class: ${student.rollClass || "-"}`, MARGIN_X, y);
    y += lineHeight(10) + 3;

    d.setFont("helvetica", "bold");
    d.setFontSize(9);
    d.text("Unexplained late arrival date(s):", MARGIN_X, y);
    y += lineHeight(9);

    d.setFont("helvetica", "normal");
    const dateText = student.unexplainedLateArrivalDates.length > 0
      ? student.unexplainedLateArrivalDates.map(formatDateForDisplay).join(", ")
      : "Date unavailable";
    const dateLines = d.splitTextToSize(dateText, MAX_TEXT_WIDTH);
    d.text(dateLines, MARGIN_X, y, { lineHeightFactor: LINE_FACTOR });
    y += dateLines.length * lineHeight(9) + 4;

    const parts = getMissedDetentionNoticeParts(student.yearGroup);

    d.setFont("helvetica", "normal");
    d.setFontSize(8.5);
    const introLines = d.splitTextToSize(parts.intro, MAX_TEXT_WIDTH);
    d.text(introLines, MARGIN_X, y, { lineHeightFactor: LINE_FACTOR });
    y += introLines.length * lineHeight(8.5) + 1.5;

    d.setFont("helvetica", "bold");
    d.setFontSize(11);
    const roomLines = d.splitTextToSize(parts.roomLine, MAX_TEXT_WIDTH);
    d.text(roomLines, MARGIN_X, y, { lineHeightFactor: LINE_FACTOR });
    y += roomLines.length * lineHeight(11) + 1.5;

    d.setFont("helvetica", "normal");
    d.setFontSize(8.5);
    const outroLines = d.splitTextToSize(parts.outro, MAX_TEXT_WIDTH);
    d.text(outroLines, MARGIN_X, y, { lineHeightFactor: LINE_FACTOR });
    y += outroLines.length * lineHeight(8.5);

    return y;
  };

  // Pass 1: measure each slip on a scratch document
  const scratch = new jsPDF({ unit: "mm", format: [RECEIPT_WIDTH, 500] });
  const pageHeights = rows.map(student => drawSlip(scratch, student) + BOTTOM_PADDING);

  // Pass 2: build the real document, one custom-height page per student
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: [RECEIPT_WIDTH, pageHeights[0]] });
  rows.forEach((student, index) => {
    if (index > 0) {
      doc.addPage([RECEIPT_WIDTH, pageHeights[index]], "portrait");
    }
    drawSlip(doc, student);
  });

  doc.save(`missed-detention-notices-${date}.pdf`);
}

function exportHistoryReport() {
  const selectedStudents = resolveHistorySelection();
  if (selectedStudents.length === 0) {
    alert("Select at least one student for the history report.");
    return;
  }

  const date = getFormattedDate();
  const historyRows = selectedStudents.flatMap(student => {
    const arrivals = [...student.lateArrivals].sort((a, b) => String(a.date).localeCompare(String(b.date)));
    return arrivals.map(arrival => ({
      studentId: student.studentId,
      surname: student.surname,
      givenName: student.givenName,
      rollClass: student.rollClass,
      date: arrival.date || '',
      arrivalTime: arrival.arrivalTime || '',
      minutesLate: arrival.minutesLate ?? '',
      shorthand: arrival.shorthand || '',
      description: arrival.description || '',
      oldestOutstandingDetention: student.detentionStatus?.oldestOutstandingDetentionDate || ''
    }));
  });

  if (historyRows.length === 0) {
    alert("No late-arrival history was found for the selected student(s).");
    return;
  }

  if (exportFormat.value === "pdf") {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.text("Student Late Arrival History", 14, 15);
    doc.autoTable({
      startY: 22,
      head: [["Student ID", "Surname", "Given Name", "Roll Class", "Date", "Arrival Time", "Minutes Late", "Shorthand", "Description"]],
      body: historyRows.map(row => [row.studentId, row.surname, row.givenName, row.rollClass, row.date, row.arrivalTime, row.minutesLate, row.shorthand, row.description]),
      styles: { fontSize: 8 }
    });
    doc.save(`student_history_${date}.pdf`);
    return;
  }

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(historyRows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Student History");
  XLSX.writeFile(workbook, `student_history_${date}.xlsx`);
}

function resolveHistorySelection() {
  const scope = historyScope.value;
  const rollClass = rollClassSelect.value;
  const yearGroup = yearGroupSelect.value;

  if (scope === "rollClass") {
    return allStudents.filter(student => !rollClass || student.rollClass === rollClass);
  }

  if (scope === "yearGroup") {
    return allStudents.filter(student => !yearGroup || student.yearGroup === yearGroup);
  }

  return allStudents.filter(student => selectedStudentIds.has(student.studentId));
}

function getFormattedDate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function getLocalDateString() {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Australia/Sydney",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });
  const parts = formatter.formatToParts(new Date());
  const year = parts.find(part => part.type === "year")?.value || "";
  const month = parts.find(part => part.type === "month")?.value || "";
  const day = parts.find(part => part.type === "day")?.value || "";
  return `${year}-${month}-${day}`;
}

function getMissedWhilePresentCount(student) {
  return 0;
}

function getUnjustifiedLateArrivalDates(student) {
  return [...student.lateArrivals]
    .filter(arrival => arrival && arrival.justified !== true)
    .map(arrival => arrival.date || '')
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function getMostRecentCompletedDetentionDate(student) {
  const servedDates = [...student.detentionHistory]
    .filter(entry => entry?.outcome === "served")
    .map(entry => entry.date || entry.scheduledForDate || '')
    .filter(Boolean)
    .sort((a, b) => b.localeCompare(a));

  return servedDates[0] || '';
}

function getLatestLateArrivalDate(student) {
  return [...student.lateArrivals]
    .filter(arrival => arrival && arrival.justified !== true)
    .map(arrival => arrival.date || '')
    .filter(Boolean)
    .sort((a, b) => b.localeCompare(a))[0] || '';
}

function getLateArrivalCount(student) {
  return getUnjustifiedLateArrivalDates(student).length || student.lateCount || 0;
}

function getDetentionsServedCount(student) {
  const historyServedCount = Array.isArray(student.detentionHistory)
    ? student.detentionHistory.filter(entry => entry?.outcome === "served").length
    : 0;
  return Math.max(historyServedCount, student.detentionsServed || 0);
}

function buildMissedDetentionNoticeRows(maximumDays) {
  const today = getLocalDateString();

  return allStudents
    .map(student => {
      const detentions = getDetentionLedger(student);
      const servedEvents = getDetentionServedEvents(student);
      const latestServedDate = getLatestServedDate(servedEvents);
      const outstandingDetentions = getOutstandingDetentions(detentions, latestServedDate);
      const oldestOutstandingDate = outstandingDetentions
        .map(detention => detention.originalScheduledForDate)
        .filter(Boolean)
        .sort()[0];
      const daysOutstanding = countSchoolDaysElapsed(oldestOutstandingDate, today);
      return {
        student,
        outstandingDetentions,
        include: oldestOutstandingDate < today && daysOutstanding <= maximumDays
      };
    })
    .filter(entry => entry.include)
    .sort((a, b) =>
      compareYearGroups(a.student.yearGroup, b.student.yearGroup) ||
      a.student.surname.localeCompare(b.student.surname) ||
      a.student.givenName.localeCompare(b.student.givenName)
    )
    .map(({ student, outstandingDetentions }) => {
      const actualUnexplainedLateArrivalDates = new Set(getUnjustifiedLateArrivalDates(student));
      const outstandingLateArrivalDates = [...new Set(
        outstandingDetentions
          .map(detention => detention.sourceLateDate)
          .filter(date => actualUnexplainedLateArrivalDates.has(date))
      )].sort();

      return {
        fullName: formatStudentFullName(student.givenName, student.surname),
        yearGroup: student.yearGroup || "",
        rollClass: student.rollClass || "",
        unexplainedLateArrivalDates: outstandingLateArrivalDates
      };
    });
}

function getMissedDetentionNoticeParts(yearGroup) {
  const todayText = getLocalDateString();
  const weekday = formatWeekday(todayText);

  const breakTimes = {
    Monday: "10:35",
    Tuesday: "10:25",
    Wednesday: "10:35",
    Thursday: "10:25",
    Friday: "10:35"
  };
  const breakTime = breakTimes[weekday];
  const atTime = breakTime ? ` at ${breakTime}` : "";
  const firstBreakText = weekday
    ? `FIRST BREAK TODAY (${weekday})`
    : "FIRST BREAK TODAY";

  const year = Number.parseInt(yearGroup, 10);
  let roomLine;

  if (year >= 7 && year <= 9) {
    roomLine = `A7${atTime}`;
  } else if (year >= 10 && year <= 12) {
    roomLine = `A9${atTime}`;
  } else {
    // Year unknown: show both rooms
    roomLine = `Year 7, Year 8 & Year 9 in A7${atTime}\nYear 10, Year 11 & Year 12 in A9${atTime}`;
  }

  return {
    intro: `This week you arrived to school late (after roll call) and did not bring a note. You had a detention scheduled, but did not attend. Please attend the detention room at ${firstBreakText}, in room:`,
    roomLine,
    outro: [
      "If you refuse to attend a detention for your late arrival to school, you may receive an after school detention.",
      "",
      "If you believe this detention is an error, you must still attend the detention room and talk to the teacher on supervision."
    ].join("\n")
  };
}

function formatStudentFullName(givenName, surname) {
  return [givenName, surname].filter(Boolean).join(" ").trim() || "Student";
}

function formatDateForDisplay(dateText) {
  const match = String(dateText || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(dateText || "");
  const [, year, month, day] = match;
  return `${day}/${month}/${year}`;
}

function buildOutstandingDetentionRows(maximumDays) {
  const today = getLocalDateString();
  return allStudents
    .flatMap(student => {
      const detentions = getDetentionLedger(student);
      const detentionServedEvents = getDetentionServedEvents(student);
      const status = student.detentionStatus || buildDetentionStatus(detentions, detentionServedEvents);
      if (!status.hasOpenDetention) return [];

      const latestServedDate = getLatestServedDate(detentionServedEvents);
      const outstanding = getOutstandingDetentions(detentions, latestServedDate);
      const oldestDate = status.oldestOutstandingDetentionDate || "";
      const daysOutstanding = countSchoolDaysElapsed(oldestDate, today);
      if (daysOutstanding > maximumDays) return [];

      return [{
        "Days Since Original Detention": daysOutstanding,
        "Oldest Outstanding Detention Date": oldestDate,
        "Newest Outstanding Detention Date": status.newestOutstandingDetentionDate || "",
        "Outstanding Detention Count": status.outstandingCount,
        Surname: student.surname,
        "Given Name": student.givenName,
        Year: student.yearGroup || '',
        "Roll Class": student.rollClass,
        "Date of Most Recent Late Arrival": getLatestLateArrivalDate(student),
        "Number of Late Arrivals": getLateArrivalCount(student),
        "Number of Late Detentions Served": getDetentionsServedCount(student),
        "Latest Served Detention Date": latestServedDate || "",
        "Open Detention Source Dates": outstanding.map(detention => detention.sourceLateDate || "").filter(Boolean).join(", "),
        studentId: student.studentId,
        sortYear: student.yearGroup || '',
        sortSurname: student.surname,
        sortGivenName: student.givenName
      }];
    })
    .sort((a, b) =>
      Number(b["Days Since Original Detention"]) - Number(a["Days Since Original Detention"]) ||
      compareYearGroups(a.sortYear, b.sortYear) ||
      a.sortSurname.localeCompare(b.sortSurname) ||
      a.sortGivenName.localeCompare(b.sortGivenName)
    )
    .map(({ studentId, sortYear, sortSurname, sortGivenName, ...reportRow }) => reportRow);
}

function getMaximumDays(input, minimum) {
  const value = Number(input.value);
  if (!Number.isInteger(value) || value < minimum) {
    alert(`Enter a whole number of at least ${minimum} day${minimum === 1 ? "" : "s"}.`);
    input.focus();
    return null;
  }
  return value;
}

function compareYearGroups(a, b) {
  const numericA = Number.parseInt(a, 10);
  const numericB = Number.parseInt(b, 10);
  const bothNumeric = !Number.isNaN(numericA) && !Number.isNaN(numericB);

  if (bothNumeric && numericA !== numericB) {
    return numericA - numericB;
  }

  return String(a || "").localeCompare(String(b || ""));
}

function formatWeekday(dateText) {
  const match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return '';

  const [, year, month, day] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  return date.toLocaleDateString('en-AU', { weekday: 'long' });
}

function getDetentionLedger(student) {
  const detentions = Array.isArray(student.detentions) ? [...student.detentions] : [];
  const activeDetention = student.activeDetention;

  if (activeDetention?.status === "open") {
    const sourceLateDate = activeDetention.createdFromLateDate || activeDetention.scheduledForDate || "";
    const detentionId = sourceLateDate || activeDetention.scheduledForDate || "legacy_active_detention";
    if (!detentions.some(detention => detention?.detentionId === detentionId)) {
      detentions.push({
        detentionId,
        sourceLateDate,
        originalScheduledForDate: activeDetention.scheduledForDate || sourceLateDate,
        createdAt: activeDetention.createdAt || "",
        createdBy: "legacy_active_detention_migration",
        sourceContext: activeDetention.sourceContext || "legacy_active_detention"
      });
    }
  }

  return detentions
    .filter(detention => detention && detention.originalScheduledForDate)
    .sort((a, b) => String(a.originalScheduledForDate).localeCompare(String(b.originalScheduledForDate)));
}

function getDetentionServedEvents(student) {
  const explicitEvents = Array.isArray(student.detentionServedEvents)
    ? [...student.detentionServedEvents]
        .filter(event => event?.servedDate)
        .sort((a, b) => String(a.servedDate).localeCompare(String(b.servedDate)) || String(a.markedAt || "").localeCompare(String(b.markedAt || "")))
    : [];

  if (explicitEvents.length > 0) {
    return explicitEvents;
  }

  const history = Array.isArray(student.detentionHistory) ? student.detentionHistory : [];
  return history
    .filter(entry => entry?.outcome === "served" && (entry.date || entry.scheduledForDate))
    .map(entry => ({
      servedDate: entry.date || entry.scheduledForDate,
      markedAt: entry.date || entry.scheduledForDate,
      markedBy: "legacy_detention_history",
      source: "legacy_detention_history"
    }))
    .sort((a, b) => String(a.servedDate).localeCompare(String(b.servedDate)));
}

function buildDetentionStatus(detentions, servedEvents) {
  const latestServedDate = getLatestServedDate(servedEvents);
  const outstanding = getOutstandingDetentions(detentions, latestServedDate);
  const dates = outstanding
    .map(detention => detention.originalScheduledForDate)
    .filter(Boolean)
    .sort();

  return {
    hasOpenDetention: outstanding.length > 0,
    outstandingCount: outstanding.length,
    latestServedDate,
    oldestOutstandingDetentionDate: dates[0] || null,
    newestOutstandingDetentionDate: dates[dates.length - 1] || null
  };
}

function getOutstandingDetentions(detentions, latestServedDate) {
  return detentions.filter(detention => {
    const scheduledDate = detention?.originalScheduledForDate;
    if (!scheduledDate) return false;
    return !latestServedDate || scheduledDate > latestServedDate;
  });
}

function getLatestServedDate(servedEvents) {
  const dates = servedEvents
    .map(event => event?.servedDate)
    .filter(Boolean)
    .sort();
  return dates[dates.length - 1] || null;
}

function countSchoolDaysElapsed(startDateText, endDateText) {
  const startDate = parseDate(startDateText);
  const endDate = parseDate(endDateText);
  if (!startDate || !endDate || startDate >= endDate) {
    return 0;
  }

  let count = 0;
  const cursor = new Date(startDate);
  cursor.setDate(cursor.getDate() + 1);
  while (cursor <= endDate) {
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) {
      count += 1;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return count;
}

function parseDate(dateText) {
  const match = String(dateText || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const [, year, month, day] = match;
  return new Date(Number(year), Number(month) - 1, Number(day));
}

function getYearGroup(rollClass) {
  const match = String(rollClass).match(/\d+/);
  return match ? match[0] : '';
}

function resolveYearGroup(student) {
  const explicitYear = normalizeYearGroupValue(student.yearGroup);
  if (explicitYear) return explicitYear;

  const lateArrivalYear = Array.isArray(student.lateArrivals)
    ? student.lateArrivals.map(entry => normalizeYearGroupValue(entry.yearGroup)).find(Boolean)
    : '';
  if (lateArrivalYear) return lateArrivalYear;

  return getYearGroup(student.rollClass || '');
}

function normalizeYearGroupValue(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  if (text.endsWith('.0')) {
    return text.slice(0, -2);
  }

  const digits = text.match(/\d+/);
  return digits ? digits[0] : text;
}
