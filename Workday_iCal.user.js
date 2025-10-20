// ==UserScript==
// @name        WPI Workday Calendar
// @namespace   Violentmonkey Scripts
// @match       *://wd5.myworkday.com/wpi/*
// @grant       none
// @require     https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @author      Hello There
// @description Generates ICS calendars from Workday schedule
// ==/UserScript==

const originalSend = XMLHttpRequest.prototype.send;
const BASE_URL = "https://wd5.myworkday.com";
const REGEX_XLSX = /^https:\/\/wd5\.myworkday\.com\/wpi\/export\/c\d*\.xlsx\?wml-xpath-filter=.*Current_Registrations_for_Student_-_WPI.*Report_Row&page-title=Registered\+Classes&clientRequestID=.*/;

// ---------------------------
// Utilities
// ---------------------------
async function xlsxToXML(url) {
  const resp = await fetch(url);
  const buf = await resp.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const sheetXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
  return new DOMParser().parseFromString(sheetXml, "application/xml");
}

function xmlToArray(xmlDom) {
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const rows = Array.from(xmlDom.getElementsByTagNameNS(NS, "row"));
  return rows.map(row => {
    const cells = Array.from(row.getElementsByTagNameNS(NS, "c"));
    return cells.map(cell => {
      if (cell.getAttribute("t") === "inlineStr") {
        const tNode = cell.querySelector("is > t");
        return tNode ? tNode.textContent : "";
      } else {
        const vNode = cell.getElementsByTagNameNS(NS, "v")[0];
        return vNode ? vNode.textContent : "";
      }
    });
  });
}

// ---------------------------
// Classes
// ---------------------------
class Course {
  constructor(term, def, format, times, location, instructor, mode_of_delivery) {
    this.term = term;
    this.def = def;
    this.format = format;
    this.times = times;
    this.location = location;
    this.instructor = instructor;
    this.mode_of_delivery = mode_of_delivery;
  }

  toICal(termStartDate, termEndDate) {
    const dayMap = { M: "MO", T: "TU", W: "WE", R: "TH", F: "FR", S: "SA", U: "SU" };
    const [daysPart, timePart] = this.times.split("|").map(s => s.trim());
    const days = daysPart.split("-").map(d => dayMap[d] || "").join(",");
    const [startTimeStr, endTimeStr] = timePart.split("-").map(s => s.trim());

    const parseTime = (timeStr, firstDay) => {
      const [time, meridiem] = timeStr.split(" ");
      let [hours, minutes] = time.split(":").map(Number);
      if (meridiem?.toUpperCase() === "PM" && hours !== 12) hours += 12;
      if (meridiem?.toUpperCase() === "AM" && hours === 12) hours = 0;
      const dt = new Date(firstDay.getFullYear(), firstDay.getMonth(), firstDay.getDate(), hours, minutes);
      const pad = n => n.toString().padStart(2, "0");
      return `${dt.getUTCFullYear()}${pad(dt.getUTCMonth() + 1)}${pad(dt.getUTCDate())}T${pad(dt.getUTCHours())}${pad(dt.getUTCMinutes())}00Z`;
    };

    const termStart = new Date(termStartDate);
    const weekdayMap = { SU:0, MO:1, TU:2, WE:3, TH:4, FR:5, SA:6 };
    const dayNums = days.split(",").map(d => weekdayMap[d]);
    let firstCourseDay = new Date(termStart);
    while (!dayNums.includes(firstCourseDay.getDay())) {
      firstCourseDay.setDate(firstCourseDay.getDate() + 1);
    }

    const dtStart = parseTime(startTimeStr, firstCourseDay);
    const dtEnd = parseTime(endTimeStr, firstCourseDay);

    return [
      "BEGIN:VEVENT",
      `SUMMARY:${this.def}`,
      `DTSTART:${dtStart}`,
      `DTEND:${dtEnd}`,
      `LOCATION:${this.location}`,
      `DESCRIPTION:${this.format} | ${this.mode_of_delivery} | Instructor: ${this.instructor}`,
      `RRULE:FREQ=WEEKLY;BYDAY=${days}`,
      "END:VEVENT"
    ].join("\r\n");
  }
}

class Schedule {
  constructor(term, startDate, endDate, events = []) {
    this.term = term;
    this.startDate = startDate;
    this.endDate = endDate;
    this.events = events;
  }

  toICal() {
    const header = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//hello.there.reader.com/acme//NONSGML v1.0//EN`;
    const eventsICal = this.events.map(e => e.toICal(this.startDate, this.endDate)).join("\r\n");
    const tail = `\r\nEND:VCALENDAR`;
    return `${header}\r\n${eventsICal}${tail}`;
  }

  addDownloadButton(parent) {
    const button = document.createElement("button");
    button.textContent = "Download ICS";
    button.style.marginLeft = "10px";
    button.onclick = () => {
      const blob = new Blob([this.toICal()], { type: "text/calendar;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${this.term.replace(/\s+/g, "_")}.ics`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      parent.remove();
    };
    parent.appendChild(button);
  }
}

// ---------------------------
// Main logic
// ---------------------------
async function callback(event) {
  if (REGEX_XLSX.test(this.responseURL)) {
    this.responseJson = JSON.parse(this.responseText);
    const fullUrl = BASE_URL + this.responseJson.docReadyUri;

    const classListXML = await xlsxToXML(fullUrl);
    const classArray = xmlToArray(classListXML);
    classArray.shift(); // remove header

    const courses = classArray.map(row => new Course(...row));

    // Create banner UI
    const banner = document.createElement("div");
    banner.style.position = "fixed";
    banner.style.top = "0";
    banner.style.left = "0";
    banner.style.right = "0";
    banner.style.backgroundColor = "#fffae6";
    banner.style.borderBottom = "1px solid #ccc";
    banner.style.padding = "10px";
    banner.style.zIndex = "9999";
    banner.style.textAlign = "center";
    banner.style.fontFamily = "sans-serif";

    banner.innerHTML = `
      <strong>Generate WPI Term Calendar:</strong>
      <label style="margin-left:10px;">Term Letter:
        <input type="text" id="termLetter" maxlength="1" style="width:40px;text-transform:uppercase;">
      </label>
      <label style="margin-left:10px;">Start:
        <input type="date" id="termStart">
      </label>
      <label style="margin-left:10px;">End:
        <input type="date" id="termEnd">
      </label>
      <button id="generateSchedule" style="margin-left:10px;">Generate Schedule</button>
      <button id="closeBanner" style="margin-left:10px;">×</button>
    `;

    document.body.appendChild(banner);

    banner.querySelector("#closeBanner").onclick = () => banner.remove();

    banner.querySelector("#generateSchedule").onclick = () => {
      const termLetter = banner.querySelector("#termLetter").value.trim().toUpperCase();
      const startDate = banner.querySelector("#termStart").value;
      const endDate = banner.querySelector("#termEnd").value;

      if (!termLetter || !startDate || !endDate) {
        alert("Please fill all fields before generating the schedule.");
        return;
      }

      const filteredCourses = courses.filter(c => c.term.toUpperCase().includes(`${termLetter} TERM`));
      if (filteredCourses.length === 0) {
        alert(`No classes found for ${termLetter} Term.`);
        return;
      }

      const schedule = new Schedule(`WPI ${termLetter} Term`, startDate, endDate, filteredCourses);
      banner.innerHTML = `<strong>${termLetter} Term schedule ready!</strong>`;
      schedule.addDownloadButton(banner);
    };
  }
}

XMLHttpRequest.prototype.send = function(...args) {
  this.addEventListener("load", callback);
  return originalSend.apply(this, args);
};
