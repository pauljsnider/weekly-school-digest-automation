/**
 * Weekly school digest to Google Doc - focused output with Programs section
 * Rules:
 *  - Smore: keep full extracted newsletter text, no trimming
 *  - Calendars: include events with title, start - end, location, and description if present
 *  - Max: drop items with "daily summary" in subject or body; if kept, include cleaned email body
 *  - Fallback: if no Smore for any child, include cleaned email body
 *  - Programs (stmichaelcp.*): single combined section, no duplication under kids
 * Sources included for last 7 days:
 *  - bluevalleyk12.org
 *  - cremedelacreme.com
 *  - online.procaresoftware.com
 *  - stmichaelcp.ccsend.com
 *  - stmichaelcp.org
 */

const LOOKBACK_DAYS = 7;
const TZ = "America/Chicago";
const TARGET_FOLDER_NAME = "School Notes";
const MEANINGFUL_TEXT_THRESHOLD = 50;

const MAX_BODY_CHARS = 1500;
const INCLUDE_EVENT_DESCRIPTION = true;

// ICS feeds
const ICS_FEEDS = [
  {
    name: "Will Soccer",
    url: "YOUR_ICS_URL_HERE",
    child: "Will"
  },
  {
    name: "Will Baseball",
    url: "YOUR_ICS_URL_HERE",
    child: "Will"
  },
  {
    name: "TeamSnap Events - Madison + Max",
    url: "YOUR_ICS_URL_HERE",
    child: "Madison/Max"
  }
];

// Teacher mapping
const TEACHER_MAP = [
  { email: "TEACHER1@bluevalleyk12.org", child: "Madison" },
  { email: "TEACHER2@bluevalleyk12.org", child: "Will" },
  { email: "TEACHER3@bluevalleyk12.org", child: "Schoolwide" }
];

// Domain fallbacks
// stmichaelcp.* routes to PROGRAMS_MW section
const DOMAIN_RULES = [
  { domain: "bluevalleyk12.org", childFallback: "Schoolwide" },
  { domain: "cremedelacreme.com", childFallback: "Max" },
  { domain: "online.procaresoftware.com", childFallback: "Max" },
  { domain: "stmichaelcp.ccsend.com", childFallback: "PROGRAMS_MW" },
  { domain: "stmichaelcp.org", childFallback: "PROGRAMS_MW" }
];

const GEMINI_API_KEY = "YOUR_API_KEY_HERE";
const EMAIL_RECIPIENT = "YOUR_EMAIL@gmail.com";
const CC_RECIPIENT = "CC_EMAIL@example.com";

function runWeeklySchoolDigestToDoc() {
  const now = new Date();
  const subjectDate = Utilities.formatDate(now, TZ, "MMM d, yyyy");
  const timestamp = Utilities.formatDate(now, TZ, "HHmm");
  const title = "Weekly School Digest - " + subjectDate + " (" + timestamp + ")";

  const query = buildQuery(LOOKBACK_DAYS);
  const threads = GmailApp.search(query, 0, 200);

  const items = [];
  threads.forEach(function(th) {
    th.getMessages().forEach(function(msg) {
      const from = msg.getFrom() || "";
      const child = resolveChild(from);
      if (!child) return;

      const subj = msg.getSubject() || "";
      const dateObj = msg.getDate();
      const dateStr = Utilities.formatDate(dateObj, TZ, "EEE, MMM d h:mm a");

      const html = msg.getBody() || "";
      const plain = msg.getPlainBody() || "";

      if (child === "Max") {
        const combined = (subj + "\n" + plain).toLowerCase();
        if (combined.indexOf("daily summary") !== -1) return;
      }

      const cleanBody = cleanAndCapBody(plain ? plain : stripHtml(html), MAX_BODY_CHARS);

      const smoreLinks = extractSmoreLinks(html);
      const smoreDetails = smoreLinks.map(function(url) { return fetchSmoreFullText(url); });
      const hasSmore = smoreDetails.some(function(sd) { return sd.status === "OK" && sd.fullText; });

      items.push({
        child: child,   // "Madison" | "Will" | "Max" | "Schoolwide" | "PROGRAMS_MW"
        from: from,
        subject: subj,
        dateStr: dateStr,
        date: dateObj,
        smoreDetails: smoreDetails,
        hasSmore: hasSmore,
        cleanBody: cleanBody
      });
    });
  });

  const events = getUpcomingEventsFromIcsFeeds(ICS_FEEDS, 14, INCLUDE_EVENT_DESCRIPTION);

  let docText = buildDoc(items, events);

  // --- AI Summary Generation ---
  let aiSummary = "";
  try {
    aiSummary = generateWeeklySummaryAI(docText);
    if (aiSummary) {
      // Prepend AI summary to the document text
      docText = "# AI Weekly Summary\n\n" + aiSummary + "\n\n---\n\n" + docText;
      Logger.log("AI Summary generated and added to doc.");
    }
  } catch (e) {
    Logger.log("Error generating AI summary: " + e.toString());
  }

  // --- Append Source List to Doc ---
  docText += "\n\n---\n\n## Source Emails\n\n";
  docText += "| Subject | Sender |\n| :--- | :--- |\n";
  items.forEach(function(it) {
    docText += "| " + (it.subject || "(no subject)") + " | " + (it.from || "(unknown)") + " |\n";
  });

  const folder = ensureFolder(TARGET_FOLDER_NAME);
  const doc = DocumentApp.create(title);
  const docId = doc.getId();
  const docUrl = "https://docs.google.com/document/d/" + docId + "/edit";

  safeWriteDocById(docId, docText);
  safeMoveFileToFolder(docId, folder);

  Logger.log("Created document: " + docUrl);

  // --- Send Email ---
  if (aiSummary) {
    try {
      let emailBody = aiSummary;
      
      // Add Doc Link
      emailBody += "<br><br><hr><br>";
      emailBody += "<p><strong>Full Document:</strong> <a href='" + docUrl + "'>" + title + "</a></p>";

      // Add Source List Table
      emailBody += "<h3>Source Emails</h3>";
      emailBody += "<table border='1' style='border-collapse: collapse; width: 100%;'>";
      emailBody += "<tr><th style='padding: 8px; text-align: left;'>Subject</th><th style='padding: 8px; text-align: left;'>Sender</th></tr>";
      items.forEach(function(it) {
        emailBody += "<tr><td style='padding: 8px;'>" + (it.subject || "(no subject)") + "</td><td style='padding: 8px;'>" + (it.from || "(unknown)") + "</td></tr>";
      });
      emailBody += "</table>";

      sendEmail("Weekly Child Summary - " + subjectDate, emailBody);
      Logger.log("AI Summary sent to email.");
    } catch (e) {
      Logger.log("Error sending email: " + e.toString());
    }
  }
}

function generateWeeklySummaryAI(text) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY;
  
  const prompt = `Read the entire file below, and create a weekly summary by child (Madison, Will, Max). 
  
  For each child, include:
  - Previous week learning/activities
  - Upcoming week focus/events
  - Other important info
  
  End with a calendar-style table of all events with date, time, child, and activity.

  Output the result in clean HTML format suitable for an email. 
  - Use <h3> for child names.
  - Use <ul> and <li> for lists.
  - Use <strong> for bold text.
  - Format the calendar as a standard HTML <table> with headers, rows, and borders.
  
  [DATA START]
  ${text}
  [DATA END]`;

  const payload = {
    "contents": [{
      "parts": [{"text": prompt}]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode >= 200 && responseCode < 300) {
    const data = JSON.parse(responseText);
    if (data.candidates && data.candidates.length > 0 && data.candidates[0].content && data.candidates[0].content.parts.length > 0) {
      return data.candidates[0].content.parts[0].text;
    }
  }
  
  Logger.log("Gemini API Error: " + responseText);
  return null;
}

function sendEmail(subject, htmlBody) {
  MailApp.sendEmail({
    to: EMAIL_RECIPIENT,
    cc: CC_RECIPIENT,
    subject: subject,
    htmlBody: htmlBody
  });
}

/* ===================== Build output ===================== */

function buildDoc(items, events) {
  var out = "# Weekly School Digest\n";
  out += "Generated: " + Utilities.formatDate(new Date(), TZ, "EEE, MMM d yyyy h:mm a") + "\n\n";

  // Events
  out += "## Events\n\n";
  if (!events || events.length === 0) {
    out += "No upcoming events found.\n\n";
  } else {
    events.forEach(function(ev) {
      const when = Utilities.formatDate(ev.when, TZ, "EEE, MMM d h:mm a");
      const end = ev.end ? Utilities.formatDate(ev.end, TZ, "h:mm a") : "";
      const timeText = end ? (when + " - " + end) : when;
      out += "- " + (ev.title || "(Untitled)") + "  (" + ev.source + ")\n";
      out += "  Date: " + timeText + "\n";
      if (ev.location) out += "  Location: " + ev.location + "\n";
      if (ev.description) out += "  Notes: " + collapseNoise(ev.description) + "\n";
      out += "\n";
    });
  }

  // Programs for Madison and Will
  const programs = (items || []).filter(function(i) { return i.child === "PROGRAMS_MW"; });
  if (programs.length > 0) {
    out += "## Programs for Madison and Will\n\n";
    const progWithSmore = [];
    const progWithoutSmore = [];
    programs.forEach(function(it) {
      const okSmore = it.smoreDetails.filter(function(sd) { return sd.status === "OK" && sd.fullText; });
      if (okSmore.length > 0) {
        okSmore.forEach(function(sd) {
          progWithSmore.push({
            subject: it.subject,
            dateStr: it.dateStr,
            url: sd.url,
            text: sd.fullText
          });
        });
      } else {
        progWithoutSmore.push({
          subject: it.subject,
          dateStr: it.dateStr,
          body: it.cleanBody
        });
      }
    });

    progWithSmore.forEach(function(s) {
      out += "### " + s.subject + "\n";
      out += "Date: " + s.dateStr + "\n";
      if (s.url) out += "Link: " + s.url + "\n";
      out += "\n";
      out += s.text + "\n\n";
      out += "---\n\n";
    });

    if (progWithoutSmore.length > 0) {
      out += "### Program emails without newsletter\n\n";
      progWithoutSmore.forEach(function(m) {
        out += "- " + m.subject + "  [" + m.dateStr + "]\n";
        if (m.body) {
          out += m.body + "\n\n";
        } else {
          out += "(no body)\n\n";
        }
      });
    }

    out += "\n";
  }

  // Per child sections excluding programs
  const byChild = groupBy(items.filter(function(i) { return i.child !== "PROGRAMS_MW"; }), function(x) { return x.child; });
  ["Madison", "Will", "Max", "Schoolwide"].forEach(function(child) {
    const list = byChild[child] || [];
    const withSmore = [];
    const withoutSmore = [];

    list.forEach(function(it) {
      const okSmore = it.smoreDetails.filter(function(sd) { return sd.status === "OK" && sd.fullText; });
      if (okSmore.length > 0) {
        okSmore.forEach(function(sd) {
          withSmore.push({
            subject: it.subject,
            dateStr: it.dateStr,
            url: sd.url,
            text: sd.fullText
          });
        });
      } else {
        withoutSmore.push({
          subject: it.subject,
          dateStr: it.dateStr,
          body: it.cleanBody
        });
      }
    });

    if (withSmore.length === 0 && withoutSmore.length === 0) return;

    out += "## " + child + "\n\n";

    withSmore.forEach(function(s) {
      out += "### " + s.subject + "\n";
      out += "Date: " + s.dateStr + "\n";
      if (s.url) out += "Link: " + s.url + "\n";
      out += "\n";
      out += s.text + "\n\n";
      out += "---\n\n";
    });

    if (withoutSmore.length > 0) {
      out += "### Email items without newsletter\n\n";
      withoutSmore.forEach(function(m) {
        out += "- " + m.subject + "  [" + m.dateStr + "]\n";
        if (m.body) {
          out += m.body + "\n\n";
        } else {
          out += "(no body)\n\n";
        }
      });
    }
  });

  return out;
}

/* ===================== Gmail helpers ===================== */

function buildQuery(days) {
  const fromClause = "("
    + "from:bluevalleyk12.org OR "
    + "from:cremedelacreme.com OR "
    + "from:online.procaresoftware.com OR "
    + "from:stmichaelcp.ccsend.com OR "
    + "from:stmichaelcp.org"
    + ")";
  return fromClause + " newer_than:" + days + "d";
}

function resolveChild(fromHeader) {
  const lower = (fromHeader || "").toLowerCase();

  for (var i = 0; i < TEACHER_MAP.length; i++) {
    var t = TEACHER_MAP[i];
    if (lower.indexOf(t.email.toLowerCase()) !== -1) return t.child;
  }

  for (var j = 0; j < DOMAIN_RULES.length; j++) {
    var rule = DOMAIN_RULES[j];
    if (lower.indexOf("@" + rule.domain) !== -1) return rule.childFallback;
  }

  return "";
}

function extractSmoreLinks(html) {
  if (!html) return [];
  var links = [];
  var re = /https?:\/\/secure\.smore\.com\/n\/[A-Za-z0-9_-]+/gi;
  var m;
  while ((m = re.exec(html)) !== null) links.push(m[0]);
  var uniq = {};
  var out = [];
  for (var i = 0; i < links.length; i++) {
    if (!uniq[links[i]]) {
      uniq[links[i]] = true;
      out.push(links[i]);
    }
  }
  return out;
}

/* ===================== Smore helpers ===================== */

function fetchSmoreFullText(url) {
  try {
    const res = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { "User-Agent": ua(), "Accept-Language": "en-US,en;q=0.9" }
    });
    const code = res.getResponseCode();
    const html = res.getContentText() || "";
    if (code >= 200 && code < 300 && html) {
      const jsData = extractEmbeddedSmoreData(html);
      if (jsData) {
        const text = smoreJsonToText(jsData);
        if (text && text.length > MEANINGFUL_TEXT_THRESHOLD) return { url: url, status: "OK", fullText: text };
      }
      const text2 = stripHtml(html);
      if (text2 && text2.length > MEANINGFUL_TEXT_THRESHOLD) return { url: url, status: "OK", fullText: text2 };
      return { url: url, status: "MANUAL", reason: "Likely JS rendered or blocked" };
    }
    return { url: url, status: "MANUAL", reason: "HTTP " + code };
  } catch (e) {
    return { url: url, status: "MANUAL", reason: "Error " + e };
  }
}

function extractEmbeddedSmoreData(html) {
  try {
    const re = /"newsletter"\s*:\s*{[\s\S]*?"js_content"\s*:\s*"({[\s\S]+?})"/i;
    const match = html.match(re);
    if (!match) return null;
    const jsonStr = match[1].replace(/\\"/g, '"').replace(/\\\\/g, '\\');
    return JSON.parse(jsonStr);
  } catch (e) {
    return null;
  }
}

function smoreJsonToText(json) {
  const bits = [];
  function push(s) { if (s && typeof s === "string" && s.trim()) bits.push(s.trim()); }

  if (json && Object.prototype.toString.call(json.blocks) === "[object Array]") {
    json.blocks.forEach(function(block) {
      if (block.title) push(block.title);
      if (block.content && Object.prototype.toString.call(block.content) === "[object Array]") {
        block.content.forEach(function(item) {
          const t = extractTextFromContentItem(item);
          if (t) push(t);
        });
      }
    });
  }
  if (json.title) push(json.title);
  if (json.subtitle) push(json.subtitle);
  if (typeof json.html === "string") push(stripHtml(json.html));

  return bits.join("\n\n").trim();
}

function extractTextFromContentItem(item) {
  if (!item) return "";
  if (typeof item === "string") return item;
  if (typeof item.text === "string") return item.text;
  if (Object.prototype.toString.call(item.c) === "[object Array]") {
    return item.c.map(extractTextFromContentItem).join(" ");
  }
  return "";
}

function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<\/?(?:br|p|div|li|ul|ol|h[1-6])[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function collapseNoise(s) {
  if (!s) return "";
  return s
    .replace(/this message.*confidential.*intended recipient.*delete/gis, "")
    .replace(/unsubscribe.*privacy.*terms/gis, "")
    .replace(/sent from.*iphone|android|mobile/gi, "")
    .replace(/do not reply.*this.*automated/gis, "")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function cleanAndCapBody(raw, cap) {
  if (!raw) return "";
  var s = collapseNoise(raw);
  if (cap && s.length > cap) {
    var candidate = s.slice(0, cap);
    var lastPeriod = candidate.lastIndexOf(".");
    if (lastPeriod > cap * 0.6) {
      s = candidate.slice(0, lastPeriod + 1) + "\n\n[truncated]";
    } else {
      s = candidate + "\n\n[truncated]";
    }
  }
  return s;
}

/* ===================== ICS helpers ===================== */

function getUpcomingEventsFromIcsFeeds(feeds, takeN, includeDescription) {
  const now = new Date();
  const events = [];

  feeds.forEach(function(f) {
    try {
      const url = f.url.indexOf("webcal://") === 0 ? "https://" + f.url.slice("webcal://".length) : f.url;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) return;

      const ics = resp.getContentText();
      const parsed = parseIcs(ics, f.name, includeDescription);
      parsed.forEach(function(ev) {
        if (ev.when && ev.when >= now) events.push(ev);
      });
    } catch (e) {}
  });

  events.sort(function(a, b) { return a.when - b.when; });
  return events.slice(0, takeN);
}

function parseIcs(icsText, sourceName, includeDescription) {
  if (!icsText) return [];
  const lines = unfoldIcsLines(icsText.split(/\r?\n/));
  const events = [];
  let current = null;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    if (line === "BEGIN:VEVENT") { current = {}; continue; }
    if (line === "END:VEVENT") {
      if (current && current.DTSTART) {
        const start = parseIcsDate(current.DTSTART);
        const end = current.DTEND ? parseIcsDate(current.DTEND) : null;
        const title = current.SUMMARY || "";
        const location = current.LOCATION || "";
        const description = includeDescription ? (current.DESCRIPTION || "") : "";
        if (start) {
          events.push({
            when: start,
            end: end,
            title: title,
            location: location,
            description: decodeIcsText(description),
            source: sourceName || ""
          });
        }
      }
      current = null;
      continue;
    }
    if (!current) continue;

    const idx = line.indexOf(":");
    if (idx === -1) continue;

    const rawKey = line.slice(0, idx);
    const value = line.slice(idx + 1);
    const key = rawKey.split(";")[0].toUpperCase();
    current[key] = value;
  }

  return events;
}

function decodeIcsText(v) {
  if (!v) return "";
  return v.replace(/\\n/g, "\n").replace(/\\,/g, ",").replace(/\\;/g, ";");
}

function unfoldIcsLines(arr) {
  const out = [];
  for (var i = 0; i < arr.length; i++) {
    var line = arr[i];
    if (line === undefined || line === null) continue;
    if (line.indexOf(" ") === 0 || line.indexOf("\t") === 0) {
      if (out.length) out[out.length - 1] += line.slice(1);
    } else {
      out.push(line);
    }
  }
  return out;
}

function parseIcsDate(val) {
  if (!val) return null;
  const parts = val.split(":");
  const raw = parts.length > 1 ? parts.slice(1).join(":") : parts[0];

  if (/^\d{8}$/.test(raw)) {
    const y = +raw.slice(0, 4);
    const m = +raw.slice(4, 6) - 1;
    const d = +raw.slice(6, 8);
    return new Date(y, m, d);
  }

  var z = raw.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/);
  if (z) {
    const y2 = +z[1];
    const m2 = +z[2] - 1;
    const d2 = +z[3];
    const hh = +z[4];
    const mm = +z[5];
    const ss = +z[6];
    return new Date(Date.UTC(y2, m2, d2, hh, mm, ss));
  }

  var l = raw.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})$/);
  if (l) {
    const y3 = +l[1];
    const m3 = +l[2] - 1;
    const d3 = +l[3];
    const hh2 = +l[4];
    const mm2 = +l[5];
    const ss2 = +l[6];
    return new Date(y3, m3, d3, hh2, mm2, ss2);
  }

  return null;
}

/* ===================== Drive and utility ===================== */

function ensureFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function safeWriteDocById(docId, text) {
  const MAX_CHUNK = 50000;
  const MAX_RETRIES = 3;

  var doc = null;
  var lastErr = null;
  for (var i = 0; i < MAX_RETRIES; i++) {
    try {
      if (i > 0) Utilities.sleep(500 * i);
      doc = DocumentApp.openById(docId);
      break;
    } catch (e) {
      lastErr = e;
    }
  }
  if (!doc) throw new Error("Could not open doc for writing: " + lastErr);

  const body = doc.getBody();
  body.clear();

  if (!text) {
    body.appendParagraph("(empty)");
    doc.saveAndClose();
    return;
  }

  for (var p = 0; p < text.length; p += MAX_CHUNK) {
    const chunk = text.substring(p, Math.min(p + MAX_CHUNK, text.length));
    body.appendParagraph(chunk);
  }
  doc.saveAndClose();
}

function safeMoveFileToFolder(fileId, folder) {
  try {
    const file = DriveApp.getFileById(fileId);
    folder.addFile(file);
    try {
      DriveApp.getRootFolder().removeFile(file);
    } catch (e) {
      // ignore
    }
  } catch (e2) {
    Logger.log("Warning: move failure: " + e2);
  }
}

function groupBy(arr, keyFn) {
  var map = {};
  for (var i = 0; i < arr.length; i++) {
    var item = arr[i];
    var k = keyFn(item);
    if (!map[k]) map[k] = [];
    map[k].push(item);
  }
  return map;
}

function ua() {
  return "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123 Safari/537.36";
}
