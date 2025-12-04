# Weekly School Digest Automation (Google Apps Script + Gemini AI)

This project automates the collection of school updates from email, newsletters, and calendars, compiles them into a Google Doc, and generates an AI-powered summary using Google's Gemini API—all within a single Google Apps Script.

---

## 🔄 End-to-End Flow Overview

**Complete Automated Workflow:**

1. **📧 Email Collection**
   - Script runs automatically (e.g., every Sunday).
   - Collects emails from the last 7 days from configured school/childcare domains.
   - Extracts full newsletter content from Smore links.
   - Pulls upcoming calendar events from ICS feeds.

2. **📄 Document Generation**
   - Generates a structured Google Doc in your Drive.
   - Sections include: AI Summary, Events, Programs, and Individual Child Updates.
   - Appends a "Source Emails" table at the bottom for reference.

3. **🤖 AI Summarization (Gemini)**
   - The script sends the digest content to Google's Gemini API.
   - Generates a concise weekly summary grouped by child.
   - Creates a calendar-style table of events.
   - Prepends this summary to the top of the Google Doc.

4. **📬 Email Delivery**
   - Sends an HTML-formatted email to parents.
   - Email includes:
     - The AI-generated summary.
     - A link to the full Google Doc.
     - A table of source emails (Subject & Sender).

**Result:** You get a comprehensive detailed digest AND a clean AI summary delivered to your inbox every week, completely automated.

---

## ⚙️ Setup Guide

### Step 1: Create the Script

1. **Open Google Apps Script**
   - Go to [https://script.google.com/](https://script.google.com/)
   - Create a new project.
   - Paste in the provided `parse_email.gs` script.

### Step 2: Configure the Script

Update the configuration constants at the top of the script:

* **`GEMINI_API_KEY`**: Your Google Cloud API key for the Generative Language API.
* **`EMAIL_RECIPIENT`**: The primary email address to receive the digest.
* **`CC_RECIPIENT`**: (Optional) Additional email address to CC.
* **`DOMAIN_RULES`**: Add your school/organization domains.
* **`TEACHER_MAP`**: Map teacher emails to children.
* **`ICS_FEEDS`**: Add your calendar feed URLs.
* **`TARGET_FOLDER_NAME`**: Set your desired Google Drive folder.

### Step 3: Get a Gemini API Key

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (or use an existing one).
3. Enable the **Generative Language API**.
4. Create an **API Key**.
5. **Security Tip**: Restrict the key to only the "Generative Language API" to prevent unauthorized use.

### Step 4: Authorize Services

1. Run the `runWeeklySchoolDigestToDoc` function manually once.
2. Accept permissions for:
   - Gmail (read-only)
   - Google Drive (create/edit files)
   - External Services (fetch URL for calendars/AI)
   - Send Email

### Step 5: Schedule the Script

1. In Apps Script, go to **Triggers** (clock icon).
2. Add a new trigger:
   - **Function**: `runWeeklySchoolDigestToDoc`
   - **Event source**: Time-driven
   - **Type**: Weekly
   - **Day**: Sunday
   - **Time**: 10:00 AM (or your preferred time)

---

## 🧩 Features

### 1. Smore Newsletter Extraction
* Detects Smore links in emails.
* Extracts the *full text content* directly from the hidden JSON data inside Smore pages.
* Falls back to email body if Smore extraction fails.

### 2. Calendar Integration
* Reads upcoming events from ICS feeds (GameChanger, TeamSnap, etc.).
* Sorts events by date and includes location/notes.

### 3. Intelligent Filtering
* Removes duplicate/unnecessary content.
* Filters out "daily summary" reports for specific children to reduce noise.
* Truncates long email bodies if no newsletter is found.

### 4. AI Summary (Gemini 2.5 Flash)
* Reads the entire digest.
* Summarizes highlights by child:
  * Previous Week Learning/Activities
  * Upcoming Week Focus/Events
* Generates a clean HTML calendar table.

---

## 🔧 Configuration Examples

### Domain Rules
```javascript
const DOMAIN_RULES = [
  { domain: "yourschool.org", childFallback: "Schoolwide" },
  { domain: "childcare.com", childFallback: "Child3" },
  { domain: "programs.yourschool.com", childFallback: "PROGRAMS_MW" }
];
```

### Teacher Mapping
```javascript
const TEACHER_MAP = [
  { email: "teacher1@yourschool.org", child: "Child1" },
  { email: "teacher2@yourschool.org", child: "Child2" }
];
```

### Calendar Feeds
```javascript
const ICS_FEEDS = [
  {
    name: "Child1 Soccer",
    url: "webcal://api.team-manager.gc.com/...",
    child: "Child1"
  }
];
```

---

## 🔒 Privacy & Security

* **Local Processing**: The script runs entirely within your Google Account.
* **API Key**: Your Gemini API key is stored in the script. Ensure you do not share the script file publicly with the key inside.
* **Scrubbed Version**: A `parse_email.gs` file is provided with placeholders for safe sharing.
* **GitIgnore**: `parse_email_w_ai.gs` (containing real keys) is ignored by git.

---

## 📞 Support

For issues:
1. Check the **Executions** tab in Apps Script for logs.
2. Verify your API key is active and has quota.
3. Ensure the script has permission to send emails (`MailApp`).
