# APF Resource Person Dashboard — App Structure

> **Single-page web app** for Academic Resource Persons (APF) to track school visits, teacher training, classroom observations, and professional growth. Built entirely in **vanilla HTML + CSS + JavaScript** with no build tools.

---

## File Overview

| File | Size | Purpose |
|---|---|---|
| `index.html` | ~365 KB, ~6,100 lines | Full HTML structure — sidebar nav, 30 content sections, 16+ modals |
| `app.js` | ~1.22 MB, ~22,890 lines | All application logic — data layer, encryption, all feature modules |
| `styles.css` | ~497 KB, ~24,955 lines | All styling — design system, component styles, responsive rules |
| `excel-analytics.js` | ~187 KB | Standalone Excel analytics module (charts, pivot tables, data quality) |
| `package.json` | 50 B | Only dependency: `xlsx` (SheetJS) for Excel read/write |

---

## Architecture

```
┌──────────────────────────────────────────────────┐
│                   index.html                      │
│  ┌─────────┐  ┌──────────────────────────────┐   │
│  │ Sidebar  │  │     Content Sections          │   │
│  │ Nav      │  │  section-dashboard            │   │
│  │ Items    │  │  section-visits               │   │
│  │          │  │  section-training             │   │
│  │ data-    │  │  section-observations         │   │
│  │ section= │  │  section-{name}  ...          │   │
│  │ "visits" │  │                               │   │
│  └─────────┘  └──────────────────────────────┘   │
│               ┌──────────────────────────────┐   │
│               │     Modal Overlays            │   │
│               │  .modal-overlay#visitModal    │   │
│               │  .modal-overlay#{name}Modal   │   │
│               └──────────────────────────────┘   │
└──────────────────────────────────────────────────┘
        │                    │
        ▼                    ▼
┌──────────────┐    ┌──────────────┐
│   app.js     │    │  styles.css  │
│              │    │              │
│  DB (memory) │    │  CSS Vars    │
│  Encryption  │    │  Components  │
│  Navigation  │    │  Responsive  │
│  All modules │    │  Dark/Light  │
└──────────────┘    └──────────────┘
```

---

## Data Layer

### In-Memory Store (`DB` object — line 59)

All data lives in memory only. **Nothing is saved to localStorage/IndexedDB directly.** Persistence is handled via encrypted `.apf` files.

```javascript
const DB = {
    _store: {},           // All data lives here
    get(key),             // Returns deep clone of data array
    set(key, data),       // Deep clones and stores
    clear(),              // Wipes all data
    generateId()          // Returns unique ID: timestamp36 + random
};
```

### Data Keys (`ENCRYPTED_DATA_KEYS` — line 466)

These are ALL the data categories stored and encrypted:

| Key | Description |
|---|---|
| `visits` | School visit records |
| `trainings` | Teacher training sessions |
| `observations` | Classroom observations (imported from DMT Excel) |
| `resources` | Resource library items |
| `notes` | Quick notes |
| `ideas` | Idea tracker entries |
| `reflections` | Monthly reflection journal entries |
| `contacts` | Contact directory |
| `plannerTasks` | Weekly planner tasks |
| `goalTargets` | Monthly goal tracker entries |
| `followupStatus` | Follow-up action items |
| `worklog` | Monthly work log entries |
| `userProfile` | User profile (name, block, cluster, etc.) — stored as object, not array |
| `meetings` | Meeting tracker entries |
| `growthAssessments` | Professional growth self-assessments |
| `growthActionPlans` | Growth action plan items |
| `maraiTracking` | MARAI stage tracking per teacher |
| `schoolWork` | School-based work tracking entries |
| `visitPlanEntries` | Visit plan vs execution entries |
| `visitPlanDropdowns` | Saved dropdown options for visit plan |
| `feedbackReports` | Bug reports / feature requests |
| `teacherRecords` | Teacher database records |
| `schoolStudentRecords` | School student learning level records |
| `selfCapacityBuilding` | Self capacity building entries |
| `teachingPractices` | Subject-wise teaching practices (serial, group, description, classes, effect tracking) |

---

## Encryption & Storage

### Encryption Flow (AES-256-GCM)

```
User Password
     │
     ▼ PBKDF2 (100K iterations, SHA-256)
  AES Key
     │
     ▼ AES-256-GCM encrypt
  Encrypted .apf file (JSON → encrypted blob → base64)
```

| Module | Line | Purpose |
|---|---|---|
| `EncryptedFileStorage` | 115 | Core AES-256-GCM encrypt/decrypt functions |
| `EncryptedCache` | 173 | Stores encrypted blob in `localStorage` as browser cache |
| `FileLink` | 265 | File System Access API — direct read/write to a `.apf` file on disk |
| `GoogleDriveSync` | 468 | Auto-backup to Google Drive via Apps Script webhook |
| `PasswordManager` | 78 | SHA-256 password hashing, stored hash management |
| `SessionPersist` | 448 | Stores session password in `sessionStorage` for page refresh survival |

### Save Triggers
- **Auto-save**: Every 30 seconds when changes exist (`performAutoSave()` — line 842)
- **On change**: `markUnsavedChanges()` is called after every data mutation
- **Before unload**: Best-effort save on page close (line 20106)
- **File link**: If a file is linked via File System Access API, writes directly to it

---

## Navigation System

### Section Routing (`navigateTo(section)` — line 1469)

Navigation is DOM-based. Clicking a sidebar nav item calls `navigateTo('sectionName')` which:
1. Removes `.active` from all `.nav-item` elements
2. Adds `.active` to the clicked nav item
3. Removes `.active` from all `.content-section` elements
4. Adds `.active` to `#section-{name}`
5. Calls `refreshSection(section)` to render data

### Section → Render Function Map (line 1492)

| Section ID | Render Function |
|---|---|
| `dashboard` | `renderDashboard()` |
| `visits` | `renderVisits()` |
| `training` | `renderTrainings()` |
| `observations` | `renderObservations()` |
| `reports` | `_reportPopulateFilters()` |
| `resources` | `renderResources()` |
| `excel` | *(handled by excel-analytics.js)* |
| `notes` | `renderNotes()` |
| `planner` | `renderPlanner()` |
| `goals` | `renderGoals()` |
| `analytics` | `renderAnalytics()` |
| `followups` | `renderFollowups()` |
| `ideas` | `renderIdeas()` |
| `schools` | `renderSchoolProfiles()` |
| `clusters` | `renderClusterProfiles()` |
| `teachers` | `renderTeacherGrowth()` |
| `marai` | `renderMaraiTracking()` |
| `schoolwork` | `renderSchoolWork()` |
| `visitplan` | `renderVisitPlan()` |
| `reflections` | `renderReflections()` |
| `contacts` | `renderContacts()` |
| `teacherrecords` | `renderTeacherRecords()` |
| `meetings` | `renderMeetings()` |
| `worklog` | `renderWorkLog()` |
| `livesync` | `renderSyncSettings()` |
| `backup` | `renderBackupInfo()` |
| `settings` | `renderSettings()` |
| `feedback` | `renderFeedbackList()` |
| `growth` | `renderGrowthFramework()` |
| `capacitybuilding` | `renderCapacityBuilding()` |
| `importguide` | `renderImportGuide()` |
| `teachingpractices` | `renderTeachingPractices()` / `renderTpAnalytics()` |

---

## Feature Modules (app.js sections)

### Core Infrastructure

| Section | Lines | Key Functions |
|---|---|---|
| License Key Protection | 17–56 | `handleLicenseSubmit()`, `isLicenseActivated()` |
| Data Layer (DB) | 58–76 | `DB.get()`, `DB.set()`, `DB.clear()`, `DB.generateId()` |
| Password Protection | 78–113 | `PasswordManager.*`, lock screen, auto-lock timer |
| Encrypted File Storage | 115–171 | `EncryptedFileStorage.encrypt()`, `.decrypt()` |
| Encrypted Cache | 173–263 | `EncryptedCache.save()`, `.load()`, `.exists()`, `.clear()` |
| File System Access | 265–466 | `FileLink.linkFile()`, `.readFromFile()`, `.writeToFile()` |
| Google Drive Backup | 468–820 | `GoogleDriveSync.backup()`, `.restore()`, `.testConnection()` |
| Auto-Save System | 824–920 | `markUnsavedChanges()`, `performAutoSave()`, `startPeriodicSave()` |
| Navigation | 1468–1525 | `navigateTo()`, `refreshSection()` |
| Mobile/Desktop Sidebar | 1527–1575 | `closeMobileSidebar()`, `toggleDesktopSidebar()` |
| Theme Toggle | 1576–1605 | `toggleTheme()`, `restoreTheme()` |
| Toast Notifications | 2565–2578 | `showToast(message, type, duration)` |
| Modal Functions | 2579–2592 | `openModal(id)`, `closeModal(id)` |
| Pagination Utility | 2758–2800 | `getPaginatedItems()`, `renderPaginationControls()` |
| Custom Confirm Popup | 15643+ | `showPopupConfirm({ title, message, icon, ... })` — async, returns Promise<boolean> |

### Feature Modules

| Module | Lines (approx) | DB Key | Key Functions |
|---|---|---|---|
| **Dashboard** | 2598–2757 | *reads all* | `renderDashboard()`, `renderDashboardCharts()` |
| **School Visits** | 2802–3566 | `visits` | `openVisitModal()`, `saveVisit()`, `renderVisits()`, `deleteVisit()` |
| **Teacher Training** | 3567–4505 | `trainings` | `openTrainingModal()`, `saveTraining()`, `renderTrainings()` |
| **Training Attendance** | 3690–4505 | Inside `trainings[].attendanceList` | `addAttendee()`, `renderAttendanceList()`, `generateAttendanceReport()` |
| **Observations** | 4506–4886 | `observations` | `renderObservations()`, tab views (card/table/analytics) |
| **DMT Excel Import** | 4887–5593 | `observations` | `importDMTExcel()`, `triggerFilteredImport()`, `executeFilteredImport()` |
| **Smart Planner** | 5594–7162 | *reads visits, observations, trainings* | AI-powered suggestion engine for next visits |
| **Observation Analytics** | 7163–7301 | `observations` | Charts, stage distributions, engagement analysis |
| **Resources** | 7302–7421 | `resources` | `renderResources()`, `saveResource()` |
| **Reports** | 7422–9321 | *reads all* | 10 report types: Monthly, Cluster, Block, District, Health, Visits, Training, School, Summary, Reflective. All types share two export buttons: `printReport()` and `exportReportPDF()` — both route through `openPdfEditor`. |
| **Quick Notes** | 9322–9436 | `notes` | `renderNotes()`, `saveNote()` |
| **Weekly Planner** | 9437–9597 | `plannerTasks` | Drag-and-drop weekly planner |
| **Goal Tracker** | 9598–9838 | `goalTargets` | Monthly goals with target/actual/status |
| **Analytics** | 9839–10389 | *reads visits, trainings, observations* | Charts (Chart.js), visit-by-school, training trends |
| **Follow-up Tracker** | 10390–10537 | `followupStatus` | Action items with status tracking |
| **Idea Tracker** | 10538–10845 | `ideas` | `renderIdeas()`, inline editing, sharing |
| **School Profile Import** | 10846–11152 | *creates schools* | `triggerSchoolProfileImport()`, from Excel |
| **School Profiles** | 11355–11646 | via `observations` | `renderSchoolProfiles()`, school cards with visit/observation data |
| **Cluster Profiles** | 11647–12154 | via `observations` | `renderClusterProfiles()`, cluster-level aggregation |
| **School Student Records** | 12155–12480 | `schoolStudentRecords` | Learning level tracking by class |
| **Meeting Tracker** | 12481–12677 | `meetings` | `renderMeetings()`, BRC/cluster/district meetings |
| **School Health Card** | 12678–12838 | *reads observations, visits* | Print-ready school health card |
| **Reflections Journal** | 12839–12973 | `reflections` | Monthly reflection journal |
| **Contact Directory** | 12974–13412 | `contacts` | `renderContacts()`, `extractContactsFromObservations()` |
| **Teachers Record** | 13413–14000 | `teacherRecords` | `renderTeacherRecords()`, `extractTeachersFromObservations()`, import/export Excel |
| **MARAI Tracking** | 14001–14880 | `maraiTracking` | Stage tracking (M/A/R/A/I) per teacher, intervention suggestions |
| **School-Based Work** | 14881–15110 | `schoolWork` | Work assigned to teachers, status tracking |
| **Visit Plan** | 15111–15642 | `visitPlanEntries`, `visitPlanDropdowns` | Plan vs execution table, Excel link (two-way sync) |
| **Data & Security** | 16485–16628 | — | Backup status UI, file link status, password UI |
| **Export All Data** | 16629–16744 | — | `exportAllData()` — Excel export of everything |
| **App Settings** | 16745–17099 | — | Theme, cards per page, sidebar visibility, user profile |
| **Sidebar Visibility** | 17100–17164 | — | Toggle which sections appear in sidebar |
| **Bug Report** | 17257–17334 | `feedbackReports` | Bug/feature request with email via Gmail link |
| **Telegram Bot** | 17335–17587 | — | Send reports/data via Telegram bot API |
| **Professional Growth** | 17588–18789 | `growthAssessments`, `growthActionPlans` | 5-dimension self-assessment, radar chart, trend chart |
| **User Profile** | 18791–18880 | `userProfile` | Name, block, cluster, role, photo |
| **Monthly Work Log** | 18881–19227 | `worklog` | Auto-generated daily field diary from visits/trainings |
| **Dashboard Alerts** | 19246–19370 | *reads all* | Smart alerts (pending follow-ups, unvisited schools, etc.) |
| **Teacher Growth** | 19413–19608 | via `observations` | Teacher performance timeline from observations |
| **Observation Feedback** | 19609–19682 | `observations` | Print formatted observation feedback |
| **Period Comparison** | 19683–19770 | *reads all* | Compare two time periods |
| **Welcome Screen** | 19881–19981 | — | First boot: load file, start fresh, or load sample data |
| **User Guide** | 19982–20104 | — | Step-by-step guided tour (11 steps) |
| **Reflective Report** | 20189–20365 | Inside `visits[]/trainings[].reflectiveReport` | Reflective reports for visits/trainings, PDF export |
| **Capacity Building** | 20366–20703 | `selfCapacityBuilding` | Reading, courses, targets, skills, reflections |
| **Teaching Practices** | 20579–21160 | `teachingPractices` | Subject-wise practice management, Practices tab (CRUD, import, groups), Analytics tab (observation frequency, teacher adoption Yes/No/All, school coverage Yes/No breakdown, gap spotting) |
| **Import Guide** | 19241+ | — | `renderImportGuide()`, `_buildImportGuideHTML()` — in-app reference for all Excel import column formats |
| **Live Sync** | 1606–2564 | — | Peer-to-peer real-time sync via PeerJS (WebRTC) |

---

## PDF Export Architecture

All PDF exports across the entire app go through a single centralized editor before printing.

### `openPdfEditor(bodyHtml, title)` — line ~20096

A full-screen modal with a CKEditor 5-style rich text toolbar. The user can edit content before clicking **Export PDF**, which opens a print window with the final edited HTML.

**Toolbar features**: Font family, font size, paragraph/heading/blockquote/code blocks, bold/italic/underline/strikethrough/superscript/subscript, 15-color text color picker, 10-color highlight picker, bullet/numbered lists, alignment (L/C/R/J), indent/outdent, insert link, insert table, insert photo (base64), horizontal rule, remove formatting, undo/redo.

### All Export Functions → `openPdfEditor`

| Function | Section | Triggered By |
|---|---|---|
| `printReport()` | Reports | Print button in Reports section |
| `exportReportPDF()` | Reports | Download PDF button in Reports section |
| `printSchoolHealthCard()` | School Profiles | Print card button per school |
| `printVisitReport()` | School Visits | Print button on visit card |
| `printObsFeedback()` | Observations | Feedback button on observation card |
| `printTeacherSummary()` | Teacher Growth | Print icon on teacher panel |
| `printGrowthReport()` | Professional Growth | Print Report button |
| `printWorkLog()` | Monthly Work Log | Print button |
| `exportReflectiveReportPDF()` | Reflective Report Modal | Export PDF button in modal |

> **Zero direct `window.print()` or `html2pdf()` calls** remain outside of `openPdfEditor` itself.

---

## index.html Structure

### Document Structure
```
<!DOCTYPE html>
├── <head> — meta, fonts (Google Fonts: Inter), Font Awesome 6, Chart.js, html2pdf, XLSX
├── <body>
│   ├── #licenseScreen — License key gate (shown first if not activated)
│   ├── #lockScreen — Password lock screen
│   ├── #welcomeScreen — First boot / file load screen
│   ├── #sidebar — Navigation sidebar
│   │   ├── .sidebar-brand — App logo/title
│   │   ├── .nav-item[data-section="..."] — Nav links
│   │   └── .sidebar-footer — Version, lock button
│   ├── #sidebarOverlay — Mobile sidebar backdrop
│   ├── .mobile-header — Top bar on mobile
│   ├── .main-content
│   │   ├── section#section-dashboard
│   │   ├── section#section-visits
│   │   ├── section#section-training
│   │   ├── section#section-observations
│   │   ├── section#section-resources
│   │   ├── section#section-excel
│   │   ├── section#section-notes
│   │   ├── section#section-planner
│   │   ├── section#section-goals
│   │   ├── section#section-analytics
│   │   ├── section#section-followups
│   │   ├── section#section-ideas
│   │   ├── section#section-schools
│   │   ├── section#section-clusters
│   │   ├── section#section-teachers
│   │   ├── section#section-marai
│   │   ├── section#section-schoolwork
│   │   ├── section#section-visitplan
│   │   ├── section#section-reflections
│   │   ├── section#section-growth
│   │   ├── section#section-contacts
│   │   ├── section#section-teacherrecords
│   │   ├── section#section-livesync
│   │   ├── section#section-worklog
│   │   ├── section#section-meetings
│   │   ├── section#section-backup
│   │   ├── section#section-settings
│   │   ├── section#section-feedback
│   │   ├── section#section-capacitybuilding
│   │   ├── section#section-teachingpractices
│   │   └── section#section-importguide
│   ├── Modals (.modal-overlay)
│   │   ├── #visitModal, #trainingModal, #observationModal
│   │   ├── #filteredImportModal, #schoolProfileImportModal, #unloadModal
│   │   ├── #reflectiveReportModal, #capacityBuildingModal
│   │   ├── #userGuideModal, #meetingModal, #tpEditModal, #tpImportModal, etc.
│   │   └── #customConfirmPopup — reusable confirm dialog
│   ├── .toast-container — Toast notifications
│   └── #quickCaptureBtn — Floating action button
├── <script src="app.js">
└── <script src="excel-analytics.js">
```

### Modal Pattern

All modals follow this pattern:
```html
<div class="modal-overlay" id="{name}Modal">
    <div class="modal">
        <div class="modal-header">
            <h2>Title</h2>
            <button class="modal-close" onclick="closeModal('{name}Modal')">×</button>
        </div>
        <div class="modal-body">
            <!-- Content / Form -->
        </div>
        <div class="modal-footer">
            <button onclick="closeModal('{name}Modal')">Cancel</button>
            <button type="submit">Save</button>
        </div>
    </div>
</div>
```

Modals are opened/closed by toggling `.active` class on `.modal-overlay`.

---

## styles.css Organization

### Design System (CSS Variables — line 1)
```css
:root {
    --bg-primary: #0f172a;     /* Dark background */
    --bg-secondary: #1e293b;   /* Card/modal background */
    --bg-card: #1a2332;        /* Elevated card */
    --text-primary: #f1f5f9;   /* Main text */
    --text-secondary: #94a3b8; /* Muted text */
    --accent: #f43f5e;         /* Primary accent (rose) */
    --accent-hover: #e11d48;
    --border: rgba(255,255,255,0.06);
    --radius: 12px;
    --radius-xl: 16px;
    --shadow-lg: 0 8px 32px rgba(0,0,0,0.3);
    --transition: all 0.2s ease;
}
```

### Light Mode (line 52)
`body.light-mode` overrides all CSS variables with light values.

### Major CSS Sections (in order)
1. **CSS Variables & Design System** (1–93)
2. **Global Checkbox Style** (94–319) — custom premium checkboxes
3. **Global Styles** (320–360) — body, scrollbar, selection
4. **Sidebar** (361–539)
5. **Mobile Header** (540–578)
6. **Main Content & Section Transitions** (579–705)
7. **Section Header & Buttons** (706–930)
8. **Dashboard Cards & Metrics** (931–1090)
9. **Visit Styles** (1091–2117) — stats, list items, calendar view, detail panel
10. **Training Styles** (2156–2774) — cards, attendance
11. **Observation Styles** (2775–3160) — cards, stats, tabs, import dropdown
12. **Cluster Checkbox Multi-Select** (3161–3310)
13. **Smart Planner** (3458–4805)
14. **Weekly Calendar** (4995–5324)
15. **Resource Library** (5463–5580)
16. **Reports Section** (5581–6242)
17. **Modals & Forms** (6243–6489)
18. **Toast Notifications** (6490–6561)
19. **Responsive Design** (6696–7431) — breakpoints at 1200px, 900px, 600px, 480px
20. **Excel Analytics** (7445–10104) — upload zone, KPI strip, charts, tables
21. **Quick Notes** (9540–9706)
22. **Weekly Planner** (10671–10922)
23. **Goal Tracker** (10923–11150)
24. **Analytics** (11151–11776)
25. **Follow-up Tracker** (11543–11753)
26. **Idea Tracker** (11800–12323)
27. **School Profiles** (12324–12928)
28. **Reflections** (13421–13559)
29. **Contacts** (13560–13962)
30. **Work Log** (13994–14382)
31. **Meeting Tracker** (14383–14642)
32. **Live Sync** (14643–15679)
33. **Growth Framework** (15680–16889)
34. **Settings** (16890–17307)
35. **Backup & Restore** (17308–17539)
36. **Password Lock / Welcome Screen** (17540–17929)
37. **Encrypted Storage Card** (17930–18241)
38. **Google Drive Card** (18242–18689)
39. **MARAI** (19120–20582)
40. **School Work** (20583–21149)
41. **Visit Plan** (21415–22502)
42. **Feedback / Bug Report** (22503–22837)
43. **User Guide** (22838–23028)
44. **Teachers Record** (23029–23369)
45. **Custom Confirm Popup** (23370–23512)
46. **Reflective Report Modal** (23513–23661)
47. **Capacity Building** (23662–end)

---

## External Dependencies

| Library | Loaded Via | Purpose |
|---|---|---|
| **Inter** (Google Fonts) | `<link>` in `<head>` | Primary font |
| **Font Awesome 6** | `<link>` CDN | Icons throughout the app |
| **Chart.js 4** | `<script>` CDN | Dashboard charts, analytics, growth radar |
| **html2pdf.js** | `<script>` CDN | PDF export for reports |
| **XLSX (SheetJS)** | `node_modules` or CDN | Excel/CSV import & export |
| **PeerJS** | `<script>` CDN | WebRTC peer-to-peer for Live Sync |
| **QRCode.js** | `<script>` CDN | QR code for sync room codes |

---

## Common Patterns

### Data CRUD Pattern
```javascript
// Read
const items = DB.get('dataKey') || [];

// Create
const newItem = { id: DB.generateId(), ...data, createdAt: new Date().toISOString() };
items.push(newItem);
DB.set('dataKey', items);
if (typeof markUnsavedChanges === 'function') markUnsavedChanges();

// Update
const idx = items.findIndex(e => e.id === id);
if (idx >= 0) items[idx] = { ...items[idx], ...updatedData };
DB.set('dataKey', items);

// Delete
const filtered = items.filter(e => e.id !== id);
DB.set('dataKey', filtered);
```

### Modal Open/Close
```javascript
openModal('modalId');    // Adds .active class
closeModal('modalId');   // Removes .active class
```

### Confirm Dialog
```javascript
const confirmed = await showPopupConfirm({
    title: 'Delete Item',
    message: 'Are you sure?',
    icon: 'fa-trash',
    confirmText: 'Delete',
    confirmColor: '#ef4444'
});
if (!confirmed) return;
```

### Pagination
```javascript
const pg = getPaginatedItems(filteredItems, 'sectionKey', pageSize);
// pg = { items, page, totalPages, total, start, end }
// Render pg.items, then append:
renderPaginationControls('sectionKey', pg, 'renderFunctionName');
```

### Toast
```javascript
showToast('Message text', 'success');  // types: success, error, info, warning
showToast('With custom duration', 'info', 5000);
```

### Excel Export
```javascript
const ws = XLSX.utils.json_to_sheet(dataArray);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet Name');
XLSX.writeFile(wb, 'filename.xlsx');
```

### HTML Escaping
```javascript
escapeHtml(str)   // Used everywhere for XSS prevention in template literals
escapeHTML(str)    // Alias used in some modules (same function)
```

---

## App Boot Sequence

```
DOMContentLoaded
    │
    ├── restoreTheme()
    │
    ├── isLicenseActivated()?
    │   ├── NO  → Show #licenseScreen → wait for key → proceedAfterLicense()
    │   └── YES → proceedAfterLicense()
    │
    └── proceedAfterLicense()
            │
            ├── Restore FileLink handle from IndexedDB
            │
            ├── PasswordManager.isPasswordSet()?
            │   ├── YES → Try SessionPersist.restore()
            │   │         ├── Valid session → auto-load data → initApp()
            │   │         └── Invalid/none → showLockScreen('unlock')
            │   └── NO  → EncryptedCache.exists()?
            │             ├── YES → clear it, showWelcomeScreen()
            │             └── NO  → restoreFromLocalStorage()
            │                       ├── Data found → initApp()
            │                       └── No data → showWelcomeScreen()
            │
            └── initApp()
                    ├── renderDashboard()
                    ├── Setup event listeners
                    ├── Restore sidebar state
                    ├── Initialize star ratings
                    ├── applyProfileToUI()
                    └── Show user guide if first use
```

## Import File Structures

All imports use **XLSX (SheetJS)** and accept `.xlsx`, `.xls`, `.csv` files.

---

### 1. DMT Observations Import (`importDMTExcel` — line 5418)

**Source**: DMT Field Notes Excel from the Pratham/APF observation system.
**DB key**: `observations` | **Dedup key**: `NID + date + practiceSerial`
**Detection**: File must contain columns like `Teacher: Teacher Name`, `Practice Type`, or `Teacher Engagement Level`.

| Excel Column Name | Maps To | Required |
|---|---|---|
| `NID` | `nid` (teacher ID) | ✓ (for dedup) |
| `Response Date` | `date` | ✓ |
| `School Name` | `school` | ✓ |
| `Teacher: Teacher Name` | `teacher` | ✓ (detection) |
| `Teacher Phone No.` | `teacherPhone` | |
| `Teacher Stage` | `teacherStage` | |
| `Cluster` | `cluster` | |
| `Block Name` | `block` | |
| `Observation` | `observationStatus` | |
| `Observed While Teaching` | `observedWhileTeaching` | |
| `Teacher Engagement Level` | `engagementLevel` | ✓ (detection) |
| `Practice Type` | `practiceType` | ✓ (detection) |
| `Practice Master: Practice Serial No` | `practiceSerial` | ✓ (for dedup) |
| `Practice` | `practice` | |
| `Group` | `group` | |
| `Subject` | `subject` | |
| `Notes` | `notes` | |
| `Actual Observer: Full Name` | `observer` | |
| `Primary Observer: Full Name` | `observer` (fallback) | |
| `Stakeholder Status` | `stakeholderStatus` | |
| `History` | `history` | |
| `District Name` | `district` | |
| `State` | `state` | |

**Import modes**: `add` (append), `replace` (clear DMT imports, keep manual), `filtered` (with State/Block/Cluster/Observer filters).

> **CRITICAL**: This is the most important import. All observation-dependent features (MARAI, Smart Planner, Teacher Growth, Analytics, School/Cluster Profiles) rely on this data.

---

### 2. School Visits Import (`importVisitsExcel` — line 3416)

**DB key**: `visits` | **No dedup** — all rows are appended.

| Excel Column | Maps To | Required |
|---|---|---|
| `School` / `school` | `school` | ✓ |
| `Date` / `date` | `date` | |
| `Block` / `block` | `block` | |
| `Cluster` / `cluster` | `cluster` | |
| `District` / `district` | `district` | |
| `Status` / `status` | `status` (completed/planned/cancelled) | |
| `Purpose` / `purpose` | `purpose` | |
| `Duration` / `duration` | `duration` | |
| `Rating` / `rating` | `rating` | |
| `People Met` / `peopleMet` | `peopleMet` | |
| `Teachers Met` / `teachersMet` | `teachersMet` | |
| `Classes Visited` / `classesVisited` | `classesVisited` | |
| `Students Observed` / `studentCount` | `studentCount` | |
| `HM Present` / `hmPresent` | `hmPresent` | |
| `HM Discussion` / `hmDiscussion` | `hmDiscussion` | |
| `Activities` | `activities` (comma-separated → array) | |
| `Notes` / `notes` | `notes` | |
| `Best Practices` / `bestPractices` | `bestPractices` | |
| `Challenges` / `challenges` | `challenges` | |
| `Infrastructure` / `infrastructure` | `infrastructure` | |
| `SDMC / Community` / `sdmc` | `sdmc` | |
| `Materials Shared` / `materialsShared` | `materialsShared` | |
| `Broader Plan / Objective` / `broaderPlan` | `broaderPlan` | |
| `Follow-up` / `followUp` | `followUp` | |
| `Next Visit Date` / `nextDate` | `nextDate` | |

---

### 3. Teacher Records Import (`importTeacherRecordsExcel` — line 13905)

**DB key**: `teacherRecords` | **Dedup**: `name + school` (case-insensitive)
**Column detection**: Auto-matches using regex patterns (supports English + Hindi headers).

| Regex Pattern | Maps To | Matches |
|---|---|---|
| `/^(teacher\s*)?name\|शिक्षक\s*का\s*नाम\|full\s*name/i` | `name` | Name, Teacher Name, Full Name, शिक्षक का नाम |
| `/^gender\|लिंग\|sex/i` | `gender` | Gender, लिंग, Sex |
| `/^school\|विद्यालय\|स्कूल\|institution/i` | `school` | School, विद्यालय, Institution |
| `/^designation\|पदनाम\|post\|position/i` | `designation` | Designation, पदनाम, Post, Position |
| `/^subject\|विषय/i` | `subject` | Subject, विषय |
| `/^class\|कक्षा\|classes?\s*taught\|grade/i` | `classesTaught` | Class, Classes Taught, Grade, कक्षा |
| `/^phone\|mobile\|मोबाइल\|contact\s*no\|tel/i` | `phone` | Phone, Mobile, Contact No, मोबाइल |
| `/^email\|ई-?मेल/i` | `email` | Email, ई-मेल |
| `/^block\|ब्लॉक\|district/i` | `block` | Block, ब्लॉक |
| `/^cluster\|संकुल\|zone/i` | `cluster` | Cluster, संकुल, Zone |
| `/^quali\|योग्यता\|education\|degree/i` | `qualification` | Qualification, Education, योग्यता |
| `/^exp\|अनुभव\|years?\s*(of\s*)?exp/i` | `experience` | Experience, अनुभव |
| `/^(date\s*(of\s*)?)?join\|नियुक्ति\|doj/i` | `joinDate` | Join Date, DOJ, नियुक्ति |
| `/^(n\.?)?id\|employee\s*id\|emp\s*id/i` | `nid` → `employeeId` | NID, Employee ID, Emp ID |
| `/^note\|remarks\|टिप्पणी/i` | `notes` | Notes, Remarks, टिप्पणी |

> **Fallback**: If no `Name` column is matched, the **first column** is used as the name.

---

### 4. Training Attendance Import (`processAttendanceExcel` — line 4095)

**DB key**: Inside `trainings[].attendanceList` | **Dedup**: `name + school` (per training)

Uses a `findCol()` helper that matches column names case-insensitively:

| `findCol` Candidates | Maps To | Required |
|---|---|---|
| `name`, `teachername`, `teacher`, `शिक्षक` | `name` | ✓ |
| `school`, `schoolname`, `विद्यालय` | `school` | |
| `phone`, `mobile`, `contact`, `फोन` | `phone` | |
| `designation`, `post`, `stage`, `पद` | `designation` | |
| `cluster`, `संकुल` | `cluster` | |
| `block`, `ब्लॉक` | `block` | |

> **Side effect**: Attendees added manually or via Excel are **also auto-added to `teacherRecords`** if not already present.

---

### 5. School Profile Import (`loadSchoolProfileImportPreview` — line 10909)

**Source**: Same DMT Field Notes Excel as observation import.
**What it does**: Extracts unique schools from the DMT data and creates school profiles.
**Detection**: Same as DMT — checks for `School Name`, `Teacher: Teacher Name`, or `Practice Type` columns.

**Columns used for school extraction**:
- `School Name` → school name
- `Block Name` → block
- `Cluster` → cluster
- `State` → state
- `District Name` → district

**Filter options**: State, Block, Cluster (checkbox multi-select), Observer — same as filtered import.

---

### 6. Visit Plan Import (`importVisitPlanExcel` — line 15424)

**DB key**: `visitPlanEntries` + `visitPlanDropdowns`
**Special**: Reads **two sheets** from the same workbook.

#### Sheet 1: Main Visit Data (first non-dropdown sheet)
Columns by **position** (index-based, not header-based):

| Column Index | Maps To | Description |
|---|---|---|
| 0 (A) | `date` / `dateSerial` | Date (Excel serial number or string) |
| 1 (B) | `day` | Day name (auto-calculated if date present) |
| 2 (C) | `time` | Time of visit |
| 3 (D) | `domain` | Plan domain |
| 4 (E) | `stakeholderType` | Stakeholder type |
| 5 (F) | `cluster` | Cluster |
| 6 (G) | `venue` | Venue / school |
| 7 (H) | `stakeholderName` | Stakeholder name |
| 8 (I) | `designation` | Designation |
| 9 (J) | `objective` | Objective |
| 10 (K) | `review` | Review notes (if filled → status = `executed`) |
| 11 (L) | `tps` | TPS notes |
| 12 (M) | `stakeholderCount` | Number of stakeholders |
| 13 (N) | `teacherComments` | Teacher comments |
| 14 (O) | `studentComments` | Student comments |
| 15 (P) | `reportSharing` | Report sharing status |

**Status logic**: If `review` column is filled → `executed`, if any data exists → `planned`, else → `empty`.

#### Sheet 2: Dropdowns (sheet named "Sheet4" or "Dropdown")
Columns by **position** (data starts from row 4, col B):

| Column Index | Maps To | Content |
|---|---|---|
| B (1) | `domains` | Plan domain options |
| C (2) | `days` | Day name options |
| D (3) | `times` | Time slot options |
| E (4) | `stakeholderTypes` | Stakeholder type options |
| F (5) | `clusters` | Cluster name options |
| G (6) | `stakeholderNames` | Stakeholder name options |
| H (7) | `venues` | Venue/school options |
| I (8) | `designations` | Designation options |

---

### Import Data Flow Summary

```
DMT Excel (.xlsx)
    │
    ├─→ observations (DB)     ← importDMTExcel / triggerFilteredImport
    ├─→ school profiles       ← loadSchoolProfileImportPreview
    ├─→ contacts (DB)         ← extractContactsFromObservations
    └─→ teacherRecords (DB)   ← extractTeachersFromObservations

Teacher Excel (.xlsx)
    └─→ teacherRecords (DB)   ← importTeacherRecordsExcel

Attendance Excel (.xlsx)
    ├─→ trainings[].attendanceList (DB) ← processAttendanceExcel
    └─→ teacherRecords (DB)   ← auto-sync

Visit Plan Excel (.xlsx)
    ├─→ visitPlanEntries (DB)  ← importVisitPlanExcel
    └─→ visitPlanDropdowns (DB) ← dropdown sheet

School Visits Excel (.xlsx)
    └─→ visits (DB)           ← importVisitsExcel
```

---

## Key Conventions

1. **No framework** — Pure vanilla JS with DOM manipulation via `document.getElementById()` and template literals
2. **Section IDs** — HTML sections use `id="section-{name}"`, nav items use `data-section="{name}"`
3. **Deep clone** — `DB.get()` always returns a deep clone to prevent mutation bugs
4. **markUnsavedChanges()** — Must be called after every `DB.set()` to trigger auto-save
5. **escapeHtml()** — Always used when inserting user data into HTML templates
6. **Render pattern** — Each section has a `render{SectionName}()` function that rebuilds the entire section HTML
7. **Filter pattern** — Filters are `<select>` or `<input>` elements with `onchange` calling the render function with `_pageState.{section}=1` to reset pagination
8. **Modals** — Opened via `openModal(id)`, closed via `closeModal(id)` or clicking ✕ / Cancel button. Escape key closes all active modals.

