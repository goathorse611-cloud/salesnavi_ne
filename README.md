# Blueprint ワークショップ

Google Apps Script (GAS) Web Application for sales team collaboration with customers.

## Overview

Blueprint ワークショップ is a workshop facilitation tool that enables sales teams to:
- Build consensus with customers
- Plan next actions
- Generate approval documents (稟議書) via Google Docs export

## Features

- **9 Screens**: Home, Project List, Project Detail, and 7 Workshop Modules
- **Project CRUD**: Create, list, view, and update projects
- **Module System**: 7 structured workshop modules with save/load/complete functionality
- **Value Tracker**: Track quantitative and qualitative benefits per use case
- **Export to Google Docs**: Generate formal approval documents
- **Permission Control**: Creator and editors only access
- **Audit Logging**: All write operations are logged

## Project Structure

```
salesnavi_ne/
├── Code.gs              # Server-side GAS code (APIs, CRUD, permissions)
├── Index.html           # Main HTML template with shared layout
├── Stylesheet.html      # CSS, Tailwind config, custom styles
├── ClientJS.html        # Client-side JavaScript (SPA routing, events)
├── Views/               # Extracted view templates
│   ├── home.html
│   ├── projectlist.html
│   ├── projectdetail.html
│   ├── vision.html
│   ├── usecaseselect.html
│   ├── 90daysplan.html
│   ├── raci.html
│   └── valuetracker.html
├── stitch_raw/          # Original Stitch View Code files (preserved as-is)
│   ├── home.html
│   ├── projectlist.html
│   ├── projectdetail.html
│   ├── vision.html
│   ├── usecaseselect.html
│   ├── 90daysplan.html
│   ├── raci.html
│   └── valuetracker.html
└── README.md            # This file
```

## Database Schema

Uses Google Spreadsheet as database with 4 sheets:

### 1. プロジェクト (Projects)
| Column | Description |
|--------|-------------|
| プロジェクトID | UUID |
| 顧客名 | Customer name |
| 作成日 | Created date |
| 作成者メール | Creator email |
| 編集者メールCSV | Editors (comma-separated) |
| 状態 | Status (下書き/確定) |
| 最終更新 | Last updated |

### 2. モジュールデータ (Module Data)
| Column | Description |
|--------|-------------|
| プロジェクトID | Project ID |
| モジュールID | Module number (1-7) |
| JSON | Module data as JSON |
| 最終更新 | Last updated |
| 更新者 | Updated by |

### 3. 価値 (Values)
| Column | Description |
|--------|-------------|
| プロジェクトID | Project ID |
| ユースケースID | Use case identifier |
| 定量効果 | Quantitative effect |
| 定性効果 | Qualitative effect |
| 証跡リンク | Evidence link |
| 次の投資判断 | Next investment decision |
| 最終更新 | Last updated |

### 4. 監査ログ (Audit Log)
| Column | Description |
|--------|-------------|
| 日時 | Timestamp |
| ユーザー | User email |
| 操作 | Operation type |
| プロジェクトID | Project ID |
| 詳細JSON | Details as JSON |

## Deployment

### Prerequisites
- Google Account with access to Google Apps Script
- Google Drive for storing the Spreadsheet database

### Steps

1. **Create a new Google Apps Script project**
   - Go to https://script.google.com
   - Create a new project
   - Name it "Blueprint ワークショップ"

2. **Upload the files**
   - Create each file in the Apps Script editor:
     - `Code.gs` (main server code)
     - `Index.html`
     - `Stylesheet.html`
     - `ClientJS.html`
   - Create `Views` folder and add all view files

3. **Configure the Spreadsheet**
   - Update `CONFIG.SPREADSHEET_ID` in `Code.gs` with your Spreadsheet ID
   - Or leave empty to auto-create on first run

4. **Initialize the database**
   - Run the `initDb()` function from the Apps Script editor
   - This creates the 4 required sheets with headers

5. **Deploy as Web App**
   - Click "Deploy" > "New deployment"
   - Select "Web app"
   - Set:
     - Execute as: "User accessing the web app"
     - Who has access: "Anyone with Google account" (or your organization)
   - Click "Deploy"
   - Copy the web app URL

6. **Test the application**
   - Open the web app URL in a browser
   - Create a test project
   - Navigate through modules
   - Test save/load functionality

## Configuration

Edit `CONFIG` object in `Code.gs`:

```javascript
const CONFIG = {
  SPREADSHEET_ID: '',  // Leave empty to auto-create, or set your Spreadsheet ID
  SHEET_NAMES: {
    PROJECTS: 'プロジェクト',
    MODULE_DATA: 'モジュールデータ',
    VALUES: '価値',
    AUDIT_LOG: '監査ログ'
  }
};
```

## API Reference

### Project APIs
- `createProject(customerName, editorsCsv)` - Create new project
- `listProjects(filter, search)` - List projects (filter: 'all', 'mine', 'shared')
- `getProject(projectId)` - Get project details
- `updateProjectStatus(projectId, newStatus)` - Update project status

### Module APIs
- `loadModule(projectId, moduleId)` - Load module data
- `saveModule(projectId, moduleId, payloadJson)` - Save module data
- `markComplete(projectId, moduleId)` - Mark module as complete
- `unmarkComplete(projectId, moduleId)` - Unmark module completion

### Value Tracker APIs
- `saveValue(projectId, useCaseId, quantitative, qualitative, evidenceLink, nextInvestment)` - Save value entry
- `listValues(projectId)` - List all values for a project

### Export APIs
- `exportProjectDoc(projectId)` - Generate Google Doc and return URL

### Utility APIs
- `getCurrentUserInfo()` - Get current user email
- `getAuditLogs(projectId)` - Get audit logs for a project

## Security

- **Permission Checks**: All APIs verify user is creator or editor
- **XSS Prevention**: HTML escaping for all user input
- **Audit Logging**: All write operations logged with user, timestamp, and details
- **HTTPS Only**: All external resources loaded via HTTPS (IFRAME sandbox requirement)

## Browser Support

- Chrome (recommended)
- Firefox
- Safari
- Edge

Note: Uses Tailwind CSS CDN and Google Material Symbols for styling.

## License

Internal use only.
