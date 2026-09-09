# Team Tracker

A web-based team management system built for Deloitte USI to monitor and manage team data, onboarding progress, and training status.

\---

## Technology Stack

|Component|Technology|
|-|-|
|Backend|ASP.NET Core Web API (.NET 8)|
|Frontend|HTML, CSS, JavaScript (Vanilla)|
|Data Store|Microsoft Excel (.xlsx) via EPPlus|
|Authentication|JWT Bearer Tokens|
|Hosting|GitHub Pages (Frontend)|
|Version Control|github.com/rk20597/TeamTracker|

\---

## Prerequisites

* Visual Studio 2022 or later
* .NET 8 SDK
* Microsoft Excel
* Git
* Chrome browser (recommended)

\---

## Setup Instructions

### 1\. Clone Repository

```
git clone https://github.com/rk20597/TeamTracker
```

### 2\. Open in Visual Studio

```
Backend/HROnboarding.API/HROnboarding.API.sln
```

### 3\. Verify Excel Files

```
Data/TeamTracker.xlsx
Data/HRData.xlsx
```

### 4\. Run API

```
Press F5 in Visual Studio
API starts at: http://localhost:5289
```

### 5\. Open Application

```
http://localhost:5289/index.html
```

\---

## Default Credentials

|Role|Username|Password|
|-|-|-|
|Admin|tmadiya@deloitte.com|Admin@123|
|Admin|kmullangi@deloitte.com|Admin2@123|
|User|karastogi@deloitte.com|User@123|

**Password Rules:** Min 8 chars, one uppercase, one number, one special character, must end with @deloitte.com

\---

## Application Tabs

|Tab|Visible To|Description|
|-|-|-|
|Active Inactive Members|Both|All team members with filters and sorting|
|Training Status|Both|Training completion with overdue highlighting|
|Onboarding Formalities ▾|Both|Dropdown with two sub-tabs|
|→ Onboarding Formalities|Both|Team-wise steps with completion tracking|
|→ Onboarding Status|Both|All candidates progress view|
|Onboarded/Offboarded|Both|Onboarded list and offboarded members|
|Management|Admin only|User management|

\---

## Role Based Access

|Feature|Admin|User|
|-|-|-|
|View Active Inactive Members|✅|✅|
|Add/Edit/Delete Members|✅|❌|
|View Training Status|✅|✅|
|Add/Edit Training Records|✅|✅|
|Delete Training Records|✅|❌|
|View Onboarding Formalities|✅|✅|
|Update Step Progress|✅|✅|
|View Onboarding Status|✅|✅|
|Edit/Delete Progress|✅|❌|
|View Onboarded/Offboarded|✅|✅|
|Add/Edit/Delete Offboarded|✅|❌|
|Management (Users)|✅|❌|

\---

## Excel Schema

### TeamTracker.xlsx

**TeamMember Sheet**

|Column|Field|Description|
|-|-|-|
|A|SrNo|Unique identifier (PK)|
|B|Name|Full name|
|C|Email|Deloitte email|
|D|Location|India / US / Mexico|
|E|Active/Inactive|Member status|
|F|Expansion|Phase 1 / Phase 2|
|G|City|City name|
|H|Level-Original|Original designation|
|I|Level|Current level|
|J|Client Level|Client level|
|K|Onboarding Date|Date of joining|
|L|Joined|Y / N|
|M|Job Family|.Net / CRM / ETL etc|
|N|GD Leader 1|Primary GD Lead|
|O|GD Leader 2|Secondary GD Lead|
|P|GD Leader 3|Third GD Lead|
|Q|Project|Assigned project|
|R|Client KT|KT status|
|S|Project Delivery|Delivery started|
|T|Tool Access|Tool access status|
|U|Datadog|Y / N / N/A|
|V|AKS Trained|Y / N / N/A|
|W|ROVO|Y / N / N/A|
|X|Claude Code Dev Day|Y / N / N/A|
|Y|AI Fluency|Y / N / N/A|
|Z|Claude 101|Y / N / N/A|
|AA|Other AI Certification|Y / N / N/A|
|AB|CoPilot Trained|Y / N / N/A|
|AC|Client Compliance|Y / N / N/A|
|AD|SciForma Access|Y / N / N/A|
|AE|Comments|Free text notes|
|AF|Offboarding Date|Date offboarded|
|AG|Mobile Number|10 digit contact|
|AH|Replacement|Replacement name|

**Users Sheet**

|Column|Field|Description|
|-|-|-|
|A|UserID|Unique identifier|
|B|UserName|Deloitte email|
|C|PasswordHash|Password|
|D|Role|Admin / User|
|E|IsActive|TRUE / FALSE|

**OnboardingSteps Sheet**

|Column|Field|Description|
|-|-|-|
|A|StepID|Unique identifier (PK)|
|B|TeamName|Project team name (FK → Project)|
|C|StepName|Step name|
|D|StepOrder|Sequential order|
|E|Description|Step description|

**OnboardingProgress Sheet**

|Column|Field|Description|
|-|-|-|
|A|ProgressID|Unique identifier|
|B|CandidateID|FK → TeamMember.SrNo|
|C|StepID|FK → OnboardingSteps.StepID|
|D|Status|NotStarted / InProgress / Completed|
|E|CompletedDate|Completion date|

**Training Status Sheet**

|Column|Field|Description|
|-|-|-|
|A|TrainingID|Unique identifier|
|B|CandidateID|FK → TeamMember.SrNo|
|C|Domain|DataDog / AKS / ROVO / Copilot / AI Fluency / Claude 101 / Client Compliance Trainings|
|D|Status|NotStarted / InProgress / Completed|
|E|DueDate|Due date|
|F|CompletedDate|Completion date|

**Offboarding Sheet**

|Column|Field|Description|
|-|-|-|
|A|CandidateId|FK → TeamMember.SrNo|
|B|Name|Member name|
|C|Email|Deloitte email|
|D|Role|Job role|
|E|Reason|Offboarding reason|

\---

## Key Features

### Overdue Highlighting

* Training records highlighted red when past due date and not completed
* Also highlights when completed after due date

### Onboarding Progress Persistence

* Selected team persists across sessions via localStorage
* Step completion status persists after logout and re-login

### Email Validation

* New users must have email existing in TeamMember sheet
* Offboarded members email must exist in TeamMember sheet
* All emails must end with @deloitte.com

### Force Logout on Role Change

* If admin changes their own role they are automatically logged out

\---

## API Endpoints

### Authentication

```
POST /api/auth/login
```

### Team Members

```
GET    /api/teammember/all
GET    /api/teammember/active
GET    /api/teammember/inactive
POST   /api/teammember              (Admin)
PATCH  /api/teammember/{srNo}       (Admin)
DELETE /api/teammember/{srNo}       (Admin)
```

### Training Status

```
GET    /api/trainingstatus
GET    /api/trainingstatus/withnames
POST   /api/trainingstatus          (Both)
PATCH  /api/trainingstatus/{id}     (Both)
DELETE /api/trainingstatus/{id}     (Admin)
```

### Onboarding

```
GET    /api/onboarding/steps
GET    /api/onboarding/teamsteps/{teamName}
POST   /api/onboarding/steps        (Admin)
PATCH  /api/onboarding/steps/{id}   (Admin)
DELETE /api/onboarding/steps/{id}   (Admin)
GET    /api/onboarding/progress
PATCH  /api/onboarding/progress     (Both)
PATCH  /api/onboarding/progress/{cId}/{sId}  (Admin)
DELETE /api/onboarding/progress/{cId}/{sId}  (Admin)
GET    /api/onboarding/statuswithnames
```

### Users (Admin Only)

```
GET    /api/users
POST   /api/users
PATCH  /api/users/{id}
DELETE /api/users/{id}
```

### Offboarded

```
GET    /api/offboarded
POST   /api/offboarded              (Admin)
PATCH  /api/offboarded/{id}         (Admin)
DELETE /api/offboarded/{id}         (Admin)
```

\---

## GitHub Pages

Frontend deployed at:

```
https://rk20597.github.io/TeamTracker
```

Note: Backend API must be running locally at http://localhost:5289

\---

## Security Notes

* All endpoints require JWT authentication
* Tokens expire after 8 hours
* Admin role required for write/delete operations
* Role changes force immediate logout
* SemaphoreSlim prevents concurrent file access
* Excel file must be closed during write operations

\---

## Known Limitations

* Excel file must be closed when API writes data
* Backend requires local server environment
* Concurrent users may experience file lock conflicts
* GitHub Pages serves frontend only

\---

## Project Structure

```
HROnboardingApp/
├── Backend/
│   └── HROnboarding.API/
│       ├── Controllers/
│       ├── Models/
│       ├── Repositories/
│       └── wwwroot/
├── Data/
│   ├── TeamTracker.xlsx
│   └── HRData.xlsx
├── docs/
├── dashboard.html
├── index.html
└── README.md
```

\---

*Team Tracker | Rohan Kulkarni | Deloitte USI | 2026*

