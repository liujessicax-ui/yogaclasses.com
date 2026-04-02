# Yoga Website Business Rules - Flowchart Documentation

## How to View These Diagrams

- **GitHub**: Push this file to a repository; GitHub renders Mermaid blocks natively.
- **VS Code**: Install the "Markdown Preview Mermaid Support" extension, then open Markdown preview (`Cmd+Shift+V`).
- **mermaid.live**: Copy any individual `mermaid` code block into [https://mermaid.live](https://mermaid.live) for live editing and PNG/SVG export.
- **Obsidian / Notion**: Both render Mermaid blocks natively in their Markdown views.

---

## 1. Main Sign-Up Flow

This is the primary diagram covering the full journey from visiting the schedule page through form submission, duplicate checks, waiver logic, and confirmation.

```mermaid
flowchart TD
    A([User visits schedule.html]) --> B[View 3 Class Cards]

    B --> B1["Sunday 6-7:15 PM PST\nOnline - Google Meet\nOpen to Everyone"]
    B --> B2["Tuesday 6-7:15 PM PST\nOnline - Google Meet\nRestorative Yoga"]
    B --> B3["Wednesday 5:15-6:30 PM PST\nCCV Clubhouse - In Person\nCCV Residents Only"]

    B1 --> C{Class within\n7-day rolling window\nAND before 6:15 PM PST\ncutoff?}
    B2 --> C
    B3 --> C

    C -- No --> CUTOFF[Sign-up closed.\nNext week's class\nimmediately appears.]
    C -- Yes --> D["User clicks 'Sign Up' button"]

    D --> E["Navigate to signup.html\nwith ?class= parameter\npre-selecting the class"]

    E --> F["User fills out form:\n- First Name\n- Last Name\n- Email\n- Class Selection\n  (multi-select checkboxes)"]

    F --> G{Form validation:\nAll fields filled?\nValid email format?}
    G -- No --> G_ERR[Show validation errors.\nUser corrects and resubmits.]
    G_ERR --> G
    G -- Yes --> H

    H["Send GET request to\nGoogle Apps Script\nfor duplicate check"]

    H --> I{Duplicate found?\nSame name + email +\nclass + class date\nin Google Sheets?}

    I -- Yes --> J["Show 'Already Signed Up' popup:\n- Already registered message\n- Option to choose another class\n- Check confirmation email\n  for cancellation link"]
    J --> F

    I -- No --> K{Any selected class\nin-person?\nWednesday CCV?}

    K -- No / Online Only --> SUBMIT["POST to Google Apps Script.\nWrite one row per class:\nTimestamp, First Name, Last Name,\nEmail, Class, Class Date,\nClass Type, Liability Waiver = N/A"]

    K -- Yes / In-Person --> L["GET request to Google Sheets:\nCheck for prior waiver\n(same first + last + email\nwith Waiver = YES)"]

    L --> M{Returning student?\nPrior waiver on file?}

    M -- Yes --> SUBMIT_RETURN["POST to Google Apps Script.\nWrite rows.\nLiability Waiver = PREVIOUSLY SIGNED"]

    M -- No / First-Time --> N["Show Liability Waiver popup:\n- Assumption of risk\n- Release of liability\n- Indemnification\n- Medical acknowledgment\n- CA governing law\n- Severability\nReferences: Jessica, CCV,\nHOA, property mgmt,\nboard members, residents"]

    N --> O{User checks\nagreement checkbox?}
    O -- No --> O_WAIT["'I Agree & Submit' button\nremains disabled"]
    O_WAIT --> O
    O -- Yes --> P["'I Agree & Submit'\nbutton enabled.\nUser clicks it."]

    P --> SUBMIT_WAIVER["POST to Google Apps Script.\nWrite rows.\nLiability Waiver = YES"]

    SUBMIT --> CONFIRM
    SUBMIT_RETURN --> CONFIRM
    SUBMIT_WAIVER --> CONFIRM

    CONFIRM([Confirmation Page:\n- Personalized thank you\n- List of classes signed up for\n- Email confirmation message])

    style A fill:#e8f5e9,stroke:#388e3c
    style CONFIRM fill:#e8f5e9,stroke:#388e3c
    style CUTOFF fill:#fff3e0,stroke:#f57c00
    style J fill:#fff3e0,stroke:#f57c00
    style G_ERR fill:#ffebee,stroke:#d32f2f
    style N fill:#e3f2fd,stroke:#1565c0
    style SUBMIT fill:#c8e6c9,stroke:#2e7d32
    style SUBMIT_RETURN fill:#c8e6c9,stroke:#2e7d32
    style SUBMIT_WAIVER fill:#c8e6c9,stroke:#2e7d32
```

---

## 2. Class Availability Logic

This diagram details the rolling 7-day window and cutoff time logic that determines which classes are available for sign-up at any given moment.

```mermaid
flowchart TD
    START([Determine available classes]) --> NOW["Get current date/time\nin PST (UTC-8)"]

    NOW --> WINDOW["Calculate 7-day\nrolling window:\nToday through Today + 6 days"]

    WINDOW --> MON{Is Sunday\nwithin window?}
    WINDOW --> TUE{Is Tuesday\nwithin window?}
    WINDOW --> WED{Is Wednesday\nwithin window?}

    MON -- Yes --> MON_CUT{Current time\npast Sunday\n6:15 PM PST?}
    MON -- No --> MON_HIDE[Sunday class\nnot shown]

    MON_CUT -- No --> MON_SHOW["Show THIS Sunday:\nSunday 6-7:15 PM PST\nOnline via Google Meet\nOpen to Everyone\nSign-up OPEN"]
    MON_CUT -- Yes --> MON_NEXT["Show NEXT Sunday:\nSign-up opens for\nnext week's class\nimmediately"]

    TUE -- Yes --> TUE_CUT{Current time\npast Tuesday\n6:15 PM PST?}
    TUE -- No --> TUE_HIDE[Tuesday class\nnot shown]

    TUE_CUT -- No --> TUE_SHOW["Show THIS Tuesday:\nTuesday 6-7:15 PM PST\nOnline via Google Meet\nRestorative Yoga\nSign-up OPEN"]
    TUE_CUT -- Yes --> TUE_NEXT["Show NEXT Tuesday:\nSign-up opens for\nnext week's class\nimmediately"]

    WED -- Yes --> WED_CUT{Current time\npast Wednesday\n5:30 PM PST?}
    WED -- No --> WED_HIDE[Wednesday class\nnot shown]

    WED_CUT -- No --> WED_SHOW["Show THIS Wednesday:\nWednesday 5:15-6:30 PM PST\nCCV Clubhouse - In Person\nCCV Residents Only\nSign-up OPEN"]
    WED_CUT -- Yes --> WED_NEXT["Show NEXT Wednesday:\nSign-up opens for\nnext week's class\nimmediately"]

    MON_SHOW --> COMPARE
    MON_NEXT --> COMPARE
    TUE_SHOW --> COMPARE
    TUE_NEXT --> COMPARE
    WED_SHOW --> COMPARE
    WED_NEXT --> COMPARE
    MON_HIDE --> COMPARE
    TUE_HIDE --> COMPARE
    WED_HIDE --> COMPARE

    COMPARE["All times compared:\nPST (UTC-8) vs\nuser's local time"]

    COMPARE --> RENDER([Render available\nclass cards on\nschedule.html])

    style START fill:#e8f5e9,stroke:#388e3c
    style RENDER fill:#e8f5e9,stroke:#388e3c
    style MON_SHOW fill:#c8e6c9,stroke:#2e7d32
    style TUE_SHOW fill:#c8e6c9,stroke:#2e7d32
    style WED_SHOW fill:#c8e6c9,stroke:#2e7d32
    style MON_NEXT fill:#fff9c4,stroke:#f9a825
    style TUE_NEXT fill:#fff9c4,stroke:#f9a825
    style WED_NEXT fill:#fff9c4,stroke:#f9a825
    style MON_HIDE fill:#eeeeee,stroke:#9e9e9e
    style TUE_HIDE fill:#eeeeee,stroke:#9e9e9e
    style WED_HIDE fill:#eeeeee,stroke:#9e9e9e
```

---

## 3. Page Structure / Sitemap

This diagram shows the website's page hierarchy, navigation structure, and how pages connect to each other.

```mermaid
flowchart TD
    NAV["Main Navigation Bar"]

    NAV --> HOME["index.html\nHome Page\n- Hero section\n- Welcome message\n- What to expect\n- Inspirational quote"]

    NAV --> ABOUT["about.html\nAbout Me\n- Background story\n- Photo placeholders\n- Contact info placeholder\n- Gallery"]

    NAV --> SCHEDULE["schedule.html\nClass Schedule\n- 3 class cards\n- Props listed inline\n- Sign Up buttons"]

    NAV --> PRIVATES["privates.html\nPrivate Sessions\n- In-person only\n- Playa Del Rey, CA"]

    NAV --> PROPS["props.html\nProps\n- Blocks\n- Straps\n- Yoga chair\n- Blanket\n- Bolsters\n- Images for each"]

    NAV --> DONATE["donations.html\nDonations\n- Classes always free\n- PayPal hosted button\n- Venmo: @Jessica-eifkc\n- Credit/debit card\n- Supporting Iyengar\n  certification journey"]

    SCHEDULE -- "Sign Up button\n?class= parameter" --> SIGNUP["signup.html\nSign-Up Form\n(NOT in main nav)\n- First Name\n- Last Name\n- Email\n- Class checkboxes"]

    SIGNUP -- "POST via\nApps Script" --> GSHEETS[("Google Sheets\n'Yoga Signup' spreadsheet\n'Sign-Ups' sheet\n\nColumns:\nTimestamp | First Name\nLast Name | Email\nClass | Class Date\nClass Type | Liability Waiver")]

    SIGNUP -- "GET via\nApps Script" --> GSHEETS

    SIGNUP -- "On success" --> CONFIRMATION["Confirmation Page\n- Personalized thank you\n- Classes signed up for\n- Email confirmation msg"]

    DONATE -- "PayPal link" --> PAYPAL["External: PayPal\nHosted donate button"]
    DONATE -- "Venmo link" --> VENMO["External: Venmo\n@Jessica-eifkc"]

    style NAV fill:#7e57c2,stroke:#4527a0,color:#fff
    style HOME fill:#e8f5e9,stroke:#388e3c
    style ABOUT fill:#e8f5e9,stroke:#388e3c
    style SCHEDULE fill:#e8f5e9,stroke:#388e3c
    style PRIVATES fill:#e8f5e9,stroke:#388e3c
    style PROPS fill:#e8f5e9,stroke:#388e3c
    style DONATE fill:#e8f5e9,stroke:#388e3c
    style SIGNUP fill:#fff3e0,stroke:#f57c00
    style CONFIRMATION fill:#c8e6c9,stroke:#2e7d32
    style GSHEETS fill:#e3f2fd,stroke:#1565c0
    style PAYPAL fill:#eeeeee,stroke:#9e9e9e
    style VENMO fill:#eeeeee,stroke:#9e9e9e
```

---

## 4. Quick Reference: Google Sheets Data Flow

```mermaid
flowchart LR
    FORM["signup.html\nForm Submission"] -- "GET: duplicate check\n+ waiver lookup" --> SCRIPT["Google Apps Script\nWeb App Endpoints"]
    FORM -- "POST: write sign-up rows\n(one row per class)" --> SCRIPT

    SCRIPT <--> SHEET[("Google Sheets\n'Yoga Signup'\n'Sign-Ups' sheet")]

    SHEET --- COLS["Columns:\n1. Timestamp\n2. First Name\n3. Last Name\n4. Email\n5. Class\n6. Class Date\n7. Class Type\n8. Liability Waiver"]

    subgraph "GET Checks"
        DUP["Duplicate Check:\nfirst + last + email\n+ class + date"]
        WAIVER["Waiver Check:\nfirst + last + email\nwith Waiver = YES"]
    end

    SCRIPT --> DUP
    SCRIPT --> WAIVER

    subgraph "POST Values per Row"
        ROW["Timestamp: auto\nFirst Name: form\nLast Name: form\nEmail: form\nClass: selected class\nClass Date: calculated\nClass Type: online/in-person\nWaiver: YES / N/A / PREV"]
    end

    SCRIPT --> ROW

    style FORM fill:#fff3e0,stroke:#f57c00
    style SCRIPT fill:#e3f2fd,stroke:#1565c0
    style SHEET fill:#e8f5e9,stroke:#388e3c
```

---

## 5. Liability Waiver Decision Tree

```mermaid
flowchart TD
    START{Does sign-up include\nan in-person class?\nWednesday CCV?} -- No --> SKIP["No waiver needed.\nSubmit directly.\nWaiver field = N/A"]

    START -- Yes --> CHECK["Query Google Sheets:\nLook up first name +\nlast name + email\nwhere Waiver = YES"]

    CHECK --> RETURNING{Match found?\nReturning in-person\nstudent?}

    RETURNING -- Yes --> SKIP_RETURN["Skip waiver.\nSubmit directly.\nWaiver field =\nPREVIOUSLY SIGNED"]

    RETURNING -- No / First Time --> SHOW["Display Liability Waiver Popup"]

    SHOW --> CONTENTS["Waiver Contents:\n1. Assumption of risk\n2. Release of liability\n3. Indemnification\n4. Medical acknowledgment\n5. California governing law\n6. Severability"]

    CONTENTS --> REFS["Referenced Parties:\n- Jessica (instructor)\n- Cross Creek Village\n- HOA\n- Property management\n- Board members\n- Residents"]

    REFS --> CHECKBOX{User checks\nagreement checkbox?}

    CHECKBOX -- No --> DISABLED["'I Agree & Submit'\nbutton stays disabled.\nCannot proceed."]
    DISABLED --> CHECKBOX

    CHECKBOX -- Yes --> ENABLED["Button enabled.\nUser clicks\n'I Agree & Submit'."]

    ENABLED --> SUBMIT["Submit to Google Sheets.\nWaiver field = YES"]

    style START fill:#e3f2fd,stroke:#1565c0
    style SKIP fill:#c8e6c9,stroke:#2e7d32
    style SKIP_RETURN fill:#c8e6c9,stroke:#2e7d32
    style SUBMIT fill:#c8e6c9,stroke:#2e7d32
    style SHOW fill:#fff3e0,stroke:#f57c00
    style DISABLED fill:#ffebee,stroke:#d32f2f
```

---

## 6. Duplicate Registration Check

```mermaid
flowchart TD
    SUBMIT["User submits\nsign-up form"] --> QUERY["GET request to\nGoogle Apps Script:\nSend first name, last name,\nemail, class, class date"]

    QUERY --> SEARCH["Apps Script searches\n'Sign-Ups' sheet for\nmatching row"]

    SEARCH --> MATCH{Row found where\nALL match?\n- First Name\n- Last Name\n- Email\n- Class Name\n- Class Date}

    MATCH -- No Match --> PROCEED["No duplicate.\nProceed with\nsubmission flow."]

    MATCH -- Match Found --> POPUP["Show 'Already Signed Up' popup:\n- You are already registered\n  for this class on this date\n- Choose a different class\n- Check your confirmation\n  email for cancellation link"]

    POPUP --> CHOICE{User action?}
    CHOICE -- "Choose another class" --> BACK["Return to form.\nUser selects\ndifferent class(es)."]
    CHOICE -- "Dismiss" --> CLOSE["Close popup.\nForm remains\nas-is."]

    style SUBMIT fill:#e3f2fd,stroke:#1565c0
    style PROCEED fill:#c8e6c9,stroke:#2e7d32
    style POPUP fill:#fff3e0,stroke:#f57c00
```
