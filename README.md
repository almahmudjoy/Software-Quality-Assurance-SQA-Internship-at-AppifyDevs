# EchoGPT SQA Internship Assignment

**Candidate:** Abdullah Al Mahmud Joy  
**Company:** AppifyDevs  
**Submission date:** 27 September 2026

This repository contains functional testing, bug reporting, UI/UX review, exploratory testing, and optional API testing for the EchoGPT Chrome Extension and Android application.

## Assignment Coverage

| Requirement | Deliverable | Coverage |
|---|---|---:|
| Functional testing | [Software Quality Assurance (SQA) Internship at AppifyDevs.xlsx](./Software%20Quality%20Assurance%20%28SQA%29%20Internship%20at%20AppifyDevs.xlsx) | 40 test cases |
| Bug reporting | Same workbook, `Bug Reports` sheet | 10 issues |
| UI/UX review | Same workbook, `UI-UX` sheet | 10 suggestions |
| Exploratory testing | [Software Quality Assurance (SQA) Internship at AppifyDevs.pdf](./Software%20Quality%20Assurance%20%28SQA%29%20Internship%20at%20AppifyDevs.pdf) | 30-45 minute sessions |
| API testing bonus | [API_TESTING.md](./API_TESTING.md) | 10 scenarios |
| API collection | [EchoGPT.postman_collection.json](./EchoGPT.postman_collection.json) | 10 Postman requests |

## Applications Tested

- Chrome Extension: [EchoGPT - Multi-AI Chat Sidebar](https://chromewebstore.google.com/detail/echogpt-multi-ai-chat-sid/negimdcamohmoheiifgecbjgjepkcfhj)
- Android Application: [EchoChat](https://play.google.com/store/apps/details?id=com.echogpt.chatapp&hl=en)

## Main Findings

The most important findings were state reset defects after New Chat, a non-working Translate language swap control, missing Read Aloud playback, Compare response failure, intermittent Settings authentication navigation, and non-specific API validation errors.

## Evidence

### Functional and bug evidence

- [TC-READ-002](./TC-READ-002.png)
- [TC-TR-006 - Basic Translation](./TC-TR-006%20%E2%80%94%20Basic%20Translation.png)
- [BUG-003](./BUG-003.png)
- [BUG-005 - New Chat attachment state](./BUG-005%20%E2%80%94%20New%20Chat%20Does%20Not%20Clear%20Previous%20Attachment.png)
- [BUG-006 - New Chat draft state](./BUG-006%20%E2%80%94%20New%20Chat%20Does%20Not%20Clear%20Unsaved%20Draft.png)
- [BUG-007 - Translation result state](./BUG-007%20%E2%80%94%20Translation%20Result%20Persists%20After%20Clearing%20Input.png)

### UI/UX evidence

Screenshots are named `UX-001` through `UX-010` and correspond to the `UI-UX` workbook sheet.

- [UX-001 Web](./UX-001-Web.png) and [UX-001 Android](./UX-001-Android.jpeg)
- [UX-002 Translate swap](./UX-002-Translate-Swap.png)
- [UX-003 Read Aloud](./UX-003-Read-Aloud.png)
- [UX-004 Compare](./UX-004-Compare.png)
- [UX-005 New Chat state](./UX-005-New-Chat-State.png)
- [UX-006 File types](./UX-006-File-Upload-Types.png)
- [UX-007 File size](./UX-007-File-Upload-Limit%20or%20Size.png)
- [UX-008 Error message](./UX-008-Error-Message.png)
- [UX-009 Translate clear](./UX-009-Translate-Clear-Button.png)
- [UX-010 Error recovery](./UX-010-Error-Recovery.png)

## API Testing

The API report is available in [API_TESTING.md](./API_TESTING.md). It documents one successful chat-generation request, nine rejected validation scenarios, two API issues, and the limitation around conclusively testing streaming.

The Postman collection has been sanitized before publication. Import [EchoGPT.postman_environment.example.json](./EchoGPT.postman_environment.example.json), provide a fresh local token in Postman, and never commit credentials.

Published Postman documentation: [View collection documentation](https://documenter.getpostman.com/view/53084089/2sBYB4K6oR)

## Repository Contents

- [README.md](./README.md) - submission overview
- [API_TESTING.md](./API_TESTING.md) - API scope, results, evidence, and issues
- [EchoGPT.postman_collection.json](./EchoGPT.postman_collection.json) - sanitized Postman collection
- [EchoGPT_SQA_Assignment_Functional_and_Bug_Reports.py](./EchoGPT_SQA_Assignment_Functional_and_Bug_Reports.py) - workbook-generation source
- [Software Quality Assurance (SQA) Internship at AppifyDevs.xlsx](./Software%20Quality%20Assurance%20%28SQA%29%20Internship%20at%20AppifyDevs.xlsx) - completed workbook
- [Software Quality Assurance (SQA) Internship at AppifyDevs.pdf](./Software%20Quality%20Assurance%20%28SQA%29%20Internship%20at%20AppifyDevs.pdf) - exploratory report

## Submission Checklist

- [x] 30+ functional test cases
- [x] 10+ genuine bug reports
- [x] 10 UI/UX suggestions
- [x] Exploratory testing report
- [x] Postman API collection and API report
- [x] Screenshots linked from the repository
- [x] Credentials removed from the published collection

GitHub repository: [almahmudjoy/Software-Quality-Assurance-SQA-Internship-at-AppifyDevs](https://github.com/almahmudjoy/Software-Quality-Assurance-SQA-Internship-at-AppifyDevs)
