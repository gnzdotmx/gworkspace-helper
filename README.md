# gworkspace-helper

A simple Go library that provides helper functions for interacting with Google Workspace APIs, including Google Docs, Google Drive, and Google Calendar.

## Features

- **Authentication**: Handles OAuth2 and Service Account authentication.
- **Google Docs Helper** (`gdocsHelper`):
  - Create, copy, rename, and delete Google Docs.
  - Add and replace text, insert tables, manage permissions, and more.
- **Google Drive Helper** (`gDriveHelper`):
  - Create, rename, delete folders and files.
  - Manage file and folder permissions.
- **Google Calendar Helper** (`gMeetHelper`):
  - Create calendar events with timezone support.
  - Add attendees and attachments to events.

## Installation

```bash
go get github.com/gnzdotmx/gworkspace-helper
```

## Token
For the `credentials.json` file, `OAuth 2.0 Client ID` has to be `Desktop` type.