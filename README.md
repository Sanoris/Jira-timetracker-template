# Jira Timesheet Tools

This repository provides two Python scripts for generating Excel timesheets from Jira data:

- **jira_timesheet_parser.py**: Converts a Jira RSS XML export into a timesheet.
- **jira_pull_timesheet.py**: Pulls issues and comments directly from the Jira Cloud API and builds a timesheet.

---

## Features

- Parse Jira RSS exports or pull directly from the Jira API.
- Group work by ticket and date.
- Estimate hours per event (default: 1.5).
- Output Excel files with clickable Jira ticket links.
- Simple configuration via `.config` files.

---

## Requirements

- Python 3.7+
- [pandas](https://pandas.pydata.org/)
- [python-dateutil](https://dateutil.readthedocs.io/en/stable/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [requests](https://requests.readthedocs.io/en/latest/) (for API pulls)

Install dependencies:

```sh
pip install pandas python-dateutil openpyxl requests
```

---

## Configuration

Copy `template.config` to `jira.config` and fill in your Jira details:

```ini
[JIRA]
base_url = https://yourcompany.atlassian.net
email = user@email.com
api_token = yourApiToken

[META]
domain = yourcompany
output_file = timesheet.xlsx
```

- `domain`: Used for building ticket links.
- `output_file`: Name of the generated Excel file.

**Note:** `jira.config` is in `.gitignore` and should not be committed.

---

## Usage

### 1. Using RSS Export

1. Export your Jira activity as RSS (XML file).
2. Run:
   ```sh
   python jira_timesheet_parser.py
   ```

   By default, it reads `jira_rss_export.xml` and outputs `callahan_timesheet.xlsx`.

### 2. Using Jira API

1. Ensure your `jira.config` is filled out.
2. Run:
   ```sh
   python jira_pull_timesheet.py
   ```

   This will pull your last month's activity and output to the configured Excel file.

---

## Output

The Excel file contains:

- `DATE`: Date of work log.
- `HOURS`: Estimated hours (default: 1.5 per event).
- `TICKET`: Jira ticket key.
- `DESCRIPTION`: Ticket summary.
- `LINK`: Clickable link to the Jira ticket.

---

## Customization

- **Estimated Hours:**
  Change the `hours = 1.5` line in the scripts to adjust the default hours per event.
- **Domain/Output File:**
  Edit the `domain` and `output_file` fields in your config file.

---

## License

MIT License

---

## Author

Matthew Hendricks
2025

---
