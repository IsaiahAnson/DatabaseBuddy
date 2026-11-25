<img src="https://raw.githubusercontent.com/IsaiahAnson/database-buddy/main/docs/images/banner.svg" width="100%" alt="Database Buddy, an automation by Isaiah Anson. Sorts OneDrive IT docs into Notion databases. Pipedream workflow, GPT-4o, 10 Notion databases.">

[![CI](https://github.com/IsaiahAnson/database-buddy/actions/workflows/ci.yml/badge.svg)](https://github.com/IsaiahAnson/database-buddy/actions/workflows/ci.yml)

An automated IT knowledge-base builder. On a schedule, it scans OneDrive for IT documentation,
classifies each new or changed document with OpenAI, and files it in one of ten Notion
databases, one per IT category.

Built as a [Pipedream](https://pipedream.com) workflow with seven custom Node.js steps. All step
code is in [`steps/`](steps/).

**Status: prototype.** The step logic is covered by unit tests that fake OneDrive, OpenAI, Notion
and the data store (see [Testing](#testing)). The workflow has not been run end to end on
Pipedream since the latest changes to the steps.

![Database Buddy pipeline: sources, seven steps, ten Notion databases](docs/images/pipeline.png)

The diagram is older than the current code: step 7 no longer moves pages under a parent, the
alias count is now 39, and nothing in this repo produces the Excel or OneNote inputs.

## Category routing

![Category routing: aliases mapped onto ten canonical IT categories](docs/images/categories.png)

Step 1 asks the model to answer with one of the ten category names. Step 4 then routes each
record in this order:

1. a canonical name, in any case (`security` → Security);
2. an exact alias, 39 in all (`azure`, `intune`, `virtual desktop`, …);
3. an alias found as a whole word in a longer label, longest alias first
   (`Azure AD Connect` → User Management, `Windows Autopilot` → MDM).

A record whose category still has no database is reported in the step's `errors` list.

## Features

- **Scheduled scanning**: walks a OneDrive folder tree through the Microsoft Graph API, follows
  Graph's paging, skips folders whose name exactly matches the exclude list, and stops at a
  file cap. This deployment runs Monday, Wednesday and Friday at 11:00 UTC.
- **Skips unchanged files**: step 1 stores each classified file's id and eTag in a Pipedream data
  store and skips files whose eTag has not changed, so later runs reach files the earlier runs
  did not get to.
- **AI classification**: GPT-4o in JSON mode returns type, key topics, technical relevance, use
  cases, a category and a confidence score for each document. The prompt lists the ten
  category names and asks for exactly one of them.
- **No duplicate pages**: before creating a page, step 4 looks for a page with the same
  `Source ID` (the OneDrive item id) and updates that page instead.
- **Deduplication**: near-duplicates are dropped by word-set similarity, keeping the
  better-extracted copy. Step 2 can also merge Excel and OneNote records, but nothing in this
  repo produces them.
- **Normalization**: ISO-8601 dates, tags split on five delimiters, and canonical status and
  priority values. Unrecognized fields are kept under `metadata`.
- **AI enhancement**: a second OpenAI pass (`gpt-4o-mini` by default, JSON mode) writes a title,
  summary, tags, priority and optional custom sections such as `action_items`.
- **Notion sync**: writes the enhanced title, summary and tags back to each page, appends or
  replaces content blocks, and can archive pages that have no matching record.

## Stack

- [Pipedream](https://pipedream.com) workflow (Node.js, ES modules)
- Microsoft Graph API for OneDrive
- OpenAI (GPT-4o for classification, `gpt-4o-mini` by default for enhancement)
- Notion API (the Pipedream Notion app currently sends `Notion-Version: 2025-09-03`)
- GitHub Actions: syntax check of every step and the unit tests

## Setup

### 1. Notion databases

Create one database per category: Network Infrastructure, VDI, Applications, Hardware,
Security, MDM, Cloud Services, Backup Recovery, User Management and Monitoring. Each needs these
properties, with these exact names:

| Property | Type | Written by |
| --- | --- | --- |
| `Name` | Title | step 4 (file name), step 7 (enhanced title) |
| `Source` | Text | step 4 |
| `Source ID` | Text | step 4 (OneDrive item id, used to find the page again) |
| `Source URL` | URL | step 4 |
| `Category` | Select | step 4 |
| `Content Type` | Select | step 4 |
| `Tags` | Multi-select | step 4, step 7 |
| `Created Date` | Date | step 4, on create only |
| `Last Modified` | Date | step 4 |
| `Content Summary` | Text | step 4, step 7 |

Step 3 can create a database from JSON property definitions. Run it once, by hand or in a
separate workflow, not in the scheduled one: each run creates another database.

### 2. Data store

Create a Pipedream data store and select it in step 1. Keys look like `onedrive:<item id>`.
To make step 1 process a file again, delete its key.

### 3. Step wiring

The step names below (`scan`, `dedupe`, …) are examples; use the names in your workflow.

| Step | Prop | Value |
| --- | --- | --- |
| 1 `scan` | `folderId`, `data_store` | your folder and data store |
| 2 `dedupe` | `onedrive_data` | `{{steps.scan.$return_value.records}}` |
| 2 `dedupe` | `max_records` | at least step 1's `maxFiles` (the default 3 is lower than step 1's default 5; records over the cap are dropped, and step 1 has already marked their files as done) |
| 4 `populate` | `cleanedData` | `{{steps.dedupe.$return_value.final_records.map(r => JSON.stringify(r))}}` |
| 4 `populate` | one database prop per category | the data source of each Notion database |
| 5 `normalize` | `records` | `{{steps.populate.$return_value.records.map(r => JSON.stringify(r))}}` |
| 6 `enhance` | `records` | `{{steps.normalize.$return_value.normalized_records.map(r => JSON.stringify(r))}}` |
| 7 `update` | `enhancedRecords` | `{{steps.enhance.$return_value.enhanced_records.map(r => JSON.stringify(r))}}` |
| 7 `update` | `createdPages` (optional) | `{{steps.populate.$return_value.created_pages.map(p => JSON.stringify(p))}}` |

### 4. Record shape

Steps 1, 2, 4, 5, 6 and 7 pass records with these fields:

| Field | Meaning |
| --- | --- |
| `id` | OneDrive item id. Stored in Notion as `Source ID`. |
| `eTag` | OneDrive eTag when the file was scanned |
| `title` | File name; step 6 replaces it with the enhanced title |
| `category` | One of the ten category names, or the model's label if it is not one |
| `content_type` | Document type from step 1, for example `Technical Guide` |
| `tags` | Array of strings; step 6 replaces them with suggested tags |
| `summary` | Key information from step 1; step 6 replaces it with a detailed summary |
| `content` | Extracted text (see the limitation below) |
| `source` | `OneDrive` |
| `webUrl` | Link to the file in OneDrive |
| `pageId` | Notion page id, added by step 4 |

Step 6 also adds `priority_level`, `categories`, `original_title`, any custom sections, and
`enhancement_error` when a record could not be enhanced.

## Quickstart

1. Open the share link: **[pipedream.com/new?h=tch_BXf12z](https://pipedream.com/new?h=tch_BXf12z)**.
   It holds the workflow as it was first published. To use the current code, paste each file in
   `steps/` into the matching code step.
2. Connect your **Microsoft OneDrive**, **OpenAI** and **Notion** accounts when prompted.
3. Follow [Setup](#setup): databases, data store and step wiring.
4. Set the schedule. This deployment runs `0 11 * * 1,3,5` (UTC).

## How it works

### 1 · OneDrive knowledge scanner
[`steps/01_comprehensive_onedrive_processor.js`](steps/01_comprehensive_onedrive_processor.js)

Walks the folder tree (`/items/{id}/children`, following `@odata.nextLink`), filters by MIME type
(Word, PDF, plain text, Markdown, OneNote) and skips files whose eTag is already in the data
store. Plain text and Markdown are read in full; other formats are described from their
metadata. Each document goes to GPT-4o for classification, and near-duplicates are dropped by
Jaccard word-set similarity. The step returns `records` in the record shape, plus one of three
report shapes (`json`, `summary`, `knowledge_base`). Files that were classified are stored in the
data store; files whose classification failed are not, so the next run tries them again.

### 2 · Clean and deduplicate
[`steps/02_clean_and_deduplicate_data.js`](steps/02_clean_and_deduplicate_data.js)

Merges records from up to three inputs (OneDrive, Excel, OneNote), normalizes text, removes
duplicates by percentage similarity, and returns at most `max_records`. With persistent storage
on, it remembers each fully handled batch by the items in it and returns only records added in
the current run.

### 3 · Create a Notion database
[`steps/03_create_specialized_it_databases.js`](steps/03_create_specialized_it_databases.js)

Creates one Notion database at the workspace level from JSON property definitions, such as
`{"name": "Name", "type": "title"}`. Any Notion property type can be given; options such as a
select's list go in the object named after the type.

### 4 · Populate by IT category
[`steps/04_populate_specialized_databases.js`](steps/04_populate_specialized_databases.js)

Routes each record to its database (see [Category routing](#category-routing)), looks for an
existing page by `Source ID`, and either creates a page with properties and a content block or
updates the existing page's properties. Returns `created_pages` and `records` with `pageId` added.

### 5 · Normalize records
[`steps/05_normalize_database_records.js`](steps/05_normalize_database_records.js)

Resolves field names case-insensitively against candidate lists (`title`/`name`/`subject`/…),
normalizes second and millisecond timestamps to ISO-8601, and maps status and priority values to
canonical vocabularies (`"done"` → `completed`, `"blocker"` → `critical`). `pageId`, `summary`,
`content_type`, `webUrl` and `eTag` are kept as they are.

### 6 · AI enhancement
[`steps/06_enhance_individual_records.js`](steps/06_enhance_individual_records.js)

Asks the model for an enhanced title, detailed summary, categories, suggested tags, priority and
content type, plus any custom sections, and writes them into `title`, `summary` and `tags`.

### 7 · Update Notion pages
[`steps/07_update_notion_pages_with_enhanced_content.js`](steps/07_update_notion_pages_with_enhanced_content.js)

Matches records to pages by page ID (or title), updates `Name`, `Content Summary` and `Tags`, and
appends or replaces content blocks. Records that step 6 could not enhance are skipped.

## Errors

- Steps collect per-item errors and keep going when some items succeed.
- A step fails the run when nothing it was given succeeded, so Pipedream's error notifications
  fire: step 1 when the top folder cannot be read or every classification failed, and steps 4,
  6 and 7 when no record or page was written.
- Step 3 fails on invalid property JSON.

## Design notes

- **No secrets in code.** Authentication goes through Pipedream's managed app connections
  (`microsoft_onedrive`, `openai`, `notion`). Nothing in this repo holds credentials.
- **Configurable.** Similarity thresholds, batch sizes, file caps, output formats and custom AI
  sections are exposed as step props.

## Known limitations

- PDF, Word and OneNote content is analyzed from metadata descriptions, not full text.
- Step 1 marks a file as done once it is classified. If a later step fails for that file, delete
  its data store key to process it again.
- The `Source ID` lookup only searches the database of the record's current category.
- When step 4 updates an existing page, it changes the properties only, not the page body.
- With `replaceContent` off, step 7 appends its blocks below the existing content.
- Step 2's persistent storage compares new records with every stored one, so a changed file can
  be dropped as a duplicate of its earlier version. Step 1 already skips unchanged files, so
  leave it off in this workflow.

## Testing

```bash
node --test test/*.test.mjs
```

[`test/steps.test.mjs`](test/steps.test.mjs) loads each step file with its package imports
stubbed and checks:

- category matching in steps 1 and 4, and that the step 1 prompt names the ten categories step 4
  routes;
- the OpenAI request (JSON mode, the step's `$`) and failure handling in steps 1 and 6;
- data store skipping, Graph paging and exact folder excludes in step 1;
- stable batch ids in step 2;
- the database body step 3 sends;
- the Notion properties, `Source ID` lookup, create-or-update and page ids in step 4;
- the fields step 5 keeps;
- a contract test that a step 6 record changes the title, summary and tags step 7 writes, and that
  step 7 sends no `parent`.

The fakes follow the Pipedream app methods (`createChatCompletion`, `queryDataSource`,
`createPage`, `updatePage`, `httpRequest`). No test calls OneDrive, OpenAI or Notion.

## Project layout

```text
steps/          source of each workflow step, exported from Pipedream
test/           unit tests for the step logic (node:test)
docs/images/    README diagrams
.github/        CI (syntax check for every step, unit tests)
```

## License

Copyright (c) 2026 Isaiah Anson. All rights reserved. You may use the released software for
personal, non-commercial use; copying, modifying or redistributing it requires written
permission. See [LICENSE](LICENSE).
