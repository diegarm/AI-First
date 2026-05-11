# Task Plan: Timesheet Daily Hour Registration Tool

## Goal
Create a Python script that reads today's git commits, extracts DevOps task info, and registers hours in "Composables Timesheet(Timesheet).csv". On first run, asks for user and saves to config. If no commits, registers N/A for the day.

## Current Phase
Phase 1

## Phases

### Phase 1: Requirements & Discovery
- [x] Read CSV structure to understand columns
- [x] Understand data format (YYYYMMDD, Pessoa, Projeto, DevOps, Tarefa, Horas, Comentários, Date, Dia Semana)
- [x] Note: File is CSV despite being called "excel" by user
- **Status:** complete

### Phase 2: Design & Planning
- [x] Define config file structure (.timesheet-config.json)
- [x] Define git log command to extract commits by date
- [x] Define DevOps ID extraction regex from commit messages
- [x] Define hours distribution logic (equal split among unique tasks)
- [x] Define N/A row structure when no commits
- **Status:** complete

### Phase 3: Implementation
- [ ] Create timesheet.py (main script)
- [ ] Create run-timesheet.ps1 (PowerShell wrapper)
- **Status:** in_progress

### Phase 4: Testing & Verification
- [ ] Verify script runs on Windows PowerShell
- [ ] Verify config is saved and reloaded correctly
- [ ] Verify CSV rows are written correctly
- **Status:** pending

### Phase 5: Delivery
- [ ] Review all output files
- [ ] Deliver to user
- **Status:** pending

## Key Questions
1. What DevOps ID pattern to extract from commits? → 4-5 digit numbers (20000-29999 range seen in CSV)
2. How to distribute hours? → 8h split equally among unique DevOps IDs per day
3. Where is the git repo? → Configurable, default "." (current dir)

## Decisions Made
| Decision | Rationale |
|----------|-----------|
| Python script (.py) | Cross-platform, handles CSV well, git subprocess |
| .timesheet-config.json | Simple JSON, easy to edit manually |
| Regex for DevOps IDs | Extract 5-digit numbers from commit messages |
| Equal hours distribution | Simple default; 8h / number of unique tasks |
| N/A row on no commits | Explicit record that day was checked but no work committed |

## Errors Encountered
| Error | Attempt | Resolution |
|-------|---------|------------|
|       | 1       |            |
