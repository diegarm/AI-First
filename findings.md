# Findings & Decisions

## Requirements
- Daily hour registration in CSV file named "Composables Timesheet(Timesheet).csv"
- First run: ask for user name, save to config
- Subsequent runs: use saved user from config
- Registrations based on git commits for the day
- No commits → register day as N/A
- Config must persist user setting across runs

## CSV Structure (discovered)
Columns: `Data YYYYMMDD`, `Pessoa`, `Projeto`, `DevOps`, `Tarefa`, `Horas`, `Comentários`, `_`, `Date`, `Dia Semana`

Sample row: `20251211,Alexandre Gago,SI Pessoas,20116,Dev,8.00,,,11-12-2025,#REF!`

Notes:
- `Dia Semana` column has `#REF!` — was a formula in Excel; script will compute actual weekday in Portuguese
- Multiple rows per day are valid (one per DevOps task)
- Hours per day typically sum to 8.00
- DevOps IDs are 5-digit numbers in the 20000-29000 range

## Technical Decisions
| Decision | Rationale |
|----------|-----------|
| Python (stdlib only) | No extra dependencies needed; csv, json, subprocess, re, datetime |
| .timesheet-config.json | Easy to read/edit; stores user, project, task, git_repo_path |
| subprocess git log | Reliable way to get commits without gitpython dep |
| Regex `\b2\d{4}\b` | Matches DevOps IDs like 20116, 20382, 21079 seen in data |
| PS1 wrapper | Easy to double-click or alias in PowerShell |

## Issues Encountered
| Issue | Resolution |
|-------|------------|
|       |            |

## Resources
- CSV file: `Composables Timesheet(Timesheet).csv`
- Config: `.timesheet-config.json`
- Git log format: `--format="%ai|%H|%s"` (author date ISO, hash, subject)
