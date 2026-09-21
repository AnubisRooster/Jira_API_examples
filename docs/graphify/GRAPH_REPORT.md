# Graph Report - Jira_API_examples  (2026-09-21)

## Corpus Check
- Corpus is ~20,192 words - fits in a single context window. You may not need a graph.

## Summary
- 60 nodes · 111 edges · 9 communities (4 shown, 5 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- graphify_pipeline.py
- jira
- All_GIT_160_Days.py
- confluence_w_epic.py
- recursive_bug_jail.py
- PMZ_Components.py
- pandas
- matplotlib_pyplot

## God Nodes (most connected - your core abstractions)
1. `Self-contained graphify pipeline for CI. Builds a knowledge graph over this…` - 1 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Import Cycles
- None detected.

## Communities (9 total, 5 thin omitted)

### Community 0 - "graphify_pipeline.py"
Cohesion: 0.15
Nodes (11): graphify_analyze, graphify_build, graphify_cluster, graphify_detect, graphify_export, graphify_extract, graphify_llm, graphify_report (+3 more)

### Community 1 - "jira"
Cohesion: 0.35
Nodes (4): datetime, jira, openpyxl, openpyxl_styles

### Community 3 - "confluence_w_epic.py"
Cohesion: 0.40
Nodes (4): atlassian, configparser, os, sys

### Community 6 - "pandas"
Cohesion: 0.60
Nodes (3): docx, docx_shared, pandas

## Knowledge Gaps
- **5 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Not enough signal to generate questions. This usually means the corpus has no AMBIGUOUS edges, no bridge nodes, no INFERRED relationships, and all communities are tightly cohesive. Add more files or run with --mode deep to extract richer edges._