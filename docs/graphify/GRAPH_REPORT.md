# Graph Report - Jira_API_examples  (2026-10-05)

## Corpus Check
- Corpus is ~25,745 words - fits in a single context window. You may not need a graph.

## Summary
- 60 nodes · 115 edges · 9 communities (0 shown, 9 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- matplotlib_pyplot
- recursive_bug_jail.py
- DB-stats.py

## God Nodes (most connected - your core abstractions)
1. `add_chart_and_metrics()` - 2 edges
2. `add_chart_and_metrics()` - 2 edges
3. `add_chart_and_metrics()` - 2 edges
4. `get_issues()` - 2 edges
5. `Self-contained graphify pipeline for CI. Builds a knowledge graph over this…` - 1 edges

## Surprising Connections (you probably didn't know these)
- `add_chart_and_metrics()` --calls--> `matplotlib_pyplot`  [EXTRACTED]
  DB-stats.py →   _Bridges community 7 → community 2_

## Import Cycles
- None detected.

## Communities (9 total, 9 thin omitted)

## Knowledge Gaps
- **9 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Not enough signal to generate questions. This usually means the corpus has no AMBIGUOUS edges, no bridge nodes, no INFERRED relationships, and all communities are tightly cohesive. Add more files or run with --mode deep to extract richer edges._