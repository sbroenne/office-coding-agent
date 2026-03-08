---
name: test-excel-prompt
description: Analyse a named Excel range and summarise the data.
agent: Test Excel Agent
argument-hint: Range name (e.g. SalesData)
---

Please analyse the Excel range "${input:rangeName}" and provide:
1. A summary of the data (row/column count, data types).
2. Key statistics (min, max, average where applicable).
3. Any obvious data quality issues.
