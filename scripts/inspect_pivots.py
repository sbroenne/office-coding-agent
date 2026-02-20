"""Inspect pivot filter and other patterns."""
import re

with open(r'D:\source\office-coding-agent\tests-e2e\src\test-taskpane.ts', 'r', encoding='utf-8') as f:
    content = f.read()

# Find all pivot filter actions
for m in re.finditer(r"action: 'filter'[^}]+?\}", content, re.DOTALL):
    print('FILTER:', repr(m.group()[:150]))
    print()

print('---')
# Find pivot sort_field_values (with valuesHierarchyName)
for m in re.finditer(r"action: 'sort'[^}]+?valuesHierarchyName[^}]+?\}", content, re.DOTALL):
    print('SORT_VALUES:', repr(m.group()[:200]))
    print()

print('---')
# Find all set_number_validation patterns (now data_validation/set)
idx = 0
while True:
    idx = content.find("action: 'set',\n", idx)
    if idx < 0:
        break
    print('DV SET:', repr(content[idx:idx+120]))
    print()
    idx += 1

print('---')
# Find all pivot manual filter
for m in re.finditer(r"pivotTableConfigs,\n\s+'pivot',\n\s+\{ action: 'filter',\n", content, re.DOTALL):
    print('PIVOT FILTER:', repr(content[m.start():m.start()+200]))
    print()
