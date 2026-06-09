import codecs, sys
sys.stdout.reconfigure(encoding='utf-8')
files = [
    r'D:\Emaar Land\Qayd\qayd_part1_core.js',
    r'D:\Emaar Land\Qayd\qayd_part2_processing.js',
    r'D:\Emaar Land\Qayd\qayd_part3_reports.js'
]
with codecs.open(r'D:\Emaar Land\Qayd\qayd_system.js', 'w', 'utf-8-sig') as out:
    for f in files:
        with codecs.open(f, 'r', 'utf-8-sig') as inp:
            out.write(inp.read())
        out.write('\n')

with open(r'D:\Emaar Land\Qayd\qayd_system.js', 'r', encoding='utf-8-sig') as f:
    lines = f.readlines()
print(f'Total: {len(lines)} lines')

bugs = []
fixes = []
for i, line in enumerate(lines, 1):
    s = line.strip()
    if 'SUMPRODUCT((A$3' in s: bugs.append(f'L{i}: O(n2) SUMPRODUCT')
    if 'getFormula().indexOf' in s: bugs.append(f'L{i}: getFormula() bug')
    if 'replace(/B/g' in s: bugs.append(f'L{i}: Fragile regex')
    if 'sumRow=4+DR+2' in s: bugs.append(f'L{i}: Buried summary')
    if "var projs=['" in s: bugs.append(f'L{i}: Hardcoded projs')
    if "var emps=['" in s: bugs.append(f'L{i}: Hardcoded emps')
    if 'SORT(FILTER' in s: fixes.append('SORT(FILTER)')
    if 'getRangeByName' in s: fixes.append('Named Range read')
    if 'var sc=12' in s: fixes.append('Side summary')
    if 'H'+"(r-1)" in s: fixes.append('O(n) balance')
    if 'newConditionalFormatRule' in s: fixes.append('CondFmt')
    if 'مردود' in s and 'SUMIFS' in s: fixes.append('مردود in balance')
    if 'replace(/B(\\d)/g' in s: fixes.append('Safe regex')

if bugs:
    print('BUGS:')
    for b in bugs: print(f'  {b}')
else:
    print('NO BUGS')
print(f'FIXES ({len(set(fixes))}):')
for f in sorted(set(fixes)): print(f'  + {f}')
