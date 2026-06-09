import sys
sys.stdout.reconfigure(encoding='utf-8')

import codecs
files=[
    r'D:\Emaar Land\Qayd\qayd_part1_core.js',
    r'D:\Emaar Land\Qayd\qayd_part2_processing.js',
    r'D:\Emaar Land\Qayd\qayd_part3_reports.js'
]
with codecs.open(r'D:\Emaar Land\Qayd\qayd_system.js','w','utf-8-sig') as out:
    for f in files:
        with codecs.open(f,'r','utf-8-sig') as inp: out.write(inp.read())
        out.write('\n')

with open(r'D:\Emaar Land\Qayd\qayd_system.js','r',encoding='utf-8-sig') as f:
    content=f.read()
    lines=content.split('\n')

print(f'Total lines: {len(lines)}')
print('All setNamedRange occurrences:')
for i,line in enumerate(lines,1):
    if 'setNamedRange' in line:
        print(f'  L{i}: {line.strip()[:120]}')
