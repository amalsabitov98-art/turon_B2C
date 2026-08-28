import json, io, sys
data = json.load(open('data/deals.json'))
tpl = open('prototype/dashboard.tpl.html', encoding='utf-8').read()
out = tpl.replace('/*__DATA__*/null', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
open('prototype/dashboard.html','w',encoding='utf-8').write(out)
print('ok', len(out), 'байт')
