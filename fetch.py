import json, csv, os, time, urllib.request, urllib.parse, datetime as dt

T = ['LYP6.DE', 'SXR8.DE', 'AAPL', '^GSPC', '^IXIC', '^STOXX', '^STOXX50E', '^GDAXI',
     '^FCHI', '^FTSE', '^N225', '^HSI', '000001.SS', 'EUNL.DE']
UA = {'User-Agent': 'Mozilla/5.0'}

def chart(t, q):
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(t)}?{q}"
    for i in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                return json.load(r)['chart']['result'][0]
        except Exception:
            time.sleep(5)
    return None

os.makedirs('data', exist_ok=True)
latest = {'fetched_at': dt.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')}
for t in T:
    r = chart(t, 'range=1d&interval=1d')
    if not r: latest[t] = None; continue
    m = r['meta']
    latest[t] = [m.get('regularMarketPrice'), m.get('chartPreviousClose'),
                 dt.datetime.utcfromtimestamp(m['regularMarketTime'] + m['gmtoffset']).strftime('%Y-%m-%d'), m.get('currency')]
json.dump(latest, open('data/latest.json', 'w'), indent=1)

# daily close history (for the weekly report)
path = 'data/history.csv'
rows = {}
if os.path.exists(path):
    for d, t, c in csv.reader(open(path)):
        if d != 'date': rows[(d, t)] = c
start = int(dt.datetime(2025, 12, 29).timestamp())
for t in T:
    r = chart(t, f'period1={start}&period2={int(time.time())}&interval=1d')
    if not r or 'timestamp' not in r: continue
    off = r['meta']['gmtoffset']; cl = r['indicators']['quote'][0]['close']
    for ts, c in zip(r['timestamp'], cl):
        if c is not None:
            rows[(dt.datetime.utcfromtimestamp(ts + off).strftime('%Y-%m-%d'), t)] = f'{c:.4f}'
with open(path, 'w', newline='') as f:
    w = csv.writer(f); w.writerow(['date', 'ticker', 'close'])
    for (d, t), c in sorted(rows.items()): w.writerow([d, t, c])
print(json.dumps(latest))
