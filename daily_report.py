"""Daily portfolio report.
usage: python3 daily_report.py --xlsx P.xlsx --fetch fetch.json --fx "2026-11-18,1.0995;2026-11-17,1.0989"
       --template Daily_Report_Template.docx --world world.json --date 2026-11-18 --out DIR
Prints a JSON summary (for the chat message) and writes DIR/Portfolio_Report_<date>.pdf.
Deps: pip install docxtpl ; LibreOffice (soffice) for PDF."""
import argparse, json, os, subprocess, datetime as dt
import openpyxl, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm

a = argparse.ArgumentParser()
for k in ('xlsx', 'fetch', 'fx', 'template', 'world', 'date', 'out'): a.add_argument('--' + k, required=True)
A = a.parse_args()
TODAY = A.date
F = json.load(open(A.fetch))
fx = [(s.split(',')[0], float(s.split(',')[1])) for s in A.fx.split(';')]  # newest first
FX_T, FX_P = fx[0][1], fx[1][1]
UP, DN, GRAY = '006300', 'B42B2B', '898781'
MON = ['januar', 'februar', 'marec', 'april', 'maj', 'junij', 'julij', 'avgust', 'september', 'oktober', 'november', 'december']
DAY = ['ponedeljek', 'torek', 'sreda', 'četrtek', 'petek', 'sobota', 'nedelja']

def num(x, d=2): return f"{x:,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')
def eur(x, d=2, sign=False): return ('+' if sign and x > 0 else '') + num(x, d) + '\u00a0€'
def pct(x, d=2, sign=True): return ('+' if sign and x > 0 else '') + num(x * 100, d) + '\u00a0%'
def rt(text, v, bold=False, size=14):
    r = RichText(); r.add(text, color=GRAY if v is None else (UP if v >= 0 else DN), bold=bold, size=size); return r

# ---------- holdings from Transactions (Buy - Sell, average cost) ----------
wb = openpyxl.load_workbook(A.xlsx, data_only=True)
meta = {r[0]: r for r in wb['Totals'].iter_rows(min_row=2, values_only=True) if r[0]}
qty, cost = {}, {}
for r in wb['Transactions'].iter_rows(min_row=2, values_only=True):
    if not r[0]: continue
    i, typ, q, amt = r[1], r[2], r[3] or 0, r[6] or 0
    if typ == 'Buy': qty[i] = qty.get(i, 0) + q; cost[i] = cost.get(i, 0) + amt
    elif typ == 'Sell' and qty.get(i):
        c = cost[i] / qty[i] * q; qty[i] -= q; cost[i] -= c
notes, rows, missing = [], [], []
tot_v = tot_prev = tot_cost = 0
for isin, q in qty.items():
    if q <= 1e-9: continue
    t = meta[isin][1]; d = F.get(t)
    if not d or d[0] is None:
        missing.append(t); notes.append(f'{t}: no data (Yahoo ni vrnil tečaja) – izključeno iz vsot.'); continue
    last, prev, ldate, cur = d
    usd = cur == 'USD'
    v = q * last / (FX_T if usd else 1)
    closed = ldate != TODAY
    pv = v if closed else q * prev / (FX_P if usd else 1)
    if closed: notes.append(f'{t}: trg danes ni trgoval, uporabljen zadnji zaključek {ldate}.')
    rows.append(dict(t=t, name=meta[isin][2], q=q, last=last, cur=cur, v=v, pv=pv, cost=cost[isin], closed=closed))
    tot_v += v; tot_prev += pv; tot_cost += cost[isin]
holdings = []
for r in sorted(rows, key=lambda r: -r['v']):
    dp = r['v'] / r['pv'] - 1; de = r['v'] - r['pv']; gl = r['v'] - r['cost']
    holdings.append(dict(name=r['name'], ticker=r['t'], qty=num(r['q'], 4),
        price=f"{num(r['last'])}\u00a0{'USD' if r['cur'] == 'USD' else '€'}", value=eur(r['v']), weight=pct(r['v'] / tot_v, 1, False),
        d_pct=rt('zaprto' if r['closed'] else pct(dp), None if r['closed'] else dp), d_eur=rt(eur(de, sign=True), de),
        gl=rt(f"{eur(gl, 0, True)}\u00a0({pct(gl / r['cost'], 1)})", gl), _dp=dp, _de=de))
chg = tot_v - tot_prev; chg_p = chg / tot_prev

# ---------- indices ----------
IDX = [('ZDA', '^GSPC', 'S&P 500', -74.0, 40.7, -112, 52), ('ZDA', '^IXIC', 'Nasdaq Composite', -74.0, 40.7, -112, 31),
       ('Evropa', '^STOXX', 'STOXX Europe 600', 4.4, 50.8, -40, 37), ('Evropa', '^STOXX50E', 'Euro Stoxx 50', 4.4, 50.8, 36, 38),
       ('Evropa', '^GDAXI', 'DAX (Frankfurt)', 8.7, 50.1, 30, 63), ('Evropa', '^FCHI', 'CAC 40 (Pariz)', 2.35, 48.85, -36, 51),
       ('Evropa', '^FTSE', 'FTSE 100 (London)', -0.1, 51.5, -16, 65), ('Azija', '^N225', 'Nikkei 225 (Tokio)', 139.7, 35.7, 152, 50),
       ('Azija', '^HSI', 'Hang Seng (Hongkong)', 114.2, 22.3, 98, 4), ('Azija', '000001.SS', 'Shanghai Composite', 121.5, 31.2, 140, 12),
       ('Svet', 'EUNL.DE', 'MSCI World (EUNL, EUR)', None, None, None, None)]
indices, pts = [], []
for reg, t, name, lon, lat, bx, by in IDX:
    d = F.get(t)
    if not d or d[0] is None:
        indices.append(dict(region=reg, name=name, last='no data', chg=rt('no data', None), date='')); notes.append(f'{name}: no data.'); continue
    c = d[0] / d[1] - 1; closed = d[2] != TODAY
    lbl = 'zaprto' if closed else pct(c)
    indices.append(dict(region=reg, name=name, last=num(d[0]), chg=rt(lbl, None if closed else c), date=f"{d[2][8:10]}.{d[2][5:7]}."))
    pts.append((name.split(' (')[0], lon, lat, bx, by, None if closed else c, d[2]))

# ---------- map ----------
world = json.load(open(A.world))
fig, ax = plt.subplots(figsize=(7.2, 3.0))
for p in world: ax.add_patch(Polygon(p, closed=True, fc='#ebeae5', ec='#ffffff', lw=0.4))
for name, lon, lat, bx, by, c, dte in pts:
    if lon is None:
        ax.text(-128, -5, f"MSCI World  {'zaprto' if c is None else pct(c)}", fontsize=8, fontweight='bold',
                color='#' + (GRAY if c is None else UP if c >= 0 else DN), ha='left', va='bottom'); continue
    col = '#' + (GRAY if c is None else UP if c >= 0 else DN); bg = '#f1f0ec' if c is None else ('#e4f2e4' if c >= 0 else '#f9e3e3')
    ax.plot(lon, lat, 'o', ms=4, color='#52514e', zorder=4)
    ax.annotate(f"{name}\n{'zaprto · ' + dte[8:10] + '.' + dte[5:7] + '.' if c is None else pct(c)}", (lon, lat), (bx, by),
                fontsize=6.8, color=col, ha='center', va='center', fontweight='bold', zorder=5,
                bbox=dict(boxstyle='round,pad=0.3', fc=bg, ec=col, lw=0.6),
                arrowprops=dict(arrowstyle='-', color='#898781', lw=0.6))
ax.set_xlim(-130, 172); ax.set_ylim(-8, 72); ax.set_aspect(1.1); ax.axis('off')
fig.subplots_adjust(0, 0, 1, 1); mp = os.path.join(A.out, 'map.png'); fig.savefig(mp, dpi=220); plt.close(fig)

# ---------- render ----------
y, m, dd = map(int, TODAY.split('-'))
dl = f"{DAY[dt.date(y, m, dd).weekday()]}, {dd}. {MON[m - 1]} {y}"
if F.get('fetched_at'): notes.append(f"Tečaji pridobljeni {F['fetched_at'].replace('T', ' ').replace('Z', ' UTC')}.")
notes.append('Viri: Yahoo Finance (zadnji tečaj in prejšnji zaključek), referenčni tečaj ECB za USD. Dnevna sprememba = zadnji tečaj / prejšnji zaključek. '
             'D/I glede na nabavno vrednost iz lista »Transactions«. Celoletni pregled (TWR, XIRR, tveganje) je v tedenskem poročilu.')
doc = DocxTemplate(A.template)
ctx = dict(date_long=dl, total=eur(tot_v), chg=rt(f"{eur(chg, sign=True)} ({pct(chg_p)})", chg, True, 26),
           chg_pct=rt(pct(chg_p), chg_p), chg_eur=rt(eur(chg, sign=True), chg), holdings=holdings,
           gl_total=rt(f"{eur(tot_v - tot_cost, 0, True)}\u00a0({pct(tot_v / tot_cost - 1, 1)})", tot_v - tot_cost),
           holdings_note=f"USD naložbe preračunane s tečajem ECB {fx[0][0]}: 1 EUR = {num(FX_T, 4)} USD (prejšnji dan {num(FX_P, 4)}).",
           map=InlineImage(doc, mp, width=Mm(180)), indices=indices,
           fx_note=f"EUR/USD (ECB): {num(FX_T, 4)} ({pct(FX_T / FX_P - 1)} glede na prejšnji fiksing).",
           notes=' '.join(notes))
doc.render(ctx, autoescape=True)
base = os.path.join(A.out, f'Portfolio_Report_{TODAY}')
doc.save(base + '.docx')
subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', A.out, base + '.docx'], check=True, capture_output=True, timeout=180)
mv = max(holdings, key=lambda h: abs(h['_dp']))
print(json.dumps(dict(pdf=base + '.pdf', total=eur(tot_v), chg=eur(chg, sign=True), chg_pct=pct(chg_p),
    biggest=f"{mv['ticker']} {pct(mv['_dp'])} ({eur(mv['_de'], sign=True)})",
    over5=[f"{h['ticker']} {pct(h['_dp'])}" for h in holdings if abs(h['_dp']) > 0.05], missing=missing), ensure_ascii=False))
