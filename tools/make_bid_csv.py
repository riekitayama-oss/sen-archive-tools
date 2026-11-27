"""入札CSV作成: 過去売価中央値 × 素材完全一致 のみで入札上限を出し、元CSVと同じ形式(Shift-JIS)に入札価格を書き込む
usage: python3 bid.py RANKING.csv OUTDIR RATE MIN_PAST_SOLD orig1.csv [orig2.csv ...]   (orig は元のShift-JIS CSV)
"""
import csv, io, math, os, re, sys
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

RANK, OUTDIR, RATE, MIN_PAST, ORIGS = sys.argv[1], sys.argv[2], float(sys.argv[3]), int(sys.argv[4]), sys.argv[5:]
src = open(os.path.join(os.path.dirname(__file__), 'match_ranking.py'), encoding='utf-8').read()
exec(src[src.index('# ---------- brand'):src.index('def ref_rows')])  # BRANDS/MODELS/MATS/IS_BAG/NOT_BAG
ALIASES = {'Epsom': ['Epsom', 'Veau Epsom', 'Vaux Epsom'], 'Vogue Gulliver': ['Vogue Gulliver', 'Gulliver']}  # 同一素材の表記ゆれのみ
MODEL_ALIASES = {'City': ['City', 'The City'], 'Muse': ['Muse', 'Muse Bag']}

def num(s):
    s = (s or '').replace(',', '').replace('¥', '').strip()
    try: return float(s)
    except: return None

rows = list(csv.reader(open(RANK, encoding='utf-8-sig')))
sec = 0; hdr = None; full = {}
for x in rows:
    if x[0][:1] in '①②③': sec = x[0][:1]
    elif x[0] == '順位': hdr = x
    elif x[0].isdigit() and hdr and sec == '③':
        d = dict(zip(hdr, x)); d['rank'] = int(x[0]); full[(x[2].upper(), x[3], x[4])] = d

TAX = 1.10  # 落札価格・手数料にかかる消費税（還付を見込まない保守的計算）
def eco_fee(b):  # エコリング落札手数料（2025年4月改定・税別）
    for lim, fee in ((10000, 500), (49999, 1000), (99999, 2000), (499999, 4000), (999999, 8000)):
        if b <= lim: return fee
    return 10000

def max_bid(cost_cap):
    """(落札価格 + 落札手数料) × 消費税 ≦ 仕入れ上限 となる最大の入札額（千円単位）"""
    b = math.floor(cost_cap / TAX / 1000) * 1000
    while b > 0 and (b + eco_fee(b)) * TAX > cost_cap: b -= 1000
    return b

def split_raw(line):
    """引用符を保ったままフィールドに分割（元の書式をそのまま残すため）"""
    out, cur, q = [], '', False
    for ch in line:
        if ch == '"': q = not q
        if ch == ',' and not q: out.append(cur); cur = ''
        else: cur += ch
    out.append(cur); return out

def evaluate(name, price):
    """入札対象なら (dict) を返す。対象外は (None, 理由)"""
    if 'JAPAN ONLY' in name: return None, 'JAPAN ONLY'
    if NOT_BAG.search(name) or not IS_BAG.search(name): return None, 'バッグ以外'
    b = next((bb for p, bb in BRANDS if re.search(p, name, re.I)), None)
    if b not in MODELS: return None, 'ランキング外'
    model = next((m for p, m in MODELS[b] if re.search(p, name)), None)
    if not model: return None, 'ランキング外'
    if b == 'CHANEL' and model in ('Classic Flap', '2.55') and re.search('セカンド|クラッチ|ポーチ|リュック|バックパック|バニティ|ウエスト|ウェスト|トート', name): return None, 'ランキング外(フラップ以外の型)'
    if model == 'Tessuto/Nylon': return None, 'PRADAナイロン(型まとめ集計のため除外)'
    mat = next((m for p, m in MATS if re.search(p, name)), None)
    if not mat: return None, '素材不明'
    row = next((full[(bb, mo, a)] for bb in BRAND_ALIASES.get(b, [b]) for mo in MODEL_ALIASES.get(model, [model]) for a in ALIASES.get(mat, [mat]) if (bb, mo, a) in full), None)
    if not row: return None, f'{model}×{mat} のデータなし'
    med, n = num(row['過去売価中央値']), int(num(row['過去Sold']) or 0)
    if not med: return None, f'{model}×{mat} 過去売価なし'
    if n < MIN_PAST: return None, f'過去Sold {n}件(<{MIN_PAST})'
    bid = max_bid(med * RATE)
    if bid <= price: return None, '現在価格が入札上限以上'
    return {'ブランド': b, 'モデル': model, '素材': mat, 'ランキング順位': row['rank'], '過去Sold件数': n,
            '過去売価中央値': int(med), '過去売価25%点': int(num(row['過去売価25%点']) or 0), '仕入れ上限': int(med * RATE), '入札価格': bid}, ''

os.makedirs(OUTDIR, exist_ok=True)
allbids, reasons = [], Counter()
for path in ORIGS:
    raw = open(path, 'rb').read().decode('cp932')
    lines = raw.split('\n')
    t = re.search(r'2026-11-27_(\d\d)_(\d\d)', path); label = f'{t.group(1)}{t.group(2)}'
    nb = 0
    for i in range(2, len(lines)):
        if not lines[i].strip(): continue
        f = split_raw(lines[i])
        name = f[2].strip('"'); price = num(f[5]) or 0
        r, why = evaluate(name, price)
        if r is None:
            reasons[why] += 1; continue
        f[7] = f"入札上限{r['入札価格']}円"; lines[i] = ','.join(f); nb += 1
        allbids.append({'オークション': f'11/27 {t.group(1)}:{t.group(2)}', '出品順': int(f[3]), '商品ランク': f[4],
                        '現在/開始価格': int(price), '商品名': name, **r})
    out = os.path.join(OUTDIR, f'入札CSV_1127_{label}.csv')
    open(out, 'wb').write('\n'.join(lines).encode('cp932'))
    print(label, '入札', nb, '件 →', out)

allbids.sort(key=lambda d: (d['オークション'], d['出品順']))
wb = Workbook(); ws = wb.active; ws.title = '入札一覧'
cols = list(allbids[0].keys()); ws.append(cols)
for d in allbids: ws.append([d[c] for c in cols])
for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='404040'); c.alignment = Alignment(wrap_text=True)
for i, w in enumerate([13, 7, 7, 11, 70, 14, 16, 16, 9, 9, 12, 12, 11, 11], 1): ws.column_dimensions[get_column_letter(i)].width = w
for r in ws.iter_rows(min_row=2):
    for k in ('現在/開始価格', '過去売価中央値', '過去売価25%点', '仕入れ上限', '入札価格'): r[cols.index(k)].number_format = '#,##0'
ws.freeze_panes = 'F2'; ws.auto_filter.ref = ws.dimensions
ws2 = wb.create_sheet('ルール')
for line in [['販売相場', 'ランキングの「過去売価中央値」のみ使用（現行Sold価格中央値は使わない）'],
             ['素材', '商品名から判定した素材と、ランキングの ブランド×モデル×素材 が完全一致した行だけ使用（別素材の価格は流用しない）'],
             ['データ件数', f'過去Soldが{MIN_PAST}件未満の行は数字がぶれるため入札しない'],
             ['仕入れ上限', f'過去売価中央値 × {int(RATE*100)}% （販売手数料20% + 目標利益{100-20-int(RATE*100)}% を差し引いた残り）'],
             ['入札価格', '(入札価格 + エコリング落札手数料) × 消費税10% が仕入れ上限に収まる最大額（千円単位）。送料は未考慮'],
             ['落札手数料', '1〜10,000円:500円 / 〜49,999円:1,000円 / 〜99,999円:2,000円 / 〜499,999円:4,000円 / 〜999,999円:8,000円 / 100万円以上:10,000円（税別、2025年4月改定）'],
             ['消費税', '輸出免税で還付を受ける場合は、入札価格を約1.1倍まで上げられる'],
             ['除外', 'JAPAN ONLY / バッグ以外 / 素材不明 / PRADAナイロン / 現在価格が入札上限以上']]:
    ws2.append(line)
ws2.column_dimensions['A'].width = 12; ws2.column_dimensions['B'].width = 120
ws3 = wb.create_sheet('除外理由')
ws3.append(['理由', '件数'])
for k, v in reasons.most_common(): ws3.append([k, v])
ws3.column_dimensions['A'].width = 50
wb.save(os.path.join(OUTDIR, '入札一覧_1127.xlsx'))
print('合計', len(allbids))
for d in allbids: print(d['オークション'][-5:], d['出品順'], d['商品ランク'], d['現在/開始価格'], '→', d['入札価格'], d['モデル'], d['素材'], d['過去Sold件数'], d['商品名'][12:60])
print(reasons.most_common(12))
