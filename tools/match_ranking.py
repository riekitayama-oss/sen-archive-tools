"""複数オークションCSV × 統合ランキング 照合（全ブランド対応）
usage: python3 match2.py RANKING.csv OUT.xlsx auc1.csv [auc2.csv ...]  (auc は UTF-8 変換済み)
"""
import csv, re, sys
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

RANK, OUT, AUCS = sys.argv[1], sys.argv[2], sys.argv[3:]

def num(s):
    s = (s or '').replace(',', '').replace('¥', '').strip()
    try: return float(s)
    except: return None

# ---------- ranking ----------
rows = list(csv.reader(open(RANK, encoding='utf-8-sig')))
sec = 0; hdr = None; top12 = set(); full = {}
for x in rows:
    if x[0][:1] in '①②③': sec = x[0][:1]
    elif x[0] == '順位': hdr = x
    elif x[0].isdigit() and hdr:
        b = x[2].upper()
        if sec in '①②': top12.add((b, x[3]))
        if sec == '③':
            d = dict(zip(hdr, x)); d['rank'] = int(x[0]); full[(b, x[3], x[4])] = d

# ---------- brand / model / material dictionaries ----------
BRANDS = [('HERMES|エルメス', 'HERMES'), ('CHANEL|シャネル', 'CHANEL'), ('LOUIS|ヴィトン', 'LOUIS VUITTON'),
          ('GUCCI|グッチ', 'GUCCI'), ('PRADA|プラダ', 'PRADA'), ('CELINE|セリーヌ', 'CELINE'), ('FENDI|フェンディ', 'FENDI'),
          ('DIOR|ディオール', 'DIOR'), ('BALENCIAGA|バレンシアガ', 'BALENCIAGA'), ('BOTTEGA|ボッテガ', 'BOTTEGA VENETA'), ('LOEWE|ロエベ', 'LOEWE'),
          ('MIU ?MIU|ミュウミュウ', 'MIU MIU'), ('SAINT ?LAURENT|サンローラン|YSL', 'SAINT LAURENT')]
BRAND_ALIASES = {'SAINT LAURENT': ['SAINT LAURENT', 'YVES SAINT LAURENT']}
MODELS = {
    'HERMES': [('バーキン', 'Birkin'), ('ケリー', 'Kelly'), ('エブリン|エヴリン', 'Evelyne'), ('ピコタン', 'Picotin'),
               ('ガーデンパーティ', 'Garden Party'), ('ボリード', 'Bolide'), ('コンスタンス', 'Constance'), ('プリュム', 'Plume')],
    'CHANEL': [('ダイアナ', 'Diana'), ('ボーイ', 'Boy'), ('2\\.55', '2.55'), ('ココハンドル', 'Coco Handle'),
               ('マトラッセ|クラシック|タイムレス|フラップ', 'Classic Flap')],
    'LOUIS VUITTON': [('スピーディ', 'Speedy'), ('ネヴァーフル|ネバーフル', 'Neverfull'), ('パピヨン', 'Papillon'),
                      ('ポシェット[・ ]?アクセソワール', 'Pochette Accessoire'), ('ノエ', 'Noé'), ('アルマ', 'Alma'), ('キーポル', 'Keepall')],
    'FENDI': [('バゲット|マンマ', 'Baguette'), ('ピーカブー', 'Peekaboo'), ('ズッカ.*ショルダー|ズッカ.*ワンショル', 'FF Zucca Pattern Shoulder Bag')],
    'DIOR': [('サドル', 'Saddle'), ('レディ ?ディオール', 'Lady Dior'), ('ブック ?トート', 'Book Tote'), ('マリス', 'Malice'),
             ('トロッター.*ショルダー', 'Trotter Pattern Shoulder Bag in'), ('トロッター.*(ハンド|トート)', 'Trotter Pattern Handbag')],
    'BALENCIAGA': [('シティ', 'City'), ('ファースト', 'The First')],
    'BOTTEGA VENETA': [('イントレ', 'Intrecciato')],
    'LOEWE': [('アマソナ', 'Amazona'), ('パズル', 'Puzzle Small 2-way Bag'), ('ハンモック', 'Hammock Small 2-Way Bag'),
              ('フラメンコ', 'Anagram Flamenco Shoulder Bag')],
    'MIU MIU': [('マドラス', 'Madras 2-way bag,'), ('ヴィッテロ|VITELLO', 'Vitello')],
    'SAINT LAURENT': [('ミューズ ?トゥ|ミューズ2', 'Muse Two'), ('ミューズ', 'Muse'), ('カバ ?シック', 'Cabas Chic 2-way Handbag')],
    'GUCCI': [('ジャッキー', 'Jackie'), ('ホースビット', 'Horsebit'), ('ソーホー', 'Soho')],
    'PRADA': [('ミニ ?ボニー', 'Mini Bonnie'), ('ボニー', 'Bonnie'), ('テスート|ナイロン', 'Tessuto/Nylon')],
    'CELINE': [('ラゲージ', 'Luggage'), ('マカダム', 'Macadam')],
}
MATS = [('トゴ', 'Togo'), ('トリヨン|クレマンス', 'Taurillon Clemence'), ('ボックスカーフ|BOX', 'Calfskin'),
        ('エプソン|エプシ', 'Epsom'), ('スイフト', 'Swift'), ('ガリバー', 'Vogue Gulliver'), ('ネゴンダ', 'Negonda'), ('バレニア', 'Barenia'),
        ('アリゲーター|クロコ|リザード|ヒマラヤ|ポロサス|オーストリッチ|パイソン', 'Exotic Leather'),
        ('(?<!グ)ラム', 'Lambskin'), ('キャビア', 'Caviar Skin'), ('ツイード', 'Tweed'), ('デニム', 'Denim'), ('ヴェルニ|エナメル|パテント', 'Patent Leather'),
        ('ベロア|ベルベット', 'Velvet'), ('コットン', 'Cotton'), ('サテン', 'Satin'), ('スエード', 'Suede'), ('ファー|ハラコ', 'Haircalf/Fur'),
        ('モノグラム|ダミエ|キャンバス|トワル|ズッカ|ズッキーノ|GG|PVC|トロッター|オブリーク|ジャガード', 'Canvas'), ('ナイロン|テスート', 'Nylon'),
        ('エピ|ノマド|タイガ|カーフ|レザー|革', 'Leather')]
MODEL_ALIASES = {'City': ['City', 'The City'], 'Muse': ['Muse', 'Muse Bag']}
ALIASES = {'Epsom': ['Epsom', 'Veau Epsom', 'Vaux Epsom'], 'Vogue Gulliver': ['Vogue Gulliver', 'Gulliver'],
           'Canvas': ['Canvas', 'Toile H'], 'Calfskin': ['Calfskin', 'Leather']}
IS_BAG = re.compile('バッグ|バック|トート|ボストン|リュック|ショルダー|ハンド|クラッチ|セカンド|アクセソワール')
NOT_BAG = re.compile('財布|ウォレット|ベルト|ポーチ|ミニミニ|レインカバー|空箱|ネックレス|ブレスレット|バングル|ピアス|イヤリング|ブローチ|リング|手袋|'
                     'チャーム|キーホルダー|キーケース|カードケース|ケリーダンス|アッカド|トゥーゴー|トゥゴー|コンパクト|ストラップ')

def ref_rows(b, model):
    return sorted([d for (bb, mo, _), d in full.items() if bb == b and mo == model], key=lambda d: d['rank'])

def market(d):
    if not d: return None, ''
    v = num(d.get('現行Sold価格中央値'))
    if v: return v, '現行Sold'
    v = num(d.get('過去売価中央値'))
    return (v, '過去Sold') if v else (None, '')

def tier(b, model, best):
    if (b, model) in top12: return 'S'
    if best and best['rank'] <= 60: return 'A'
    if best and best['rank'] <= 100: return 'B'
    return 'C'

out = []
for path in AUCS:
    for x in list(csv.reader(open(path, encoding='utf-8')))[2:]:
        name = x[2]
        if NOT_BAG.search(name) or not IS_BAG.search(name): continue
        b = next((bb for p, bb in BRANDS if re.search(p, name)), None)
        if b not in MODELS: continue
        model = next((m for p, m in MODELS[b] if re.search(p, name)), None)
        if not model: continue
        cands = ref_rows(b, model)
        if not cands: continue
        best = cands[0]
        mat = next((m for p, m in MATS if re.search(p, name)), None)
        mr = None
        if mat:
            for a in ALIASES.get(mat, [mat]):
                if (b, model, a) in full: mr = full[(b, model, a)]; break
        if b == 'CHANEL' and model == 'Classic Flap' and re.search('セカンド|クラッチ|ポーチ', name): continue
        # 価格の根拠: 素材別の行 → 同モデルで販売データ(実測+過去Sold)が最も多い行
        def nd(d): return (num(d['実測Sold総数']) or 0) + (num(d['過去Sold']) or 0)
        pool = sorted([c for c in cands if market(c)[0]], key=lambda d: (-nd(d), d['rank']))
        ref = mr if (mr and market(mr)[0]) else (pool[0] if pool else None)
        mk, src = market(ref)
        price = num(x[5]) or 0
        ratio = price / mk if mk and price else None
        note = []
        if 'JAPAN ONLY' in name: note.append('JAPAN ONLY: 海外販売不可')
        if b == 'HERMES' and re.search('ミニケリー|ポシェット|マイクロ|ミクロ', name): note.append('ミニ系: 相場はサイズ違いの参考値')
        out.append({
            'tier': tier(b, model, best), '終了': x[8][5:16], '出品順': int(x[3]), '商品ランク': x[4], '現在価格': int(price),
            '商品名': name, 'ブランド': b, 'モデル': model, '素材': mat or '',
            'モデル最上位順位': best['rank'], '素材別順位': mr['rank'] if mr else '',
            '15日ST率(素材別)': mr['15日Sell-through率'] if mr else '', '実測Sold(素材別)': mr['実測Sold総数'] if mr else '',
            '売却日数中央値': (ref or {}).get('売却日数中央値', ''), '相場中央値': int(mk) if mk else '',
            '相場の根拠': f"{ref['モデル']}/{ref['素材']} {src}" if ref else '',
            '現在価格÷相場': round(ratio, 2) if ratio is not None else '', '注意': ' / '.join(note),
        })

order = {'S': 0, 'A': 1, 'B': 2, 'C': 3}
out.sort(key=lambda d: (order[d['tier']], 'JAPAN' in d['注意'], d['現在価格÷相場'] if d['現在価格÷相場'] != '' else 9, d['終了'], d['出品順']))

wb = Workbook(); ws = wb.active; ws.title = '仕入れ候補'
cols = list(out[0].keys())
ws.append(['優先度'] + cols[1:])
fill = {'S': 'FCE4D6', 'A': 'FFF2CC', 'B': 'E2EFDA', 'C': 'F2F2F2'}
for d in out:
    ws.append([d[c] for c in cols])
    for c in ws[ws.max_row]: c.fill = PatternFill('solid', fgColor=fill[d['tier']])
for c in ws[1]:
    c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='404040'); c.alignment = Alignment(wrap_text=True, vertical='center')
widths = [6, 12, 7, 7, 11, 70, 14, 16, 16, 9, 9, 10, 9, 9, 12, 34, 9, 30]
for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
ci = {c: i for i, c in enumerate(cols)}
for r in ws.iter_rows(min_row=2):
    r[ci['現在価格']].number_format = '#,##0'; r[ci['相場中央値']].number_format = '#,##0'; r[ci['現在価格÷相場']].number_format = '0%'
ws.freeze_panes = 'C2'; ws.auto_filter.ref = ws.dimensions

ws2 = wb.create_sheet('見方')
for line in [
    ['優先度S', 'ランキング①(高回転×継続需要) または ②(販売量×継続需要) に入っているモデル'],
    ['優先度A', '統合ランキング60位以内のモデル'],
    ['優先度B', '統合ランキング61〜100位'],
    ['優先度C', 'それ以下（ランキングには載っているが売却実績が少ない）'],
    ['現在価格', 'CSV出力時点の現在価格。終了までにさらに上がる'],
    ['現在価格÷相場', '低いほど利幅の余地あり（手数料・送料・落札手数料は未考慮）'],
    ['PRADAナイロン', 'ランキングのTessuto/Nylonは型をまとめた中央値。型番・状態で相場差が大きいので要個別確認'],
    ['相場中央値', '素材別の行に価格があればそれ、なければ同モデルで販売データ（実測Sold＋過去Sold）が最も多い行の現行Sold中央値（なければ過去売価中央値）'],
    ['対象外', 'バッグ以外（財布・小物・アクセ）と、ランキングに載っていないモデルは除外'],
]:
    ws2.append(line)
ws2.column_dimensions['A'].width = 16; ws2.column_dimensions['B'].width = 110
wb.save(OUT)

print(len(out), Counter(d['tier'] for d in out), Counter((d['終了'][-5:], d['ブランド']) for d in out))
for d in out:
    if d['tier'] in 'SAB' and 'JAPAN' not in d['注意'] and d['現在価格÷相場'] != '' and d['現在価格÷相場'] <= 0.5:
        print(d['tier'], d['終了'][-5:], d['出品順'], d['商品ランク'], d['現在価格'], d['現在価格÷相場'], d['相場中央値'], d['モデル'], d['素材'], d['商品名'][12:75])
