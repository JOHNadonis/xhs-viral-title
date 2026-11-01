# -*- coding: utf-8 -*-
"""xhs-viral-title 工具链：筛选结构 / 装配 Excel / 校验。

子命令：
  list      列出全部结构
  show      查看单个结构详情
  pick      筛选结构，导出待填模板（给模型填文案）
  check     校验：字数、去重、空值
  assemble  把填好的文案装配成最终 Excel
"""
import argparse
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STRUCT_PATH = os.path.join(BASE, 'assets', 'structures.json')
TITLE_LIMIT = 20


def load_structures():
    with open(STRUCT_PATH, encoding='utf-8') as f:
        return json.load(f)


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def cmd_list(args):
    data = load_structures()
    rows = data['structures']
    print(f"共 {len(rows)} 个结构")
    print(f"{'rank':>4}  {'语气':<8}{'品类':<6}{'相似数':>6}  结构")
    for it in rows:
        print(f"{it['rank']:>4}  {it['tone']:<9}{it['fit']:<7}{it['similar_count']:>5}  {it['structure']}")


def cmd_pick(args):
    data = load_structures()
    rows = data['structures']
    sel = []
    for it in rows:
        if args.category:
            fits = it['fit']
            if fits != '*' and args.category.upper() not in fits.upper():
                continue
        if args.tone and it['tone'] not in args.tone.split(','):
            continue
        if args.rank:
            want = {int(x) for x in args.rank.split(',')}
            if it['rank'] not in want:
                continue
        if args.min_likes and (it['likes_median'] or 0) < args.min_likes:
            continue
        sel.append(it)

    if args.sort == 'likes':
        sel.sort(key=lambda x: (x['likes_median'] or 0), reverse=True)
    else:
        sel.sort(key=lambda x: x['rank'])
    if args.limit:
        sel = sel[:args.limit]

    payload = {
        'instruction': '为每个 structure 的 slots 填入产品信息，生成 per_row 条互不相同的标题，每条不超过 20 字',
        'product': args.product or '',
        'per_row': args.per_row,
        'structures': [
            {
                'rank': it['rank'],
                'structure': it['structure'],
                'slots': it['slots'],
                'examples': it['examples'],
                'tone': it['tone'],
                'likes_median': it['likes_median'],
            } for it in sel
        ],
    }
    out = json.dumps(payload, ensure_ascii=False, indent=1)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(out)
        print(f'已导出 {len(sel)} 个结构 -> {args.out}')
    else:
        print(out)


def cmd_show(args):
    data = load_structures()
    for it in data['structures']:
        if it['rank'] == args.rank:
            print(json.dumps(it, ensure_ascii=False, indent=1))
            return
    print(f'未找到 rank={args.rank}', file=sys.stderr)
    sys.exit(1)


def cmd_check(args):
    """校验标题集合：字数、重复、空值。不做内容审查，只报事实。"""
    titles = load_titles(args.titles)
    problems = []
    seen = {}
    for i, t in enumerate(titles):
        s = norm(t)
        if not s:
            problems.append((i, 'EMPTY', t))
            continue
        if len(s) > TITLE_LIMIT:
            problems.append((i, f'LONG({len(s)})', t))
        if s in seen:
            problems.append((i, f'DUP with #{seen[s]}', t))
        else:
            seen[s] = i

    print(f'共 {len(titles)} 条，去重后 {len(seen)} 条，问题 {len(problems)} 条')
    for p in problems:
        print('  ', p[1], '|', p[2])
    if problems and args.strict:
        sys.exit(2)
    return problems


def load_titles(path):
    with open(path, encoding='utf-8') as f:
        raw = f.read().strip()
    if not raw:
        return []
    # 优先整体按 JSON 解析，兼容三种形态：
    #   ["标题", ...]
    #   [{"rank":1,"titles":[...]}, ...]
    #   {"rows":[{"rank":1,"titles":[...]}, ...]}
    try:
        obj = json.loads(raw)
    except json.JSONDecodeError:
        obj = None
    if obj is not None:
        return _flatten(obj)
    out = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            o = json.loads(line)
        except json.JSONDecodeError:
            out.append(line)
            continue
        out.extend(_flatten(o))
    return out


def _flatten(obj):
    out = []
    if isinstance(obj, str):
        return [obj]
    if isinstance(obj, dict):
        if 'rows' in obj:
            return _flatten(obj['rows'])
        if 'titles' in obj:
            return list(obj['titles'])
        return out
    if isinstance(obj, list):
        for item in obj:
            out.extend(_flatten(item))
    return out


def cmd_assemble(args):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    data = load_structures()
    by_rank = {it['rank']: it for it in data['structures']}
    with open(args.filled, encoding='utf-8') as f:
        filled = json.load(f)
    if isinstance(filled, dict):
        filled = filled.get('rows', [])

    wb = Workbook()
    ws = wb.active
    ws.title = args.sheet or '爆款标题文案'

    head_fill = PatternFill('solid', fgColor='DCEBF7')
    thin = Side(style='thin', color='BFBFBF')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)

    per = max((len(r.get('titles', [])) for r in filled), default=3)
    per = max(per, 1)
    headers = ['序号', '原表排名', '爆款句式结构', '相似标题数', '点赞中位数'] + \
              [f'文案{i + 1}' for i in range(per)]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, size=11)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bd

    seq = 0
    for r in filled:
        rank = int(r['rank'])
        it = by_rank.get(rank)
        if not it:
            print(f'跳过未知 rank={rank}', file=sys.stderr)
            continue
        seq += 1
        titles = list(r.get('titles', []))
        titles += [''] * (per - len(titles))
        ws.append([seq, rank, it['structure'], it['similar_count'],
                   it['likes_median']] + titles)

    widths = [6, 8, 46, 12, 12] + [52] * per
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = bd
            cell.alignment = Alignment(
                vertical='center', wrap_text=True,
                horizontal='center' if cell.column <= 5 else 'left')
        ws.row_dimensions[row[0].row].height = 34
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{ws.max_row}'

    # 产品信息卡
    if args.product_card and os.path.exists(args.product_card):
        with open(args.product_card, encoding='utf-8') as f:
            card = json.load(f)
        ws2 = wb.create_sheet('产品信息卡')
        for c in range(1, 3):
            cell = ws2.cell(row=1, column=c, value=['项目', '内容'][c - 1])
            cell.font = Font(bold=True, size=11)
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for k, v in card.items():
            ws2.append([k, v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)])
        ws2.column_dimensions['A'].width = 22
        ws2.column_dimensions['B'].width = 100
        for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=2):
            for cell in row:
                cell.border = bd
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            ws2.row_dimensions[row[0].row].height = 40

    wb.save(args.out)
    print(f'已生成 {seq} 行 x {per} 条 = {seq * per} 条 -> {args.out}')


def main():
    ap = argparse.ArgumentParser(prog='xvt', description='小红书爆款标题工具链')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('list', help='列出全部结构')
    p.set_defaults(func=cmd_list)

    p = sub.add_parser('show', help='查看单个结构详情')
    p.add_argument('--rank', type=int, required=True)
    p.set_defaults(func=cmd_show)

    p = sub.add_parser('pick', help='筛选结构并导出待填模板')
    p.add_argument('--category', help='品类代码 F/U/H/G/D/E/K')
    p.add_argument('--tone', help='语气，逗号分隔 price,hype,story,pain,soft,social')
    p.add_argument('--rank', help='指定 rank，逗号分隔')
    p.add_argument('--min-likes', type=float, default=0)
    p.add_argument('--limit', type=int, default=0)
    p.add_argument('--per-row', type=int, default=3, help='每个结构生成几条（默认3）')
    p.add_argument('--sort', choices=['rank', 'likes'], default='rank')
    p.add_argument('--product', help='产品名，写入模板')
    p.add_argument('--out', help='输出 json 路径，不给则打印')
    p.set_defaults(func=cmd_pick)

    p = sub.add_parser('check', help='校验字数/重复/空值')
    p.add_argument('--titles', required=True)
    p.add_argument('--strict', action='store_true', help='有问题时返回非零退出码')
    p.set_defaults(func=cmd_check)

    p = sub.add_parser('assemble', help='装配 Excel')
    p.add_argument('--filled', required=True, help='填好的 JSON: [{rank, titles[]}]')
    p.add_argument('--out', required=True)
    p.add_argument('--product-card', help='产品信息卡 JSON 路径，会作为第二个 sheet')
    p.add_argument('--sheet', help='主 sheet 名')
    p.set_defaults(func=cmd_assemble)

    args = ap.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
