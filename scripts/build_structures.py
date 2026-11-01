# -*- coding: utf-8 -*-
"""从原始排行榜 xlsx 生成 structures.json，并注入人工审校的槽位/品类元数据。"""
import json, os, re
import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, 'assets', '爆款标题排行榜1026终版.xlsx')
OUT = os.path.join(BASE, 'assets', 'structures.json')

# ---- 人工审校元数据：rank -> (槽位名列表, 品类倾向, 语气, 风险标记) ----
# 品类倾向: F=服饰 U=美妆个护 H=家居百货 G=食品饮品 D=数码 E=母婴 K=知识课程 *=通用
# 语气: price=价格冲击 hype=夸张赞叹 story=故事自述 pain=痛点共情 soft=温柔治愈 social=社交背书
# 风险: 空=可直接用; "story"=需真实故事; "claim"=易触医疗/极限词红线; "seller"=需确认卖家身份
META = {
1:(["寻找经历","对象"],"FUHGDE","story", ""),
2:(["范围","评价","对象"],"*","hype",""),
3:(["对象","评价"],"*","hype",""),
4:(["卖点","强调"],"*","price",""),
5:(["反向优点"],"*","soft",""),
6:(["新版本"],"FUD","hype","seller"),
7:(["痛点场景","问题"],"*","pain",""),
8:(["价格","规格"],"*","price","seller"),
9:(["对象","场景"],"*","story",""),
10:(["特征对象"],"*","soft",""),
11:(["场景","对象"],"*","hype",""),
12:(["平台","数量"],"*","social","seller"),
13:(["礼物"],"*","soft",""),
14:(["对象","正面体验"],"FUHGDE","soft",""),
15:(["对象","选择结论"],"*","pain",""),
16:(["季节","量词","对象"],"FU","soft",""),
17:(["次数","表现"],"FUHGD","social",""),
18:(["对象"],"*","hype",""),
19:(["原预期","实际反转"],"*","price",""),
20:(["群体","数量"],"FUHGD","social",""),
21:(["场景","对象","认可"],"*","hype",""),
22:(["痛点"],"*","pain",""),
23:(["评价","对象"],"*","hype",""),
24:(["对象"],"*","pain",""),
25:(["评价"],"*","hype","claim"),
26:(["对象","依据"],"*","pain",""),
27:(["季节","对象"],"*","soft",""),
28:(["载体","意象"],"F","soft",""),
29:(["家人","规格"],"*","story","seller"),
30:(["次数","对象"],"*","story","seller"),
31:(["卖点"],"*","hype",""),
32:(["季节","好处"],"GH","soft","claim"),
33:(["对象","人群","次数"],"*","social",""),
34:(["对象","意外效果"],"*","story","claim"),
35:(["改版点"],"*","hype","seller"),
36:(["对象"],"*","hype",""),
37:(["需求人群"],"*","pain",""),
38:(["生活方式"],"*","soft",""),
39:(["价格"],"*","price",""),
40:(["对象"],"*","story","seller"),
41:(["对象"],"*","story","seller"),
42:(["人群","金额"],"*","price",""),
43:(["卖点"],"*","hype",""),
44:(["经营铺垫"],"*","story","seller"),
45:(["对象","场景"],"*","hype",""),
46:(["价格"],"*","price",""),
47:(["新规格"],"*","soft","seller"),
48:(["人群","行动"],"*","pain",""),
49:(["价格","限量"],"*","price","seller"),
50:(["人群称呼"],"*","price","seller"),
51:(["对象","体验"],"*","soft",""),
52:(["高价选项","替代选项"],"*","price",""),
53:(["性格镜像"],"*","story","claim"),
54:(["对象","夸张认可"],"*","story","seller"),
55:(["对象"],"*","soft",""),
56:(["地区","特色"],"GH","hype",""),
57:(["方案1","方案2","回应"],"*","price","story"),
58:(["高品质对象","替代品"],"FUHGDE","soft",""),
59:(["对象"],"*","hype",""),
60:(["对标品牌","卖点"],"*","price",""),
61:(["动作"],"*","hype",""),
62:(["旧风格","新风格"],"FU","soft",""),
63:(["对象"],"*","hype",""),
64:(["投入","对象"],"*","story","seller"),
65:(["对象"],"*","hype",""),
66:(["原价","新价"],"*","price",""),
67:(["旧认知","新认知"],"*","pain",""),
68:(["直供价"],"GH","price","seller"),
69:(["对象"],"*","hype",""),
70:(["平台"],"*","social",""),
71:(["对标品牌","价格"],"*","price",""),
72:(["价格","昂贵档次"],"*","price",""),
73:(["对象","评价"],"*","soft",""),
74:(["身材数据","反馈"],"F","pain",""),
75:(["变化","人群状态"],"*","social",""),
76:(["现象"],"*","hype",""),
77:(["收礼对象","表现"],"*","soft",""),
78:(["行动","正面结果"],"GH","pain","claim"),
79:(["价格","时长"],"*","price",""),
80:(["对象"],"*","soft",""),
81:(["批次"],"*","story","seller"),
82:(["权益"],"*","price","seller"),
83:(["金额"],"*","price",""),
84:(["说话者","对象"],"*","social",""),
85:(["低价或配置"],"*","price",""),
86:(["客户返图"],"*","story","seller"),
87:(["特征"],"FUHGDE","hype",""),
88:(["人群"],"*","pain",""),
89:(["风格/配色"],"*","soft",""),
90:(["对象","周期","次数"],"*","hype",""),
91:(["动作","结果"],"*","price",""),
92:(["功能绰号"],"*","hype","claim"),
93:(["店铺"],"*","social","seller"),
94:(["对象"],"*","hype",""),
95:(["习惯","变化"],"GH","pain","claim"),
96:(["原价","新价"],"*","price",""),
97:(["价格","规格","权益"],"*","price",""),
98:(["礼物"],"*","soft",""),
99:(["年龄","行动"],"*","pain","claim"),
100:(["家人","产品"],"GH","story","seller"),
101:(["评价","对象"],"*","hype",""),
102:(["对象"],"*","soft",""),
103:(["对象"],"*","story",""),
104:(["想法"],"*","pain",""),
105:(["对象","理由"],"*","soft",""),
106:(["使用时长"],"*","soft",""),
107:(["痛点","人群"],"*","pain",""),
108:(["对比经历"],"*","social",""),
109:(["产品"],"GH","story","seller"),
110:(["场景A","场景B"],"FU","soft",""),
111:(["权益"],"*","price",""),
112:(["低价","额外权益"],"*","price",""),
113:(["对象"],"*","soft",""),
114:(["普通状态","惊喜表现"],"*","soft",""),
115:(["产品"],"*","story","seller"),
116:(["对照选项","表现"],"*","hype","claim"),
}

# 原表里部分结构含「/」并列的口语，做细微正字化
NORMALIZE = {
    2: "（个人认为/目前为止）[范围]最[评价]的[对象]（没有之一）",
    5: "唯一的缺点：[反向优点/限制/遗憾]",
    30: "改了又改[次数]，终于做出[对象]",
    40: "终于做出[对象]",
    62: "放弃[旧风格]，要大大方方[新风格]",
}

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['句式排行']
items = []
for r in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=9, values_only=True):
    rank, structure, cnt, cat, authors, likes = r[0], r[1], r[2], r[3], r[4], r[5]
    if rank is None or structure is None:
        continue
    rank = int(rank)
    slots, fit, tone, _ = META.get(rank, ([], "*", "hype", ""))
    items.append({
        "rank": rank,
        "structure": NORMALIZE.get(rank, structure),
        "examples": [r[6], r[7], r[8]],
        "similar_count": cnt,
        "category_count": cat,
        "author_count": authors,
        "likes_median": likes,
        "slots": slots,
        "fit": fit,
        "tone": tone,
    })

items.sort(key=lambda x: x["rank"])
data = {
    "source": "爆款标题排行榜1026终版.xlsx",
    "total_structures": len(items),
    "fit_legend": {"F":"服饰","U":"美妆个护","H":"家居百货","G":"食品饮品","D":"数码","E":"母婴","K":"知识课程","*":"通用"},
    "tone_legend": {"price":"价格冲击","hype":"夸张赞叹","story":"故事自述","pain":"痛点共情","soft":"温柔治愈","social":"社交背书"},
    "structures": items,
}
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
print('structures:', len(items))
miss = [i['rank'] for i in items if not i['slots']]
print('缺槽位标注:', miss)
