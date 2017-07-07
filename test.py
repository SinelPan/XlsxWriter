import xlsxwriter

workbook = xlsxwriter.Workbook('conditional_format.xlsx')
worksheet1 = workbook.add_worksheet()

keys = [
    "title",
    "idx",
    "publish_date",
    "reach_hour",
    "reach_minute",
    "target_user",
    "int_page_read_user",
    "int_page_from_session_read_user",
    "int_page_from_feed_read_user",
    "int_page_from_friends_read_user",
    "int_page_from_hist_msg_read_user",
    "share_user",
    "feed_share_from_session_user",
    "feed_share_from_feed_user",
    "feed_share_from_other_user",
    "ori_page_read_user",
    "add_to_fav_user",
    "like_num",
    "comment_user",
    "mp_open_rate",
    "share_rate",
    "feed_open_rate",
    "like_rate",
    "comment_rate",
    "ori_open_rate"
]

percent_keys = [
    "mp_open_rate",
    "share_rate",
    "like_rate",
    "comment_rate",
    "ori_open_rate"
]

chinese_keys = [
    "标题",
    "图文位置",
    "日期",
    "小时",
    "分钟",
    "送达人数",
    "阅读量",
    "公众号会话阅读",
    "朋友圈阅读",
    "好友分享阅读",
    "历史页阅读",
    "分享量",
    "会话分享至朋友圈",
    "朋友圈二次传播",
    "其他渠道分享至朋友圈",
    "原文页阅读人数",
    "收藏数",
    "点赞数",
    "留言数",
    "打开率",
    "分享率",
    "朋友圈打开比例",
    "点赞率",
    "留言率",
    "原文打开率"
]

datas = [
    chinese_keys,
    [
        "有书创始人说","1","2017-05-06","23","35",20000,
        2000,1000,840,155,5,400,5,50,
        0,100,50,100,70,0.05,0.2,0.02,
        0.05,0.035,0.05
    ],
    [
        "有书创始人说222","1","2017-05-06","23","35",20000,
        2000,10000,1840,155,5,400,5,50,
        0,1020,504,1200,70,0.05,0.2,0.02,
        0.05,0.035,0.05
    ]
]
length = len(datas)+1
percent_fmt = workbook.add_format({'num_format': '0.00%'})
for column,data in enumerate(datas):
    for row, row_data in enumerate(data):
        # print(row,row_data)
        if (type(row_data) == int or type(row_data) == float) and row_data < 1:
            worksheet1.write(column, row,row_data,percent_fmt)
        else:
            worksheet1.write(column, row,row_data)

worksheet1.freeze_panes(1, 1)

worksheet1.autofilter('A2:')
worksheet1.autofilter('B2:')
worksheet1.autofilter('C2:')
worksheet1.autofilter('D2:')
worksheet1.autofilter('E2:')

worksheet1.write_comment('G1',"阅读量为阅读人数，非阅读次数。")
worksheet1.write_comment('L1',"分享量为分享人数，非分享次数。")
worksheet1.write_comment('M1',"从公众号会话（或好友或群聊会话）分享至朋友圈的人数。\n会话的定义：到底是公众号会话还是好友会话，微信给出该数据时未详细说明会话的定义，大家自行判定")
worksheet1.write_comment('N1',"从朋友圈阅读并分享到朋友圈")
worksheet1.write_comment('O1',"其他来源的渠道分享至朋友圈，如手机APP的内置分享到朋友圈功能")
worksheet1.write_comment('T1',"阅读量/送达人数\n综合反映粉丝的活跃度、标题的吸引力、推送的时间节点是否恰当。")
worksheet1.write_comment('U1',"分享量/阅读量\n反映文章的传播能力。")
worksheet1.write_comment('V1',"朋友圈阅读UV/分享量\n即：平均每一次分享可以带来多少个朋友圈阅读\n反映文章标题在朋友圈的吸引力")
worksheet1.write_comment('W1',"点赞数/阅读量\n反映粉丝对文章的喜爱程度、作者的认可程度。")
worksheet1.write_comment('X1',"留言数/阅读量")
worksheet1.write_comment('Y1',"原文页阅读人数/阅读量")

databar_column = ["G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y"]
for c in databar_column:
    worksheet1.conditional_format('%s2:%s%s'%(c,c,length),{'type': 'data_bar','bar_color':'#d5e9d6'})

# merge_format = workbook.add_format({'align': 'center'})
# worksheet1.merge_range('B3:D3', 'Merged Cells', merge_format)
workbook.close()
