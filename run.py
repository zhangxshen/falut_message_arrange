import pandas as pd
import re
from datetime import datetime


def nb_kx_jf(path_target):
    print("正在进行故障信息自动整理，请稍候……")

    # 源数据文件前两行以及后两列为多余数据，且包含某些问号字符，先将源文件进行处理
    text = pd.read_excel(path_target)
    if text.columns[0] != '发送时间':  # 先判断是否为未处理过的源数据
        text = text.drop([0, 1])  # 去除前两行数据
        text = text.drop(['Unnamed: 2', 'Unnamed: 3'], axis=1)  # 去除后两列数据
        text.columns = ['发送时间', '短信内容']
        text['短信内容'].replace(' ）', ' ', regex=True, inplace=True)
        text['短信内容'].replace('？', ' ', regex=True, inplace=True)  # 把快讯中的问号字符去掉

    # 将非故障快讯单独保存
    other_text = text[text['短信内容'].str.contains('请审核|互联互通|】测试|演练|领导值班|突发事件|地震')]

    # 去除非故障快讯
    text = text[~ text['短信内容'].str.contains('请审核|互联互通|】测试|演练|领导值班|突发事件|地震')]

    # 值班长列表
    monitor = ['刘浩', '张海鹏', '冯春雨', '陈俊鑫', '张振斌', '梁国贤', '梅坚', '郭润海', '冯轶颖', '周永德', '王华', '周华造']

    # 地市列表
    all_city = ['广州', '深圳', '佛山', '东莞', '汕头', '珠海', '惠州', '中山', '江门', '湛江', '韶关', '河源', '梅州', '汕尾',
                '阳江', '茂名', '肇庆', '清远', '潮州', '揭阳', '云浮']

    # 内部通报数据
    nb_result = pd.DataFrame(columns=['序号', '月份', '是否已恢复', '负责单位', '网络类型', '所属专业', '设备类型', '影响业务种类', '维护单位（统计主动监控）',
                                      '故障标题', '故障发生时间', '故障报障专业室时间', '故障处理过程', '故障销除时间', '故障历时（分钟）', '故障处理历时（分钟）',
                                      '业务影响历时（分钟）', '故障定位历时（分钟）', '专业定位时长（分钟）', '影响业务范围', '投诉用户数', '故障原因分析',
                                      '故障原因分类', '故障原因细分', '上报集团', '上报管局', '故障级别', '采用应急容灾情况', '采用应急方案情况', '是否有告警',
                                      '是否省监控主动发现', '网管告警缺漏的原因', '故障网元', '设备厂家', '故障发生地市', '故障影响地市', '客服预警级别',
                                      '区域客服信息发布内容', '实际影响用户数', '故障后评估', '故障剖析情况', '值班长', '故障通报短信', '首条故障短信发布时间',
                                      '故障存在问题（资源、处理过程等）', '问题闭环情况'])

    # 快讯数据
    kx_result = pd.DataFrame(columns=['序号', '月份', '是否已恢复', '负责单位', '网络类型', '所属专业', '设备类型', '影响业务种类', '维护单位（统计主动监控）',
                                      '故障标题', '故障发生时间', '故障销除时间', '故障处理时长', '业务影响历时', '故障处理过程',
                                      '投诉用户数', '故障原因', '故障原因分类', '故障原因细分', '是否有告警', '是否省监控主动发现',
                                      '网管告警缺漏的原因', '故障网元', '设备厂家', '故障发生地市', '故障影响地市', '值班长', '首条故障短信发布时间',
                                      '故障存在问题（资源、处理过程等）'])

    # 重点机房及机楼数据
    jf_result = pd.DataFrame(columns=['序号', '月份', '是否已恢复', '负责单位', '网络类型', '所属专业', '设备类型', '影响业务种类',
                                      '故障标题', '故障发生时间', '故障销除时间', '通报信息', '类型', '故障原因分析',
                                      '故障原因分类', '故障原因细分', '是否有告警', '是否省监控主动发现', '网管告警缺漏的原因',
                                      '故障网元', '设备厂家', '故障发生地市', '故障影响地市', '值班长', '首条故障短信发布时间', '故障存在问题（资源、处理过程等）'])

    # 管控类信息字段
    guankong_result = pd.DataFrame(columns=['序号', '月份', '是否已恢复', '负责单位', '网络类型', '所属专业', '设备类型', '影响业务种类',
                                            '维护单位（统计主动监控）', '故障标题', '故障发生时间', '故障销除时间', '故障处理时长',
                                            '业务影响历时', '通报信息', '投诉用户数', '故障原因分析', '故障原因分类', '故障原因细分',
                                            '是否有告警', '是否省监控主动发现', '网管告警缺漏的原因', '故障网元', '设备厂家',
                                            '故障发生地市', '故障影响地市', '值班长', '首条故障短信发布时间', '故障存在问题（资源、处理过程等）'
                                            ])

    # 外部门故障字段
    waibu_result = pd.DataFrame(columns=['网络类型', '所属专业', '设备类型', '影响业务种类', '故障标题', '故障发生时间', '故障报障专业室时间',
                                         '故障处理过程', '故障销除时间', '故障历时（分钟）', '故障处理历时（分钟）', '业务影响历时（分钟）',
                                         '故障定位历时（分钟）', '影响业务范围', '投诉用户数', '故障原因分析', '故障原因分类', '故障原因细分',
                                         '上报集团', '上报管局', '故障级别', '采用应急容灾情况', '采用应急方案情况',
                                         '是否有告警', '是否省监控主动发现', '网管告警缺漏的原因', '故障网元', '设备厂家', '故障发生地市', '故障影响地市',
                                         '客服预警级别', '区域客服信息发布内容', '实际影响用户数', '故障后评估', '故障剖析情况', '值班长', '是否网管网存在问题'
                                                                                                      '故障通报短信'])

    # 恢复和阶段快讯
    hf = pd.DataFrame(columns=['恢复短信内容', '标题'])
    jd = pd.DataFrame(columns=['阶段短信内容', '标题'])

    def get_kx(kx_df):
        tmp_kx_title = re.search(r'【.*】', kx_df)
        if "【新" in tmp_kx_title.group():
            kx_title = tmp_kx_title.group().split('|')[3].replace('】', '').replace('故障', '').replace('事件', '')
        elif "【新" in tmp_kx_title.group() and "专线中断" in tmp_kx_title.group():
            kx_title = tmp_kx_title.group().split('|')[3].replace('】', '')
        elif "专线中断" in tmp_kx_title.group() or "故障】" in tmp_kx_title.group():
            kx_title = tmp_kx_title.group().split('|')[1].replace('】', '')
        else:
            kx_title = tmp_kx_title.group().split('|')[1].replace('】', '').replace('故障', '').replace('事件', '')
        return kx_title

    def get_hf(hf_df):
        tmp_hf_title = re.search(r'【.*】', hf_df)
        if "【销" in tmp_hf_title.group():
            hf_title = tmp_hf_title.group().split('|')[3].replace('】', '')
        elif "期）" in tmp_hf_title.group():
            hf_title = tmp_hf_title.group().split('|')[1].replace('已恢复】', '').replace('事件', '')
        elif "重点管控" in tmp_hf_title.group():
            hf_title = tmp_hf_title.group().split('|')[1].replace('已恢复】', '').replace('事件', '')
        else:
            hf_title = tmp_hf_title.group().split('|')[1].replace('已恢复】', '').replace('事件', '').split('（')[0]
        return hf_title

    def get_jd(jd_df):
        tmp_jd_title = re.search(r'【.*】', jd_df)
        if "【阶段" in tmp_jd_title.group():
            jd_title = tmp_jd_title.group().split('|')[3].replace('】', '')
        return jd_title

    for j in text.itertuples():
        message = getattr(j, '短信内容')
        if '已恢复' in message or '【销' in message:
            hf = hf.append({'恢复短信内容': message}, ignore_index=True)
        elif '【阶段' in message:
            jd = jd.append({'阶段短信内容': message}, ignore_index=True)

    hf['标题'] = hf['恢复短信内容'].apply(get_hf)
    text['标题'] = text['短信内容'].apply(get_kx)
    jd['标题'] = jd['阶段短信内容'].apply(get_jd)

    tmp_text = pd.merge(text, hf,
                        how='left',
                        left_on='标题',
                        right_on='标题')
    new_text = pd.merge(tmp_text, jd,
                        how='left',
                        left_on='标题',
                        right_on='标题')
    new_text = new_text.drop('标题', axis=1)

    for j in new_text.itertuples():
        message = getattr(j, '短信内容')  # 获取快讯内容
        hf_message = getattr(j, '恢复短信内容')  # 获取恢复快讯内容
        jd_message = getattr(j, '阶段短信内容')  # 获取恢复快讯内容
        # 如果快讯是突发事件、阶段进展或者其他类型的快讯直接跳过
        if '已恢复' in message or '阶段' in message or '【销' in message \
                or '后评估' in message:
            continue
        if pd.isnull(hf_message):
            recover = '否'
        else:
            recover = '是'
        # 获取快讯发送时间
        send_time = getattr(j, '发送时间')
        # 用正则获取快讯标题
        tmp_title = re.search(r'【.*】', message)
        title = tmp_title.group()
        # 获取城市
        tmp_message = message[-20:-1]
        tmp_city = []
        for c in all_city:
            if c in tmp_message:
                tmp_city.append(c)
        if len(tmp_city) == 1:
            city = tmp_city[0]
        elif len(tmp_city) == 2:
            city = tmp_city[0] + '、' + tmp_city[1]
        else:
            city = ''
        # 获取故障时间
        match = re.search(r'\d{4}/\d{1,2}/\d{1,2} \d{1,2}:\d{2}', message)
        fault_time = datetime.strptime(match.group(), '%Y/%m/%d %H:%M')
        year = str(fault_time.year)
        month = str(fault_time.month)
        day = str(fault_time.day)
        # 获取恢复时间
        if not pd.isnull(hf_message):
            hf_match = re.compile(r'\d{1,2}:\d{2}')
            hf_time = year + '-' + month + '-' + day + ' ' + hf_match.findall(hf_message)[-1]
            hf_time = datetime.strptime(hf_time, '%Y-%m-%d %H:%M')
            dif_time = hf_time - fault_time
            influence_time = dif_time.total_seconds() / 60
            if influence_time < 0:
                hf_time = year + '-' + month + '-' + str(fault_time.day + 1) + ' ' + hf_match.findall(hf_message)[-1]
                hf_time = datetime.strptime(hf_time, '%Y-%m-%d %H:%M')
                dif_time = hf_time - fault_time
                influence_time = dif_time.total_seconds() / 60
        else:
            hf_time = None
            influence_time = None
            # 获取业务影响时间
        yw_influence_time = influence_time
        # 获取故障原因
        tmp_reason = re.search(r'故?障?发?生?原因.*处理情况', str(hf_message))
        tmp_jf_reason = re.search(r'停电原因.*停电期间', str(hf_message))
        tmp_jf_reason2 = re.search(r'停电原因.*停电后', str(message))
        if tmp_reason is not None:
            reason = tmp_reason.group()[5:-5]
        else:
            reason = ''
        if tmp_jf_reason is not None:
            jf_reason = tmp_jf_reason.group()[5:-5]
        else:
            jf_reason = ''
        if tmp_jf_reason2 is not None:
            jf_reason2 = tmp_jf_reason2.group()[5:-4]
        else:
            jf_reason2 = ''
        if "停电" in str(hf_message) and "重点管控" not in str(hf_message):
            tmp_jl_reason = re.search(r'原因.*处理情况', str(hf_message))
            if tmp_jl_reason is not None:
                jl_reason = tmp_jl_reason.group()[3:-5]
        if "计划内" in str(hf_message):
            jl_reason = "供电局计划内停电"
        else:
            jl_reason = ""
        # 如果有投诉则获取投诉量，如果没有获取到则投诉量为0
        if hf_message:
            match1 = re.search(r'投诉总?量?累?计?共?\d{1,4}宗', str(hf_message))
        if match1:
            match1 = re.search(r'\d{1,4}', match1.group())
            complaint = match1.group()
        else:
            complaint = 0
        # 使用"|"作分隔符对标题进行分割，因为内部通报格式和快讯不同，所以分别判断
        if "内部通报" in title:
            a = title.split('|')[3].replace('】', '')
        else:
            a = title.split('|')[1].replace('】', '')
        gz_title = (year + '年' + month + '月' + day + '日' + a).replace('事件', '')
        # 获取值班长
        for m in monitor:
            if m in message:
                lead = m

        if "国际" in title or "一干" in title:
            net_type = '传输网'
            major = '一干传输'
            device_type = '一干光缆'
            business = '不影响业务'
            unit = '传动室'
        elif "二干" in title:
            net_type = '传输网'
            major = '二干传输'
            device_type = '二干光缆'
            business = '不影响业务'
            unit = '传动室'
        elif "5G基站退服超门限" in title:
            net_type = '无线网'
            major = '无线5G'
            business = '数据业务-5G'
            device_type = ''
            unit = ''
        elif "4G基站退服超门限" in title or "LTE-RRU退服超门限" in title:
            net_type = '无线网'
            major = '无线4G'
            business = '数据业务-4G'
            device_type = ''
            unit = ''
        elif "2G基站退服超门限" in title or "GSM-RRU退服超门限" in title:
            net_type = '无线网'
            major = '无线2G'
            business = '数据业务-2G'
            device_type = ''
            unit = ''
        elif "VVIP" in title:
            net_type = '无线网'
            major = '无线4G'
            device_type = 'ENODEB'
            business = '数据业务-4G'
        elif "AAA专线" in title or "3A专线" in title:
            net_type = '传输网'
            major = '本地传输'
            device_type = '本地光缆'
            business = '集客业务'
            unit = '地市-本地传动'
        elif "重点管控" in title:
            net_type = '动环设备'
            major = '动力设备'
            device_type = '动力设备'
            jf_type = '汇聚机房'
            business = '不影响业务'
        elif "停电" in title and "重点管控" not in title:
            net_type = '动环设备'
            major = '动力设备'
            device_type = '动力设备'
            jf_type = '核心机楼'
            business = '不影响业务'
        elif "OLT" in title and "管控" not in title:
            net_type = '传输网'
            major = '本地传输'
            device_type = 'OLT'
            business = '家宽业务'
        elif "一干" in title and "SDH单板" in title:
            net_type = '传输网'
            major = '一干传输'
            device_type = 'SDH'
            business = '不影响业务'
        elif "无告警上报" in title:
            net_type = '网管网'
            major = '网管网'
            device_type = '网管系统'
            business = '不影响业务'
        elif "4G上网" in title:
            net_type = '核心网'
            major = '电路域'
            device_type = 'MME'
            business = '数据业务-4G'
        elif "节电小区" in title:
            net_type = '无线网'
            major = '无线5G'
            device_type = '5G网管'
            business = '不影响业务'
        elif "IPTV业务" in title:
            net_type = '数据网'
            major = '数据网'
            device_type = '业务平台'
            business = 'IPTV'
        elif "UPS" in title:
            net_type = '动环设备'
            major = '动力设备'
            device_type = '动力设备'
            business = '不影响业务'
        elif "掌上运维" in title:
            net_type = '网管网'
            major = '网管网'
            device_type = '防火墙'
            business = '不影响业务'
        elif "网络云" in title:
            net_type = '网络云'
            major = '网络云'
            device_type = '服务器'
            business = '不影响业务'
        elif "短信收发" in title:
            net_type = '核心网'
            major = '电路域'
            device_type = 'EIR'
            business = '短信业务'
        elif "短号呼入呼出" in title:
            net_type = '核心网'
            major = '电路域'
            device_type = '其他'
            business = '短号业务'
        else:
            net_type = None
            major = None
            device_type = None
            business = None
            unit = None

        if business == '不影响业务':
            yw_influence_time = 0

        if "内部通报" in title:
            nb_result = nb_result.append({'月份': month,
                                          '是否已恢复': recover,
                                          '负责单位': city + '分公司',
                                          '网络类型': net_type,
                                          '所属专业': major,
                                          '设备类型': device_type,
                                          '影响业务种类': business,
                                          '故障标题': gz_title,
                                          '故障发生时间': fault_time,
                                          '故障销除时间': hf_time,
                                          '故障历时（分钟）': influence_time,
                                          '故障处理历时（分钟）': influence_time,
                                          '投诉用户数': complaint,
                                          '故障处理过程': message + '\n' + str(jd_message) + '\n' + str(hf_message),
                                          '故障级别': '升级的一般快讯',
                                          '是否有告警': '是',
                                          '是否省监控主动发现': '是',
                                          '网管告警缺漏的原因': '无',
                                          '故障发生地市': city,
                                          '故障影响地市': city,
                                          '上报集团': '否',
                                          '上报管局': '否',
                                          '采用应急容灾情况': '否',
                                          '值班长': lead,
                                          '首条故障短信发布时间': send_time},
                                         ignore_index=True)

        if "【快讯" in title and "停电" not in title and "数据中心" not in title \
                and "异常事件管控" not in title:
            kx_result = kx_result.append({'月份': month,
                                          '是否已恢复': recover,
                                          '负责单位': city + '分公司',
                                          '网络类型': net_type,
                                          '所属专业': major,
                                          '设备类型': device_type,
                                          '影响业务种类': business,
                                          '故障处理过程': message + '\n' + str(hf_message),
                                          '故障标题': gz_title,
                                          '投诉用户数': complaint,
                                          '故障发生时间': fault_time,
                                          '故障销除时间': hf_time,
                                          '故障处理时长': influence_time,
                                          '故障原因': reason,
                                          '维护单位（统计主动监控）': unit,
                                          '是否有告警': '是',
                                          '是否省监控主动发现': '是',
                                          '网管告警缺漏的原因': '无',
                                          '业务影响历时': yw_influence_time,
                                          '故障发生地市': city,
                                          '故障影响地市': city,
                                          '值班长': lead,
                                          '首条故障短信发布时间': send_time},
                                         ignore_index=True)

        if "停电" in title or "数据中心" in title:
            jf_result = jf_result.append({'月份': month,
                                          '是否已恢复': recover,
                                          '负责单位': city + '分公司',
                                          '网络类型': net_type,
                                          '所属专业': major,
                                          '设备类型': device_type,
                                          '影响业务种类': business,
                                          '类型': jf_type,
                                          '故障标题': gz_title,
                                          '故障网元': "机房市电",
                                          '故障发生时间': fault_time,
                                          '故障销除时间': hf_time,
                                          '故障发生地市': city,
                                          '故障影响地市': city,
                                          '通报信息': message + '\n' + str(hf_message),
                                          '是否有告警': '是',
                                          '是否省监控主动发现': '是',
                                          '网管告警缺漏的原因': '无',
                                          '故障原因分析': jf_reason + jl_reason + jf_reason2,
                                          '故障原因分类': '动环原因',
                                          '故障原因细分': '其他原因',
                                          '设备厂家': "其他",
                                          '值班长': lead,
                                          '首条故障短信发布时间': send_time},
                                         ignore_index=True)

        if "异常事件管控" in title or "家宽" in title:
            guankong_result = guankong_result.append({'月份': month,
                                                      '是否已恢复': recover,
                                                      '负责单位': city + '分公司',
                                                      '网络类型': net_type,
                                                      '所属专业': major,
                                                      '设备类型': device_type,
                                                      '影响业务种类': business,
                                                      '投诉用户数': complaint,
                                                      '故障标题': gz_title,
                                                      '故障发生时间': fault_time,
                                                      '故障销除时间': hf_time,
                                                      '故障处理时长': influence_time,
                                                      '业务影响历时': None,
                                                      '故障发生地市': city,
                                                      '故障影响地市': city,
                                                      '通报信息': message + '\n' + str(hf_message),
                                                      '是否有告警': '是',
                                                      '是否省监控主动发现': '是',
                                                      '网管告警缺漏的原因': '无',
                                                      '值班长': lead,
                                                      '首条故障短信发布时间': send_time}, ignore_index=True)

    nb_result['序号'] = range(1, len(nb_result) + 1)
    kx_result['序号'] = range(1, len(kx_result) + 1)
    jf_result['序号'] = range(1, len(jf_result) + 1)
    guankong_result['序号'] = range(1, len(guankong_result) + 1)

    writer = pd.ExcelWriter('故障信息整理.xlsx')
    nb_result.to_excel(writer, "升级一般故障", index=None)
    kx_result.to_excel(writer, "故障快讯", index=None)
    jf_result.to_excel(writer, "机楼和汇聚机房停电快讯", index=None)
    guankong_result.to_excel(writer, "管控类快讯", index=None)
    other_text.to_excel(writer, "非故障快讯", index=None)
    writer.save()
    print('自动整理已完成！')
