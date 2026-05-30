"""
生成示例测试用例 Excel 文件
运行：python examples/generate_sample.py
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

# ═══ 颜色定义 ═══
PRIORITY_COLORS = {
    'P0': 'FF4444',  # 红色 — 核心功能
    'P1': 'FF8800',  # 橙色 — 重要功能
    'P2': '3399FF',  # 蓝色 — 次要功能
    'P3': '999999',  # 灰色 — 低频场景
}
HEADER_FILL = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
BODY_FONT = Font(size=10)
THIN_BORDER = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC'),
)
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_ALIGN = Alignment(horizontal='left', vertical='center', wrap_text=True)

# ═══ 测试用例数据 ═══
HEADERS = ['编号', '模块', '测试点', '用例标题', '前置条件', '操作步骤', '输入数据', '预期结果', '优先级', '用例类型', '设计方法']
COL_WIDTHS = [10, 10, 12, 25, 18, 30, 20, 25, 8, 10, 12]

TEST_CASES = [
    ['TC_001', '登录', '登录功能', '正确手机号+密码登录成功', '已注册账号且状态正常',
     '1.打开登录页\n2.输入手机号\n3.输入密码\n4.点击登录按钮',
     '手机号:13800001111\n密码:Abc12345',
     '跳转首页，显示用户名', 'P0', '功能', '场景法'],

    ['TC_002', '登录', '登录功能', '正确邮箱+密码登录成功', '已注册账号且状态正常',
     '1.打开登录页\n2.输入邮箱\n3.输入密码\n4.点击登录按钮',
     '邮箱:user@test.com\n密码:Abc12345',
     '跳转首页', 'P0', '功能', '场景法'],

    ['TC_003', '登录', '用户名校验', '用户名为5位（少于下限6位）', '未登录状态',
     '1.打开登录页\n2.输入5位字符\n3.输入密码\n4.点击登录',
     '用户名:abc12\n密码:Abc12345',
     '提示"用户名至少6位"', 'P1', '边界', '边界值'],

    ['TC_004', '登录', '用户名校验', '用户名为21位（超过上限20位）', '未登录状态',
     '1.打开登录页\n2.输入21位字符\n3.输入密码\n4.点击登录',
     '用户名:a' * 21 + '\n密码:Abc12345',
     '提示"用户名不超过20位"', 'P1', '边界', '边界值'],

    ['TC_005', '登录', '用户名校验', '用户名为空', '未登录状态',
     '1.打开登录页\n2.不输入用户名\n3.输入密码\n4.点击登录',
     '用户名:\n密码:Abc12345',
     '提示"请输入用户名"', 'P1', '异常', '等价类'],

    ['TC_006', '登录', '密码校验', '密码为7位（少于下限8位）', '已注册账号',
     '1.打开登录页\n2.输入正确用户名\n3.输入7位密码\n4.点击登录',
     '用户名:13800001111\n密码:Abc1234',
     '提示"密码至少8位"', 'P1', '边界', '边界值'],

    ['TC_007', '登录', '密码校验', '密码不含字母', '已注册账号',
     '1.打开登录页\n2.输入正确用户名\n3.输入纯数字密码\n4.点击登录',
     '用户名:13800001111\n密码:12345678',
     '提示"密码需包含字母和数字"', 'P2', '异常', '等价类'],

    ['TC_008', '登录', '登录失败', '错误的密码登录', '已注册账号',
     '1.打开登录页\n2.输入正确用户名\n3.输入错误密码\n4.点击登录',
     '用户名:13800001111\n密码:WrongPwd1',
     '提示"用户名或密码错误"', 'P1', '异常', '错误推测'],

    ['TC_009', '登录', '登录失败', '不存在的用户名登录', '未注册状态',
     '1.打开登录页\n2.输入未注册手机号\n3.输入任意密码\n4.点击登录',
     '用户名:13999999999\n密码:Abc12345',
     '提示"用户名或密码错误"', 'P1', '异常', '等价类'],

    ['TC_010', '登录', '锁定规则', '连续输错5次密码', '已注册账号',
     '1.打开登录页\n2.连续输入5次错误密码',
     '正确用户名+错误密码×5',
     '提示"账号已锁定，请30分钟后重试"', 'P1', '异常', '错误推测'],

    ['TC_011', '登录', '锁定规则', '锁定期间尝试登录', '账号已被锁定（30分钟内）',
     '1.打开登录页\n2.输入正确用户名+密码\n3.点击登录',
     '用户名:13800001111\n密码:Abc12345',
     '提示"账号已锁定，剩余XX分钟"', 'P2', '异常', '场景法'],

    ['TC_012', '登录', '锁定规则', '解锁后正确登录', '锁定30分钟后',
     '1.等待30分钟\n2.打开登录页\n3.输入正确凭据\n4.点击登录',
     '用户名:13800001111\n密码:Abc12345',
     '登录成功，跳转首页', 'P2', '异常', '场景法'],

    ['TC_013', '登录', '验证码', '输错3次后出现验证码', '已注册账号',
     '1.打开登录页\n2.连续输错3次密码\n3.观察登录按钮上方',
     '正确用户名+错误密码×3',
     '出现验证码输入框', 'P2', '功能', '场景法'],

    ['TC_014', '登录', '验证码', '验证码输入错误', '已触发验证码',
     '1.输入正确用户名密码\n2.输入错误验证码\n3.点击登录',
     '用户名:13800001111\n密码:Abc12345\n验证码:0000',
     '提示"验证码错误"', 'P2', '异常', '等价类'],

    ['TC_015', '登录', '会话管理', '登录后30分钟无操作', '已登录状态',
     '1.登录成功\n2.等待30分钟不操作\n3.点击任意功能',
     '—',
     '会话超时，跳转登录页', 'P2', '异常', '场景法'],

    ['TC_016', '登录', 'UI', '密码切换可见/隐藏', '登录页打开',
     '1.输入密码\n2.点击密码框右侧"眼睛"图标',
     '密码:Abc12345',
     '密码明文/密文切换显示', 'P3', 'UI', '场景法'],

    ['TC_017', '登录', '安全性', 'SQL注入尝试', '未登录状态',
     '1.在用户名框输入SQL注入语句\n2.输入任意密码\n3.点击登录',
     "用户名: ' OR 1=1 --",
     '系统未异常，提示登录失败', 'P1', '安全', '错误推测'],

    ['TC_018', '登录', '兼容性', 'Chrome浏览器登录', '已注册账号',
     '在Chrome中执行正确登录流程',
     '正确凭据',
     '登录成功，布局正常', 'P2', '兼容性', '正交实验'],

    ['TC_019', '登录', '兼容性', 'Firefox浏览器登录', '已注册账号',
     '在Firefox中执行正确登录流程',
     '正确凭据',
     '登录成功，布局正常', 'P2', '兼容性', '正交实验'],

    ['TC_020', '登录', '性能', '100并发同时登录', '测试环境准备',
     'JMeter配置100线程同时请求登录接口',
     '100组不同账号',
     '平均响应<2s，错误率<1%', 'P1', '性能', '场景法'],
]


def create_sheet(ws, title, headers, data, col_widths):
    """创建格式化表格"""
    ws.title = title

    # 写入表头
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # 写入数据
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = BODY_FONT
            cell.alignment = LEFT_ALIGN if col_idx in (5, 6, 7) else CENTER_ALIGN
            cell.border = THIN_BORDER

            # 优先级颜色标注
            if col_idx == 9:  # 优先级列
                color = PRIORITY_COLORS.get(value, 'FFFFFF')
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
                cell.font = Font(color='FFFFFF', bold=True, size=10) if value in ('P0', 'P1') else BODY_FONT

    # 设置列宽
    for col_idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # 冻结首行
    ws.freeze_panes = 'A2'

    # 添加筛选器
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{len(data) + 1}'


# ════════════════════════════════════════════
# Sheet 1: 测试用例明细
# ════════════════════════════════════════════
ws1 = wb.active
create_sheet(ws1, '测试用例明细', HEADERS, TEST_CASES, COL_WIDTHS)

# ════════════════════════════════════════════
# Sheet 2: 测试策略概览
# ════════════════════════════════════════════
ws2 = wb.create_sheet('测试策略概览')
overview = [
    ['项目', '内容'],
    ['模块名称', '用户登录'],
    ['模块类型', '功能测试 / 接口测试 / UI 测试 / 性能测试'],
    ['测试范围', '用户登录全流程：输入校验 → 验证码 → 登录 → 锁定 → 会话管理'],
    ['设计方法', '等价类划分、边界值分析、场景法、正交实验法、错误推测法'],
    ['用例总数', '20 条（功能 6 + 边界 4 + 异常 6 + UI 1 + 安全 1 + 兼容 2 + 性能 1）'],
    ['P0（核心冒烟）', '2 条 — 通过即代表登录基本功能可用'],
    ['P1（重要功能）', '8 条 — 覆盖核心业务场景'],
    ['P2（次要功能）', '8 条 — 异常场景、兼容性'],
    ['P3（低频场景）', '2 条 — UI 细节、低频边界'],
    ['测试风险', '验证码依赖第三方服务，需 mock 测试；锁定定时依赖系统时间'],
]
for row_idx, row_data in enumerate(overview, 1):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=value)
        cell.border = THIN_BORDER
        if row_idx == 1:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER_ALIGN
        else:
            cell.font = BODY_FONT
            cell.alignment = LEFT_ALIGN
ws2.column_dimensions['A'].width = 18
ws2.column_dimensions['B'].width = 60

# ════════════════════════════════════════════
# Sheet 3: 测试数据设计
# ════════════════════════════════════════════
ws3 = wb.create_sheet('测试数据设计')
data_headers = ['数据类型', '输入字段', '具体数据', '设计意图']
data_rows = [
    ['正常数据', '用户名', '13800001111', '满足格式要求的正常手机号'],
    ['正常数据', '用户名', 'user@test.com', '满足格式要求的正常邮箱'],
    ['正常数据', '密码', 'Abc12345', '满足8-16位含字母数字要求'],
    ['边界数据', '用户名', 'abc12（5位）', '低于下限（6位）'],
    ['边界数据', '用户名', 'a' * 21 + '（21位）', '超过上限（20位）'],
    ['边界数据', '密码', 'Abc1234（7位）', '低于下限（8位）'],
    ['边界数据', '密码', 'Abc' * 6 + '（18位）', '超过上限（16位）'],
    ['异常数据', '用户名', "' OR 1=1 --", 'SQL注入攻击'],
    ['异常数据', '密码', '12345678', '不含字母的密码'],
    ['异常数据', '用户名', '<script>alert(1)</script>', 'XSS攻击'],
]
create_sheet(ws3, '测试数据设计', data_headers, data_rows, [12, 14, 30, 30])

# 保存
output_path = 'D:/VibeCoding/ai-testcase-generator/examples/登录模块-测试用例.xlsx'
wb.save(output_path)
print(f'[OK] 已生成: {output_path}')
