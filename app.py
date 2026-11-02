import streamlit as st
import pandas as pd
import numpy as np
import io

# 设置浏览器标签页名称
st.set_page_config(page_title="日语投放数据自动看板", layout="wide")

# CSS 样式渲染（包含深蓝表头、淡黄总计、粉红高亮等全局样式）
st.markdown("""
<style>
    .board-header-container { display: flex; justify-content: space-between; align-items: center; margin-top: 22px; margin-bottom: 8px; }
    .board-title { font-size: 18px; font-weight: bold; color: #1f2937; border-left: 5px solid #3b82f6; padding-left: 10px; margin: 0; }
    .excel-table { width: 100%; border-collapse: collapse; font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; margin-bottom: 25px; }
    .excel-table th { background-color: #2b3a87; color: #ffffff; font-weight: bold; border: 1px solid #1a2558; padding: 7px 10px; text-align: center; }
    .excel-table td { border: 1px solid #d9d9d9; padding: 6px 10px; color: #000000; text-align: center; }
    .excel-table tr.total-row { background-color: #fef08a; font-weight: bold; border-top: 2px solid #ca8a04; }
    .excel-table tr.blue-total-row { background-color: #d9e1f2; font-weight: bold; border-top: 2px solid #8ea9db; }
    .highlight-pink { background-color: #fecdd3 !important; color: #9f1239; font-weight: bold; }
    .text-red { color: #dc2626 !important; font-weight: bold; }
    .sub-roas { font-size: 11px; color: #6b7280; display: block; }
    .budget-col { background-color: #e0f2fe; }
    .copy-btn {
        background-color: #10b981; color: white; border: none; padding: 5px 12px;
        font-size: 12px; font-weight: bold; border-radius: 4px; cursor: pointer;
        transition: background-color 0.2s;
    }
    .copy-btn:hover { background-color: #059669; }
</style>

<script>
function copyTableToClipboard(tableId, btnObj) {
    var table = document.getElementById(tableId);
    if (!table) return;
    var range = document.createRange();
    range.selectNode(table);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    try {
        document.execCommand('copy');
        var origText = btnObj.innerText;
        btnObj.innerText = '✅ 已复制！';
        btnObj.style.backgroundColor = '#059669';
        setTimeout(function() {
            btnObj.innerText = origText;
            btnObj.style.backgroundColor = '#10b981';
        }, 2000);
    } catch (err) {
        alert('复制失败，请手动选中表格后复制');
    }
    window.getSelection().removeAllRanges();
}
</script>
""", unsafe_allow_html=True)

st.title("📊 日语投放数据自动看板")

uploaded_file = st.file_uploader("请上传你的原始推广数据 Excel (.xlsx)", type=["xlsx", "xls"])

table_counter = 0

def render_board_section(title_text, header_html, rows_html, total_html="", total_top=False):
    global table_counter
    table_counter += 1
    table_id = f"custom_table_{table_counter}"
    
    tbody = f"{total_html}{rows_html}" if total_top else f"{rows_html}{total_html}"
    
    html = f"""
    <div class='board-header-container'>
        <div class='board-title'>{title_text}</div>
        <button class='copy-btn' onclick="copyTableToClipboard('{table_id}', this)">📋 复制表格数据</button>
    </div>
    <table id='{table_id}' class='excel-table'>
        <thead>{header_html}</thead>
        <tbody>{tbody}</tbody>
    </table>
    """
    st.markdown(html, unsafe_allow_html=True)


if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    all_cols = df_raw.columns.tolist()

    def match(kws):
        for kw in kws:
            for c in all_cols:
                if kw.lower() in c.lower(): return c
        return None

    c_date = match(["时间", "日期", "date"]) or all_cols[0]
    c_platform = match(["投放平台", "平台", "channel", "platform"])
    c_account = match(["账户", "投放账户", "账户名称", "account"])
    c_pixel = match(["pixel名称", "pixel"])
    c_target = match(["推广目标", "目标", "objective"])
    c_opt_way = match(["优化方式", "优化方式/出价", "bidding"])
    c_bid = match(["手动出价的金额/系数", "出价金额", "出价系数", "bid"])
    c_opt = match(["优化师", "优化", "optimizer"])
    c_genre = match(["书籍类型", "类型", "category"])
    c_zh_name = match(["中文名称", "中文推广剧目/书籍名称", "中文"])
    c_en_name = match(["英文名称", "英文推广剧目/书籍名称", "英文"])
    c_book_name = match(["推广剧目/书籍名称", "剧目名称", "书籍名称"])
    c_book_id = match(["剧目id", "书籍id", "book_id", "id"])
    c_dub = match(["是否配音", "配音"])
    c_cost = match(["花费", "消耗", "cost", "spend"])
    c_d0_rev = match(["d0_revenue", "d0 revenue", "d0收入"])
    c_d0_tar = match(["d0目标收入", "目标收入"])
    c_act = match(["首日激活用户数", "激活", "act"])
    c_pay = match(["首日付费用户数", "付费", "pay"])
    c_reg_r = match(["新注册用户占比", "新注册占比"])
    c_pay_r = match(["付费用户新注册用户占比", "付费新注册占比"])
    c_kpi = match(["周 kpi", "周kpi", "kpi"])
    c_pay_rate = match(["付费率", "pay_rate"])
    c_arppu = match(["arppu"])

    # 规范化日期列
    df_raw[c_date] = pd.to_datetime(df_raw[c_date], errors='coerce')
    valid_dates = df_raw[c_date].dropna().sort_values().unique()

    # ==================== 📅 新增：日期范围选择框 ====================
    st.markdown("### 📅 日期范围筛选")
    if len(valid_dates) > 0:
        min_d = valid_dates[0].date()
        max_d = valid_dates[-1].date()
        
        date_range = st.date_input(
            "📅 请选择需要统计的日期区间：",
            value=(min_d, max_d),
            min_value=min_d,
            max_value=max_d
        )
        
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            df = df_raw[(df_raw[c_date].dt.date >= start_date) & (df_raw[c_date].dt.date <= end_date)].copy()
        else:
            df = df_raw.copy()
    else:
        df = df_raw.copy()

    # 数值清洗
    num_cols = [c_cost, c_d0_rev, c_d0_tar, c_act, c_pay, c_reg_r, c_pay_r, c_kpi, c_pay_rate, c_arppu, c_bid]
    for c in num_cols:
        if c and c in df.columns and df[c].dtype == 'object':
            s = df[c].astype(str).str.rstrip('%').str.replace(',', '').str.replace('$', '').str.strip()
            v = pd.to_numeric(s, errors='coerce')
            if df[c].astype(str).str.contains('%').any(): df[c] = v / 100.0
            else: df[c] = v

    # =========================================================================
    # 🌟 模块 0：优化师专属个人面板（独立个人工作区）
    # =========================================================================
    st.markdown("---")
    st.markdown("### 👤 优化师个人面板（个人独立分析与日预算表）")

    if c_opt and c_opt in df.columns:
        opts = [o for o in df[c_opt].dropna().unique() if str(o).strip() != ""]
        selected_opt = st.selectbox("请选择优化师姓名：", opts)

        if selected_opt:
            opt_df = df[df[c_opt] == selected_opt].copy()
            opt_df[c_date] = pd.to_datetime(opt_df[c_date], errors='coerce').dt.strftime('%Y-%m-%d')

            # --- 0.1 优化师账户综合透视表 ---
            if c_platform in opt_df.columns and c_account in opt_df.columns and c_target in opt_df.columns:
                def join_unique(series):
                    vals = [str(x).strip() for x in series.dropna().unique() if str(x).strip() != ""]
                    return " / ".join(vals) if vals else "-"

                g_summary = opt_df.groupby([c_platform, c_target, c_account], as_index=False).agg({
                    c_cost: 'sum', 
                    c_d0_rev: 'sum', 
                    c_d0_tar: 'sum', 
                    c_act: 'sum', 
                    c_pay: 'sum',
                    c_opt_way: join_unique if c_opt_way else 'count'
                })

                g_summary['target_total_cost'] = g_summary.groupby([c_platform, c_target])[c_cost].transform('sum')
                g_summary = g_summary.sort_values(
                    by=[c_platform, 'target_total_cost', c_target, c_cost], 
                    ascending=[True, False, True, False]
                )

                h_sum = "<tr><th>投放平台</th><th>推广目标</th><th>账户</th><th>优化方式</th><th>花费</th><th>D0 达标率</th><th>ROAS</th><th>CPI</th><th>CPP</th><th>付费率</th><th>ARPPU</th></tr>"
                r_sum = ""
                last_p_sum, last_t_sum = None, None

                for _, row in g_summary.iterrows():
                    p_val = str(row[c_platform]) if pd.notna(row[c_platform]) else ""
                    t_val = str(row[c_target]) if c_target and pd.notna(row[c_target]) else ""
                    a_val = str(row[c_account]) if pd.notna(row[c_account]) else ""
                    w_val = str(row[c_opt_way]) if c_opt_way else "-"

                    dis_p = p_val if p_val != last_p_sum else ""
                    dis_t = t_val if (p_val != last_p_sum or t_val != last_t_sum) else ""
                    last_p_sum, last_t_sum = p_val, t_val

                    cost = row[c_cost] if c_cost else 0
                    d0_rev = row[c_d0_rev] if c_d0_rev else 0
                    d0_rate = d0_rev / row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
                    roas = (d0_rev / cost * 100) if cost else 0
                    cpi = cost / row[c_act] if c_act and row[c_act] else np.nan
                    cpp = cost / row[c_pay] if c_pay and row[c_pay] else np.nan
                    p_rate = (row[c_pay] / row[c_act] * 100) if c_act and row[c_act] else np.nan
                    arppu = (d0_rev / row[c_pay]) if c_pay and row[c_pay] else np.nan

                    pink_cls = "highlight-pink" if d0_rate >= 0.85 else ""
                    cpi_s = f"{cpi:.2f}" if pd.notna(cpi) else "#DIV/0!"
                    cpp_s = f"{cpp:.2f}" if pd.notna(cpp) else "#DIV/0!"
                    pr_s = f"{p_rate:.2f}%" if pd.notna(p_rate) else "0.00%"
                    arp_s = f"{arppu:.2f}" if pd.notna(arppu) else "#DIV/0!"

                    r_sum += f"<tr><td><b>{dis_p}</b></td><td><b>{dis_t}</b></td><td>{a_val}</td><td>{w_val}</td><td>{cost:,.2f}</td><td class='{pink_cls}'>{d0_rate*100:.2f}%</td><td>{roas:.2f}%</td><td>{cpi_s}</td><td>{cpp_s}</td><td>{pr_s}</td><td>{arp_s}</td></tr>"

                tot_cost_s = opt_df[c_cost].sum() if c_cost else 0
                tot_d0_rev_s = opt_df[c_d0_rev].sum() if c_d0_rev else 0
                tot_d0_tar_s = opt_df[c_d0_tar].sum() if c_d0_tar else 0
                tot_d0_rate_s = tot_d0_rev_s / tot_d0_tar_s if tot_d0_tar_s else 0
                tot_roas_s = (tot_d0_rev_s / tot_cost_s * 100) if tot_cost_s else 0
                tot_act_s = opt_df[c_act].sum() if c_act else 0
                tot_pay_s = opt_df[c_pay].sum() if c_pay else 0

                t_cpi_s = tot_cost_s / tot_act_s if tot_act_s else np.nan
                t_cpp_s = tot_cost_s / tot_pay_s if tot_pay_s else np.nan
                t_prate_s = (tot_pay_s / tot_act_s * 100) if tot_act_s else np.nan
                t_arppu_s = (tot_d0_rev_s / tot_pay_s) if tot_pay_s else np.nan

                t_pink_s = "highlight-pink" if tot_d0_rate_s >= 0.85 else ""
                t_cpi_str = f"{t_cpi_s:.2f}" if pd.notna(t_cpi_s) else "#DIV/0!"
                t_cpp_str = f"{t_cpp_s:.2f}" if pd.notna(t_cpp_s) else "#DIV/0!"
                t_pr_str = f"{t_prate_s:.2f}%" if pd.notna(t_prate_s) else "0.00%"
                t_arp_str = f"{t_arppu_s:.2f}" if pd.notna(t_arppu_s) else "#DIV/0!"

                t_sum = f"<tr class='blue-total-row'><td colspan='4'>总计</td><td>{tot_cost_s:,.2f}</td><td class='{t_pink_s}'>{tot_d0_rate_s*100:.2f}%</td><td>{tot_roas_s:.2f}%</td><td>{t_cpi_str}</td><td>{t_cpp_str}</td><td>{t_pr_str}</td><td>{t_arp_str}</td></tr>"
                
                render_board_section("📌 优化师账户综合透视 (按平台 / 目标归类，首行显示一次)", h_sum, r_sum, total_html=t_sum)

            # --- 0.2 账户 & 出价策略深度透视表 ---
            acc_cross_cols = [c for c in [c_platform, c_account, c_target, c_opt_way, c_bid] if c and c in opt_df.columns]
            if len(acc_cross_cols) >= 2:
                g_acc = opt_df.groupby(acc_cross_cols, as_index=False).agg({
                    c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_act: 'sum', c_pay: 'sum'
                })
                g_acc['acc_total_cost'] = g_acc.groupby(c_account)[c_cost].transform('sum')
                g_acc = g_acc.sort_values(by=[c_platform, 'acc_total_cost', c_account, c_cost], ascending=[True, False, True, False])

                h_acc = "<tr><th>投放平台</th><th>账户</th><th>推广目标</th><th>优化方式</th><th>手动出价的金额/系数</th><th>花费</th><th>D0 达标率</th><th>ROAS</th><th>CPI</th><th>CPP</th><th>付费率</th><th>ARPPU</th></tr>"
                r_acc = ""
                last_p, last_a, last_t, last_w = None, None, None, None

                for _, row in g_acc.iterrows():
                    p_val = str(row[c_platform]) if c_platform and pd.notna(row[c_platform]) else ""
                    a_val = str(row[c_account]) if c_account and pd.notna(row[c_account]) else ""
                    t_val = str(row[c_target]) if c_target and pd.notna(row[c_target]) else ""
                    w_val = str(row[c_opt_way]) if c_opt_way and pd.notna(row[c_opt_way]) else ""
                    b_raw = row[c_bid] if c_bid else 0
                    b_val = f"{float(b_raw):.4f}" if pd.notna(b_raw) else "0.0000"

                    dis_p = p_val if p_val != last_p else ""
                    dis_a = a_val if (p_val != last_p or a_val != last_a) else ""
                    dis_t = t_val if (p_val != last_p or a_val != last_a or t_val != last_t) else ""
                    dis_w = w_val if (p_val != last_p or a_val != last_a or t_val != last_t or w_val != last_w) else ""
                    last_p, last_a, last_t, last_w = p_val, a_val, t_val, w_val

                    cost = row[c_cost] if c_cost else 0
                    d0_rev = row[c_d0_rev] if c_d0_rev else 0
                    d0_rate = d0_rev / row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
                    roas = (d0_rev / cost * 100) if cost else 0
                    cpi = cost / row[c_act] if c_act and row[c_act] else np.nan
                    cpp = cost / row[c_pay] if c_pay and row[c_pay] else np.nan
                    p_rate = (row[c_pay] / row[c_act] * 100) if c_act and row[c_act] else np.nan
                    arppu = (d0_rev / row[c_pay]) if c_pay and row[c_pay] else np.nan

                    pink_cls = "highlight-pink" if d0_rate >= 0.85 else ""
                    cpi_s = f"{cpi:.2f}" if pd.notna(cpi) else "#DIV/0!"
                    cpp_s = f"{cpp:.2f}" if pd.notna(cpp) else "#DIV/0!"
                    pr_s = f"{p_rate:.2f}%" if pd.notna(p_rate) else "0.00%"
                    arp_s = f"{arppu:.2f}" if pd.notna(arppu) else "#DIV/0!"

                    r_acc += f"<tr><td><b>{dis_p}</b></td><td><b>{dis_a}</b></td><td>{dis_t}</td><td><b>{dis_w}</b></td><td>{b_val}</td><td>{cost:,.2f}</td><td class='{pink_cls}'>{d0_rate*100:.2f}%</td><td>{roas:.2f}%</td><td>{cpi_s}</td><td>{cpp_s}</td><td>{pr_s}</td><td>{arp_s}</td></tr>"

                tot_cost = opt_df[c_cost].sum() if c_cost else 0
                tot_d0_rev = opt_df[c_d0_rev].sum() if c_d0_rev else 0
                tot_d0_tar = opt_df[c_d0_tar].sum() if c_d0_tar else 0
                tot_d0_rate = tot_d0_rev / tot_d0_tar if tot_d0_tar else 0
                tot_roas = (tot_d0_rev / tot_cost * 100) if tot_cost else 0
                tot_act = opt_df[c_act].sum() if c_act else 0
                tot_pay = opt_df[c_pay].sum() if c_pay else 0

                t_cpi = tot_cost / tot_act if tot_act else np.nan
                t_cpp = tot_cost / tot_pay if tot_pay else np.nan
                t_prate = (tot_pay / tot_act * 100) if tot_act else np.nan
                t_arppu = (tot_d0_rev / tot_pay) if tot_pay else np.nan

                t_pink = "highlight-pink" if tot_d0_rate >= 0.85 else ""
                t_cpi_s = f"{t_cpi:.2f}" if pd.notna(t_cpi) else "#DIV/0!"
                t_cpp_s = f"{t_cpp:.2f}" if pd.notna(t_cpp) else "#DIV/0!"
                t_pr_s = f"{t_prate:.2f}%" if pd.notna(t_prate) else "0.00%"
                t_arp_s = f"{t_arppu:.2f}" if pd.notna(t_arppu) else "#DIV/0!"

                t_acc = f"<tr class='blue-total-row'><td colspan='5'>总计</td><td>{tot_cost:,.2f}</td><td class='{t_pink}'>{tot_d0_rate*100:.2f}%</td><td>{tot_roas:.2f}%</td><td>{t_cpi_s}</td><td>{t_cpp_s}</td><td>{t_pr_s}</td><td>{t_arp_s}</td></tr>"
                
                render_board_section("📌 个人账户 & 出价策略深度透视 (按出价系数拆分，参考图二效果)", h_acc, r_acc, total_html=t_acc)

            # --- 0.3 近 3 天在投剧集日预算规划表 ---
            recent_dates = sorted([d for d in opt_df[c_date].dropna().unique()], reverse=True)[:3]

            if recent_dates:
                latest_date = recent_dates[0]
                latest_day_str = latest_date.split('-')[-1] + "日预算"
                sub_3d = opt_df[opt_df[c_date].isin(recent_dates)].copy()

                group_keys = [k for k in [c_en_name, c_book_name, c_book_id] if k and k in sub_3d.columns]
                if not group_keys: group_keys = [sub_3d.columns[0]]

                latest_day_cost = sub_3d[sub_3d[c_date] == latest_date].groupby(group_keys)[c_cost].sum().reset_index()
                latest_day_cost.rename(columns={c_cost: 'latest_cost'}, inplace=True)

                g_3d = sub_3d.groupby(group_keys + [c_date], as_index=False).agg({
                    c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_act: 'sum'
                })

                g_3d = pd.merge(g_3d, latest_day_cost, on=group_keys, how='left')
                g_3d['latest_cost'] = g_3d['latest_cost'].fillna(0)
                g_3d = g_3d.sort_values(by=['latest_cost'] + group_keys + [c_date], ascending=[False] + [True]*len(group_keys) + [False])

                h_3d = f"<tr><th>英文推广剧目/书籍名称</th><th>推广剧目/书籍名称</th><th>剧目ID</th><th>日期</th><th>花费</th><th>D0 达标率</th><th>ROAS</th><th>CPI</th><th style='background-color:#0284c7;'>{latest_day_str}</th><th style='background-color:#0284c7;'>备注</th></tr>"
                r_3d = ""
                last_book_key = None
                export_rows = []

                for _, row in g_3d.iterrows():
                    book_key = tuple(row[k] for k in group_keys)
                    is_new_book = (book_key != last_book_key)
                    last_book_key = book_key

                    en_val = str(row[c_en_name]) if c_en_name and is_new_book and pd.notna(row[c_en_name]) else ""
                    b_val = str(row[c_book_name]) if c_book_name and is_new_book and pd.notna(row[c_book_name]) else ""
                    id_val = str(row[c_book_id]) if c_book_id and is_new_book and pd.notna(row[c_book_id]) else ""

                    date_val = str(row[c_date])
                    cost = row[c_cost] if c_cost else 0
                    d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
                    roas = (row[c_d0_rev]/cost*100) if cost else 0
                    cpi = cost / row[c_act] if c_act and row[c_act] else 0

                    red_cls = "highlight-pink" if d0_rate >= 0.85 else ""

                    r_3d += f"<tr><td><b>{en_val}</b></td><td>{b_val}</td><td>{id_val}</td><td>{date_val}</td><td>{cost:,.2f}</td><td class='{red_cls}'>{d0_rate*100:.2f}%</td><td>{roas:.2f}%</td><td>{cpi:.2f}</td><td class='budget-col'></td><td></td></tr>"

                    export_rows.append({
                        "英文推广剧目/书籍名称": str(row[c_en_name]) if c_en_name else "",
                        "推广剧目/书籍名称": str(row[c_book_name]) if c_book_name else "",
                        "剧目ID": str(row[c_book_id]) if c_book_id else "",
                        "日期": date_val,
                        "花费": round(cost, 2),
                        "D0 达标率": f"{d0_rate*100:.2f}%",
                        "ROAS": f"{roas:.2f}%",
                        "CPI": round(cpi, 2),
                        latest_day_str: "",
                        "备注": ""
                    })

                render_board_section("📌 近 3 天在投剧集日预算规划表 (按前一天花费降序)", h_3d, r_3d)

                out_df = pd.DataFrame(export_rows)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    out_df.to_excel(writer, sheet_name=f'{selected_opt}_日预算规划', index=False)
                
                st.download_button(
                    label=f"📥 导出【{selected_opt}】的日预算规划表 (Excel)",
                    data=buffer.getvalue(),
                    file_name=f"{selected_opt}_近3天在投剧集日预算规划表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    # =========================================================================
    # 📊 全局自动看板
    # =========================================================================
    st.markdown("---")
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>🌐 周报全量数据透视看板</h2>", unsafe_allow_html=True)

    # --- 1. 本周每日数据趋势透视 ---
    if c_date in df.columns:
        df_date = df.copy()
        df_date[c_date] = pd.to_datetime(df_date[c_date], errors='coerce').dt.strftime('%Y-%m-%d')
        g_date = df_date.groupby(c_date, as_index=False).agg({
            c_cost: 'sum' if c_cost else 'count',
            c_d0_rev: 'sum' if c_d0_rev else 'count',
            c_d0_tar: 'sum' if c_d0_tar else 'count',
            c_pay_r: 'mean' if c_pay_r else 'count',
            c_reg_r: 'mean' if c_reg_r else 'count'
        }).sort_values(by=c_cost, ascending=False)
        
        tot_cost = df[c_cost].sum() if c_cost else 1
        tot_d0_rev = df[c_d0_rev].sum() if c_d0_rev else 0
        tot_d0_rate = df[c_d0_rev].sum()/df[c_d0_tar].sum() if c_d0_tar and df[c_d0_tar].sum() else 0
        tot_roas = (tot_d0_rev / tot_cost * 100) if tot_cost else 0
        tot_pay_p = df[c_pay_r].mean() if c_pay_r else 0
        tot_reg_p = df[c_reg_r].mean() if c_reg_r else 0
        
        d0_rate_cls = "text-red" if tot_d0_rate >= 0.85 else ""
        t_top = f"<tr class='total-row'><td>总计</td><td>{tot_cost:,.2f}</td><td class='{d0_rate_cls}'>{tot_d0_rate*100:.2f}%</td><td>{tot_d0_rev:,.2f}<span class='sub-roas'>{tot_roas:.2f}%</span></td><td>{tot_pay_p*100:.2f}%</td><td>{tot_reg_p*100:.2f}%</td></tr>"
        
        h = "<tr><th>时间</th><th>花费</th><th>D0达标率</th><th>d0(revenue/roas)</th><th>付费用户新注册用户占比</th><th>新注册用户占比</th></tr>"
        r = ""
        for _, row in g_date.iterrows():
            cost = row[c_cost] if c_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            d0_rev = row[c_d0_rev] if c_d0_rev else 0
            roas = (d0_rev / cost * 100) if cost else 0
            pay_p = row[c_pay_r] if c_pay_r else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            red_cls = "text-red" if d0_rate >= 0.85 else ""
            r += f"<tr><td>{row[c_date]}</td><td>{cost:,.2f}</td><td class='{red_cls}'>{d0_rate*100:.2f}%</td><td>{d0_rev:,.2f}<span class='sub-roas'>{roas:.2f}%</span></td><td>{pay_p*100:.2f}%</td><td>{reg_p*100:.2f}%</td></tr>"
            
        render_board_section("📌 本周每日数据趋势透视 (按花费降序)", h, r, total_html=t_top, total_top=True)

    # --- 2. 【板块一】渠道维度双周环比对比 ---
    if c_platform and c_platform in df.columns:
        g1 = df.groupby(c_platform, as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_act: 'sum', c_pay: 'sum',
            c_reg_r: 'mean', c_pay_r: 'mean'
        }).sort_values(by=c_cost, ascending=False)
        tot_cost = df[c_cost].sum() if c_cost else 1
        
        h = "<tr><th>投放平台</th><th>花费 ($)</th><th>消耗占比</th><th>D0 达标率</th><th>ROAS</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比</th><th>付费新用户占比</th></tr>"
        r = ""
        for _, row in g1.iterrows():
            cost = row[c_cost] if c_cost else 0
            ratio = cost / tot_cost if tot_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            cpi = cost / row[c_act] if c_act and row[c_act] else 0
            cpp = cost / row[c_pay] if c_pay and row[c_pay] else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            pay_p = row[c_pay_r] if c_pay_r else 0
            r += f"<tr><td><b>{row[c_platform]}</b></td><td>${cost:,.2f}</td><td>{ratio*100:.2f}%</td><td>{d0_rate*100:.2f}%</td><td>-</td><td>{cpi:.2f}</td><td>{cpp:.2f}</td><td>{reg_p*100:.2f}%</td><td>{pay_p*100:.2f}%</td></tr>"
        
        tot_d0 = df[c_d0_rev].sum()/df[c_d0_tar].sum() if c_d0_tar and df[c_d0_tar].sum() else 0
        tot_cpi = tot_cost / df[c_act].sum() if c_act and df[c_act].sum() else 0
        tot_cpp = tot_cost / df[c_pay].sum() if c_pay and df[c_pay].sum() else 0
        tot_reg = df[c_reg_r].mean() if c_reg_r else 0
        tot_pay = df[c_pay_r].mean() if c_pay_r else 0
        t = f"<tr class='total-row'><td>总计</td><td>${tot_cost:,.2f}</td><td>100.00%</td><td>{tot_d0*100:.2f}%</td><td>-</td><td>{tot_cpi:.2f}</td><td>{tot_cpp:.2f}</td><td>{tot_reg*100:.2f}%</td><td>{tot_pay*100:.2f}%</td></tr>"
        
        render_board_section("📌 【板块一】渠道维度双周环比对比 (FB / TT) (按花费降序)", h, r, t)

    # --- 3. 【板块二】推广目标维度对比 ---
    if c_target and c_target in df.columns:
        g2 = df.groupby(c_target, as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_pay_r: 'mean', c_reg_r: 'mean'
        }).sort_values(by=c_cost, ascending=False)
        tot_cost = df[c_cost].sum() if c_cost else 1
        
        h = "<tr><th>推广目标</th><th>花费</th><th>消耗占比</th><th>D0达标率</th><th>d0(revenue/roas)</th><th>付费用户新注册用户占比</th><th>新注册用户占比</th></tr>"
        r = ""
        for _, row in g2.iterrows():
            cost = row[c_cost] if c_cost else 0
            ratio = cost / tot_cost if tot_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            d0_rev = row[c_d0_rev] if c_d0_rev else 0
            roas = (d0_rev / cost * 100) if cost else 0
            d0_combined = f"${d0_rev:,.2f} ({roas:.2f}%)"
            pay_p = row[c_pay_r] if c_pay_r else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            r += f"<tr><td><b>{row[c_target]}</b></td><td>${cost:,.2f}</td><td>{ratio*100:.2f}%</td><td>{d0_rate*100:.2f}%</td><td>{d0_combined}</td><td>{pay_p*100:.2f}%</td><td>{reg_p*100:.2f}%</td></tr>"
        
        tot_d0_rev = df[c_d0_rev].sum() if c_d0_rev else 0
        tot_roas = (tot_d0_rev / tot_cost * 100) if tot_cost else 0
        tot_d0_combined = f"${tot_d0_rev:,.2f} ({tot_roas:.2f}%)"
        tot_d0_rate = df[c_d0_rev].sum()/df[c_d0_tar].sum() if c_d0_tar and df[c_d0_tar].sum() else 0
        tot_pay_p = df[c_pay_r].mean() if c_pay_r else 0
        tot_reg_p = df[c_reg_r].mean() if c_reg_r else 0
        t = f"<tr class='total-row'><td>总计</td><td>${tot_cost:,.2f}</td><td>100.00%</td><td>{tot_d0_rate*100:.2f}%</td><td>{tot_d0_combined}</td><td>{tot_pay_p*100:.2f}%</td><td>{tot_reg_p*100:.2f}%</td></tr>"
        
        render_board_section("📌 【板块二】推广目标维度双周环比对比 (w2a / app) (按花费降序)", h, r, t)

    # --- 4. 优化方式 & 出价透视表 (全局通用) ---
    cross_cols = [c for c in [c_platform, c_target, c_opt_way] if c and c in df.columns]
    if len(cross_cols) >= 2:
        g4 = df.groupby(cross_cols, as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum',
            c_act: 'sum', c_pay: 'sum',
            c_pay_rate: 'mean' if c_pay_rate else 'count',
            c_arppu: 'mean' if c_arppu else 'count'
        }).sort_values(by=c_cost, ascending=False)
        
        h = "<tr><th>投放平台</th><th>推广目标</th><th>优化方式</th><th>花费</th><th>D0 达标率</th><th>CPI</th><th>CPP</th><th>付费率</th><th>ARPPU</th><th>ROAS</th></tr>"
        r = ""
        last_plat, last_target = None, None
        for _, row in g4.iterrows():
            plat = str(row[c_platform]) if c_platform and pd.notna(row[c_platform]) else ""
            tgt = str(row[c_target]) if c_target and pd.notna(row[c_target]) else ""
            opt_w = str(row[c_opt_way]) if c_opt_way and pd.notna(row[c_opt_way]) else "-"
            dis_plat = plat if plat != last_plat else ""
            dis_tgt = tgt if (plat != last_plat or tgt != last_target) else ""
            last_plat, last_target = plat, tgt
            cost = row[c_cost] if c_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            cpi = cost / row[c_act] if c_act and row[c_act] else 0
            cpp = cost / row[c_pay] if c_pay and row[c_pay] else 0
            p_rate = row[c_pay_rate] if c_pay_rate and pd.notna(row[c_pay_rate]) else (row[c_pay]/row[c_act] if c_act and row[c_act] else 0)
            arppu_v = row[c_arppu] if c_arppu and pd.notna(row[c_arppu]) else (row[c_d0_rev]/row[c_pay] if c_pay and row[c_pay] else 0)
            roas_v = (row[c_d0_rev] / cost * 100) if cost else 0
            pink_cls = "highlight-pink" if d0_rate >= 0.85 else ""
            r += f"<tr><td><b>{dis_plat}</b></td><td><b>{dis_tgt}</b></td><td>{opt_w}</td><td>{cost:,.2f}</td><td class='{pink_cls}'>{d0_rate*100:.2f}%</td><td>{cpi:.2f}</td><td>{cpp:.2f}</td><td>{p_rate*100:.2f}%</td><td>{arppu_v:.2f}</td><td>{roas_v:.2f}%</td></tr>"
        
        tot_cost = df[c_cost].sum() if c_cost else 0
        tot_d0_rate = df[c_d0_rev].sum()/df[c_d0_tar].sum() if c_d0_tar and df[c_d0_tar].sum() else 0
        tot_cpi = tot_cost / df[c_act].sum() if c_act and df[c_act].sum() else 0
        tot_cpp = tot_cost / df[c_pay].sum() if c_pay and df[c_pay].sum() else 0
        tot_prate = df[c_pay].sum()/df[c_act].sum() if c_act and df[c_act].sum() else 0
        tot_arppu = df[c_d0_rev].sum()/df[c_pay].sum() if c_pay and df[c_pay].sum() else 0
        tot_roas = (df[c_d0_rev].sum()/tot_cost*100) if tot_cost else 0
        tot_pink = "highlight-pink" if tot_d0_rate >= 0.85 else ""
        t = f"<tr class='total-row'><td colspan='3'>总计</td><td>{tot_cost:,.2f}</td><td class='{tot_pink}'>{tot_d0_rate*100:.2f}%</td><td>{tot_cpi:.2f}</td><td>{tot_cpp:.2f}</td><td>{tot_prate*100:.2f}%</td><td>{tot_arppu:.2f}</td><td>{tot_roas:.2f}%</td></tr>"
        
        render_board_section("📌 4. 全局优化方式 & 出价透视表", h, r, t)

    # --- 5. 本周优化师数据汇总透视表 ---
    if c_opt and c_opt in df.columns:
        agg_dict = {c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_pay_r: 'mean', c_reg_r: 'mean'}
        if c_kpi and c_kpi in df.columns: agg_dict[c_kpi] = 'mean'
        g_opt = df.groupby(c_opt, as_index=False).agg(agg_dict).sort_values(by=c_cost, ascending=False)
        h = "<tr><th>优化师</th><th>花费</th><th>D0达标率</th><th>d0_roas</th><th>付费用户新注册用户占比</th><th>新注册用户占比</th><th>周 KPI (7天)</th><th>完成率</th></tr>"
        r = ""
        tot_cost = df[c_cost].sum() if c_cost else 0
        tot_d0_rev = df[c_d0_rev].sum() if c_d0_rev else 0
        tot_roas = (tot_d0_rev / tot_cost * 100) if tot_cost else 0
        tot_pay_p = df[c_pay_r].mean() if c_pay_r else 0
        tot_reg_p = df[c_reg_r].mean() if c_reg_r else 0
        tot_d0_rate = df[c_d0_rev].sum()/df[c_d0_tar].sum() if c_d0_tar and df[c_d0_tar].sum() else 0
        
        for _, row in g_opt.iterrows():
            cost = row[c_cost] if c_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            d0_rev = row[c_d0_rev] if c_d0_rev else 0
            roas = (d0_rev / cost * 100) if cost else 0
            pay_p = row[c_pay_r] if c_pay_r else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            kpi = row[c_kpi] if c_kpi and c_kpi in row and pd.notna(row[c_kpi]) else 0
            comp = (cost / kpi * 100) if kpi else 0
            kpi_str = f"${kpi:,.2f}" if kpi else "—"
            comp_str = f"{comp:.2f}%" if kpi else "—"
            red_cls = "text-red" if d0_rate >= 0.85 else ""
            r += f"<tr><td><b>{row[c_opt]}</b></td><td>{cost:,.2f}</td><td class='{red_cls}'>{d0_rate*100:.2f}%</td><td>{roas:.2f}%</td><td>{pay_p*100:.2f}%</td><td>{reg_p*100:.2f}%</td><td>{kpi_str}</td><td>{comp_str}</td></tr>"
        
        t = f"<tr class='total-row'><td>总计</td><td>{tot_cost:,.2f}</td><td>{tot_d0_rate*100:.2f}%</td><td>{tot_roas:.2f}%</td><td>{tot_pay_p*100:.2f}%</td><td>{tot_reg_p*100:.2f}%</td><td>—</td><td>—</td></tr>"
        
        render_board_section("📌 本周优化师数据汇总透视表（含 KPI 完成率，按花费降序）", h, r, t)

    # --- 6. 本周投放平台 & Pixel 名称数据透视 ---
    if c_pixel and c_pixel in df.columns:
        g3 = df.groupby([c_platform, c_pixel], as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum', c_act: 'sum',
            c_pay_r: 'mean', c_reg_r: 'mean'
        }).sort_values(by=c_cost, ascending=False)
        h = "<tr><th>投放平台</th><th>pixel名称</th><th>花费</th><th>D0达标率</th><th>d0_roas</th><th>付费用户新注册用户占比</th><th>新注册用户占比</th><th>CPI</th></tr>"
        r = ""
        last_platform = None
        for _, row in g3.iterrows():
            curr_platform = str(row[c_platform]) if pd.notna(row[c_platform]) else ""
            display_platform = curr_platform if curr_platform != last_platform else ""
            last_platform = curr_platform
            cost = row[c_cost] if c_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            cls = "highlight-pink" if d0_rate >= 0.85 else ""
            cpi = cost / row[c_act] if c_act and row[c_act] else 0
            pay_p = row[c_pay_r] if c_pay_r else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            r += f"<tr><td><b>{display_platform}</b></td><td>{row[c_pixel]}</td><td>{cost:,.2f}</td><td class='{cls}'>{d0_rate*100:.2f}%</td><td>-</td><td>{pay_p*100:.2f}%</td><td>{reg_p*100:.2f}%</td><td>{cpi:.4f}</td></tr>"
            
        render_board_section("📌 本周投放平台 & Pixel 名称数据透视 (按消耗降序)", h, r)

    # --- 7. 本周剧目测出率透视分析 ---
    if c_genre and c_book_name and c_genre in df.columns and c_book_name in df.columns:
        b_df = df.groupby([c_genre, c_book_name], as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum'
        })
        b_df["d0_rate"] = b_df[c_d0_rev] / b_df[c_d0_tar]
        h = "<tr><th>书籍类型</th><th>花费 ($)</th><th>消耗占比</th><th>测试剧目数 (花费>$100)</th><th>测出剧目数</th><th>测出率</th></tr>"
        r = ""
        tot_cost = df[c_cost].sum()
        t_test, t_pass = 0, 0
        genre_summary = []
        for genre, sub in b_df.groupby(c_genre):
            g_cost = sub[c_cost].sum()
            ratio = g_cost / tot_cost if tot_cost else 0
            test_cnt = len(sub[sub[c_cost] > 100])
            pass_cnt = len(sub[(sub[c_cost] > 100) & (sub["d0_rate"] >= 0.85)])
            rate = pass_cnt / test_cnt if test_cnt else 0
            genre_summary.append({
                'genre': genre, 'cost': g_cost, 'ratio': ratio,
                'test_cnt': test_cnt, 'pass_cnt': pass_cnt, 'rate': rate
            })
            t_test += test_cnt
            t_pass += pass_cnt

        genre_df = pd.DataFrame(genre_summary).sort_values(by='cost', ascending=False)
        for _, row in genre_df.iterrows():
            r += f"<tr><td>{row['genre']}</td><td>${row['cost']:,.2f}</td><td>{row['ratio']*100:.2f}%</td><td>{row['test_cnt']}</td><td>{row['pass_cnt']}</td><td>{row['rate']*100:.2f}%</td></tr>"
        tot_rate = t_pass / t_test if t_test else 0
        t = f"<tr class='total-row'><td>总计</td><td>${tot_cost:,.2f}</td><td>100.00%</td><td>{t_test}</td><td>{t_pass}</td><td>{tot_rate*100:.2f}%</td></tr>"
        
        render_board_section("📌 本周剧目测出率透视分析（限制条件：仅计算花费 > $100 剧目，D0 达标率 > 85% 算测出）", h, r, t)

    # --- 8. 本周每部书籍/剧集数据明细透视 (全量) ---
    group_cols_book = [c for c in [c_genre, c_zh_name, c_en_name, c_book_name, c_dub] if c and c in df.columns]
    if group_cols_book:
        g8 = df.groupby(group_cols_book, as_index=False).agg({
            c_cost: 'sum', c_d0_rev: 'sum', c_d0_tar: 'sum',
            c_act: 'sum', c_pay: 'sum', c_reg_r: 'mean', c_pay_r: 'mean'
        }).sort_values(by=c_cost, ascending=False)
        h = "<tr><th>书籍类型</th><th>中文推广剧目/书籍名称</th><th>英文推广剧目/书籍名称</th><th>推广剧目/书籍名称</th><th>是否配音</th><th>花费</th><th>D0 达标率</th><th>CPI</th><th>CPP</th><th>注册拉新占比</th><th>拉新付费占比</th></tr>"
        r = ""
        last_genre = None
        for _, row in g8.iterrows():
            curr_genre = str(row[c_genre]) if c_genre and pd.notna(row[c_genre]) else ""
            display_genre = curr_genre if curr_genre != last_genre else ""
            last_genre = curr_genre
            cost = row[c_cost] if c_cost else 0
            d0_rate = row[c_d0_rev]/row[c_d0_tar] if c_d0_tar and row[c_d0_tar] else 0
            cpi = cost / row[c_act] if c_act and row[c_act] else 0
            cpp = cost / row[c_pay] if c_pay and row[c_pay] else 0
            reg_p = row[c_reg_r] if c_reg_r else 0
            pay_p = row[c_pay_r] if c_pay_r else 0
            zh_n = str(row[c_zh_name]) if c_zh_name and pd.notna(row[c_zh_name]) else "-"
            en_n = str(row[c_en_name]) if c_en_name and pd.notna(row[c_en_name]) else "-"
            b_n = str(row[c_book_name]) if c_book_name and pd.notna(row[c_book_name]) else "-"
            dub = str(row[c_dub]) if c_dub and pd.notna(row[c_dub]) else "-"
            red_cls = "text-red" if d0_rate >= 0.85 else ""
            r += f"<tr><td><b>{display_genre}</b></td><td>{zh_n}</td><td>{en_n}</td><td>{b_n}</td><td>{dub}</td><td>{cost:,.2f}</td><td class='{red_cls}'>{d0_rate*100:.2f}%</td><td>{cpi:.2f}</td><td>{cpp:.2f}</td><td>{reg_p*100:.2f}%</td><td>{pay_p*100:.2f}%</td></tr>"
            
        render_board_section("📌 本周每部书籍/剧集数据明细透视 (全量，按花费降序)", h, r)

else:
    st.info("👈 请在上方上传包含原始数据的 Excel 文件以自动生成全套周报透视表。")