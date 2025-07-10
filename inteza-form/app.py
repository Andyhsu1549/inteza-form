import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from io import BytesIO
import xlsxwriter

MACHINE_CODES = ['ZL-01', 'ZL-02', 'ZL-03', 'ZL-04', 'ZL-05', 'ZL-07', 'ZL-08', 'ZL-09', 'ZL-10', 'ZL-11',
                 'DL-03', 'DL-04', 'DL-05', 'DL-10', 'DL-13']

EVALUATION_SECTIONS = {
    '觸感體驗': ['可於座位調整重量片', '整體做動穩定有質感', '承靠舒適度佳', '抓握舒適度佳'],
    '人因調整': ['調整把手容易調整', '承靠墊位置符合需求', '坐墊位置調整符合需求', '握把位置與抓握角度 / 或踏板位置符合需求', '關節可對齊軸點'],
    '力線評估': ['起始重量', '行程重量變化'],
    '運動軌跡': ['可完成全行程訓練', '符合關節角度', '運動軌跡可完全刺激目標肌群']
}

st.markdown("<h1 style='text-align: center; color: #4CAF50;'>INTEZA 人因評估系統</h1>", unsafe_allow_html=True)

app_mode = st.sidebar.selectbox('選擇功能', ['表單填寫工具', '分析工具'])

if app_mode == '表單填寫工具':
    if 'records' not in st.session_state:
        st.session_state.records = []
    if 'current_machine_index' not in st.session_state:
        st.session_state.current_machine_index = 0
    if 'tester_name' not in st.session_state:
        st.session_state.tester_name = ''

    if st.session_state.tester_name == '':
        tester_input = st.text_input('請輸入測試者姓名')
        if st.button('✅ 確認提交姓名'):
            if tester_input.strip() != '':
                st.session_state.tester_name = tester_input.strip()
                st.rerun()
            else:
                st.warning('請先輸入姓名再提交')
        st.stop()
    else:
        st.success(f"✅ 目前測試者姓名：{st.session_state.tester_name}（請確認無誤）")
        if st.button('🔄 重新輸入姓名'):
            st.session_state.tester_name = ''
            st.rerun()

    current_machine = MACHINE_CODES[st.session_state.current_machine_index]
    st.header(f'🚀 目前機器：{current_machine}')

    data_list = []
    date_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for section, items in EVALUATION_SECTIONS.items():
        st.subheader(f'🔹 {section}')
        section_notes = []
        for item in items:
            key_result = f'{section}_{item}_result'
            if key_result not in st.session_state:
                st.session_state[key_result] = None

            st.markdown(f"**{item}**")
            col1, col2 = st.columns([0.48, 0.48])
            with col1:
                if st.button('✅ Pass', key=f'{section}_{item}_pass'):
                    st.session_state[key_result] = 'Pass'
            with col2:
                if st.button('❌ NG', key=f'{section}_{item}_ng'):
                    st.session_state[key_result] = 'NG'

            current_selection = st.session_state[key_result]
            if current_selection:
                st.write(f"👉 已選擇：**{current_selection}**")
            note = st.text_input(f'{item} Note', key=f'{section}_{item}_note')
            if note.strip() != '':
                section_notes.append(f'{item}: {note}')
            data_list.append({
                '測試者': st.session_state.tester_name,
                '機器代碼': current_machine,
                '區塊': section,
                '項目': item,
                'Pass/NG': current_selection if current_selection else '未選擇',
                'Note': note,
                '日期時間': date_str
            })

        combined_note = '; '.join(section_notes)
        summary_note = st.text_area(f'💬 {section} 區塊總結 Note（以下為細項 Note 整理供參考）\n{combined_note}', key=f'{section}_summary_note')
        data_list.append({
            '測試者': st.session_state.tester_name,
            '機器代碼': current_machine,
            '區塊': section,
            '項目': '區塊總結 Note',
            'Pass/NG': 'N/A',
            'Note': summary_note,
            '日期時間': date_str
        })

    score = st.slider('⭐ 整體評分（1~5分）', 1, 5, 3)
    data_list.append({
        '測試者': st.session_state.tester_name,
        '機器代碼': current_machine,
        '區塊': '整體評估',
        '項目': '整體評分',
        'Pass/NG': 'N/A',
        'Note': str(score),
        '日期時間': date_str
    })

    if st.button('✅ 完成本機台並儲存，進入下一台'):
        st.session_state.records.extend(data_list)
        st.session_state.current_machine_index += 1
        if st.session_state.current_machine_index >= len(MACHINE_CODES):
            st.success('🎉 所有機台填寫完成！請至側邊欄下載資料')
        else:
            st.rerun()

    st.sidebar.header('✅ 已完成機台')
    completed_machines = sorted(set([r['機器代碼'] for r in st.session_state.records]), key=lambda x: MACHINE_CODES.index(x))
    for m in completed_machines:
        st.sidebar.write(m)
    progress = len(completed_machines) / len(MACHINE_CODES)
    st.sidebar.progress(progress)

    if st.session_state.records:
        df = pd.DataFrame(st.session_state.records)
        with st.expander('🔍 預覽目前已填寫資料'):
            st.dataframe(df)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='評估結果')
            workbook = writer.book
            worksheet = writer.sheets['評估結果']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
            worksheet.freeze_panes(1, 0)
        output.seek(0)
        filename = f'評估結果_INTEZA_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx'
        st.sidebar.download_button('📥 下載 Excel 檔案', output, file_name=filename)
    else:
        st.sidebar.write('尚無資料')

elif app_mode == '分析工具':
    uploaded_file = st.sidebar.file_uploader("📂 上傳整合資料檔（Excel）", type=['xlsx'])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.success("✅ 資料上傳成功！")

        ng_data = df[df['Pass/NG'] == 'NG']
        score_data = df[df['項目'] == '整體評分'].copy()
        score_data['整體評分'] = pd.to_numeric(score_data['Note'], errors='coerce')

        summary_list = []
        SECTION_ORDER = ['觸感體驗', '人因調整', '力線評估', '運動軌跡', '整體評估']

        for machine in MACHINE_CODES:
            machine_df = df[df['機器代碼'] == machine]
            for section in SECTION_ORDER:
                sec_df = machine_df[machine_df['區塊'] == section]
                if len(sec_df) == 0:
                    continue
                pass_count = (sec_df['Pass/NG'] == 'Pass').sum()
                ng_count = (sec_df['Pass/NG'] == 'NG').sum()
                total = pass_count + ng_count
                pass_rate = (pass_count / total * 100) if total > 0 else None
                notes = sec_df[(sec_df['項目'] == '區塊總結 Note') & (sec_df['Note'] != '')]
                combined_notes = '; '.join([f"{n}（{t}）" for n, t in zip(notes['Note'], notes['測試者'])])
                summary_list.append({'區塊': section, '項目': '通過率 (%)', machine: f"{pass_rate:.1f}%" if pass_rate is not None else 'N/A'})
                summary_list.append({'區塊': section, '項目': '區塊總結 Note', machine: combined_notes if combined_notes else '無'})

            avg_score = score_data[score_data['機器代碼'] == machine]['整體評分'].mean()
            summary_list.append({'區塊': '整體評估', '項目': '總體評分', machine: f"{avg_score:.1f}" if not pd.isna(avg_score) else 'N/A'})

        ng_summary = ng_data.groupby(['機器代碼', '區塊', '項目']).size().reset_index(name='NG次數')
        for machine in MACHINE_CODES:
            machine_ng = ng_summary[ng_summary['機器代碼'] == machine].sort_values('NG次數', ascending=False)
            for _, row in machine_ng.iterrows():
                summary_list.append({'區塊': f"NG：{row['區塊']}", '項目': row['項目'], machine: f"{row['NG次數']} 次"})

        summary_df = pd.DataFrame(summary_list)
        for machine in MACHINE_CODES:
            if machine not in summary_df.columns:
                summary_df[machine] = None

        final_df = summary_df.pivot_table(index=['區塊', '項目'], values=MACHINE_CODES, aggfunc='first').reset_index()

        ng_sections = sorted([s for s in final_df['區塊'].unique() if s.startswith('NG：')])
        section_order_full = SECTION_ORDER + ng_sections
        final_df['區塊'] = pd.Categorical(final_df['區塊'], categories=section_order_full, ordered=True)
        final_df = final_df.sort_values(['區塊', '項目']).reset_index(drop=True)

        st.markdown("### 📊 分析結果預覽")
        st.dataframe(final_df)

        # 匯出 Excel with 分區標題美化
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('分析報告')

            header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
            section_format = workbook.add_format({'bold': True, 'bg_color': '#2196F3', 'font_color': 'white', 'align': 'left'})
            cell_format = workbook.add_format({'align': 'center'})

            worksheet.write_row(0, 0, final_df.columns, header_format)
            worksheet.set_column(0, len(final_df.columns) - 1, 20)
            worksheet.freeze_panes(1, 0)

            row_idx = 1
            section_map = {
                '🔹 1️⃣ 前面四大區塊': SECTION_ORDER,
                '🔹 2️⃣ 整體評估': ['整體評估'],
                '🔹 3️⃣ NG 項目（分區塊標示）': ng_sections
            }

            for section_title, section_names in section_map.items():
                worksheet.write(row_idx, 0, section_title, section_format)
                row_idx += 1
                sub_df = final_df[final_df['區塊'].isin(section_names)]
                for _, row in sub_df.iterrows():
                    worksheet.write_row(row_idx, 0, row.values, cell_format)
                    row_idx += 1

        output.seek(0)
        filename = f'分析報告_INTEZA_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
        st.sidebar.download_button('📥 下載分析報告 Excel', output, file_name=filename)
    else:
        st.info("請在側邊欄上傳資料檔案以開始分析。")