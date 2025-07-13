import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import xlsxwriter

ZL_MACHINES = ['ZL-01', 'ZL-02', 'ZL-03', 'ZL-04', 'ZL-05', 'ZL-07', 'ZL-08', 'ZL-09', 'ZL-10', 'ZL-11']
DL_MACHINES = ['DL-03', 'DL-04', 'DL-05', 'DL-10', 'DL-13']

EVALUATION_SECTIONS = {
    '觸感體驗': ['座位調整重量片是否方便？', '整體動作是否穩定有質感？', '承靠部位是否舒適？', '抓握部分是否符合手感？'],
    '人因調整': ['把手調整是否容易？', '承靠墊位置是否符合需求？', '坐墊位置是否調整方便？', '握把／踏板位置與角度是否符合需求？', '使用時關節是否可對齊軸點？'],
    '力線評估': ['起始重量是否恰當？', '動作過程中重量變化是否流暢？'],
    '運動軌跡': ['是否能完成全行程訓練？', '關節活動角度是否自然？', '運動軌跡是否能完全刺激目標肌群？'],
    '心理感受': ['使用後的滿意度如何？', '是否有願意推薦給他人的意願？'],
    '價值感受': ['你認為我們品牌在傳遞什麼形象？', '你估算這台機器價值多少？']
}

st.set_page_config(layout='wide')
st.markdown("<h1 style='text-align: center; color: #4CAF50;'>INTEZA 人因評估系統</h1>", unsafe_allow_html=True)

app_mode = st.sidebar.selectbox('選擇功能', ['表單填寫工具', '分析工具'])

if app_mode == '表單填寫工具':
    if 'records' not in st.session_state:
        st.session_state.records = []
    if 'current_machine_index' not in st.session_state:
        st.session_state.current_machine_index = 0
    if 'tester_name' not in st.session_state:
        st.session_state.tester_name = ''
    if 'selected_series' not in st.session_state:
        st.session_state.selected_series = None

    MACHINE_CODES = []
    current_machine = None
    if st.session_state.selected_series:
        MACHINE_CODES = ZL_MACHINES if st.session_state.selected_series == 'ZL 系列' else DL_MACHINES
        if st.session_state.current_machine_index < len(MACHINE_CODES):
            current_machine = MACHINE_CODES[st.session_state.current_machine_index]

    # 顯示測試者姓名與目前機台在側邊欄最上方
    st.sidebar.success(f"✅ 目前測試者姓名：{st.session_state.tester_name or '未輸入'}")
    if current_machine:
        st.sidebar.info(f"🚀 **目前驗證中機台：{current_machine}**")

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
        if st.button('🔄 重新輸入姓名'):
            st.session_state.tester_name = ''
            st.session_state.selected_series = None
            st.session_state.current_machine_index = 0
            st.rerun()

    if st.session_state.selected_series is None:
        series_choice = st.radio('請選擇要開始的系列', ['ZL 系列', 'DL 系列'])
        if st.button('✅ 確認系列'):
            st.session_state.selected_series = series_choice
            st.session_state.current_machine_index = 0
            st.rerun()
        st.stop()

    all_machines = ZL_MACHINES + DL_MACHINES
    completed_machines = sorted(set([r['機器代碼'] for r in st.session_state.records]), key=lambda x: all_machines.index(x))

    st.sidebar.header('✅ 已完成機台')
    for m in completed_machines:
        if st.sidebar.button(f'{m} 修正'):
            st.session_state.records = [r for r in st.session_state.records if r['機器代碼'] != m]
            st.session_state.selected_series = 'ZL 系列' if m.startswith('ZL') else 'DL 系列'
            st.session_state.current_machine_index = ZL_MACHINES.index(m) if m.startswith('ZL') else DL_MACHINES.index(m)
            st.experimental_rerun()
        st.sidebar.write(m)
    st.sidebar.progress(len(completed_machines) / len(all_machines))

    if current_machine is None:
        st.success(f'🎉 {st.session_state.selected_series} 填寫完成！請至側邊欄下載資料或選擇另一系列繼續填寫')
    else:
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
                    '分數': None,
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
                '分數': None,
                '日期時間': date_str
            })

        score = st.radio('⭐ 整體評分（1~5分）', [1, 2, 3, 4, 5], index=2)
        data_list.append({
            '測試者': st.session_state.tester_name,
            '機器代碼': current_machine,
            '區塊': '整體評估',
            '項目': '整體評分',
            'Pass/NG': 'N/A',
            'Note': '',
            '分數': score,
            '日期時間': date_str
        })

        if st.button('✅ 完成本機台並儲存，進入下一台'):
            st.session_state.records.extend(data_list)
            st.session_state.current_machine_index += 1
            if st.session_state.current_machine_index >= len(MACHINE_CODES):
                st.success(f'🎉 {st.session_state.selected_series} 填寫完成！請至側邊欄下載資料或選擇另一系列繼續填寫')
            else:
                st.rerun()

    if st.session_state.records:
        df = pd.DataFrame(st.session_state.records)
        with st.expander('🔍 預覽目前已填寫資料'):
            st.dataframe(df)

        def create_excel(df_input):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_input.to_excel(writer, index=False, sheet_name='評估結果')
                workbook = writer.book
                worksheet = writer.sheets['評估結果']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
                for col_num, value in enumerate(df_input.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 20)
                worksheet.freeze_panes(1, 0)
            output.seek(0)
            return output

        st.sidebar.download_button('📥 下載 全系列 Excel 檔案', create_excel(df), file_name=f'評估結果_INTEZA_全系列_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx')

        df_zl = df[df['機器代碼'].str.startswith('ZL')]
        if not df_zl.empty:
            st.sidebar.download_button('📥 下載 ZL 系列 Excel 檔案', create_excel(df_zl), file_name=f'評估結果_INTEZA_ZL系列_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx')

        df_dl = df[df['機器代碼'].str.startswith('DL')]
        if not df_dl.empty:
            st.sidebar.download_button('📥 下載 DL 系列 Excel 檔案', create_excel(df_dl), file_name=f'評估結果_INTEZA_DL系列_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx')
    else:
        st.sidebar.write('尚無資料')

# 分析工具區塊省略，如需我幫你整合完整分析工具，請直接說：「幫我整合分析工具區」！


elif app_mode == '分析工具':
    uploaded_files = st.sidebar.file_uploader("📂 上傳整合資料檔（Excel）", type=['xlsx'], accept_multiple_files=True)

    if uploaded_files:
        df_list = [pd.read_excel(file) for file in uploaded_files]
        df = pd.concat(df_list, ignore_index=True)
        st.success(f"✅ 已整合 {len(uploaded_files)} 個檔案，共 {len(df)} 筆資料！")

        ng_data = df[df['Pass/NG'] == 'NG']
        score_data = df[df['項目'] == '整體評分'].copy()
        score_data['整體評分'] = pd.to_numeric(score_data['分數'], errors='coerce')

        summary_list = []
        SECTION_ORDER = list(EVALUATION_SECTIONS.keys()) + ['整體評估']
        MACHINE_CODES_ALL = ZL_MACHINES + DL_MACHINES

        for machine in MACHINE_CODES_ALL:
            machine_df = df[df['機器代碼'] == machine]
            for section in SECTION_ORDER:
                sec_df = machine_df[machine_df['區塊'] == section]
                if sec_df.empty:
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
        for machine in MACHINE_CODES_ALL:
            machine_ng = ng_summary[ng_summary['機器代碼'] == machine].sort_values('NG次數', ascending=False)
            for _, row in machine_ng.iterrows():
                summary_list.append({'區塊': f"NG：{row['區塊']}", '項目': row['項目'], machine: f"{row['NG次數']} 次"})

        summary_df = pd.DataFrame(summary_list)
        for machine in MACHINE_CODES_ALL:
            if machine not in summary_df.columns:
                summary_df[machine] = None

        final_df = summary_df.pivot_table(index=['區塊', '項目'], values=MACHINE_CODES_ALL, aggfunc='first').reset_index()
        ng_sections = sorted([s for s in final_df['區塊'].unique() if s.startswith('NG：')])
        section_order_full = SECTION_ORDER + ng_sections
        final_df['區塊'] = pd.Categorical(final_df['區塊'], categories=section_order_full, ordered=True)
        final_df = final_df.sort_values(['區塊', '項目']).reset_index(drop=True)

        st.markdown("### 📊 分析結果預覽")
        st.dataframe(final_df)

        def create_analysis_excel(df_input):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_input.to_excel(writer, index=False, sheet_name='分析報告')
                workbook = writer.book
                worksheet = writer.sheets['分析報告']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
                for col_num, value in enumerate(df_input.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 20)
                worksheet.freeze_panes(1, 0)
            output.seek(0)
            return output

        st.sidebar.download_button(
            '📥 下載分析報告 Excel',
            create_analysis_excel(final_df),
            file_name=f'分析報告_INTEZA_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
        )
    else:
        st.info("請在側邊欄上傳資料檔案以開始分析。")
