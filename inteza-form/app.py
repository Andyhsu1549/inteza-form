import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import xlsxwriter

ZL_MACHINES = ['ZL-01', 'ZL-02', 'ZL-03', 'ZL-04', 'ZL-05', 'ZL-07', 'ZL-08', 'ZL-09', 'ZL-10', 'ZL-11']
DL_MACHINES = ['DL-03', 'DL-04', 'DL-05', 'DL-10', 'DL-13']

FIBO_QUESTIONS = {
    'DL-03': ['覺得整體重量會太輕嗎？'],
    'DL-04': ['覺得輕的好還是重的好？'],
    'ZL-01': ['座椅目前夠低嗎？'],
    'ZL-02': ['椅背會太低嗎？'],
    'ZL-07': ['腰帶會很不舒服嗎？'],
    'ZL-08': ['會覺得很難上機嗎？'],
    'ZL-09': ['壓腿滾筒會不會太硬很不舒服？']
}

EVALUATION_SECTIONS = {
    '觸感體驗': ['座位調整重量片是否方便？', '整體動作是否穩定有質感？', '承靠部位是否舒適？', '抓握部分是否符合手感？'],
    '人因調整': ['把手調整是否容易？', '承靠墊位置是否符合需求？', '坐墊位置是否調整方便？', '握把／踏板位置與角度是否符合需求？', '使用時關節是否可對齊軸點？'],
    '力線評估': ['起始重量是否恰當？', '動作過程中重量變化是否流暢？'],
    '運動軌跡': ['是否能完成全行程訓練？', '關節活動角度是否自然？', '運動軌跡是否能完全刺激目標肌群？'],
    '心理感受': ['使用後的滿意度如何？', '是否有願意推薦給他人的意願？'],
    '價值感受': ['你認為我們品牌在傳遞什麼形象？', '你估算這台機器價值多少？']
}

st.set_page_config(layout='wide')
st.markdown("<h1 style='text-align: center; color: #4CAF50;'>INTENZA 人因評估系統</h1>", unsafe_allow_html=True)

# 強力穩定版：每次刷新後自動回到頁面頂端
st.markdown("""
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            window.scrollTo(0, 0);
        });
    </script>
""", unsafe_allow_html=True)
app_mode = st.sidebar.selectbox('選擇功能', ['表單填寫工具', '分析工具'])

if app_mode == '表單填寫工具':
    if 'records' not in st.session_state:
        st.session_state.records = []
    if 'fibo_records' not in st.session_state:
        st.session_state.fibo_records = []
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
            st.session_state.fibo_records = [r for r in st.session_state.fibo_records if r['機器代碼'] != m]
            st.session_state.selected_series = 'ZL 系列' if m.startswith('ZL') else 'DL 系列'
            st.session_state.current_machine_index = ZL_MACHINES.index(m) if m.startswith('ZL') else DL_MACHINES.index(m)
            st.experimental_rerun()
        st.sidebar.write(m)
    st.sidebar.progress(len(completed_machines) / len(all_machines))

    if current_machine is None:
        st.success(f'🎉 {st.session_state.selected_series} 填寫完成！請至側邊欄下載資料或選擇另一系列繼續填寫')
    else:
        data_list = []
        fibo_data_list = []
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
        # Fibo 問題區塊（只在指定機台出現，並標註為 Fibo問題）
        if current_machine in FIBO_QUESTIONS:
            st.subheader('🔹 Fibo問題追蹤')
            for item in FIBO_QUESTIONS[current_machine]:
                display_item = f'{item} （Fibo問題）'
                key_result = f'Fibo_{item}_result'
                if key_result not in st.session_state:
                    st.session_state[key_result] = None
                st.markdown(f"**{display_item}**")
                col1, col2 = st.columns([0.48, 0.48])
                with col1:
                    if st.button('✅ Yes', key=f'Fibo_{item}_yes'):
                        st.session_state[key_result] = 'Yes'
                with col2:
                    if st.button('❌ No', key=f'Fibo_{item}_no'):
                        st.session_state[key_result] = 'No'

                current_selection = st.session_state[key_result]
                if current_selection:
                    st.write(f"👉 已選擇：**{current_selection}**")
                note = st.text_input(f'{display_item} Note', key=f'Fibo_{item}_note')
                fibo_data_list.append({
                    '測試者': st.session_state.tester_name,
                    '機器代碼': current_machine,
                    '項目': display_item,
                    'Yes/No': current_selection if current_selection else '未選擇',
                    'Note': note,
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
            st.session_state.fibo_records.extend(fibo_data_list)
            for key in list(st.session_state.keys()):
                if key.endswith('_result') or key.endswith('_note') or key.endswith('_summary_note'):
                    del st.session_state[key]
            st.session_state.current_machine_index += 1
            st.success("已儲存，正在切換到下一台...")
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

            df_fibo = pd.DataFrame(st.session_state.fibo_records)
            if not df_fibo.empty:
                st.sidebar.download_button('📥 下載 Fibo問題追蹤 Excel 檔案', create_excel(df_fibo), file_name=f'Fibo問題追蹤_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx')
        else:
            st.sidebar.write('尚無資料')
