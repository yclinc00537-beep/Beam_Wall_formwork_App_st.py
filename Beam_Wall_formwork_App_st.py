import os
import io
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

# 設定網頁標題與外觀
st.set_page_config(page_title="牆與樑模板支撐結構檢核系統", layout="wide")

# ==========================================
# 1. 複合型結構計算核心模組
# ==========================================

def calc_wall_pressure_aci(height, rate, T, gamma_c=24.0):
    """牆體側壓：依據 ACI 347 公式隨深度分層計算"""
    p_hydro = gamma_c * height
    temp_factor = 100 / (T + 10)
    p_aci = gamma_c * (0.22 * rate * temp_factor)
    return min(p_hydro, max(15.0, p_aci))

def calc_beam_side_pressure(beam_d, gamma_c=24.0):
    """樑側側壓：流體靜壓力計算 (γc * D)"""
    return gamma_c * beam_d

def calc_beam_bottom_load(beam_d, gamma_c=24.0, live_load=2.5):
    """樑底垂直荷重：混凝土自重 + 施工活載重 (γc * D + LL)"""
    return (gamma_c * beam_d) + live_load

def check_composite_waler(p, spacing, span, waler_type, b_m=0.045, h_m=0.092, E_wood=1.0e7, fb_wood=10000.0, fv_wood=1000.0, is_beam_bottom=False):
    """支援【木材 / C型輕型鋼】雙軌橫檔結構檢核函式 (ASTM A36 規範)"""
    w = p * spacing
    coeff_m = 8.0 if is_beam_bottom else 10.0
    coeff_d_denom = 384.0 if is_beam_bottom else 128.0
    delta_numer_factor = 5.0 if is_beam_bottom else 1.0

    M = w * (span**2) / coeff_m
    
    if waler_type == "木橫檔 (自訂尺寸)":
        b, h = b_m, h_m
        E_run = E_wood
        fb_allow = fb_wood
        fv_allow = fv_wood
        I_val = b * (h**3) / 12
        Z_val = b * (h**2) / 6
        area = b * h
    elif waler_type == "C100x50x2.0 (輕型鋼)":
        area = 4.54 * 1e-4
        I_val = 71.4 * 1e-8
        Z_val = I_val / (0.10 / 2)
        E_run = 2.0e8
        fb_allow = 149400.0
        fv_allow = 100000.0
    else:  # "C100x50x2.3 (輕型鋼)"
        area = 5.17 * 1e-4
        I_val = 80.7 * 1e-8
        Z_val = I_val / (0.10 / 2)
        E_run = 2.0e8
        fb_allow = 149400.0
        fv_allow = 100000.0

    ratio_b = (M / Z_val) / fb_allow if fb_allow > 0 else 0
    delta = (delta_numer_factor * w * (span**4)) / (coeff_d_denom * E_run * I_val) if (E_run * I_val) > 0 else 0
    ratio_d = delta / (span / 270) if span > 0 else 0
    fv_act = 1.5 * (w * span / 2) / area if area > 0 else 0
    ratio_v = fv_act / fv_allow if fv_allow > 0 else 0
    
    return {"r_bending": ratio_b, "r_deflection": ratio_d, "r_shear": ratio_v, "r_max": max(ratio_b, ratio_d, ratio_v)}

def check_generic_member(p, spacing, span, b, h, E, allow, Fv, member_type="stud"):
    """通用面板與木角材檢核：抗彎、剪力、撓度 L/270"""
    w = p * spacing
    coeff_m = 8 if member_type == "panel" else 10
    M = w * (span**2) / coeff_m
    
    Z = b * (h**2) / 6
    ratio_b = (M / Z) / allow if allow > 0 else 0
    I = b * (h**3) / 12
    coeff_d = 384 if member_type == "panel" else 128
    delta = (5 if member_type == "panel" else 1) * w * (span**4) / (coeff_d * E * I) if (E*I) > 0 else 0
    ratio_d = delta / (span / 270) if span > 0 else 0
    fv = 1.5 * (w * span / 2) / (b * h) if (b*h) > 0 else 0
    ratio_v = fv / Fv if Fv > 0 else 0
    
    return {"r_bending": ratio_b, "r_deflection": ratio_d, "r_shear": ratio_v, "r_max": max(ratio_b, ratio_d, ratio_v)}

# ==========================================
# 2. Word 送審計算書輸出模組 (修改為記憶體串流輸出)
# ==========================================

def export_word_report_stream(cfg, df_wall, side_data, bot_data):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    
    p_title = doc.add_heading(cfg['project_info']['name'], 0)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading("一、設計依據", 1)
    doc.add_paragraph("1. ACI 347 混凝土模板設計規範\n2. CNS 混凝土工程施工規範\n3. 容許撓度限制採 L/270 (鋼型鋼同採此標準校核)")
    
    doc.add_heading("二、設計條件與構件規格明細", 1)
    wp = cfg['wall_params']
    ws = cfg['wall_section']
    bs = cfg['beam_side_section']
    bb = cfg['beam_bot_section']
    bg = cfg['beam_geo']
    
    doc.add_paragraph(f"• 牆最大高度 H = {wp['height']:.2f} m\n• 澆置速度 R = {wp['pour_rate']:.2f} m/hr")
    doc.add_paragraph(f"• 設計溫度 T = {wp['temperature']:.1f} °C\n• 混凝土重 γc = {wp['gamma_c']:.1f} kN/m³")
    doc.add_paragraph(f"• 預定施作大樑尺寸：寬 B = {bg['beam_b']*100:.1f} cm × 深 D = {bg['beam_d']*100:.1f} cm")
    
    doc.add_heading("1. 牆模板系統構件尺寸與配置規格：", 2)
    waler_w_str = f"規格種類：{ws['waler_type']}" if "輕型鋼" in ws['waler_type'] else f"木橫檔 寬 {ws['waler_b']*100:.1f} cm × 高 {ws['waler_h']*100:.1f} cm"
    doc.add_paragraph(
        f"• 面板 (Panel) 厚度：{ws['panel_t_mm']} mm (設計跨距 {cfg['wall_spacing']['panel_span']:.2f} m)\n"
        f"• 內撐角材 (Joist/Stud)：寬 {ws['stud_b']*100:.1f} cm × 高 {ws['stud_h']*100:.1f} cm (間距 {cfg['wall_spacing']['stud_spacing']:.2f} m, 跨距 {cfg['wall_spacing']['stud_span']:.2f} m)\n"
        f"• 外橫檔 (Waler)：{waler_w_str} (間距 {cfg['wall_spacing']['waler_spacing']:.2f} m, 跨距 {cfg['wall_spacing']['waler_span']:.2f} m)\n"
        f"• Form Tie 螺桿：直徑 {ws['tie_dia_mm']} mm (水平 Sx = {cfg['wall_spacing']['tie_sx']:.2f} m, 垂直 Sy = {cfg['wall_spacing']['tie_sy']:.2f} m, 容許拉力 {cfg['capacity']['tie_allowable']:.1f} kN)"
    )
    
    doc.add_heading("2. 樑模板系統構件尺寸與配置規格：", 2)
    waler_b_str = f"規格種類：{bb['waler_type']} (符合 ASTM A36 規範)" if "輕型鋼" in bb['waler_type'] else f"木橫檔 寬 {bb['waler_b']*100:.1f} cm × 高 {bb['waler_h']*100:.1f} cm"
    prop_type_str = "單排中央支撐" if bg['beam_b'] <= 0.3 else "雙排對稱分荷支撐"
    doc.add_paragraph(
        f"• 側面板 (Side Panel) 厚度：{bs['panel_t_mm']} mm (設計跨距 {cfg['beam_side_spacing']['panel_span']:.2f} m)\n"
        f"• 側內撐 (Side Joist)：寬 {bs['stud_b']*100:.1f} cm × 高 {bs['stud_h']*100:.1f} cm (間距 {cfg['beam_side_spacing']['stud_spacing']:.2f} m, 跨距 {cfg['beam_side_spacing']['stud_span']:.2f} m)\n"
        f"• 側外檔 (Side Waler)：寬 {bs['waler_b']*100:.1f} cm × 高 {bs['waler_h']*100:.1f} cm (間距 {cfg['beam_side_spacing']['waler_spacing']:.2f} m, 跨距 {cfg['beam_side_spacing']['waler_span']:.2f} m)\n"
        f"• 樑側 Form Tie 螺桿：直徑 {bs['side_tie_dia_mm']} mm (水平 Sx = {cfg['beam_side_spacing']['tie_sx']:.2f} m, 垂直 Sy = {cfg['beam_side_spacing']['tie_sy']:.2f} m, 容許拉力 {cfg['capacity']['side_tie_allowable']:.1f} kN)\n"
        f"• 底面板 (Bot Panel) 厚度：{bb['panel_t_mm']} mm (設計跨距 {cfg['beam_bot_spacing']['panel_span']:.2f} m)\n"
        f"• 底外橫檔 (Bot Waler)：{waler_b_str} (跨距採 B-5cm = {cfg['beam_bot_spacing']['waler_span']:.2f} m, 間距 {cfg['beam_bot_spacing']['waler_spacing']:.2f} m)\n"
        f"• Steel Prop 鋼管支撐：設計容許承載力 Pa = 10.0 kN ({prop_type_str}, 縱向佈置跨距 = {cfg['beam_bot_spacing']['prop_span']:.2f} m)"
    )
    
    doc.add_heading("三、牆模板系統深度分層檢核", 1)
    table_w = doc.add_table(rows=1, cols=7)
    table_w.style = 'Table Grid'
    headers_w = ["深度 (m)", "側壓(kN/m²)", "Panel利用率", "Stud利用率", "Waler利用率", "Tie螺桿", "結果"]
    for i, text in enumerate(headers_w):
        table_w.rows[0].cells[i].text = text
        
    for _, r in df_wall.iterrows():
        row = table_w.add_row().cells
        row[0].text = f"{r['level']:.2f}"
        row[1].text = f"{r['p']:.2f}"
        row[2].text = f"{r['panel']:.3f}"
        row[3].text = f"{r['stud']:.3f}"
        row[4].text = f"{r['waler']:.3f}"
        row[5].text = f"{r['tie']:.3f}"
        row[6].text = r["Result"]
        
    doc.add_heading("四、樑側與樑底模板檢核 (複合型材料與智慧支撐架構)", 1)
    doc.add_paragraph(f"💡 樑側設計側壓：p_side = γc * D = {cfg['beam_calc_params']['p_side']:.2f} kN/m²。")
    doc.add_paragraph(f"💡 樑底設計垂直荷重：p_bot = {cfg['beam_calc_params']['p_bot']:.2f} kN/m²。")
    
    # 樑側表格輸出
    doc.add_heading("1. 樑側模結構檢核結果表 (Side Form)", 2)
    t_side = doc.add_table(rows=1, cols=6)
    t_side.style = 'Table Grid'
    headers_s = ["構件項目", "尺寸規格", "彎曲利用率", "剪力利用率", "撓度利用率", "綜合判定"]
    for i, text in enumerate(headers_s):
        t_side.rows[0].cells[i].text = text
        
    for name, size, r_b, r_s, r_d, r_m in [
        ("Panel", f"t={bs['panel_t_mm']} mm", side_data['Panel']['r_bending'], side_data['Panel']['r_shear'], side_data['Panel']['r_deflection'], side_data['Panel']['r_max']),
        ("Joist", f"{bs['stud_b']*100:.1f}x{bs['stud_h']*100:.1f} cm", side_data['Joist']['r_bending'], side_data['Joist']['r_shear'], side_data['Joist']['r_deflection'], side_data['Joist']['r_max']),
        ("Waler", f"{bs['waler_b']*100:.1f}x{bs['waler_h']*100:.1f} cm", side_data['Waler']['r_bending'], side_data['Waler']['r_shear'], side_data['Waler']['r_deflection'], side_data['Waler']['r_max']),
        ("Form Tie", f"Φ {bs['side_tie_dia_mm']} mm", 0.0, 0.0, side_data['Form_Tie'], side_data['Form_Tie'])
    ]:
        row = t_side.add_row().cells
        row[0].text = name
        row[1].text = size
        row[2].text = f"{r_b:.2f}" if name != "Form Tie" else "-"
        row[3].text = f"{r_s:.2f}" if name != "Form Tie" else "-"
        row[4].text = f"{r_d:.2f}"
        row[5].text = "OK" if r_m <= 1.0 else "NG"

    # 樑底表格輸出
    doc.add_heading("2. 樑底模結構檢核結果表 (Bottom Form)", 2)
    t_bot = doc.add_table(rows=1, cols=6)
    t_bot.style = 'Table Grid'
    for i, text in enumerate(headers_s):
        t_bot.rows[0].cells[i].text = text
        
    for name, size, r_b, r_s, r_d, r_m in [
        ("Panel", f"t={bb['panel_t_mm']} mm", bot_data['Panel']['r_bending'], bot_data['Panel']['r_shear'], bot_data['Panel']['r_deflection'], bot_data['Panel']['r_max']),
        ("Joist", f"{bb['stud_b']*100:.1f}x{bb['stud_h']*100:.1f} cm", bot_data['Joist']['r_bending'], bot_data['Joist']['r_shear'], bot_data['Joist']['r_deflection'], bot_data['Joist']['r_max']),
        ("Waler", f"{bb['waler_type']} (L=B-5cm)", bot_data['Waler']['r_bending'], bot_data['Waler']['r_shear'], bot_data['Waler']['r_deflection'], bot_data['Waler']['r_max']),
        ("Steel Prop", f"Pa=10kN ({'單管全荷' if bg['beam_b']<=0.3 else '雙管分荷'})", 0.0, 0.0, bot_data['Steel_Prop'], bot_data['Steel_Prop'])
    ]:
        row = t_bot.add_row().cells
        row[0].text = name
        row[1].text = size
        row[2].text = f"{r_b:.2f}" if name != "Steel Prop" else "-"
        row[3].text = f"{r_s:.2f}" if name != "Steel Prop" else "-"
        row[4].text = f"{r_d:.2f}"
        row[5].text = "OK" if r_m <= 1.0 else "NG"

        doc.add_heading("五、結論", 1)
        wall_final_ok = (df_wall['Result'] == "OK").all()
        side_final_ok = (side_data['Panel']['r_max'] <= 1.0) and (side_data['Form_Tie'] <= 1.0)
        bot_final_ok = (bot_data['Panel']['r_max'] <= 1.0) and (bot_data['Steel_Prop'] <= 1.0)
        all_final_ok = wall_final_ok and side_final_ok and bot_final_ok

        if all_final_ok:
            doc.add_paragraph("本工程牆模板與大樑模板支撐結構(含牆樑外側鋼支撐與智慧型樑底雙排鋼管分荷系統)，經獨立演算強度均安全達標，准予送審施工。")
        else:
            doc.add_paragraph("⚠️ 注意：部分複合構件利用率已超載 (NG)，請現場重整排列間距或更改輕型鋼厚度規格。")

        # 自動備份至指定本機工作目錄 
        target_dir = r"D:\PythonTraining\Beam and Wall _Formwork_app\CMD Beam and Wall_Formwork_App"

        # 1. 先嘗試進行本機實體存檔
        try:
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
            doc.save(os.path.join(target_dir, "牆與梁模結構計算書（送審版）.docx"))
        except Exception:
            pass # 防止目錄權限問題導致網頁中斷

        # 2. 正常執行：不論本機存檔成功與否，皆產出記憶體串流供網頁端下載
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
# ==========================================
# 3. Streamlit 網頁主介面佈局
# ==========================================

st.title("🏗️ 牆與樑模板支撐結構檢核系統")
st.markdown("##### 智慧分荷與輕型鋼複合選單網頁版 App")

tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ 工程資訊 & 施工條件", 
    "2️⃣ 🧱 牆模板系統配置", 
    "3️⃣ 📐 樑側模板系統配置", 
    "4️⃣ 🔩 樑底模板系統配置"
])

with tab1:
    st.header("工程基本資訊")
    p_name = st.text_input("工程名稱:", value="牆與樑模板系統結構計算報告(送審版)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("混凝土與施工條件")
        h = st.number_input("牆最大高度 H (m):", min_value=0.0, value=3.5, step=0.1)
        r = st.number_input("澆置速度 R (m/hr):", min_value=0.0, value=2.0, step=0.1)
        t_c = st.number_input("設計溫度 T (°C):", min_value=-10.0, value=25.0, step=0.5)
        g_c = st.number_input("混凝土重 γc (kN/m³):", min_value=0.0, value=24.0, step=0.5)
    
    with col2:
        st.subheader("木材基本強度參數 (kN/m²)")
        E = st.number_input("彈性模數 E:", min_value=0.0, value=1.0e7, format="%.1e")
        fb = st.number_input("容許抗彎 fb:", min_value=0.0, value=8000.0, step=100.0)
        fb_w = st.number_input("橫檔抗彎 fb_waler:", min_value=0.0, value=10000.0, step=100.0)
        fv = st.number_input("容許剪力 fv:", min_value=0.0, value=1000.0, step=50.0)

with tab2:
    st.header("🧱 牆模板系統配置")
    col1, col2 = st.columns(2)
    with col1:
        w_panel_t_mm = float(st.selectbox("牆面板厚度 (mm):", ["12", "15", "18", "21"], index=1))
        w_p_span = st.number_input("牆面板跨距 (m):", value=0.30, step=0.05)
        w_s_b = st.number_input("牆角材寬度 b (cm):", value=9.0, step=0.5) / 100.0
        w_s_h = st.number_input("牆角材高度 h (cm):", value=4.5, step=0.5) / 100.0
        w_s_space = st.number_input("牆角材間距 (m):", value=0.30, step=0.05)
        w_s_span = st.number_input("牆角材跨距 (m):", value=0.45, step=0.05)
        
    with col2:
        waler_w_type = st.selectbox("牆橫檔種類 (Waler):", ["木橫檔 (自訂尺寸)", "C100x50x2.0 (輕型鋼)", "C100x50x2.3 (輕型鋼)"], index=0)
        w_w_b = st.number_input("牆橫檔寬度 b (cm):", value=4.5, step=0.5) / 100.0
        w_w_h = st.number_input("牆橫檔高度 h (cm):", value=9.0, step=0.5) / 100.0
        w_w_space = st.number_input("牆橫檔間距 (m):", value=0.45, step=0.05)
        w_w_span = st.number_input("牆橫檔跨距 (m):", value=0.60, step=0.05)
        
    st.subheader("牆螺桿與檢核深度設定")
    col3, col4 = st.columns(2)
    with col3:
        w_tx = st.number_input("牆螺桿水平 Sx (m):", value=0.45, step=0.05)
        w_ty = st.number_input("牆螺桿垂直 Sy (m):", value=0.60, step=0.05)
        w_tie_dia = st.number_input("牆螺桿直徑 (mm):", value=12.0, step=1.0)
        w_tie_allow = st.number_input("牆螺桿容許拉力(kN):", value=14.0, step=1.0)
    with col4:
        levels_str = st.text_input("牆檢核深度(m, 逗號隔開):", value="0.9, 1.8, 2.7, 3.2")
        levels = [float(x.strip()) for x in levels_str.split(",") if x.strip()]

with tab3:
    st.header("📐 樑側模板系統配置")
    col1, col2 = st.columns(2)
    with col1:
        beam_b_cm = st.number_input("大樑寬 B (cm):", value=40.0, step=5.0)
        beam_d_cm = st.number_input("大樑深 D (cm):", value=60.0, step=5.0)
        beam_b = beam_b_cm / 100.0
        beam_d = beam_d_cm / 100.0
        
        bs_panel_t_mm = float(st.selectbox("樑側面板厚度 (mm):", ["12", "15", "18", "21"], index=1))
        bs_p_span = st.number_input("樑側面板跨距 (m):", value=0.30, step=0.05)
        bs_s_b = st.number_input("樑側角材寬 b (cm):", value=9.0, step=0.5) / 100.0
        bs_s_h = st.number_input("樑側角材高 h (cm):", value=4.5, step=0.5) / 100.0
        
    with col2:
        bs_s_space = st.number_input("樑側角材間距 (m):", value=0.30, step=0.05)
        bs_s_span = st.number_input("樑側角材跨距 (m):", value=0.40, step=0.05)
        bs_w_b = st.number_input("樑側橫檔寬 b (cm):", value=4.5, step=0.5) / 100.0
        bs_w_h = st.number_input("樑側橫檔高 h (cm):", value=9.0, step=0.5) / 100.0
        bs_w_space = st.number_input("樑側橫檔間距 (m):", value=0.40, step=0.05)
        bs_w_span = st.number_input("樑側橫檔跨距 (m):", value=0.55, step=0.05)

    st.subheader("樑側螺桿設定")
    col3, col4 = st.columns(2)
    with col3:
        bs_tx = st.number_input("樑側螺桿水平 Sx (m):", value=0.40, step=0.05)
        bs_ty = st.number_input("樑側螺桿垂直 Sy (m):", value=0.55, step=0.05)
    with col4:
        bs_tie_dia = st.number_input("樑側螺桿直徑 (mm):", value=12.0, step=1.0)
        bs_tie_allow = st.number_input("樑側螺桿容許拉力 (kN):", value=14.0, step=1.0)

with tab4:
    st.header("🔩 樑底模板系統配置 (含鋼管支撐 Pa=10kN)")
    col1, col2 = st.columns(2)
    with col1:
        bb_panel_t_mm = float(st.selectbox("樑底面板厚度 (mm):", ["12", "15", "18", "21"], index=1))
        bb_p_span = st.number_input("樑底面板跨距 (m):", value=0.30, step=0.05)
        bb_s_b = st.number_input("樑底角材寬 b (cm):", value=9.0, step=0.5) / 100.0
        bb_s_h = st.number_input("樑底角材高 h (cm):", value=4.5, step=0.5) / 100.0
        bb_s_space = st.number_input("樑底角材間距 (m):", value=0.30, step=0.05)
        bb_s_span = st.number_input("樑底角材跨距 (m):", value=0.60, step=0.05)
        
    with col2:
        combo_bot_waler = st.selectbox("樑底橫檔種類 (Waler):", ["木橫檔 (自訂尺寸)", "C100x50x2.0 (輕型鋼)", "C100x50x2.3 (輕型鋼)"], index=0)
        bb_w_b = st.number_input("樑底橫檔寬 b (cm):", value=4.5, step=0.5) / 100.0
        bb_w_h = st.number_input("樑底橫檔高 h (cm):", value=9.0, step=0.5) / 100.0
        bb_w_space = st.number_input("樑底橫檔間距 (m):", value=0.60, step=0.05)
        bb_prop_span = st.number_input("鋼管支撐跨距 (m):", value=0.60, step=0.05)
        
        st.info("💡 **系統智慧防錯**：樑底橫檔結構跨距將由大樑寬 B - 5cm 全自動反算")
        bb_waler_span_actual = max(0.05, beam_b - 0.05)
# ==========================================
# 4. 網頁分析觸發按鈕與渲染邏輯
# ==========================================

st.markdown("---")
if st.button("🚀 執行全獨立三軌鋼木複合檢核並生成送審報告", type="primary", width="stretch"):
    try:
        # 面板厚度轉換為公尺
        w_p_t = w_panel_t_mm / 1000.0
        bs_p_t = bs_panel_t_mm / 1000.0
        bb_p_t = bb_panel_t_mm / 1000.0

        # 封裝全參數 Config
        cfg = {
            "project_info": {"name": p_name},
            "wall_params": {"height": h, "pour_rate": r, "temperature": t_c, "gamma_c": g_c},
            "beam_geo": {"beam_b": beam_b, "beam_d": beam_d},
            "material": {"E": E, "fb_standard": fb, "fb_waler": fb_w, "fv_standard": fv},
            "wall_spacing": {"panel_span": w_p_span, "stud_spacing": w_s_space, "stud_span": w_s_span, "waler_spacing": w_w_space, "waler_span": w_w_span, "tie_sx": w_tx, "tie_sy": w_ty},
            "wall_section": {"panel_t_mm": w_panel_t_mm, "panel_t": w_p_t, "stud_b": w_s_b, "stud_h": w_s_h, "waler_b": w_w_b, "waler_h": w_w_h, "tie_dia_mm": w_tie_dia, "waler_type": waler_w_type},
            "beam_side_spacing": {"panel_span": bs_p_span, "stud_spacing": bs_s_space, "stud_span": bs_s_span, "waler_spacing": bs_w_space, "waler_span": bs_w_span, "tie_sx": bs_tx, "tie_sy": bs_ty},
            "beam_side_section": {"panel_t_mm": bs_panel_t_mm, "panel_t": bs_p_t, "stud_b": bs_s_b, "stud_h": bs_s_h, "waler_b": bs_w_b, "waler_h": bs_w_h, "side_tie_dia_mm": bs_tie_dia},
            "beam_bot_spacing": {"panel_span": bb_p_span, "stud_spacing": bb_s_space, "stud_span": bb_s_span, "waler_spacing": bb_w_space, "waler_span": bb_waler_span_actual, "prop_span": bb_prop_span},
            "beam_bot_section": {"panel_t_mm": bb_panel_t_mm, "panel_t": bb_p_t, "stud_b": bb_s_b, "stud_h": bb_s_h, "waler_b": bb_w_b, "waler_h": bb_w_h, "waler_type": combo_bot_waler},
            "capacity": {"tie_allowable": w_tie_allow, "side_tie_allowable": bs_tie_allow, "prop_allowable": 10.0}
        }

        # 牆體分層檢核
        results_wall = []
        for lv in levels:
            p_lv = calc_wall_pressure_aci(lv, r, t_c, g_c)
            rp = check_generic_member(p_lv, 1.0, w_p_span, 1.0, w_p_t, E, fb, fv, "panel")["r_max"]
            rs = check_generic_member(p_lv, w_s_space, w_s_span, w_s_b, w_s_h, E, fb, fv, "stud")["r_max"]
            rw = check_composite_waler(p_lv, w_w_space, w_w_span, waler_w_type, w_w_b, w_w_h, E, fb_w, fv)["r_max"]
            rt = (p_lv * w_tx * w_ty) / w_tie_allow
            max_r = max(rp, rs, rw, rt)
            results_wall.append({
                "level": lv, "p": p_lv, "panel": rp, "stud": rs, "waler": rw, "tie": rt,
                "Result": "OK" if max_r <= 1.0 else "NG"
            })
        df_wall = pd.DataFrame(results_wall)

        # 大樑檢核與智慧型鋼管支撐 1/2 荷重分配判定
        p_beam_side = calc_beam_side_pressure(beam_d, g_c)
        p_beam_bot = calc_beam_bottom_load(beam_d, g_c, live_load=2.5)
        cfg["beam_calc_params"] = {"p_side": p_beam_side, "p_bot": p_beam_bot}
        
        # 智慧分荷係數判定 (B <= 30cm 採單管全荷=1.0, B >= 40cm 採雙排分荷=0.5)
        prop_load_factor = 1.0 if beam_b <= 0.3 else 0.5

        side_data = {
            "Panel": check_generic_member(p_beam_side, 1.0, bs_p_span, 1.0, bs_p_t, E, fb, fv, "panel"),
            "Joist": check_generic_member(p_beam_side, bs_s_space, bs_s_span, bs_s_b, bs_s_h, E, fb, fv, "stud"),
            "Waler": check_composite_waler(p_beam_side, bs_w_space, bs_w_span, "木橫檔 (自訂尺寸)", bs_w_b, bs_w_h, E, fb_w, fv),
            "Form_Tie": (p_beam_side * bs_tx * bs_ty) / bs_tie_allow
        }

        bot_data = {
            "Panel": check_generic_member(p_beam_bot, 1.0, bb_p_span, 1.0, bb_p_t, E, fb, fv, "panel"),
            "Joist": check_generic_member(p_beam_bot, bb_s_space, bb_s_span, bb_s_b, bb_s_h, E, fb, fv, "stud"),
            "Waler": check_composite_waler(p_beam_bot, bb_w_space, bb_waler_span_actual, combo_bot_waler, bb_w_b, bb_w_h, E, fb_w, fv, is_beam_bottom=True),
            "Steel_Prop": (p_beam_bot * beam_b * bb_prop_span * prop_load_factor) / 10.0
        }

        # 渲染網頁結果
        st.success("✅ 結構強度檢核計算完成！")
        
        # 1. 牆模板結果
        st.subheader("📋 三、牆模板系統深度分層檢核結果")
        st.dataframe(df_wall.style.format({
            "p": "{:.2f}", "panel": "{:.3f}", "stud": "{:.3f}", "waler": "{:.3f}", "tie": "{:.3f}"
        }), width="stretch")

        # 2. 樑側與樑底結論整理成 DataFrame 於網頁展示
        st.subheader("📋 四、樑側與樑底模板檢核結果 (複合型材料)")
        col_res1, col_res2 = st.columns(2)
        
        with col_res1:
            st.markdown("**1. 樑側模結構檢核表 (Side Form)**")
            side_summary = [
                {"構件項目": "Panel", "尺寸規格": f"t={bs_panel_t_mm} mm", "最大利用率": f"{side_data['Panel']['r_max']:.2f}"},
                {"構件項目": "Joist", "尺寸規格": f"{bs_s_b*100:.1f}x{bs_s_h*100:.1f} cm", "最大利用率": f"{side_data['Joist']['r_max']:.2f}"},
                {"構件項目": "Waler", "尺寸規格": f"{bs_w_b*100:.1f}x{bs_w_h*100:.1f} cm", "最大利用率": f"{side_data['Waler']['r_max']:.2f}"},
                {"構件項目": "Form Tie", "尺寸規格": f"Φ {bs_tie_dia} mm", "最大利用率": f"{side_data['Form_Tie']:.2f}"}
            ]
            st.dataframe(pd.DataFrame(side_summary), width="stretch")

        with col_res2:
            st.markdown("**2. 樑底模結構檢核表 (Bottom Form)**")
            bot_summary = [
                {"構件項目": "Panel", "尺寸規格": f"t={bb_panel_t_mm} mm", "最大利用率": f"{bot_data['Panel']['r_max']:.2f}"},
                {"構件項目": "Joist", "尺寸規格": f"{bb_s_b*100:.1f}x{bb_s_h*100:.1f} cm", "最大利用率": f"{bot_data['Joist']['r_max']:.2f}"},
                {"構件項目": "Waler", "尺寸規格": f"{combo_bot_waler}", "最大利用率": f"{bot_data['Waler']['r_max']:.2f}"},
                {"構件項目": "Steel Prop", "尺寸規格": f"Pa=10kN ({'單管' if beam_b<=0.3 else '雙管'})", "最大利用率": f"{bot_data['Steel_Prop']:.2f}"}
            ]
            st.dataframe(pd.DataFrame(bot_summary), width="stretch")

        # 3. 綜合結論宣告
        st.subheader("💡 五、綜合判定結論")
        wall_final_ok = (df_wall['Result'] == "OK").all()
        side_final_ok = (side_data['Panel']['r_max'] <= 1.0) and (side_data['Form_Tie'] <= 1.0)
        bot_final_ok = (bot_data['Panel']['r_max'] <= 1.0) and (bot_data['Steel_Prop'] <= 1.0)
        all_final_ok = wall_final_ok and side_final_ok and bot_final_ok
        
        if all_final_ok:
            st.balloons()
            st.info("🟢 本工程牆模板與大樑模板支撐結構，經獨立演算強度均安全達標，准予送審施工。")
        else:
            st.error("⚠️ 注意：部分複合構件利用率已超載 (NG)，請現場重整排列間距或更改輕型鋼厚度規格。")

        # 4. 原生線上記憶體下載按鈕 (免除本機路徑存檔權限問題)
        st.subheader("📥 下載審查文件夾")
        col_dl1, col_dl2 = st.columns(2)
        
        word_stream = export_word_report_stream(cfg, df_wall, side_data, bot_data)
        with col_dl1:
            st.download_button(
                label="📝 下載 Word 送審計算書",
                data=word_stream,
                file_name="牆與樑模結構計算書（送審版）.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch"
            )
            
        excel_stream = io.BytesIO()
        df_wall.to_excel(excel_stream, index=False)
        excel_stream.seek(0)
        with col_dl2:
            st.download_button(
                label="📊 下載 Excel 牆體摘要數據",
                data=excel_stream,
                file_name="牆模板檢核摘要.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch"
            )

    except Exception as e:
        st.error(f"❌ 發生錯誤，請確認輸入欄位格式是否正確。錯誤訊息: {str(e)}")    