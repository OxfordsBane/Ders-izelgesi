import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V45 - Esnek Çözüm", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V45 - Danışman Açığı Çözümü)")
st.warning("""
**Durum Analizi:**
Elinizdeki 'Danışman' sayısı, toplam sınıf sayısından az. Bu yüzden sistem bazı sınıflara 'Destek' hocalarını Danışman olarak atamak zorunda kalacak.
Bu sürümde, boş ders kalmaması için tüm kurallar esnetilebilir hale getirildi.
""")

# --- YAN PANEL ---
st.sidebar.header("⚙️ Genel Ayarlar")
max_teachers_per_class = st.sidebar.slider("Sınıf Başına Max Hoca", 1, 6, 3)
allow_native_advisor = st.sidebar.checkbox("Native Hocalar Danışman Olabilir mi?", value=False)
allow_empty_slots = st.sidebar.checkbox("Sıkışınca Boş Ders Bırak", value=True)

st.sidebar.markdown("---")
st.sidebar.header("🏫 Sınıf ve Zaman Ayarları")

col1, col2 = st.sidebar.columns(2)
with col1:
    count_a1 = st.number_input("A1 Sayısı", 0, 20, 4)
    time_a1 = st.selectbox("A1 Zamanı", ["Sabah", "Öğle"], key="t_a1")
    count_a2 = st.number_input("A2 Sayısı", 0, 20, 4)
    time_a2 = st.selectbox("A2 Zamanı", ["Sabah", "Öğle"], key="t_a2")
    count_pre = st.number_input("PreFac Sayısı", 0, 10, 0)
    time_pre = st.selectbox("PreFac Zamanı", ["Sabah", "Öğle"], key="t_pre")

with col2:
    count_b1 = st.number_input("B1 Sayısı", 0, 20, 4)
    time_b1 = st.selectbox("B1 Zamanı", ["Sabah", "Öğle"], key="t_b1")
    count_b2 = st.number_input("B2 Sayısı", 0, 20, 2)
    time_b2 = st.selectbox("B2 Zamanı", ["Sabah", "Öğle"], key="t_b2")

# --- SINIF OLUŞTURMA ---
def create_automated_classes():
    class_list = []
    config = [
        (count_a1, "A1", 0 if time_a1 == "Sabah" else 1),
        (count_a2, "A2", 0 if time_a2 == "Sabah" else 1),
        (count_b1, "B1", 0 if time_b1 == "Sabah" else 1),
        (count_b2, "B2", 0 if time_b2 == "Sabah" else 1),
        (count_pre, "PreFaculty", 0 if time_pre == "Sabah" else 1),
    ]
    for count, lvl, time_code in config:
        for i in range(1, count + 1):
            class_name = f"{lvl}.{i:02d}"
            class_list.append({"Sınıf Adı": class_name, "Seviye": lvl, "Zaman Kodu": time_code})
    return pd.DataFrame(class_list)

# --- EXCEL ŞABLONU ---
def generate_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_teachers = pd.DataFrame({
            'Ad Soyad': ['Ahmet Hoca', 'Sarah (Native)', 'Mehmet (Danışman)', 'Ayşe Hoca'],
            'Rol': ['Destek', 'Native', 'Danışman', 'Ek Görevli'],
            'Hedef Ders Sayısı': [4, 4, 3, 2],
            'Tercih (Sabah/Öğle)': ['Sabah', 'Farketmez', 'Sabah', 'Öğle'],
            'Yasaklı Günler': ['Cuma', 'Çarşamba', '', 'Pazartesi,Salı'],
            'Sabit Sınıf': ['', '', 'A1.01', ''],
            'Yetkinlik (Seviyeler)': ['A1,A2,B1', 'Hepsi', 'A1,A2', 'B1,B2'],
            'İstenmeyen Partner': ['', '', 'Ayşe Hoca', 'Mehmet (Danışman)']
        })
        df_teachers.to_excel(writer, sheet_name='Ogretmenler', index=False)
        
        workbook = writer.book
        worksheet = workbook.add_worksheet('NASIL KULLANILIR')
        header_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#D3D3D3', 'border': 1})
        text_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        worksheet.write('A1', 'PROGRAM KULLANIM KILAVUZU', header_fmt)
        worksheet.set_column('A:A', 100)
        
        instructions = [
            "1. ROL SÜTUNU NEDİR?",
            "   - Destek: İhtiyaç halinde her yere girebilirler. Danışman açığını kapatırlar.",
            "   - Native: Yabancı hocalar. A1 seviyesine girmezler.",
            "   - Danışman: Sınıf sorumlularıdır.",
            "   - Ek Görevli: İdari/Özel görevi olanlar. Sınıf Danışmanı olamazlar.",
            "",
            "2. SÜTUNLAR NASIL DOLDURULUR?",
            "   - Hedef Ders Sayısı: Hocanın o hafta gireceği toplam 'oturum' sayısı.",
            "   - Tercih: 'Sabah', 'Öğle' veya 'Farketmez'.",
            "   - Yasaklı Günler: Hoca o gün ASLA gelmez. Virgülle ayırın.",
            "   - Sabit Sınıf: Hocanın kesin atanacağı sınıf (Örn: A2.01).",
            "   - Yetkinlik: Hocanın girebileceği seviyeler. 'Hepsi' yazarsanız her yere girer.",
        ]
        row = 1
        for line in instructions:
            worksheet.write(row, 0, line, text_fmt)
            row += 1
            
    return output.getvalue()

st.sidebar.markdown("---")
st.sidebar.download_button("📥 Kılavuzlu Şablonu İndir", generate_template(), "ogretmen_listesi.xlsx")

# --- ANALİZ ---
def analyze_data(teachers, classes):
    warnings = []
    errors = []
    
    for t in teachers:
        role = str(t['Rol']).upper()
        fixed_class = str(t['Sabit Sınıf']).strip()
        
        if "DESTEK" in role and fixed_class:
             # Destek hocası artık sabit sınıf alabilir (Danışman açığını kapatmak için)
             pass 
             
        if not allow_native_advisor:
            if "NATIVE" in role and fixed_class:
                errors.append(f"🛑 **{t['Ad Soyad']}**: Native hocaların danışman olması engellendi. (Sabit Sınıfı silin).")

        if fixed_class:
            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class), None)
            if not target_class:
                errors.append(f"❌ **{t['Ad Soyad']}**: Atandığı '{fixed_class}' sınıfı sistemde yok.")

    return errors, warnings

# --- ANA PROGRAM ---
uploaded_file = st.file_uploader("Öğretmen Listesini Yükle", type=["xlsx"])

if uploaded_file:
    df_teachers = pd.read_excel(uploaded_file, sheet_name='Ogretmenler').fillna("")
    if 'Hedef Ders Sayısı' not in df_teachers.columns and 'Hedef Gün Sayısı' in df_teachers.columns:
        df_teachers.rename(columns={'Hedef Gün Sayısı': 'Hedef Ders Sayısı'}, inplace=True)
        
    df_classes = create_automated_classes()
    
    teachers_list = df_teachers.to_dict('records')
    classes_list = df_classes.to_dict('records')

    logic_errors, logic_warnings = analyze_data(teachers_list, classes_list)
    
    if logic_errors:
        st.error("🛑 Lütfen aşağıdaki hataları düzeltip dosyayı tekrar yükleyin:")
        for e in logic_errors: st.markdown(e)
    else:
        # --- İSTATİSTİKLERİ HESAPLA ---
        total_needed = len(classes_list) * 5
        total_cap = sum(t['Hedef Ders Sayısı'] for t in teachers_list)
        
        num_danisman = sum(1 for t in teachers_list if 'Danışman' in str(t['Rol']))
        num_native = sum(1 for t in teachers_list if 'Native' in str(t['Rol']))
        num_destek = sum(1 for t in teachers_list if 'Destek' in str(t['Rol']))
        num_ek = sum(1 for t in teachers_list if 'Ek Görevli' in str(t['Rol']))
        
        num_sabah_sinif = sum(1 for c in classes_list if c['Zaman Kodu'] == 0)
        num_ogle_sinif = sum(1 for c in classes_list if c['Zaman Kodu'] == 1)
        
        # --- İSTATİSTİK PANELİ ---
        st.markdown("### 📊 Durum Analizi")
        
        # 1. Satır: Kapasite
        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam Sınıf", len(classes_list))
        c2.metric("İhtiyaç Duyulan Ders", total_needed)
        c3.metric("Hoca Kapasitesi", total_cap, delta=total_cap - total_needed)
        
        # 2. Satır: Sınıf Dağılımı ve Kadro
        c4, c5, c6, c7, c8, c9 = st.columns(6)
        c4.metric("Sabah Sınıfı", num_sabah_sinif)
        c5.metric("Öğle Sınıfı", num_ogle_sinif)
        c6.metric("Danışman", num_danisman)
        c7.metric("Native", num_native)
        c8.metric("Destek", num_destek)
        c9.metric("Ek Görevli", num_ek)
        
        st.divider()

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Optimizasyon yapılıyor... (Boşlukları doldurmak için kurallar esnetiliyor...)"):
                
                model = cp_model.CpModel()
                days = range(5)
                day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                sessions = range(2)
                
                x = {}
                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        for d in days:
                            for s in sessions:
                                x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                # --- KISITLAMALAR ---
                for c_idx, c_data in enumerate(classes_list):
                    req_session = c_data['Zaman Kodu']
                    other_session = 1 - req_session
                    for d in days:
                        if allow_empty_slots:
                            model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
                        else:
                            model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 1)
                        model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)

                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)
                
                for t_idx, t in enumerate(teachers_list):
                    target = int(t['Hedef Ders Sayısı'])
                    total_assignments = []
                    for c in range(len(classes_list)):
                        for d in days:
                            for s in sessions: total_assignments.append(x[(t_idx, c, d, s)])
                    model.Add(sum(total_assignments) <= target)

                for c_idx in range(len(classes_list)):
                    teachers_here = []
                    for t in range(len(teachers_list)):
                        teach = model.NewBoolVar(f'tch_{t}_{c_idx}')
                        model.AddMaxEquality(teach, [x[(t, c_idx, d, s)] for d in days for s in sessions])
                        teachers_here.append(teach)
                    model.Add(sum(teachers_here) <= max_teachers_per_class)

                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c_idx, c_data in enumerate(classes_list):
                            if c_data['Seviye'] == 'A1':
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

                for c_idx, c_data in enumerate(classes_list):
                    natives_in_class = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_{t_idx}_{c_idx}')
                            model.AddMaxEquality(is_present, [x[(t_idx, c_idx, d, s)] for d in days for s in sessions])
                            natives_in_class.append(is_present)
                    model.Add(sum(natives_in_class) <= 1) 

                for t_idx, t in enumerate(teachers_list):
                    if 'Ek Görevli' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            lessons_in_class = []
                            for d in days:
                                for s in sessions:
                                    lessons_in_class.append(x[(t_idx, c_idx, d, s)])
                            model.Add(sum(lessons_in_class) <= 1)

                # --- PUANLAMA ---
                objective = []
                # Temel ders atama puanını çok artırdık: Boş ders kalmasın!
                objective.append(sum(x.values()) * 10000000) 

                for t_idx, t in enumerate(teachers_list):
                    fill_score = 100000
                    if 'Ek Görevli' in str(t['Rol']): fill_score = 50000 
                    objective.append(sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions]) * fill_score)

                for t_idx, t in enumerate(teachers_list):
                    target = int(t['Hedef Ders Sayısı'])
                    current_load = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    # Danışman veya Destek fark etmez, herkesi doldurmaya çalış
                    objective.append(current_load * 50000)

                # ZAMAN TERCİHİ (Yumuşatıldı - Sınıf dolması daha önemli)
                for t_idx, t in enumerate(teachers_list):
                    pref = str(t['Tercih (Sabah/Öğle)'])
                    if pref == "Sabah":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 1)] * -500000) # Ceza düşürüldü
                    elif pref == "Öğle":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 0)] * -500000)

                # YASAKLI GÜNLER (Yumuşatıldı - Zorundaysan del)
                for t_idx, t in enumerate(teachers_list):
                    forbidden = str(t['Yasaklı Günler'])
                    for d_idx, d_name in enumerate(day_names):
                        if d_name in forbidden:
                            for c in range(len(classes_list)):
                                for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -1000000)

                for t_idx, t in enumerate(teachers_list):
                    allowed = str(t['Yetkinlik (Seviyeler)'])
                    if "Hepsi" not in allowed:
                        for c_idx, c in enumerate(classes_list):
                            if c['Seviye'] not in allowed:
                                for d in days:
                                    for s in sessions: objective.append(x[(t_idx, c, d, s)] * -40000000)

                for t_idx, t in enumerate(teachers_list):
                    if t['Sabit Sınıf']:
                        fixed_c = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c is not None:
                            req_s = classes_list[fixed_c]['Zaman Kodu']
                            days_in_class = []
                            for d in days:
                                present = model.NewBoolVar(f'pres_{t_idx}_{d}')
                                model.AddMaxEquality(present, [x[(t_idx, fixed_c, d, s)] for s in sessions])
                                days_in_class.append(present)
                            
                            # Sabit sınıfı varsa orada tutmaya çalış
                            objective.append(sum(days_in_class) * 5000000)
                            objective.append(x[(t_idx, fixed_c, 0, req_s)] * 5000000)

                for t_idx, t in enumerate(teachers_list):
                    unw = str(t['İstenmeyen Partner'])
                    if len(unw) > 2:
                        p_idx = next((i for i, tea in enumerate(teachers_list) if tea['Ad Soyad'] == unw), None)
                        if p_idx:
                            for c in range(len(classes_list)):
                                t1 = model.NewBoolVar(f't1_{c}')
                                t2 = model.NewBoolVar(f't2_{c}')
                                model.AddMaxEquality(t1, [x[(t_idx, c, d, s)] for d in days for s in sessions])
                                model.AddMaxEquality(t2, [x[(p_idx, c, d, s)] for d in days for s in sessions])
                                conflict = model.NewBoolVar(f'conflict_{t_idx}_{c}')
                                model.Add(t1 + t2 == 2).OnlyEnforceIf(conflict)
                                model.Add(t1 + t2 < 2).OnlyEnforceIf(conflict.Not())
                                objective.append(conflict * -3000000)
                
                # DESTEK PAZARTESİ (Ceza neredeyse yok)
                for t_idx, t in enumerate(teachers_list):
                    if 'Destek' in str(t['Rol']):
                        for c in range(len(classes_list)):
                            for s in sessions: objective.append(x[(t_idx, c, 0, s)] * -1000)

                # --- ÇÖZÜM ---
                model.Maximize(sum(objective))
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 120.0
                status = solver.Solve(model)

                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.balloons()
                    
                    violations = []
                    res_data = []
                    native_names = [t['Ad Soyad'] for t in teachers_list if 'Native' in str(t['Rol'])]
                    
                    advisor_map = {t['Sabit Sınıf']: t['Ad Soyad'] for t in teachers_list if t['Sabit Sınıf']}

                    for c_idx, c in enumerate(classes_list):
                        c_name = c['Sınıf Adı']
                        s_req = c['Zaman Kodu']
                        
                        teacher_counts = {}
                        for t_idx, t in enumerate(teachers_list):
                            count = 0
                            for d_idx, d_name in enumerate(day_names):
                                if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                    count += 1
                                    if d_name in str(t['Yasaklı Günler']):
                                        violations.append({"Hoca": t['Ad Soyad'], "Hata": f"Yasaklı Gün ({d_name})", "Sınıf": c_name})
                                    pref = str(t['Tercih (Sabah/Öğle)'])
                                    real_time = "Sabah" if s_req == 0 else "Öğle"
                                    if pref != "Farketmez" and pref != real_time:
                                        violations.append({"Hoca": t['Ad Soyad'], "Hata": f"Tercih İhlali ({pref})", "Sınıf": c_name})
                            if count > 0: teacher_counts[t['Ad Soyad']] = count

                        adv_disp = "-"
                        if c_name in advisor_map:
                            adv_disp = advisor_map[c_name]
                        elif teacher_counts:
                            # Destek hocası da Danışman görünebilir
                            cands = {n: c for n, c in teacher_counts.items() 
                                     if not any(t['Ad Soyad'] == n and 'Ek Görevli' in str(t['Rol']) for t in teachers_list)
                                     and n not in advisor_map.values()}
                            if not allow_native_advisor:
                                cands = {n: c for n, c in cands.items() if n not in native_names}
                            
                            if cands:
                                max_v = max(cands.values())
                                adv_disp = " / ".join([n for n, c in cands.items() if c == max_v])

                        row = {
                            "Sınıf": c_name, "Seviye": c['Seviye'], "Sınıf Danışmanı": adv_disp,
                            "Zaman": "Sabah" if s_req == 0 else "Öğle"
                        }
                        for d_idx, d_name in enumerate(day_names):
                            val = "🔴 BOŞ"
                            for t_idx, t in enumerate(teachers_list):
                                if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                    val = t['Ad Soyad']
                                    break
                            row[d_name] = val
                        res_data.append(row)

                    # İstatistik
                    stats = []
                    for t_idx, t in enumerate(teachers_list):
                        assigned = sum([solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions])
                        target = int(t['Hedef Ders Sayısı'])
                        diff = assigned - target
                        stat = "Tamam"
                        if diff > 0: stat = f"+{diff} Fazla"
                        elif diff < 0: stat = f"{diff} Eksik"
                        stats.append({"Hoca Adı": t['Ad Soyad'], "Hedef": target, "Atanan": assigned, "Durum": stat})

                    df_res = pd.DataFrame(res_data)
                    df_stats = pd.DataFrame(stats)
                    df_violations = pd.DataFrame(violations).drop_duplicates() if violations else pd.DataFrame()

                    if not df_violations.empty:
                        st.warning(f"⚠️ {len(df_violations)} adet kural esnetildi.")
                        st.table(df_violations)
                    else:
                        st.success("✅ Kusursuz Çözüm!")

                    st.dataframe(df_res)
                    st.dataframe(df_stats)

                    # --- EXCEL ---
                    output_res = io.BytesIO()
                    with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
                        df_res.to_excel(writer, index=False, sheet_name="Program")
                        df_stats.to_excel(writer, index=False, sheet_name="Istatistikler")
                        if not df_violations.empty: df_violations.to_excel(writer, index=False, sheet_name="Ihlal_Raporu")
                        
                        wb = writer.book
                        ws_prog = writer.sheets['Program']
                        ws_stat = writer.sheets['Istatistikler']
                        
                        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
                        fmt_gold = wb.add_format(dict(base_fmt, bg_color='#FFD700'))
                        fmt_orange = wb.add_format(dict(base_fmt, bg_color='#FFA500'))
                        fmt_maroon = wb.add_format(dict(base_fmt, bg_color='#800000', font_color='white'))
                        fmt_green = wb.add_format(dict(base_fmt, bg_color='#006400', font_color='white'))
                        fmt_blue = wb.add_format(dict(base_fmt, bg_color='#ADD8E6')) 
                        fmt_default = wb.add_format(base_fmt)
                        fmt_stat_missing = wb.add_format(dict(base_fmt, bg_color='#FF9999'))
                        fmt_stat_ok = wb.add_format(dict(base_fmt, bg_color='#CCFFCC'))

                        ws_prog.set_column('A:B', 8)
                        ws_prog.set_column('C:C', 20)
                        ws_prog.set_column('E:I', 12)
                        ws_prog.set_row(0, 20)

                        for r, row in df_res.iterrows():
                            excel_r = r + 1
                            ws_prog.set_row(excel_r, 20)
                            lvl = str(row['Seviye'])
                            ws_prog.write(excel_r, 0, row['Sınıf'], fmt_gold if lvl=="A1" else (fmt_orange if lvl=="A2" else (fmt_maroon if lvl=="B1" else fmt_green)))
                            ws_prog.write(excel_r, 1, row['Seviye'], fmt_default)
                            ws_prog.write(excel_r, 2, row['Sınıf Danışmanı'], fmt_default)
                            ws_prog.write(excel_r, 3, row['Zaman'], fmt_default)
                            
                            for c in range(4, 9):
                                val = row.iloc[c]
                                f = fmt_default
                                if val in native_names: f = fmt_blue
                                ws_prog.write(excel_r, c, val, f)

                        for r, row in df_stats.iterrows():
                            excel_r = r + 1
                            status = str(row['Durum'])
                            stat_fmt = fmt_default
                            if "Eksik" in status: stat_fmt = fmt_stat_missing
                            elif "Tamam" in status: stat_fmt = fmt_stat_ok
                            ws_stat.write(excel_r, 0, row['Hoca Adı'], fmt_default)
                            ws_stat.write(excel_r, 1, row['Hedef'], fmt_default)
                            ws_stat.write(excel_r, 2, row['Atanan'], fmt_default)
                            ws_stat.write(excel_r, 3, status, stat_fmt)

                    st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                else:
                    st.error("❌ Çözüm Bulunamadı.")
