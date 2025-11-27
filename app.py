import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V44 - Disiplinli", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V44 - Disiplin ve Adalet)")
st.info("""
**Kesinleşen Kurallar:**
1. ⚓ **Danışman Pazartesi:** Tartışmaya kapalıdır. Danışman Pazartesi günü sınıfındadır.
2. 🏠 **Sınıf Sahiplenme:** Danışmanlar, kendi sınıflarını doldurmadan başka sınıfa gitmezler.
3. ⚖️ **Adalet:** Her hocaya mutlaka ders yazılır (Boşta hoca kalmaz).
4. 🌙 **Vardiya Koruması:** 'Farketmez' diyenler bile aynı gün çift vardiya (Sabah+Öğle) yapmaz (Çok zorda kalınmadıkça).
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
            "   - Destek: Esnek hocalardır.",
            "   - Native: Yabancı hocalar.",
            "   - Danışman: Sınıf sahipleridir. PAZARTESİ MÜSAİT OLMALIDIRLAR.",
            "   - Ek Görevli: İdari görevi olanlar.",
            "",
            "2. SÜTUNLAR:",
            "   - Hedef Ders Sayısı: Haftalık toplam oturum sayısı.",
            "   - Sabit Sınıf: Danışman ise sınıfını yazın.",
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
        forbidden = str(t['Yasaklı Günler'])
        
        # Danışman Pazartesi Kontrolü
        if 'DANIŞMAN' in role and fixed_class:
            if "Pazartesi" in forbidden:
                errors.append(f"🛑 **{t['Ad Soyad']}**: Danışman olduğu için Pazartesi günü yasaklı OLAMAZ. Lütfen Excel'i düzeltin.")
        
        if "DESTEK" in role and fixed_class:
             errors.append(f"🛑 **{t['Ad Soyad']}**: 'Destek' hocası sabit sınıf alamaz.")

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
        # --- İSTATİSTİKLER ---
        total_needed = len(classes_list) * 5
        total_cap = sum(t['Hedef Ders Sayısı'] for t in teachers_list)
        
        num_danisman = sum(1 for t in teachers_list if 'Danışman' in str(t['Rol']))
        num_native = sum(1 for t in teachers_list if 'Native' in str(t['Rol']))
        
        num_sabah = sum(1 for c in classes_list if c['Zaman Kodu'] == 0)
        num_ogle = sum(1 for c in classes_list if c['Zaman Kodu'] == 1)
        
        st.markdown("### 📊 Durum Analizi")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Toplam Sınıf", len(classes_list))
        c2.metric("İhtiyaç", total_needed)
        c3.metric("Kapasite", total_cap, delta=total_cap - total_needed)
        c4.metric("Danışman", num_danisman)
        st.divider()

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Optimizasyon yapılıyor... (Disiplinli Mod Aktif)"):
                
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

                # --- KISITLAMALAR (HARD) ---
                
                # 1. Sınıf Doluluğu
                for c_idx, c_data in enumerate(classes_list):
                    req_s = c_data['Zaman Kodu']
                    other_s = 1 - req_s
                    for d in days:
                        if allow_empty_slots:
                            model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) <= 1)
                        else:
                            model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) == 1)
                        model.Add(sum(x[(t, c_idx, d, other_s)] for t in range(len(teachers_list))) == 0)

                # 2. Hoca Tekilliği
                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

                # 3. Hoca Adaleti (Hiçbir hoca boş kalamaz)
                for t_idx, t in enumerate(teachers_list):
                    target = int(t['Hedef Ders Sayısı'])
                    if target > 0: # Hedefi 0 olanlar hariç
                        # En az 1 ders almalı
                        model.Add(sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions]) >= 1)
                
                # 4. Hedef Ders Sayısı (Üst Limit)
                for t_idx, t in enumerate(teachers_list):
                    target = int(t['Hedef Ders Sayısı'])
                    total_assign = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    model.Add(total_assign <= target)

                # 5. Sınıf Başı Max Hoca
                for c_idx in range(len(classes_list)):
                    teachers_here = []
                    for t in range(len(teachers_list)):
                        teach = model.NewBoolVar(f'tch_{t}_{c_idx}')
                        model.AddMaxEquality(teach, [x[(t, c_idx, d, s)] for d in days for s in sessions])
                        teachers_here.append(teach)
                    model.Add(sum(teachers_here) <= max_teachers_per_class)

                # 6. Native A1 Yasağı
                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c_idx, c_data in enumerate(classes_list):
                            if c_data['Seviye'] == 'A1':
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

                # 7. Native Tekilliği
                for c_idx, c_data in enumerate(classes_list):
                    natives_in_class = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_{t_idx}_{c_idx}')
                            model.AddMaxEquality(is_present, [x[(t_idx, c_idx, d, s)] for d in days for s in sessions])
                            natives_in_class.append(is_present)
                    model.Add(sum(natives_in_class) <= 1) 

                # 8. Ek Görevli Gezici
                for t_idx, t in enumerate(teachers_list):
                    if 'Ek Görevli' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            lessons_in_class = []
                            for d in days:
                                for s in sessions:
                                    lessons_in_class.append(x[(t_idx, c_idx, d, s)])
                            model.Add(sum(lessons_in_class) <= 1)

                # 9. DANIŞMAN PAZARTESİ KİLİDİ (ASLA DELİNEMEZ)
                for t_idx, t in enumerate(teachers_list):
                    if 'Danışman' in str(t['Rol']) and t['Sabit Sınıf']:
                        fixed_c = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c is not None:
                            req_s = classes_list[fixed_c]['Zaman Kodu']
                            # Pazartesi = 0. Gün. Burada hoca olmak ZORUNDA.
                            model.Add(x[(t_idx, fixed_c, 0, req_s)] == 1)

                # --- PUANLAMA (SOFT) ---
                objective = []
                
                # A. Atama Puanı
                for t_idx, t in enumerate(teachers_list):
                    # Her atama +100 Bin
                    objective.append(sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions]) * 100000)

                # B. Hedef Doldurma (Danışman > Standart)
                for t_idx, t in enumerate(teachers_list):
                    current_load = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    if 'Danışman' in str(t['Rol']):
                        objective.append(current_load * 10000000) # Danışmanı doldurmak çok önemli
                    else:
                        objective.append(current_load * 5000)

                # C. Danışman Sahiplenme (Kendi Sınıfı)
                for t_idx, t in enumerate(teachers_list):
                    if 'Danışman' in str(t['Rol']) and t['Sabit Sınıf']:
                        fixed_c = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c is not None:
                            # Kendi sınıfındaki her ders +10 MİLYON
                            # Bu sayede hoca başka sınıfa gitmek yerine kendi sınıfını doldurmak ister.
                            load_in_class = sum([x[(t_idx, fixed_c, d, s)] for d in days for s in sessions])
                            objective.append(load_in_class * 10000000)

                # D. Çift Vardiya Cezası (Aynı gün Sabah+Öğle)
                # Farketmez diyenler bile bunu yapmamalı.
                for t_idx, t in enumerate(teachers_list):
                    for d in days:
                        is_morning = model.NewBoolVar(f'm_{t_idx}_{d}')
                        is_afternoon = model.NewBoolVar(f'a_{t_idx}_{d}')
                        model.AddMaxEquality(is_morning, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                        model.AddMaxEquality(is_afternoon, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                        
                        double_shift = model.NewBoolVar(f'dbl_{t_idx}_{d}')
                        model.Add(is_morning + is_afternoon == 2).OnlyEnforceIf(double_shift)
                        model.Add(is_morning + is_afternoon < 2).OnlyEnforceIf(double_shift.Not())
                        
                        # Çok büyük ceza (-5 Milyon)
                        objective.append(double_shift * -5000000)

                # E. Zaman Tercihi
                for t_idx, t in enumerate(teachers_list):
                    pref = str(t['Tercih (Sabah/Öğle)'])
                    if pref == "Sabah":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 1)] * -2000000)
                    elif pref == "Öğle":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 0)] * -2000000)

                # F. Yasaklı Günler (-50M)
                for t_idx, t in enumerate(teachers_list):
                    forbidden = str(t['Yasaklı Günler'])
                    for d_idx, d_name in enumerate(day_names):
                        if d_name in forbidden:
                            for c in range(len(classes_list)):
                                for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -50000000)

                # G. Native Dağılımı (Şelale)
                for c_idx, c_data in enumerate(classes_list):
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_sc_{t_idx}_{c_idx}')
                            model.AddMaxEquality(is_present, [x[(t_idx, c_idx, d, s)] for d in days for s in sessions])
                            lvl = c_data['Seviye']
                            score = 10000 if lvl == "A2" else (50000 if lvl == "B1" else (100000 if lvl == "B2" else 0))
                            objective.append(is_present * score)

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
                        
                        t_counts = {}
                        for t_idx, t in enumerate(teachers_list):
                            cnt = 0
                            for d_idx, d_name in enumerate(day_names):
                                if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                    cnt += 1
                                    if d_name in str(t['Yasaklı Günler']):
                                        violations.append({"Hoca": t['Ad Soyad'], "Hata": f"Yasaklı Gün ({d_name})", "Sınıf": c_name})
                                    pref = str(t['Tercih (Sabah/Öğle)'])
                                    real_time = "Sabah" if s_req == 0 else "Öğle"
                                    if pref != "Farketmez" and pref != real_time:
                                        violations.append({"Hoca": t['Ad Soyad'], "Hata": f"Tercih İhlali ({pref})", "Sınıf": c_name})
                            if cnt > 0: t_counts[t['Ad Soyad']] = cnt

                        # Danışman Belirleme (Akıllı)
                        adv_disp = "-"
                        if c_name in advisor_map:
                            adv_disp = advisor_map[c_name]
                        elif t_counts:
                            # Sabit değilse, o sınıfta en çok dersi olan (Ek Görevli Hariç)
                            cands = {n: c for n, c in t_counts.items() 
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
                            ws_stat.write(excel_r, 0, row['Hoca Adı'], fmt_default)
                            ws_stat.write(excel_r, 1, row['Hedef'], fmt_default)
                            ws_stat.write(excel_r, 2, row['Atanan'], fmt_default)
                            ws_stat.write(excel_r, 3, status, fmt_stat_missing if "Eksik" in status else fmt_stat_ok)

                    st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                else:
                    st.error("❌ Çözüm Bulunamadı.")
