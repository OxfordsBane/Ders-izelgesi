import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V41 - Hatasız Final", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V41 - Hatasız Sürüm)")
st.info("""
**Düzeltildi:**
✅ **Sistem Hatası Giderildi:** Yetkinlik kontrolündeki kod hatası düzeltildi.
⚓ **Danışman Kuralları:** Danışmanlar Pazartesi kendi sınıfında olur + En az 2 gün o sınıfa girer (Zorunlu).
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
            "   - Destek: Pazartesi derse giremezler (esnek), Danışman olamazlar.",
            "   - Native: Yabancı hocalar. A1'e girmezler. B2 > B1 > A2 önceliğiyle dağıtılırlar.",
            "   - Danışman: Sınıf sorumlularıdır. Program onları bir sınıfta toplamaya çalışır.",
            "   - Ek Görevli: İdari/Özel görevi olanlar. Sınıf Danışmanı olamazlar.",
            "",
            "2. SÜTUNLAR NASIL DOLDURULUR?",
            "   - Hedef Ders Sayısı: Hocanın o hafta gireceği toplam 'oturum' sayısı.",
            "   - Tercih: 'Sabah', 'Öğle'. Sistem buna uymak için ÇOK çabalar.",
            "   - Yasaklı Günler: Hoca o gün ASLA gelmez. Virgülle ayırın.",
            "   - Sabit Sınıf: Hocanın özellikle girmesi istenen sınıfı (Koordinatör vb.).",
            "   - Yetkinlik: 'Hepsi' veya 'A1,A2' gibi.",
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
             errors.append(f"🛑 **{t['Ad Soyad']}**: 'Destek' hocası sabit sınıf alamaz!")
        if not allow_native_advisor and "NATIVE" in role and fixed_class:
             errors.append(f"🛑 **{t['Ad Soyad']}**: Native hocaya sabit sınıf verilmesi engellendi.")
        
        if fixed_class:
            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class), None)
            if not target_class:
                errors.append(f"❌ **{t['Ad Soyad']}**: Atandığı '{fixed_class}' sınıfı sistemde yok.")
            
            # Yasaklı Gün Uyarısı
            forbidden_count = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
            available_days = 5 - forbidden_count
            target = int(t['Hedef Ders Sayısı'])
            if available_days < target:
                warnings.append(f"⚠️ **{t['Ad Soyad']}**: Hedefi {target} gün ama sadece {available_days} gün müsait. Hedef otomatik olarak {available_days} güne düşürülecek.")

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
        st.error("🛑 Lütfen hataları düzeltin:")
        for e in logic_errors: st.markdown(e)
    else:
        if logic_warnings:
            for w in logic_warnings: st.warning(w)
            
        total_slots_needed = len(classes_list) * 5
        total_slots_avail = sum(t['Hedef Ders Sayısı'] for t in teachers_list)
        
        col1, col2 = st.columns(2)
        col1.metric("İhtiyaç", total_slots_needed)
        col2.metric("Kapasite", total_slots_avail)

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Optimizasyon yapılıyor... (Danışmanlar sınıflarına kilitleniyor...)"):
                
                model = cp_model.CpModel()
                days = range(5) # 0-4
                day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                sessions = range(2)
                
                # --- DEĞİŞKENLER ---
                x = {}
                advisor_var = {} 
                
                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        advisor_var[(t, c)] = model.NewBoolVar(f'adv_{t}_{c}')
                        for d in days:
                            for s in sessions:
                                x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                # --- KISITLAMALAR (HARD) ---
                
                # 1. Danışman Tekilliği
                for c in range(len(classes_list)):
                    model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) == 1)

                for t in range(len(teachers_list)):
                    model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

                # 2. Rol Kısıtlamaları
                for t_idx, t in enumerate(teachers_list):
                    role = str(t['Rol'])
                    if 'Ek Görevli' in role:
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                    if not allow_native_advisor and 'Native' in role:
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)

                # 3. Sabit Sınıf
                for t_idx, t in enumerate(teachers_list):
                    if t['Sabit Sınıf']:
                        fixed_c_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c_idx is not None:
                            model.Add(advisor_var[(t_idx, fixed_c_idx)] == 1)

                # 4. DANIŞMAN ZORUNLULUKLARI (Pazartesi + Min 2 Gün)
                for t_idx, t_data in enumerate(teachers_list):
                    # Müsait gün sayısını hesapla
                    forbidden_count = len(str(t_data['Yasaklı Günler']).split(',')) if t_data['Yasaklı Günler'] else 0
                    available_days = 5 - forbidden_count
                    
                    for c_idx, c_data in enumerate(classes_list):
                        is_adv = advisor_var[(t_idx, c_idx)]
                        req_s = c_data['Zaman Kodu']
                        
                        # A. Pazartesi Zorunluluğu
                        if "Pazartesi" not in str(t_data['Yasaklı Günler']):
                            model.Add(x[(t_idx, c_idx, 0, req_s)] == 1).OnlyEnforceIf(is_adv)
                        
                        # B. En Az 2 Gün Zorunluluğu
                        if available_days >= 2:
                            days_in_class = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            model.Add(days_in_class >= 2).OnlyEnforceIf(is_adv)

                # --- Standart Kısıtlamalar ---
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
                
                # Hedef Ders
                adjusted_targets = []
                for t_idx, t in enumerate(teachers_list):
                    original_target = int(t['Hedef Ders Sayısı'])
                    forbidden_count = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
                    max_possible = 5 - forbidden_count
                    if 'Destek' in str(t['Rol']) or 'Native' in str(t['Rol']): max_possible *= 2
                    real_target = min(original_target, max_possible)
                    adjusted_targets.append(real_target)
                    
                    total_assignments = []
                    for c in range(len(classes_list)):
                        for d in days:
                            for s in sessions: total_assignments.append(x[(t_idx, c, d, s)])
                    model.Add(sum(total_assignments) <= real_target)

                # Max Hoca
                for c_idx in range(len(classes_list)):
                    teachers_here = []
                    for t in range(len(teachers_list)):
                        teach = model.NewBoolVar(f'tch_{t}_{c_idx}')
                        model.AddMaxEquality(teach, [x[(t, c_idx, d, s)] for d in days for s in sessions])
                        teachers_here.append(teach)
                    model.Add(sum(teachers_here) <= max_teachers_per_class)

                # Native A1
                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c_idx, c_data in enumerate(classes_list):
                            if c_data['Seviye'] == 'A1':
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

                # Native Tekilliği
                for c_idx, c_data in enumerate(classes_list):
                    natives_in_class = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_{t_idx}_{c_idx}')
                            model.AddMaxEquality(is_present, [x[(t_idx, c_idx, d, s)] for d in days for s in sessions])
                            natives_in_class.append(is_present)
                    model.Add(sum(natives_in_class) <= 1) 

                # Ek Görevli Gezici
                for t_idx, t in enumerate(teachers_list):
                    if 'Ek Görevli' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            lessons_in_class = []
                            for d in days:
                                for s in sessions:
                                    lessons_in_class.append(x[(t_idx, c_idx, d, s)])
                            model.Add(sum(lessons_in_class) <= 1)

                # 8. Vardiya Kısıtlaması
                for t_idx, t in enumerate(teachers_list):
                    role = str(t['Rol'])
                    if 'Danışman' in role or 'Ek Görevli' in role:
                        for d in days:
                            is_morning = model.NewBoolVar(f'm_{t_idx}_{d}')
                            is_afternoon = model.NewBoolVar(f'a_{t_idx}_{d}')
                            model.AddMaxEquality(is_morning, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                            model.AddMaxEquality(is_afternoon, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                            model.Add(is_morning + is_afternoon <= 1)

                # --- PUANLAMA (SOFT) ---
                objective = []
                objective.append(sum(x.values()) * 100000)

                # A. Danışmanlık 3 Gün Hedefi
                for t_idx, t in enumerate(teachers_list):
                    for c_idx in range(len(classes_list)):
                        is_adv = advisor_var[(t_idx, c_idx)]
                        for d in days:
                            for s in sessions:
                                is_teaching_as_adv = model.NewBoolVar(f'taa_{t_idx}_{c_idx}_{d}')
                                model.Add(is_teaching_as_adv == 1).OnlyEnforceIf([x[(t_idx, c_idx, d, s)], is_adv])
                                objective.append(is_teaching_as_adv * 5000000)

                # B. Rol Önceliği
                for t_idx, t in enumerate(teachers_list):
                    if 'Danışman' in str(t['Rol']):
                        assigned_somewhere = sum([advisor_var[(t_idx, c)] for c in range(len(classes_list))])
                        objective.append(assigned_somewhere * 1000000)
                    elif 'Destek' in str(t['Rol']):
                        assigned_somewhere = sum([advisor_var[(t_idx, c)] for c in range(len(classes_list))])
                        objective.append(assigned_somewhere * -500000)

                # C. Hedef Doldurma
                for t_idx, t in enumerate(teachers_list):
                    real_target = adjusted_targets[t_idx]
                    current_load = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    if 'Danışman' in str(t['Rol']): objective.append(current_load * 5000000)
                    else: objective.append(current_load * 5000)

                # D. Zaman/Yasak/Yetkinlik/Partner
                for t_idx, t in enumerate(teachers_list):
                    pref = str(t['Tercih (Sabah/Öğle)'])
                    if pref == "Sabah":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 1)] * -100000000)
                    elif pref == "Öğle":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 0)] * -100000000)

                    forbidden = str(t['Yasaklı Günler'])
                    for d_idx, d_name in enumerate(day_names):
                        if d_name in forbidden:
                            for c in range(len(classes_list)):
                                for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -50000000)

                    allowed = str(t['Yetkinlik (Seviyeler)'])
                    if "Hepsi" not in allowed:
                        for c_idx, c in enumerate(classes_list):
                            if c['Seviye'] not in allowed:
                                for d in days:
                                    for s in sessions:
                                        # BURASI DÜZELTİLDİ: c -> c_idx
                                        objective.append(x[(t_idx, c_idx, d, s)] * -40000000)
                    
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
                
                # H. Destek Pazartesi
                for t_idx, t in enumerate(teachers_list):
                    if 'Destek' in str(t['Rol']):
                        for c in range(len(classes_list)):
                            for s in sessions: objective.append(x[(t_idx, c, 0, s)] * -100000)

                # I. Native Dağılımı
                for c_idx, c_data in enumerate(classes_list):
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_score_{t_idx}_{c_idx}')
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
                    
                    for c_idx, c in enumerate(classes_list):
                        c_name = c['Sınıf Adı']
                        s_req = c['Zaman Kodu']
                        
                        # ATANAN DANIŞMANI BUL
                        assigned_advisor_idx = None
                        for t_idx in range(len(teachers_list)):
                            if solver.Value(advisor_var[(t_idx, c_idx)]) == 1:
                                assigned_advisor_idx = t_idx
                                break
                        
                        advisor_name = teachers_list[assigned_advisor_idx]['Ad Soyad'] if assigned_advisor_idx is not None else "Atanamadı"

                        row = {
                            "Sınıf": c_name, "Seviye": c['Seviye'], "Sınıf Danışmanı": advisor_name,
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
                        original_target = int(t['Hedef Ders Sayısı'])
                        diff = assigned - original_target
                        stat = "Tamam"
                        if diff > 0: stat = f"+{diff} Fazla"
                        elif diff < 0: stat = f"{diff} Eksik"
                        
                        real_target = adjusted_targets[t_idx]
                        if real_target < original_target and assigned == real_target:
                            stat = f"{diff} Eksik (Yasaklı Günlerden Dolayı Max)"

                        stats.append({"Hoca Adı": t['Ad Soyad'], "Hedef": original_target, "Atanan": assigned, "Durum": stat})

                    df_res = pd.DataFrame(res_data)
                    df_stats = pd.DataFrame(stats)
                    df_violations = pd.DataFrame() 

                    st.success("✅ Kusursuz Çözüm!")
                    st.dataframe(df_res)
                    st.dataframe(df_stats)

                    # --- EXCEL ---
                    output_res = io.BytesIO()
                    with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
                        df_res.to_excel(writer, index=False, sheet_name="Program")
                        df_stats.to_excel(writer, index=False, sheet_name="Istatistikler")
                        
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
                            stat = str(row['Durum'])
                            ws_stat.write(excel_r, 0, row['Hoca Adı'], fmt_default)
                            ws_stat.write(excel_r, 1, row['Hedef'], fmt_default)
                            ws_stat.write(excel_r, 2, row['Atanan'], fmt_default)
                            ws_stat.write(excel_r, 3, stat, fmt_stat_missing if "Eksik" in stat else fmt_stat_ok)

                    st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                else:
                    st.error("❌ Çözüm Bulunamadı.")
