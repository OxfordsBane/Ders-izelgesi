import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V44 - Tek Vardiya", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V44 - Tek Vardiya & Hiyerarşi)")
st.info("""
**Kesinleşen Kurallar:**
1. ⛔ **Tek Vardiya:** Hiçbir hoca aynı gün hem sabah hem öğle derse giremez.
2. 🌍 **Native Kuralı:** Native hocalar bir sınıfa haftada **en fazla 1 kez** girer.
3. 👑 **Önce Danışman:** Danışmanlar sınıflarına kilitlenir, ardından Native dağıtılır.
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
            "   - Destek: Joker elemandır. Gerektiğinde Danışman olur.",
            "   - Native: Haftada 1 sınıfa YALNIZCA 1 KEZ girer.",
            "   - Danışman: Sınıfına en az 3 gün girer (PreFac hariç).",
            "   - Ek Görevli: İdari görevli. Sınıf Danışmanı olamaz.",
            "",
            "2. KURALLAR",
            "   - Tek Vardiya: Hiçbir hoca aynı gün hem sabah hem öğle çalışmaz.",
            "   - Hedef Ders Sayısı: Sistem gerekirse bunu 1 saat azaltarak herkesi sığdırır.",
            "   - Sabit Sınıf: Hocanın kesin atanacağı sınıf.",
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
        
        if not allow_native_advisor and "NATIVE" in role and fixed_class:
             errors.append(f"🛑 **{t['Ad Soyad']}**: Native hocaya sabit sınıf verilmesi engellendi.")
        
        if fixed_class:
            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class), None)
            if not target_class:
                errors.append(f"❌ **{t['Ad Soyad']}**: Atandığı '{fixed_class}' sınıfı sistemde yok.")
            
            forbidden_count = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
            available_days = 5 - forbidden_count
            is_prefac = "PreFaculty" in (target_class['Seviye'] if target_class else "")
            
            if available_days < 3 and "DANIŞMAN" in role and not is_prefac:
                warnings.append(f"⚠️ **{t['Ad Soyad']}**: Danışman ama sadece {available_days} gün müsait. 3 gün kuralı uygulanamayabilir.")

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
            
        total_slots_needed = 0
        for c in classes_list:
            if c['Seviye'] == 'PreFaculty': total_slots_needed += 3 
            else: total_slots_needed += 5
        
        # Kapasite Hesabı (Tek Vardiya Kuralına Göre)
        raw_demand = 0
        for t in teachers_list:
            forbidden_cnt = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
            # ARTIK HERKES TEK VARDİYA -> Max 5 gün
            max_cap = 5 - forbidden_cnt 
            raw_demand += min(int(t['Hedef Ders Sayısı']), max_cap)

        col1, col2 = st.columns(2)
        col1.metric("Sınıf İhtiyacı (Oturum)", total_slots_needed)
        col2.metric("Hoca Kapasitesi (Tek Vardiya)", raw_demand)
        
        reduce_mode = False
        if raw_demand > total_slots_needed:
            st.info("ℹ️ Hoca kapasitesi fazla. Gerekirse hedeflerden 1'er saat kırpılacak.")
            reduce_mode = True

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Optimizasyon yapılıyor... (Tek Vardiya & Hiyerarşik Yerleşim...)"):
                
                model = cp_model.CpModel()
                days = range(5)
                day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                sessions = range(2)
                
                x = {}
                advisor_var = {} 
                
                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        advisor_var[(t, c)] = model.NewBoolVar(f'adv_{t}_{c}')
                        for d in days:
                            for s in sessions:
                                x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                # --- 1. DANIŞMANLIK İSKELETİ ---
                
                # A. Tekillik
                for c in range(len(classes_list)):
                    model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) == 1)
                for t in range(len(teachers_list)):
                    model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

                # B. Rol Kısıtlamaları (Ek Görevli Olamaz)
                for t_idx, t in enumerate(teachers_list):
                    if 'Ek Görevli' in str(t['Rol']):
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                    if not allow_native_advisor and 'Native' in str(t['Rol']):
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)

                # C. Sabit Sınıf Ataması
                for t_idx, t in enumerate(teachers_list):
                    if t['Sabit Sınıf']:
                        fixed_c_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c_idx is not None:
                            model.Add(advisor_var[(t_idx, fixed_c_idx)] == 1)

                # D. DANIŞMAN ŞARTLARI
                for t_idx, t_data in enumerate(teachers_list):
                    forbidden_days = str(t_data['Yasaklı Günler'])
                    forbidden_count = len(forbidden_days.split(',')) if t_data['Yasaklı Günler'] else 0
                    available_days = 5 - forbidden_count
                    
                    for c_idx, c_data in enumerate(classes_list):
                        is_adv = advisor_var[(t_idx, c_idx)]
                        req_s = c_data['Zaman Kodu']
                        
                        # 1. Pazartesi Kuralı
                        if "Pazartesi" not in forbidden_days:
                            model.Add(x[(t_idx, c_idx, 0, req_s)] == 1).OnlyEnforceIf(is_adv)
                        
                        # 2. 3 Gün Kuralı (PreFaculty HARİÇ)
                        if c_data['Seviye'] != "PreFaculty":
                            days_in_class = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            if available_days >= 3:
                                model.Add(days_in_class >= 3).OnlyEnforceIf(is_adv)
                            elif available_days == 2:
                                model.Add(days_in_class >= 2).OnlyEnforceIf(is_adv)

                # --- 2. NATIVE KISITLAMASI ---
                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            # Native bir sınıfa haftada MAX 1 kez girer
                            # Native Danışman olsa bile (izin varsa) 1 gün girer mantığı doğru değil,
                            # Danışmansa 3 gün girmeli. O yüzden danışman değilse 1 gün.
                            is_not_advisor = advisor_var[(t_idx, c_idx)].Not()
                            class_total = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            model.Add(class_total <= 1).OnlyEnforceIf(is_not_advisor)

                # --- 3. GENEL KISITLAMALAR ---
                for c_idx, c_data in enumerate(classes_list):
                    req_session = c_data['Zaman Kodu']
                    other_session = 1 - req_session
                    for d in days:
                        if allow_empty_slots:
                            model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
                        else:
                            if c_data['Seviye'] == "PreFaculty" and d >= 3:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 0)
                            else:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 1)
                        model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)

                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)
                
                # --- 4. PREFACULTY KAPAMA ---
                for c_idx, c_data in enumerate(classes_list):
                    if c_data['Seviye'] == "PreFaculty":
                        for t_idx in range(len(teachers_list)):
                            for s in sessions:
                                model.Add(x[(t_idx, c_idx, 3, s)] == 0)
                                model.Add(x[(t_idx, c_idx, 4, s)] == 0)

                # --- 5. HEDEF DENGELEME (AKILLI HEDEF) ---
                adjusted_targets = []
                for t_idx, t in enumerate(teachers_list):
                    original_target = int(t['Hedef Ders Sayısı'])
                    forbidden_count = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
                    
                    # TEK VARDİYA: Max kapasite 5 - yasaklı gün
                    max_possible = 5 - forbidden_count
                    
                    if reduce_mode and original_target > 2:
                        target_to_use = original_target - 1
                    else:
                        target_to_use = original_target
                        
                    real_target = min(target_to_use, max_possible)
                    adjusted_targets.append(real_target)
                    
                    total_assignments = []
                    for c in range(len(classes_list)):
                        for d in days:
                            for s in sessions: total_assignments.append(x[(t_idx, c, d, s)])
                    
                    model.Add(sum(total_assignments) <= real_target)
                    # Hiç boş kalmasın
                    if real_target > 0:
                        model.Add(sum(total_assignments) >= 1)

                # Max Hoca
                for c_idx in range(len(classes_list)):
                    teachers_here = []
                    for t in range(len(teachers_list)):
                        teach = model.NewBoolVar(f'tch_{t}_{c_idx}')
                        model.AddMaxEquality(teach, [x[(t, c_idx, d, s)] for d in days for s in sessions])
                        teachers_here.append(teach)
                    model.Add(sum(teachers_here) <= max_teachers_per_class)

                # Native A1 Yasağı
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
                            model.Add(sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions) <= 1)

                # --- 6. TEK VARDİYA (HERKES İÇİN) ---
                # Hiçbir hoca aynı gün hem sabah hem öğle alamaz
                for t_idx, t in enumerate(teachers_list):
                    for d in days:
                        is_morning = model.NewBoolVar(f'm_{t_idx}_{d}')
                        is_afternoon = model.NewBoolVar(f'a_{t_idx}_{d}')
                        model.AddMaxEquality(is_morning, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                        model.AddMaxEquality(is_afternoon, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                        model.Add(is_morning + is_afternoon <= 1)

                # --- PUANLAMA ---
                objective = []
                objective.append(sum(x.values()) * 100000)

                # Hedef Doldurma
                for t_idx, t in enumerate(teachers_list):
                    current_load = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    if 'Danışman' in str(t['Rol']): objective.append(current_load * 5000000)
                    else: objective.append(current_load * 5000)

                # Zaman/Yasak/Yetkinlik
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
                                    for s in sessions: objective.append(x[(t_idx, c_idx, d, s)] * -40000000)
                    
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
                
                for t_idx, t in enumerate(teachers_list):
                    if 'Destek' in str(t['Rol']):
                        for c in range(len(classes_list)):
                            for s in sessions: objective.append(x[(t_idx, c, 0, s)] * -100000)

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
                    
                    res_data = []
                    native_names = [t['Ad Soyad'] for t in teachers_list if 'Native' in str(t['Rol'])]
                    
                    for c_idx, c in enumerate(classes_list):
                        c_name = c['Sınıf Adı']
                        s_req = c['Zaman Kodu']
                        
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
                            if c['Seviye'] == "PreFaculty" and d_idx >= 3:
                                val = "⛔ KAPALI"
                            else:
                                for t_idx, t in enumerate(teachers_list):
                                    if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                        val = t['Ad Soyad']
                                        break
                            row[d_name] = val
                        res_data.append(row)

                    stats = []
                    for t_idx, t in enumerate(teachers_list):
                        assigned = sum([solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions])
                        original_target = int(t['Hedef Ders Sayısı'])
                        real_target = adjusted_targets[t_idx]
                        
                        stat = "Tamam"
                        diff = assigned - original_target
                        if diff > 0: stat = f"+{diff} Fazla"
                        elif diff < 0: stat = f"{diff} Eksik"
                        
                        if real_target < original_target and assigned == real_target:
                            stat = f"Tamam (Hedef {original_target-real_target} kırpıldı)"

                        stats.append({"Hoca Adı": t['Ad Soyad'], "Hedef (İlk)": original_target, "Atanan": assigned, "Durum": stat})

                    df_res = pd.DataFrame(res_data)
                    df_stats = pd.DataFrame(stats)
                    df_violations = pd.DataFrame()

                    st.success("✅ Kusursuz Çözüm!")
                    st.dataframe(df_res)
                    st.dataframe(df_stats)

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
                            lvl_fmt = fmt_default
                            if lvl == "A1": lvl_fmt = fmt_gold
                            elif lvl == "A2": lvl_fmt = fmt_orange
                            elif lvl == "B1": lvl_fmt = fmt_maroon
                            elif lvl == "B2": lvl_fmt = fmt_green
                            elif lvl == "PreFaculty": lvl_fmt = wb.add_format(dict(base_fmt, bg_color='#E0E0E0'))

                            ws_prog.write(excel_r, 0, row['Sınıf'], lvl_fmt)
                            ws_prog.write(excel_r, 1, row['Seviye'], lvl_fmt)
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
                            ws_stat.write(excel_r, 1, row['Hedef (İlk)'], fmt_default)
                            ws_stat.write(excel_r, 2, row['Atanan'], fmt_default)
                            ws_stat.write(excel_r, 3, stat, fmt_stat_missing if "Eksik" in stat else fmt_stat_ok)

                    st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                else:
                    st.error("❌ Çözüm Bulunamadı.")
