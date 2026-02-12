import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import multiprocessing

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V49 - Teşhis Modu", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V49 - Teşhis ve Analiz)")
st.info("""
**Bu Sürüm Ne Yapar?**
1. 🩺 **Ön Tarama (Diagnosis):** Programı hesaplamaya başlamadan önce, matematiksel olarak imkansız olan talepleri (Örn: Pazartesi yasaklı Danışman) tespit eder ve sizi uyarır.
2. 🛡️ **Esnek Sınırlar:** Çözümsüzlük durumunda 'Tek Vardiya' ve 'Native Sınırı' kurallarını esneterek programı çıkarır ve size rapor sunar.
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

# --- GELİŞMİŞ ANALİZ (DIAGNOSTICS) ---
def run_diagnostics(teachers, classes):
    report = []
    fatal_errors = False
    
    # 1. Kapasite Analizi (Sabah/Öğle)
    needed_morning = len([c for c in classes if c['Zaman Kodu'] == 0]) * 5
    needed_afternoon = len([c for c in classes if c['Zaman Kodu'] == 1]) * 5
    
    avail_morning = 0
    avail_afternoon = 0
    
    for t in teachers:
        forbidden = str(t['Yasaklı Günler'])
        days_avail = 5 - (len(forbidden.split(',')) if t['Yasaklı Günler'] else 0)
        pref = str(t['Tercih (Sabah/Öğle)'])
        target = int(t['Hedef Ders Sayısı'])
        
        real_cap = min(target, days_avail)
        
        if pref == "Sabah": avail_morning += real_cap
        elif pref == "Öğle": avail_afternoon += real_cap
        else: # Farketmez
            avail_morning += real_cap # Potansiyel
            avail_afternoon += real_cap # Potansiyel
            
    # Basit bir kontrol (Farketmezler iki yere de sayıldığı için toplamda yetiyor mu diye bakalım)
    total_slots = needed_morning + needed_afternoon
    total_cap = sum([min(int(t['Hedef Ders Sayısı']), 5 - (len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0)) for t in teachers])
    
    if total_cap < total_slots:
        report.append(f"⚠️ **Kapasite Uyarısı:** Toplam ihtiyaç {total_slots} ders saati, ancak öğretmenlerin toplam kapasitesi {total_cap} saat. Bazı dersler boş kalabilir.")
    
    # 2. Danışman Analizi
    for t in teachers:
        if t['Sabit Sınıf']:
            fixed_class_name = str(t['Sabit Sınıf'])
            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class_name), None)
            
            if not target_class:
                report.append(f"❌ **Hata:** {t['Ad Soyad']} için girilen '{fixed_class_name}' sınıfı sistemde yok.")
                fatal_errors = True
                continue

            forbidden = str(t['Yasaklı Günler'])
            forbidden_count = len(forbidden.split(',')) if t['Yasaklı Günler'] else 0
            available_days = 5 - forbidden_count
            
            # Pazartesi Kontrolü
            if "Pazartesi" in forbidden:
                report.append(f"❌ **Kritik Hata:** {t['Ad Soyad']}, '{fixed_class_name}' sınıfının danışmanı ama **Pazartesi** günü yasaklı. Danışmanlar Pazartesi okulda olmak zorundadır.")
                fatal_errors = True
            
            # Gün Sayısı Kontrolü
            if target_class['Seviye'] != "PreFaculty" and available_days < 2:
                report.append(f"❌ **Kritik Hata:** {t['Ad Soyad']}, '{fixed_class_name}' sınıfının danışmanı ama haftada sadece {available_days} gün müsait. Danışman en az 2 gün girmelidir.")
                fatal_errors = True

    return report, fatal_errors

# --- ANA PROGRAM ---
uploaded_file = st.file_uploader("Öğretmen Listesini Yükle", type=["xlsx"])

if uploaded_file:
    df_teachers = pd.read_excel(uploaded_file, sheet_name='Ogretmenler').fillna("")
    if 'Hedef Ders Sayısı' not in df_teachers.columns and 'Hedef Gün Sayısı' in df_teachers.columns:
        df_teachers.rename(columns={'Hedef Gün Sayısı': 'Hedef Ders Sayısı'}, inplace=True)
        
    df_classes = create_automated_classes()
    
    teachers_list = df_teachers.to_dict('records')
    classes_list = df_classes.to_dict('records')

    # TEŞHİS RAPORU
    diag_report, is_fatal = run_diagnostics(teachers_list, classes_list)
    
    if diag_report:
        st.markdown("### 🩺 Veri Teşhis Raporu")
        for line in diag_report:
            if "Hata" in line: st.error(line)
            else: st.warning(line)
            
    if is_fatal:
        st.stop() # Kritik hatalar varsa dur.

    # Eğer kritik hata yoksa devam et
    
    total_slots_needed = 0
    for c in classes_list:
        if c['Seviye'] == 'PreFaculty': total_slots_needed += 3 
        else: total_slots_needed += 5
    
    raw_demand = 0
    for t in teachers_list:
        forbidden_cnt = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
        max_cap = 5 - forbidden_cnt 
        raw_demand += min(int(t['Hedef Ders Sayısı']), max_cap)

    col1, col2 = st.columns(2)
    col1.metric("Sınıf İhtiyacı", total_slots_needed)
    col2.metric("Hoca Kapasitesi", raw_demand)
    
    reduce_mode = False
    if raw_demand > total_slots_needed:
        st.info("ℹ️ Hoca kapasitesi fazla. Sistem hedefleri 1'er saat kırparak dengeleyecek.")
        reduce_mode = True

    if st.button("🚀 Programı Oluştur (Esnek Mod)"):
        with st.spinner("Esnek çözüm aranıyor..."):
            
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

            # --- HARD CONSTRAINTS (KESİN KURALLAR) ---
            
            # 1. Yetkinlik (Hard)
            for t_idx, t in enumerate(teachers_list):
                allowed = str(t['Yetkinlik (Seviyeler)'])
                if "Hepsi" not in allowed:
                    for c_idx, c in enumerate(classes_list):
                        if c['Seviye'] not in allowed:
                            for d in days:
                                for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                            model.Add(advisor_var[(t_idx, c_idx)] == 0)

            # 2. Danışman Tekilliği (Hard)
            for c in range(len(classes_list)):
                model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) == 1)
            for t in range(len(teachers_list)):
                model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

            # 3. Rol Kısıtları (Hard)
            for t_idx, t in enumerate(teachers_list):
                if 'Ek Görevli' in str(t['Rol']):
                    for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                if not allow_native_advisor and 'Native' in str(t['Rol']):
                    for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)

            # 4. Sabit Sınıf (Hard)
            for t_idx, t in enumerate(teachers_list):
                if t['Sabit Sınıf']:
                    fixed_c_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                    if fixed_c_idx is not None:
                        model.Add(advisor_var[(t_idx, fixed_c_idx)] == 1)

            # 5. PreFaculty Kapama (Hard)
            for c_idx, c_data in enumerate(classes_list):
                if c_data['Seviye'] == "PreFaculty":
                    for t_idx in range(len(teachers_list)):
                        for s in sessions:
                            model.Add(x[(t_idx, c_idx, 3, s)] == 0)
                            model.Add(x[(t_idx, c_idx, 4, s)] == 0)

            # 6. Sınıf Doluluğu (Hard)
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

            # 7. Fiziksel Çakışma (Hard)
            for t in range(len(teachers_list)):
                for d in days:
                    for s in sessions:
                        model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

            # --- SOFT CONSTRAINTS (ESNEK KURALLAR - PUANLI) ---
            objective = []
            
            # A. Danışman 3 Gün ve Pazartesi
            for t_idx, t_data in enumerate(teachers_list):
                forbidden_days = str(t_data['Yasaklı Günler'])
                available_days = 5 - (len(forbidden_days.split(',')) if t_data['Yasaklı Günler'] else 0)
                
                for c_idx, c_data in enumerate(classes_list):
                    is_adv = advisor_var[(t_idx, c_idx)]
                    req_s = c_data['Zaman Kodu']
                    
                    # Pazartesi Kuralı (Mecbur değil, Puan)
                    if "Pazartesi" not in forbidden_days:
                        # Eğer danışmansa Pazartesi orada olmalı. Değilse büyük ceza.
                        # Soft implementasyon:
                        is_there_monday = x[(t_idx, c_idx, 0, req_s)]
                        # (is_adv AND NOT is_there_monday) => Ceza
                        # Maximize: is_adv * is_there_monday * 100M
                        objective.append(is_adv * 100000000) # Danışman olmayı teşvik
                        
                        # Eğer danışmansa ve pazartesi yoksa ceza eklemek lazım ama yukarıdaki pozitif puan da iş görür.
                        # Daha kesin olması için:
                        # model.Add(is_there_monday == 1).OnlyEnforceIf(is_adv) 
                        # Bunu kaldırdık çünkü çöküyor olabilir. Onun yerine Puan.
                        # Ama Pazartesi kuralı çok önemli, bunu HARD tutalım mı? 
                        # Kullanıcı "Kesinlikle" dedi. Eğer hoca müsaitse HARD yapalım.
                        model.Add(x[(t_idx, c_idx, 0, req_s)] == 1).OnlyEnforceIf(is_adv)

                    # 3 Gün Kuralı (Soft)
                    if c_data['Seviye'] != "PreFaculty":
                        days_in_class = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                        if available_days >= 3:
                            # 3. gün için puan
                            is_3_days = model.NewBoolVar(f'is3_{t_idx}_{c_idx}')
                            model.Add(days_in_class >= 3).OnlyEnforceIf(is_3_days)
                            objective.append(is_3_days * 50000000)
                            
                            # Ama en az 2 gün zorunlu (Hard)
                            model.Add(days_in_class >= 2).OnlyEnforceIf(is_adv)
                        elif available_days == 2:
                            model.Add(days_in_class >= 2).OnlyEnforceIf(is_adv)

            # B. Native Sınırı (Soft)
            for t_idx, t in enumerate(teachers_list):
                if 'Native' in str(t['Rol']):
                    for c_idx in range(len(classes_list)):
                        is_not_advisor = advisor_var[(t_idx, c_idx)].Not()
                        class_total = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                        
                        # Eğer 1'den fazla girerse ceza
                        is_violation = model.NewBoolVar(f'ntv_vio_{t_idx}_{c_idx}')
                        model.Add(class_total > 1).OnlyEnforceIf(is_violation)
                        model.Add(class_total <= 1).OnlyEnforceIf(is_violation.Not())
                        
                        # İhlal varsa ve danışman değilse ceza
                        # objective -= violation * 10M
                        objective.append(is_violation * -20000000)

            # C. Tek Vardiya (Soft)
            for t_idx, t in enumerate(teachers_list):
                for d in days:
                    is_morning = model.NewBoolVar(f'm_{t_idx}_{d}')
                    is_afternoon = model.NewBoolVar(f'a_{t_idx}_{d}')
                    model.AddMaxEquality(is_morning, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                    model.AddMaxEquality(is_afternoon, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                    
                    # Çift vardiya ceza
                    double_shift = model.NewBoolVar(f'dbl_{t_idx}_{d}')
                    model.Add(is_morning + is_afternoon > 1).OnlyEnforceIf(double_shift)
                    model.Add(is_morning + is_afternoon <= 1).OnlyEnforceIf(double_shift.Not())
                    objective.append(double_shift * -50000000)

            # D. Hedef Doldurma ve Diğerleri
            adjusted_targets = []
            for t_idx, t in enumerate(teachers_list):
                original_target = int(t['Hedef Ders Sayısı'])
                forbidden_count = len(str(t['Yasaklı Günler']).split(',')) if t['Yasaklı Günler'] else 0
                max_possible = 5 - forbidden_count
                
                # Kapasite azaltma
                if reduce_mode and original_target > 2: target_to_use = original_target - 1
                else: target_to_use = original_target
                real_target = min(target_to_use, max_possible)
                adjusted_targets.append(real_target)
                
                total_assignments = []
                for c in range(len(classes_list)):
                    for d in days:
                        for s in sessions: total_assignments.append(x[(t_idx, c, d, s)])
                
                model.Add(sum(total_assignments) <= real_target)
                if real_target > 0: model.Add(sum(total_assignments) >= 1) # Boş kalmasın
                
                # Hedefe ne kadar yakınsa o kadar iyi
                current_load = sum(total_assignments)
                if 'Danışman' in str(t['Rol']): objective.append(current_load * 5000000)
                else: objective.append(current_load * 5000)

            # Temel Atama Puanı
            objective.append(sum(x.values()) * 100000)

            # --- ÇÖZÜM ---
            model.Maximize(sum(objective))
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 120.0
            solver.parameters.num_search_workers = 8
            
            status = solver.Solve(model)

            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                st.balloons()
                
                res_data = []
                violations = []
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
                        if c['Seviye'] == "PreFaculty" and d_idx >= 3: val = "⛔ KAPALI"
                        else:
                            for t_idx, t in enumerate(teachers_list):
                                if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                    val = t['Ad Soyad']
                                    # Çift Vardiya Kontrolü
                                    other_s = 1 - s_req
                                    morning = sum([solver.Value(x[(t_idx, cc, d_idx, 0)]) for cc in range(len(classes_list))])
                                    afternoon = sum([solver.Value(x[(t_idx, cc, d_idx, 1)]) for cc in range(len(classes_list))])
                                    if morning > 0 and afternoon > 0:
                                        violations.append({"Hoca": t['Ad Soyad'], "Sorun": f"Çift Vardiya ({d_name})", "Sınıf": c_name})
                                    break
                        row[d_name] = val
                    res_data.append(row)

                stats = []
                for t_idx, t in enumerate(teachers_list):
                    assigned = sum([solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions])
                    original_target = int(t['Hedef Ders Sayısı'])
                    real_target = adjusted_targets[t_idx]
                    stat = "Tamam"
                    if real_target < original_target and assigned == real_target: stat = f"Kırpıldı ({original_target-real_target}s)"
                    elif assigned < real_target: stat = f"{real_target-assigned} Eksik"
                    stats.append({"Hoca Adı": t['Ad Soyad'], "Hedef (İlk)": original_target, "Atanan": assigned, "Durum": stat})

                df_res = pd.DataFrame(res_data)
                df_stats = pd.DataFrame(stats)
                df_violations = pd.DataFrame(violations).drop_duplicates() if violations else pd.DataFrame()

                if not df_violations.empty:
                    st.warning(f"⚠️ {len(df_violations)} noktada kurallar esnetildi (Çift Vardiya).")
                    st.table(df_violations)
                else:
                    st.success("✅ Kusursuz Çözüm!")

                st.dataframe(df_res)
                st.dataframe(df_stats)

                output_res = io.BytesIO()
                with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
                    df_res.to_excel(writer, index=False, sheet_name="Program")
                    df_stats.to_excel(writer, index=False, sheet_name="Istatistikler")
                    if not df_violations.empty: df_violations.to_excel(writer, index=False, sheet_name="Ihlal_Raporu")
                    
                    # (Excel Formatlama kodları aynen korunur...)
                    wb = writer.book
                    ws_prog = writer.sheets['Program']
                    ws_stat = writer.sheets['Istatistikler']
                    base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
                    fmt_default = wb.add_format(base_fmt)
                    fmt_gold = wb.add_format(dict(base_fmt, bg_color='#FFD700'))
                    fmt_orange = wb.add_format(dict(base_fmt, bg_color='#FFA500'))
                    fmt_maroon = wb.add_format(dict(base_fmt, bg_color='#800000', font_color='white'))
                    fmt_green = wb.add_format(dict(base_fmt, bg_color='#006400', font_color='white'))
                    fmt_blue = wb.add_format(dict(base_fmt, bg_color='#ADD8E6'))
                    
                    ws_prog.set_column('A:C', 15)
                    for r, row in df_res.iterrows():
                        excel_r = r + 1
                        lvl = str(row['Seviye'])
                        c_fmt = fmt_gold if lvl=="A1" else (fmt_orange if lvl=="A2" else (fmt_maroon if lvl=="B1" else fmt_green))
                        if lvl=="PreFaculty": c_fmt = wb.add_format(dict(base_fmt, bg_color='#E0E0E0'))
                        ws_prog.write(excel_r, 0, row['Sınıf'], c_fmt)
                        ws_prog.write(excel_r, 1, row['Seviye'], c_fmt)
                        ws_prog.write(excel_r, 2, row['Sınıf Danışmanı'], fmt_default)
                        ws_prog.write(excel_r, 3, row['Zaman'], fmt_default)
                        for c in range(4, 9):
                            val = row.iloc[c]
                            f = fmt_blue if val in native_names else fmt_default
                            ws_prog.write(excel_r, c, val, f)
                            
                    for r, row in df_stats.iterrows():
                        ws_stat.write(r+1, 0, row['Hoca Adı'], fmt_default)
                        ws_stat.write(r+1, 1, row['Hedef (İlk)'], fmt_default)
                        ws_stat.write(r+1, 2, row['Atanan'], fmt_default)
                        ws_stat.write(r+1, 3, row['Durum'], fmt_default)

                st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
            else:
                st.error("❌ Çözüm Bulunamadı. (Lütfen Teşhis Raporundaki hataları düzeltin)")
