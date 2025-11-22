import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import xlsxwriter

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V51", layout="wide")

st.title("🛡️ Hazırlık Ders Programı V51")

# --- YAN PANEL ---
st.sidebar.header("⚙️ Genel Ayarlar")
max_teachers_per_class = st.sidebar.slider("Sınıf Başına Max Hoca", 1, 6, 3)
allow_native_advisor = st.sidebar.checkbox("Native Hocalar Danışman Olabilir mi?", value=False)
allow_empty_slots = st.sidebar.checkbox("Sıkışınca Boş Ders Bırak (Normal Sınıflar)", value=True)

st.sidebar.markdown("---")
st.sidebar.header("🏫 Sınıf ve Zaman Ayarları")

col1, col2 = st.sidebar.columns(2)
with col1:
    count_a1 = st.number_input("A1 Sayısı", 0, 20, 4)
    time_a1 = st.selectbox("A1 Zamanı", ["Sabah", "Öğle"], key="t_a1")
    count_a2 = st.number_input("A2 Sayısı", 0, 20, 4)
    time_a2 = st.selectbox("A2 Zamanı", ["Sabah", "Öğle"], key="t_a2")
    count_pre = st.number_input("PreFac Sayısı", 0, 10, 2)
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
            'Ad Soyad': ['Ahmet Hoca', 'Sarah (Native)', 'Mehmet (Danışman)', 'Ayşe Hoca', 'Destek Hoca 1'],
            'Rol': ['Kadrolu', 'Native', 'Kadrolu', 'Ek Görevli', 'Destek'],
            'Hedef Ders Sayısı': [18, 18, 16, 8, 20],
            'Tercih (Sabah/Öğle)': ['Sabah', 'Farketmez', 'Sabah', 'Öğle', 'Farketmez'],
            'Yasaklı Günler': ['Cuma', 'Çarşamba', '', 'Pazartesi,Salı', ''],
            'Sabit Sınıf': ['', '', 'A1.01', '', ''],
            'Yetkinlik (Seviyeler)': ['A1,A2,B1', 'Hepsi', 'A1,A2', 'B1,B2', 'Hepsi'],
            'İstenmeyen Partner': ['', '', 'Ayşe Hoca', 'Mehmet (Danışman)', '']
        })
        df_teachers.to_excel(writer, sheet_name='Ogretmenler', index=False)
        
        workbook = writer.book
        worksheet = workbook.add_worksheet('NASIL KULLANILIR')
        
        # Formatlar
        header_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
        sub_header_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': '#DCE6F1', 'border': 1})
        text_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
        
        worksheet.set_column('A:A', 120)
        
        instructions = [
            "PROGRAM KULLANIM KILAVUZU",
            "",
            "1. SÜTUNLAR NASIL DOLDURULUR?",
            "• Ad Soyad: Hocanın sistemde görünecek adı.",
            "• Rol: Hocanın statüsü (Aşağıdaki 'Roller' bölümüne bakınız).",
            "• Hedef Ders Sayısı: Hocanın o hafta girmesi planlanan toplam ders saati (oturum sayısı).",
            "• Tercih: 'Sabah', 'Öğle' veya 'Farketmez'. (Not: Destek ve Native hocalar gerekirse tercihlerinin dışına yazılabilir).",
            "• Yasaklı Günler: Hocanın asla gelemeyeceği günler. Virgülle ayırın (Örn: Pazartesi,Cuma).",
            "• Sabit Sınıf: Eğer bir hoca bir sınıfın 'Danışmanı' ise, sınıfın adını buraya yazın (Örn: A1.01).",
            "• Yetkinlik: Hocanın girebileceği seviyeler. Hepsine girerse 'Hepsi' yazın.",
            "• İstenmeyen Partner: Aynı sınıfa girmesi istenmeyen hocanın tam adı.",
            "",
            "2. ROLLER VE ÖZELLİKLERİ",
            "★ KADROLU / DANIŞMAN:",
            "   - Eğer 'Sabit Sınıf' sütunu doluysa, o sınıfın danışmanı kabul edilir.",
            "   - KURAL: Danışmanlar, kendi sınıflarına PAZARTESİ girmek ZORUNDADIR.",
            "   - KURAL: Danışmanlar, kendi sınıflarına CUMA günü girmek için TEŞVİK EDİLİR (Sistem öncelik verir).",
            "   - KURAL: Haftalık ders yükü müsaitse, kendi sınıfına en az 3 farklı gün girmesi sağlanır.",
            "",
            "★ NATIVE (YABANCI HOCA):",
            "   - A1 seviyesindeki sınıflara ders verilmez.",
            "   - KURAL: Mümkün olduğunca PAZARTESİ günleri derse yazılmaz (Danışman değilse).",
            "   - Bir sınıfa haftada en fazla 1 kez Native hoca girer.",
            "",
            "★ DESTEK (DSÜ):",
            "   - Programdaki boşlukları doldurmak için kullanılır.",
            "   - KURAL: Bir sınıfa ya '1 kez' (yama olarak) ya da '3 ve üzeri kez' (danışman yardımcısı gibi) girer.",
            "   - ÖNEMLİ: Bir sınıfa haftada tam olarak '2 kez' girmesi yasaklanmıştır.",
            "",
            "★ EK GÖREVLİ:",
            "   - İdari görevi olan hocalardır.",
            "   - Bir sınıfa haftada en fazla 1 oturum ders verirler.",
            "",
            "3. OTOMATİK SİSTEM KURALLARI",
            "• Pre-Faculty Sınıfları: Bu sınıflar sadece Pazartesi, Salı ve Çarşamba günleri ders yapar. Perşembe/Cuma boştur.",
            "• Yasaklı Günler: Bu kural en katı kuraldır, sistem asla delmez.",
            "• Danışman Atama: Bir sınıfa 'Sabit Sınıf' ile hoca atanmamışsa, sistem o sınıfa en çok giren hocayı otomatik olarak 'Sınıf Danışmanı' ilan eder."
        ]
        
        row = 0
        for line in instructions:
            if "PROGRAM KULLANIM" in line:
                f = header_fmt
                worksheet.set_row(row, 30)
            elif line.startswith("1.") or line.startswith("2.") or line.startswith("3."):
                f = sub_header_fmt
                worksheet.set_row(row, 25)
            elif line.strip() == "":
                f = workbook.add_format({}) # Boş satır formatsız
            else:
                f = text_fmt
                worksheet.set_row(row, 20)
                
            worksheet.write(row, 0, line, f)
            row += 1
            
    return output.getvalue()

st.sidebar.markdown("---")
st.sidebar.download_button("📥 Kılavuzlu Şablonu İndir", generate_template(), "ogretmen_listesi_v51.xlsx")

# --- ANALİZ ---
def analyze_data(teachers, classes):
    warnings = []
    errors = []
    
    for t in teachers:
        role = str(t['Rol']).upper()
        fixed_class = str(t['Sabit Sınıf']).strip()
        
        if "DESTEK" in role and fixed_class:
             errors.append(f"🛑 **{t['Ad Soyad']}**: 'Destek' hocası danışman olamaz! Sabit Sınıf hücresini boşaltın.")
             
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
        # --- İSTATİSTİKLER ---
        total_needed_normal = sum(5 for c in classes_list if c['Seviye'] != "PreFaculty")
        total_needed_pre = sum(3 for c in classes_list if c['Seviye'] == "PreFaculty") 
        total_needed = total_needed_normal + total_needed_pre

        total_cap = sum(t['Hedef Ders Sayısı'] for t in teachers_list)
        
        num_native = sum(1 for t in teachers_list if 'Native' in str(t['Rol']))
        num_destek = sum(1 for t in teachers_list if 'Destek' in str(t['Rol']))
        num_ek = sum(1 for t in teachers_list if 'Ek Görevli' in str(t['Rol']))
        num_danisman = sum(1 for t in teachers_list if 'Danışman' in str(t['Rol']) or str(t['Sabit Sınıf']).strip() != "")

        num_sabah_sinif = sum(1 for c in classes_list if c['Zaman Kodu'] == 0)
        num_ogle_sinif = sum(1 for c in classes_list if c['Zaman Kodu'] == 1)
        
        # --- PANEL ---
        st.markdown("### 📊 Genel Durum Paneli")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Toplam Sınıf", len(classes_list))
        k2.metric("Sabah / Öğle", f"{num_sabah_sinif} / {num_ogle_sinif}")
        k3.metric("İhtiyaç (PreFac 3 Gün)", total_needed)
        k4.metric("Kapasite Durumu", f"{total_cap}", delta=total_cap - total_needed)
        
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Danışman Hoca", num_danisman)
        r2.metric("Native Hoca", num_native)
        r3.metric("Destek Hoca", num_destek)
        r4.metric("Ek Görevli", num_ek)
        
        st.divider()

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Optimizasyon yapılıyor..."):
                
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
                
                # 1. Sınıf Doluluğu (PreFaculty Pzt-Sal-Çar Kuralı)
                for c_idx, c_data in enumerate(classes_list):
                    req_session = c_data['Zaman Kodu']
                    other_session = 1 - req_session
                    
                    if c_data['Seviye'] == "PreFaculty":
                        for d in days:
                            if d <= 2: # Pzt, Sal, Çar
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 1)
                            else: # Per, Cum
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 0)
                            model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)
                    
                    else:
                        total_assigned_days = []
                        for d in days:
                            if allow_empty_slots:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
                            else:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 1)
                            model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)

                # 2. Hoca Çakışması
                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)
                
                # 3. Hedef Ders Yükü
                for t_idx, t in enumerate(teachers_list):
                    target = int(t['Hedef Ders Sayısı'])
                    total_assignments = []
                    for c in range(len(classes_list)):
                        for d in days:
                            for s in sessions: total_assignments.append(x[(t_idx, c, d, s)])
                    model.Add(sum(total_assignments) <= target)

                # 4. Sınıf Başına Max Hoca
                for c_idx in range(len(classes_list)):
                    teachers_here = []
                    for t in range(len(teachers_list)):
                        teach = model.NewBoolVar(f'tch_{t}_{c_idx}')
                        model.AddMaxEquality(teach, [x[(t, c_idx, d, s)] for d in days for s in sessions])
                        teachers_here.append(teach)
                    model.Add(sum(teachers_here) <= max_teachers_per_class)

                # 5. Native - A1 Yasağı
                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c_idx, c_data in enumerate(classes_list):
                            if c_data['Seviye'] == 'A1':
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

                # 6. Sınıf Başına Max 1 Native
                for c_idx, c_data in enumerate(classes_list):
                    natives_in_class = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'Native' in str(t['Rol']):
                            is_present = model.NewBoolVar(f'ntv_{t_idx}_{c_idx}')
                            model.AddMaxEquality(is_present, [x[(t_idx, c_idx, d, s)] for d in days for s in sessions])
                            natives_in_class.append(is_present)
                    model.Add(sum(natives_in_class) <= 1) 

                # 7. Ek Görevli (Haftada 1)
                for t_idx, t in enumerate(teachers_list):
                    if 'Ek Görevli' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            lessons_in_class = []
                            for d in days:
                                for s in sessions:
                                    lessons_in_class.append(x[(t_idx, c_idx, d, s)])
                            model.Add(sum(lessons_in_class) <= 1)
                
                # 8. Destek Hoca Kuralı (2 ders yasak)
                for t_idx, t in enumerate(teachers_list):
                    if 'Destek' in str(t['Rol']):
                        for c_idx in range(len(classes_list)):
                            total_lessons = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            is_two = model.NewBoolVar(f'is_2_{t_idx}_{c_idx}')
                            model.Add(total_lessons == 2).OnlyEnforceIf(is_two)
                            model.Add(total_lessons != 2).OnlyEnforceIf(is_two.Not())
                            model.Add(is_two == 0)

                # --- PUANLAMA ---
                objective = []
                objective.append(sum(x.values()) * 100000) 

                # Ek Görevli Maliyeti
                for t_idx, t in enumerate(teachers_list):
                    fill_score = 100000
                    if 'Ek Görevli' in str(t['Rol']): fill_score = 50000 
                    objective.append(sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions]) * fill_score)

                # Danışman Doldurma Önceliği
                for t_idx, t in enumerate(teachers_list):
                    current_load = sum([x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions])
                    if 'Danışman' in str(t['Rol']): objective.append(current_load * 100000000)
                    else: objective.append(current_load * 5000)

                # Tercih Yönetimi
                for t_idx, t in enumerate(teachers_list):
                    pref = str(t['Tercih (Sabah/Öğle)'])
                    role = str(t['Rol'])
                    is_flexible_role = "Destek" in role or "Native" in role
                    penalty_score = -5000 if is_flexible_role else -500000000

                    if pref == "Sabah":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 1)] * penalty_score)
                    elif pref == "Öğle":
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 0)] * penalty_score)

                # Yasaklı Günler
                for t_idx, t in enumerate(teachers_list):
                    forbidden = str(t['Yasaklı Günler'])
                    for d_idx, d_name in enumerate(day_names):
                        if d_name in forbidden:
                            for c in range(len(classes_list)):
                                for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -50000000)

                # Yetkinlik
                for t_idx, t in enumerate(teachers_list):
                    allowed = str(t['Yetkinlik (Seviyeler)'])
                    if "Hepsi" not in allowed:
                        for c_idx, c in enumerate(classes_list):
                            if c['Seviye'] not in allowed:
                                for d in days:
                                    for s in sessions: 
                                        objective.append(x[(t_idx, c_idx, d, s)] * -40000000)

                # NATIVE PAZARTESİ KORUMASI
                for t_idx, t in enumerate(teachers_list):
                    if 'Native' in str(t['Rol']):
                        for c in range(len(classes_list)):
                            for s in sessions: 
                                objective.append(x[(t_idx, c, 0, s)] * -50000)

                # --- SABİT SINIF (DANIŞMANLIK) ---
                for t_idx, t in enumerate(teachers_list):
                    if t['Sabit Sınıf']:
                        fixed_c = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                        if fixed_c is not None:
                            req_s = classes_list[fixed_c]['Zaman Kodu']
                            target_load = int(t['Hedef Ders Sayısı'])
                            
                            # 1. KURAL: Pazartesi KESİNLİKLE oradadır. (Hard Constraint)
                            model.Add(x[(t_idx, fixed_c, 0, req_s)] == 1)

                            # 2. TEŞVİK: Cuma günü de orada olsun. (Soft Constraint - Yüksek Puan)
                            # Not: PreFac için Cuma olmadığı için bu puan sadece normal sınıflarda işe yarar.
                            objective.append(x[(t_idx, fixed_c, 4, req_s)] * 2000000)
                            
                            # 3. KURAL: Hedef ders sayısı >= 3 ise en az 3 GÜN.
                            days_in_class = []
                            for d in days:
                                present = model.NewBoolVar(f'pres_{t_idx}_{d}')
                                model.AddMaxEquality(present, [x[(t_idx, fixed_c, d, s)] for s in sessions])
                                days_in_class.append(present)
                            
                            if target_load >= 3:
                                model.Add(sum(days_in_class) >= 3)
                            else:
                                model.Add(sum(days_in_class) >= target_load)
                            
                            objective.append(x[(t_idx, fixed_c, 0, req_s)] * 5000000)

                # İstenmeyen Partner
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
                
                # Destek Hoca Pazartesi
                for t_idx, t in enumerate(teachers_list):
                    if 'Destek' in str(t['Rol']):
                        for c in range(len(classes_list)):
                            for s in sessions: objective.append(x[(t_idx, c, 0, s)] * -100000)

                # Native Puanları
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
                    res = []
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
                                    is_flexible = "Destek" in str(t['Rol']) or "Native" in str(t['Rol'])
                                    if not is_flexible:
                                        if pref != "Farketmez" and pref != real_time:
                                            violations.append({"Hoca": t['Ad Soyad'], "Hata": f"Tercih İhlali ({pref})", "Sınıf": c_name})
                            
                            if count > 0: teacher_counts[t['Ad Soyad']] = count

                        # --- Danışman Hesaplama ---
                        adv_disp = "-"
                        if c_name in advisor_map:
                            adv_disp = advisor_map[c_name]
                        elif teacher_counts:
                            cands = {n: c for n, c in teacher_counts.items() 
                                     if not any(t['Ad Soyad'] == n and 'Ek Görevli' in str(t['Rol']) for t in teachers_list)
                                     and n not in advisor_map.values()}
                            if not allow_native_advisor:
                                cands = {n: c for n, c in cands.items() if n not in native_names}
                            
                            if cands:
                                max_v = max(cands.values())
                                winners = [n for n, c in cands.items() if c == max_v]
                                adv_disp = " / ".join(winners)

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
                        res.append(row)
                    
                    # --- EKRANA YAZDIRMA ---
                    if violations:
                        st.warning(f"⚠️ {len(violations)} adet kural esnetildi.")
                        df_violations = pd.DataFrame(violations).drop_duplicates()
                        st.table(df_violations)
                    else:
                        st.success("✅ Kurallar Tamam")
                        df_violations = pd.DataFrame()

                    df_res = pd.DataFrame(res)
                    st.dataframe(df_res)

                    # İstatistikler
                    stats_data = []
                    for t_idx, t in enumerate(teachers_list):
                        assigned_count = 0
                        for c in range(len(classes_list)):
                            for d in range(5):
                                for s in range(2):
                                    if solver.Value(x[(t_idx, c, d, s)]) == 1: assigned_count += 1
                        target = int(t['Hedef Ders Sayısı'])
                        diff = assigned_count - target
                        status_text = "Tamam"
                        if diff > 0: status_text = f"+{diff} Fazla"
                        elif diff < 0: status_text = f"{diff} Eksik"
                        stats_data.append({"Hoca Adı": t['Ad Soyad'],"Hedef": target,"Atanan": assigned_count,"Durum": status_text})
                    
                    df_stats = pd.DataFrame(stats_data)
                    st.dataframe(df_stats)

                    # --- EXCEL ÇIKTISI ---
                    output_res = io.BytesIO()
                    with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
                        df_res.to_excel(writer, index=False, sheet_name="Program")
                        df_stats.to_excel(writer, index=False, sheet_name="Istatistikler")
                        if not df_violations.empty: df_violations.to_excel(writer, index=False, sheet_name="Ihlal_Raporu")
                        
                        wb = writer.book
                        ws_prog = writer.sheets['Program']
                        
                        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
                        fmt_gold = wb.add_format(dict(base_fmt, bg_color='#FFD700'))
                        fmt_orange = wb.add_format(dict(base_fmt, bg_color='#FFA500'))
                        fmt_maroon = wb.add_format(dict(base_fmt, bg_color='#800000', font_color='white'))
                        fmt_green = wb.add_format(dict(base_fmt, bg_color='#006400', font_color='white'))
                        fmt_blue = wb.add_format(dict(base_fmt, bg_color='#ADD8E6', bold=True)) 
                        fmt_default = wb.add_format(base_fmt)

                        ws_prog.set_column('A:B', 8)
                        ws_prog.set_column('C:C', 20)
                        ws_prog.set_column('E:I', 12)
                        ws_prog.set_row(0, 20)

                        for r, row in df_res.iterrows():
                            excel_r = r + 1
                            ws_prog.set_row(excel_r, 20)
                            lvl = str(row['Seviye'])
                            if lvl == "PreFaculty": f_row = fmt_default
                            elif lvl == "A1": f_row = fmt_gold
                            elif lvl == "A2": f_row = fmt_orange
                            elif lvl == "B1": f_row = fmt_maroon
                            else: f_row = fmt_green
                            
                            ws_prog.write(excel_r, 0, row['Sınıf'], f_row)
                            ws_prog.write(excel_r, 1, row['Seviye'], fmt_default)
                            ws_prog.write(excel_r, 2, row['Sınıf Danışmanı'], fmt_default)
                            ws_prog.write(excel_r, 3, row['Zaman'], fmt_default)
                            
                            for c in range(4, 9):
                                val = row.iloc[c]
                                f = fmt_default
                                if val in native_names: f = fmt_blue
                                ws_prog.write(excel_r, c, val, f)

                    st.download_button("Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                else:

                    st.error("❌ Çözüm Bulunamadı.")


