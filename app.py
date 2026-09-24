import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import collections

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Ders Programı V61 - Tam Entegre Motor", layout="wide")

st.title("🛡️ Hazırlık Ders Programı (V61 - Tam Entegre Motor)")
st.info("""
**V61 Güncellemeleri:**
- **Sınıf Başına Max Hoca** kısıtı matematiksel modele tam bağlandı.
- **Sıkışınca Boş Ders Bırak** ayarı aktif kural haline getirildi (Kapatılırsa tüm derslerin dolması zorunludur).
- **İstenmeyen Partner** kuralı devreye alındı (Birbirini istemeyen hocalar aynı sınıfa giremez).
- **Yasaklı Günler** için kesin (Hard) veya esnek (Soft) kural seçeneği eklendi.
""")

# --- YAN PANEL ---
st.sidebar.header("⚙️ Genel Ayarlar")
max_teachers_per_class = st.sidebar.slider("Sınıf Başına Max Hoca", 1, 6, 3, 
                                           help="Bir sınıfa hafta boyunca girebilecek maksimum farklı öğretmen sayısı.")
allow_native_advisor = st.sidebar.checkbox("Native Hocalar Danışman Olabilir mi?", value=False)
allow_empty_slots = st.sidebar.checkbox("Sıkışınca Boş Ders Bırak", value=True, 
                                        help="Kapatılırsa tüm derslerin dolu olması zorunlu olur (Öğretmen yetersizse çözüm bulunamaz).")
strict_forbidden_days = st.sidebar.checkbox("Yasaklı Günleri Kesin Kural Yap (Hard)", value=False, 
                                           help="İşaretlenirse yasaklı günlere ASLA ders konulamaz. İşaretsizse ağır ceza puanı ile esnetilebilir.")

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
    return output.getvalue()

st.sidebar.markdown("---")
st.sidebar.download_button("📥 Kılavuzlu Şablonu İndir", generate_template(), "ogretmen_listesi.xlsx")

# --- GÜVENLİ STRİNG OKUYUCULAR ---
def get_role(t):
    return str(t.get('Rol', '')).upper().replace('İ', 'I').replace('i', 'I').replace('ı', 'I')

def get_pref(t):
    return str(t.get('Tercih (Sabah/Öğle)', 'Farketmez')).upper().replace('Ö', 'O').replace('ö', 'O').replace('Ğ', 'G').replace('ğ', 'G').strip()

def get_partners(t):
    val = str(t.get('İstenmeyen Partner', '')).strip()
    if not val or val.lower() == 'nan':
        return []
    return [p.strip() for p in val.split(',') if p.strip()]

# --- ANALİZ VE DOĞRULAMA ---
def analyze_data(teachers, classes):
    warnings = []
    errors = []
    
    assigned_fixed = []
    teacher_names = [str(t.get('Ad Soyad', '')).strip() for t in teachers]

    for t in teachers:
        t_name = str(t.get('Ad Soyad', '')).strip()
        role = get_role(t)
        fixed_class = str(t.get('Sabit Sınıf', '')).strip()
        forbidden_str = str(t.get('Yasaklı Günler', ''))

        # Sabit Sınıf Kontrolleri
        if fixed_class:
            if not allow_native_advisor and "NATIVE" in role:
                 errors.append(f"🛑 **{t_name}**: Native hocaya sabit sınıf verilmesi engellendi (Ayarlardan açabilirsiniz).")
            if "EK GÖREVL" in role:
                 errors.append(f"🛑 **{t_name}**: Ek Görevli rolündekilere sabit sınıf verilemez.")

            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class), None)
            if not target_class:
                errors.append(f"❌ **{t_name}**: Atandığı '{fixed_class}' sınıfı sistemde bulunamadı.")
            else:
                assigned_fixed.append(fixed_class)

            if "Pazartesi" in forbidden_str:
                warnings.append(f"⚠️ **Uyarı ({t_name}):** '{fixed_class}' danışmanı olarak belirlenmiş ancak Pazartesi günü yasaklı.")

        # İstenmeyen Partner Kontrolleri
        partners = get_partners(t)
        for p in partners:
            if p not in teacher_names:
                warnings.append(f"⚠️ **Uyarı ({t_name}):** İstenmeyen partner olarak belirtilen '{p}' listede bulunamadı.")
            elif p == t_name:
                errors.append(f"❌ **{t_name}**: Kendisini istenmeyen partner olarak seçemez.")

    dupes = [item for item, count in collections.Counter(assigned_fixed).items() if count > 1]
    if dupes:
        errors.append(f"❌ **ÇAKIŞMA HATASI:** {', '.join(dupes)} sınıfına 1'den fazla öğretmen sabitlenmiş!")

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
        st.error("🛑 Lütfen Excel'deki mantıksal hataları düzeltin:")
        for e in logic_errors: st.markdown(e)
    else:
        if logic_warnings:
            for w in logic_warnings: st.warning(w)

        # İhtiyaçlar ve Kapasite
        morning_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 0])
        afternoon_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 1])
        total_slots_needed = morning_needs + afternoon_needs

        morning_cap, afternoon_cap, farketmez_cap = 0, 0, 0
        base_targets = []
        
        for t in teachers_list:
            forbidden_cnt = len(str(t['Yasaklı Günler']).split(',')) if str(t['Yasaklı Günler']).strip() else 0
            teacher_cap = min(int(t['Hedef Ders Sayısı']), 5 - forbidden_cnt)
            base_targets.append(teacher_cap)
            
            pref = get_pref(t)
            if 'SABAH' in pref: morning_cap += teacher_cap
            elif 'OGLE' in pref: afternoon_cap += teacher_cap
            else: farketmez_cap += teacher_cap

        raw_demand = sum(base_targets)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Sabah İhtiyacı", morning_needs)
        col2.metric("Sabah Kapasitesi", f"{morning_cap} (+{farketmez_cap})")
        col3.metric("Öğle İhtiyacı", afternoon_needs)
        col4.metric("Öğle Kapasitesi", f"{afternoon_cap} (+{farketmez_cap})")

        # HİYERARŞİK KIRPMA MANTIĞI (NATIVE KORUMALI)
        adjusted_targets = list(base_targets)
        excess_capacity = raw_demand - total_slots_needed

        if excess_capacity > 0:
            st.info(f"ℹ️ Öğretmen kapasitesi {excess_capacity} saat fazla. Ek Görevli ve Destek hocalarından adil kırpma yapılacaktır (Native Hariç).")
            
            ek_idx = [i for i, t in enumerate(teachers_list) if 'EK GÖREVL' in get_role(t)]
            destek_idx = [i for i, t in enumerate(teachers_list) if 'DESTEK' in get_role(t)]
            other_idx = [i for i, t in enumerate(teachers_list) if i not in ek_idx and i not in destek_idx and 'NATIVE' not in get_role(t)]

            exc = excess_capacity
            while exc > 0:
                trimmed = False
                
                if any(adjusted_targets[i] > 0 for i in ek_idx):
                    for i in ek_idx:
                        if exc == 0: break
                        if adjusted_targets[i] > 0:
                            adjusted_targets[i] -= 1
                            exc -= 1
                            trimmed = True
                    if trimmed: continue 
                
                if any(adjusted_targets[i] > (1 if str(teachers_list[i].get('Sabit Sınıf','')).strip() else 0) for i in destek_idx):
                    for i in destek_idx:
                        if exc == 0: break
                        min_req = 1 if str(teachers_list[i].get('Sabit Sınıf','')).strip() else 0
                        if adjusted_targets[i] > min_req:
                            adjusted_targets[i] -= 1
                            exc -= 1
                            trimmed = True
                    if trimmed: continue
                
                if any(adjusted_targets[i] > (1 if str(teachers_list[i].get('Sabit Sınıf','')).strip() else 0) for i in other_idx):
                    for i in other_idx:
                        if exc == 0: break
                        min_req = 1 if str(teachers_list[i].get('Sabit Sınıf','')).strip() else 0
                        if adjusted_targets[i] > min_req:
                            adjusted_targets[i] -= 1
                            exc -= 1
                            trimmed = True
                    if trimmed: continue
                
                if not trimmed: break 
                
        elif excess_capacity < 0:
            if not allow_empty_slots:
                st.error(f"❌ Kapasite {abs(excess_capacity)} saat yetersiz ve 'Boş Ders Bırak' ayarı KAPALI. Bu ayarlarla program oluşturulamaz!")
            else:
                st.warning(f"⚠️ Kapasite yetersiz. {abs(excess_capacity)} ders saati zorunlu olarak boş kalacak.")

        if st.button("🚀 Programı Oluştur"):
            with st.spinner("Matematiksel model kuruluyor ve çözülüyor..."):
                model = cp_model.CpModel()
                days = range(5)
                day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                sessions = range(2)

                x = {}
                advisor_var = {}
                teacher_in_class = {}

                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        advisor_var[(t, c)] = model.NewBoolVar(f'adv_{t}_{c}')
                        teacher_in_class[(t, c)] = model.NewBoolVar(f'tinclass_{t}_{c}')
                        for d in days:
                            for s in sessions:
                                x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                # --- HOCA SINIFTA VAR MI DEĞİŞKENİ BAĞLANTISI ---
                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        # teacher_in_class[(t, c)] == 1 ancak ve ancak t hocası c sınıfına en az 1 ders giriyorsa
                        model.AddMaxEquality(teacher_in_class[(t, c)], [x[(t, c, d, s)] for d in days for s in sessions])

                # --- 1. KESİN KURALLAR (HARD CONSTRAINTS) ---

                # 1.1 Sınıf Başına Max Hoca Kuralı (Slider Entegrasyonu)
                for c in range(len(classes_list)):
                    model.Add(sum(teacher_in_class[(t, c)] for t in range(len(teachers_list))) <= max_teachers_per_class)

                # 1.2 İstenmeyen Partner Kuralı (Aynı sınıfa girememe)
                name_to_idx = {str(t['Ad Soyad']).strip(): idx for idx, t in enumerate(teachers_list)}
                processed_pairs = set()
                for t_idx, t in enumerate(teachers_list):
                    t_name = str(t['Ad Soyad']).strip()
                    partners = get_partners(t)
                    for p in partners:
                        if p in name_to_idx:
                            p_idx = name_to_idx[p]
                            pair = tuple(sorted([t_idx, p_idx]))
                            if pair not in processed_pairs and t_idx != p_idx:
                                processed_pairs.add(pair)
                                for c in range(len(classes_list)):
                                    # İki partner aynı sınıfta birlikte bulunamaz
                                    model.Add(teacher_in_class[(t_idx, c)] + teacher_in_class[(p_idx, c)] <= 1)

                # 1.3 Öğretmen Çakışması (Aynı anda 1 sınıfta)
                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

                # 1.4 Sınıf Zaman Kodu & Sıkışınca Boş Ders Bırak Kuralı
                for c_idx, c_data in enumerate(classes_list):
                    req_session = c_data['Zaman Kodu']
                    other_session = 1 - req_session
                    for d in days:
                        if c_data['Seviye'] == "PreFaculty" and d >= 3:
                            model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 0)
                        else:
                            if allow_empty_slots:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
                            else:
                                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 1)

                        model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)

                # 1.5 Yetkinlik Seviyeleri Kısıtı
                for t_idx, t in enumerate(teachers_list):
                    allowed = str(t.get('Yetkinlik (Seviyeler)', '')).strip()
                    if allowed == "": allowed = "Hepsi"
                        
                    if "Hepsi" not in allowed:
                        for c_idx, c in enumerate(classes_list):
                            if c['Seviye'] not in allowed:
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                                model.Add(advisor_var[(t_idx, c_idx)] == 0)

                # 1.6 Danışman Tekillikleri
                for c in range(len(classes_list)):
                    model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) <= 1)
                for t in range(len(teachers_list)):
                    model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

                # 1.7 Sabit Sınıf Ataması
                for t_idx, t in enumerate(teachers_list):
                    fixed_c_name = str(t.get('Sabit Sınıf', '')).strip()
                    if fixed_c_name:
                        fixed_c_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == fixed_c_name), None)
                        if fixed_c_idx is not None:
                            model.Add(advisor_var[(t_idx, fixed_c_idx)] == 1)

                # 1.8 Rol ve Danışmanlık Kısıtları
                for t_idx, t in enumerate(teachers_list):
                    role = get_role(t)
                    if 'EK GÖREVL' in role:
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                    if not allow_native_advisor and 'NATIVE' in role:
                        for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)

                # 1.9 Native A1 Yasağı
                for t_idx, t in enumerate(teachers_list):
                    if 'NATIVE' in get_role(t):
                        for c_idx, c_data in enumerate(classes_list):
                            if c_data['Seviye'] == 'A1':
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

                # 1.10 Ek Görevli: Aynı Sınıfa Haftada Max 1 Gün
                for t_idx, t in enumerate(teachers_list):
                    if 'EK GÖREVL' in get_role(t):
                        for c_idx in range(len(classes_list)):
                            model.Add(sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions) <= 1)

                # 1.11 Tavan Hedef Kısıtı
                for t_idx, t in enumerate(teachers_list):
                    real_target = adjusted_targets[t_idx]
                    total_assignments = sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions)
                    model.Add(total_assignments <= real_target)

                # 1.12 Yasaklı Günler (Hard Constraint Seçilmişse)
                if strict_forbidden_days:
                    for t_idx, t in enumerate(teachers_list):
                        forbidden = str(t.get('Yasaklı Günler', ''))
                        for d_idx, d_name in enumerate(day_names):
                            if d_name in forbidden:
                                for c in range(len(classes_list)):
                                    for s in sessions:
                                        model.Add(x[(t_idx, c, d_idx, s)] == 0)

                # --- 2. PUANLI KURALLAR VE ŞELALELER (OBJECTIVE) ---
                objective = []
                objective.append(sum(x.values()) * 100000000)

                # NATIVE ŞELALESİ VE SINIF BAŞINA MAX 1 NATIVE KURALI
                for c_idx, c_data in enumerate(classes_list):
                    native_in_this_class = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'NATIVE' in get_role(t):
                            is_present = teacher_in_class[(t_idx, c_idx)]
                            native_in_this_class.append(is_present)
                            
                            lvl = c_data['Seviye']
                            if lvl == "B2": score = 10000000
                            elif lvl == "B1": score = 1000000
                            elif lvl == "A2": score = 100000
                            elif lvl == "PreFaculty": score = 10000
                            else: score = 0 
                            
                            objective.append(is_present * score)
                            
                    if native_in_this_class:
                        model.Add(sum(native_in_this_class) <= 1)

                # Danışmanlık ve Pazartesi Teşvikleri
                for t_idx, t_data in enumerate(teachers_list):
                    forbidden_days = str(t_data.get('Yasaklı Günler', ''))
                    for c_idx, c_data in enumerate(classes_list):
                        is_adv = advisor_var[(t_idx, c_idx)]
                        req_s = c_data['Zaman Kodu']

                        if "Pazartesi" not in forbidden_days:
                            pzt_var = x[(t_idx, c_idx, 0, req_s)]
                            adv_pzt = model.NewBoolVar(f'ap_{t_idx}_{c_idx}')
                            model.AddImplication(adv_pzt, pzt_var)
                            model.AddImplication(adv_pzt, is_adv)
                            objective.append(adv_pzt * 50000000)

                        if c_data['Seviye'] != "PreFaculty":
                            days_in_class = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            
                            is_2plus = model.NewBoolVar(f'is2_{t_idx}_{c_idx}')
                            model.Add(days_in_class >= 2).OnlyEnforceIf(is_2plus)
                            model.Add(days_in_class <= 1).OnlyEnforceIf(is_2plus.Not())
                            
                            is_3plus = model.NewBoolVar(f'is3_{t_idx}_{c_idx}')
                            model.Add(days_in_class >= 3).OnlyEnforceIf(is_3plus)
                            model.Add(days_in_class <= 2).OnlyEnforceIf(is_3plus.Not())

                            adv_2days = model.NewBoolVar(f'adv2_{t_idx}_{c_idx}')
                            model.AddImplication(adv_2days, is_2plus)
                            model.AddImplication(adv_2days, is_adv)
                            objective.append(adv_2days * 20000000)

                            adv_3days = model.NewBoolVar(f'adv3_{t_idx}_{c_idx}')
                            model.AddImplication(adv_3days, is_3plus)
                            model.AddImplication(adv_3days, is_adv)
                            objective.append(adv_3days * 20000000)

                # Sabah / Öğle Tercih Cezaları
                for t_idx, t in enumerate(teachers_list):
                    pref = get_pref(t)
                    if 'SABAH' in pref:
                        for c in range(len(classes_list)):
                            for d in days:
                                objective.append(x[(t_idx, c, d, 1)] * -10000000)
                    elif 'OGLE' in pref:
                        for c in range(len(classes_list)):
                            for d in days:
                                objective.append(x[(t_idx, c, d, 0)] * -10000000)

                # Native Sınıf Başına 1 Gün Sınırı (Ceza)
                for t_idx, t in enumerate(teachers_list):
                    if 'NATIVE' in get_role(t):
                        for c_idx in range(len(classes_list)):
                            class_total = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            is_violation = model.NewBoolVar(f'ntv_vio_{t_idx}_{c_idx}')
                            model.Add(class_total <= 1 + 5 * is_violation)
                            objective.append(is_violation * -20000000)

                # Çift Vardiya (Split Shift) Cezası
                for t_idx, t in enumerate(teachers_list):
                    for d in days:
                        is_morning = model.NewBoolVar(f'm_{t_idx}_{d}')
                        is_afternoon = model.NewBoolVar(f'a_{t_idx}_{d}')
                        model.AddMaxEquality(is_morning, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                        model.AddMaxEquality(is_afternoon, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                        double_shift = model.NewBoolVar(f'dbl_{t_idx}_{d}')
                        model.Add(is_morning + is_afternoon - 1 <= double_shift)
                        objective.append(double_shift * -50000000)

                # Hedefe Ulaşma Teşviki ve Esnek Yasaklı Gün Cezası
                for t_idx, t in enumerate(teachers_list):
                    total_assignments = sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions)
                    objective.append(total_assignments * 5000000)

                    if not strict_forbidden_days:
                        forbidden = str(t.get('Yasaklı Günler', ''))
                        for d_idx, d_name in enumerate(day_names):
                            if d_name in forbidden:
                                for c in range(len(classes_list)):
                                    for s in sessions:
                                        objective.append(x[(t_idx, c, d_idx, s)] * -500000000)

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
                    native_names = [t['Ad Soyad'] for t in teachers_list if 'NATIVE' in get_role(t)]

                    # Kontroller ve İhlal Dedektifi
                    for t_idx, t in enumerate(teachers_list):
                        if 'NATIVE' in get_role(t):
                            assigned = sum([solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions])
                            real_target = adjusted_targets[t_idx]
                            if assigned < real_target:
                                violations.append({"Hoca": t['Ad Soyad'], "Sorun": f"Boşta Kaldı ({real_target - assigned} Saat)", "Sınıf": "Kapasite/Şelale"})

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
                                        
                                        # İhlal Taramaları
                                        pref = get_pref(t)
                                        if 'SABAH' in pref and s_req == 1:
                                            violations.append({"Hoca": t['Ad Soyad'], "Sorun": "Ters Vardiya (Tercih: Sabah)", "Sınıf": c_name})
                                        elif 'OGLE' in pref and s_req == 0:
                                            violations.append({"Hoca": t['Ad Soyad'], "Sorun": "Ters Vardiya (Tercih: Öğle)", "Sınıf": c_name})

                                        m = sum([solver.Value(x[(t_idx, cc, d_idx, 0)]) for cc in range(len(classes_list))])
                                        a = sum([solver.Value(x[(t_idx, cc, d_idx, 1)]) for cc in range(len(classes_list))])
                                        if m > 0 and a > 0:
                                            violations.append({"Hoca": t['Ad Soyad'], "Sorun": f"Çift Vardiya ({d_name})", "Sınıf": c_name})
                                        if d_name in str(t.get('Yasaklı Günler', '')):
                                            violations.append({"Hoca": t['Ad Soyad'], "Sorun": f"Yasaklı Güne Atama ({d_name})", "Sınıf": c_name})
                                        break
                                
                                if val == "🔴 BOŞ":
                                    violations.append({"Hoca": "Boş Kalan Ders", "Sorun": f"Ders Boş Kaldı ({d_name})", "Sınıf": c_name})
                                    
                            row[d_name] = val
                        res_data.append(row)

                    stats = []
                    for t_idx, t in enumerate(teachers_list):
                        assigned = sum([solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions])
                        original_target = int(t['Hedef Ders Sayısı'])
                        real_target = adjusted_targets[t_idx]

                        stat = "Tamam"
                        if real_target < original_target: 
                            stat = f"Kırpıldı ({original_target-real_target} saat eksildi)"
                        elif assigned < real_target: 
                            stat = f"{real_target-assigned} Ders Boş Kaldı"

                        stats.append({"Hoca Adı": t['Ad Soyad'], "Hedef (İlk)": original_target, "Güncel Hedef": real_target, "Atanan": assigned, "Durum": stat})

                    df_res = pd.DataFrame(res_data)
                    df_stats = pd.DataFrame(stats)
                    df_violations = pd.DataFrame(violations).drop_duplicates() if violations else pd.DataFrame()

                    if not df_violations.empty:
                        st.warning(f"⚠️ İhtiyaç veya kısıt çatışmaları nedeniyle {len(df_violations)} noktada istisnai durum/ihlal oluştu:")
                        st.table(df_violations)
                    else:
                        st.success("✅ Kusursuz Çözüm! Hiçbir kural veya tercih ihlal edilmedi.")

                    st.dataframe(df_res)
                    st.dataframe(df_stats)

                    # Excel İndirme
                    output_res = io.BytesIO()
                    with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
                        df_res.to_excel(writer, index=False, sheet_name="Program")
                        df_stats.to_excel(writer, index=False, sheet_name="Istatistikler")
                        if not df_violations.empty: df_violations.to_excel(writer, index=False, sheet_name="Ihlal_Raporu")

                        wb = writer.book
                        ws_prog = writer.sheets['Program']
                        
                        base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
                        fmt_default = wb.add_format(base_fmt)
                        
                        fmt_a1 = wb.add_format(dict(base_fmt, bg_color='#FFD700', font_color='black', bold=True))
                        fmt_a2 = wb.add_format(dict(base_fmt, bg_color='#FFA500', font_color='white', bold=True))
                        fmt_b1 = wb.add_format(dict(base_fmt, bg_color='#800000', font_color='white', bold=True))
                        fmt_b2 = wb.add_format(dict(base_fmt, bg_color='#006400', font_color='white', bold=True))
                        fmt_pre = wb.add_format(dict(base_fmt, bg_color='#604878', font_color='white', bold=True))
                        fmt_blue = wb.add_format(dict(base_fmt, bg_color='#ADD8E6')) 

                        ws_prog.set_column('A:B', 12)
                        ws_prog.set_column('C:C', 20)
                        ws_prog.set_column('E:I', 14)
                        ws_prog.set_row(0, 20)

                        for r, row in df_res.iterrows():
                            excel_r = r + 1
                            ws_prog.set_row(excel_r, 20)
                            lvl = str(row['Seviye'])
                            
                            if lvl == "A1": c_fmt = fmt_a1
                            elif lvl == "A2": c_fmt = fmt_a2
                            elif lvl == "B1": c_fmt = fmt_b1
                            elif lvl == "B2": c_fmt = fmt_b2
                            elif lvl == "PreFaculty": c_fmt = fmt_pre
                            else: c_fmt = fmt_default

                            ws_prog.write(excel_r, 0, row['Sınıf'], c_fmt)
                            ws_prog.write(excel_r, 1, row['Seviye'], c_fmt)
                            ws_prog.write(excel_r, 2, row['Sınıf Danışmanı'], fmt_default)
                            ws_prog.write(excel_r, 3, row['Zaman'], fmt_default)

                            for c in range(4, 9):
                                val = row.iloc[c]
                                f = fmt_blue if val in native_names else fmt_default
                                ws_prog.write(excel_r, c, val, f)

                    st.download_button("📥 Excel İndir", output_res.getvalue(), "ders_programi_final.xlsx")
                
                elif status == cp_model.UNKNOWN:
                    st.error("⏳ **Zaman Aşımı (Timeout):** Sistem 120 saniyede ideal çözüme ulaşamadı. Kısıtları biraz gevşetebilirsiniz.")
                else:
                    st.error("❌ **Çözüm Bulunamadı (Infeasible):** Kısıtlar birbiriyle çelişiyor. 'Sınıf Başına Max Hoca' sayısını artırmayı veya 'Boş Ders Bırak' seçeneğini açmayı deneyebilirsiniz.")
