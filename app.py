import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import collections

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Hazırlık Ders Programı Koordinatör Kokpiti (V70)",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SESSION STATE BAŞLATMA ---
if 'solution_output' not in st.session_state:
    st.session_state['solution_output'] = None

# --- YARDIMCI FONKSİYONLAR ---
def get_role(t):
    return str(t.get('Rol', '')).upper().replace('İ', 'I').replace('i', 'I').replace('ı', 'I').strip()

def get_pref(t):
    return str(t.get('Tercih (Sabah/Öğle)', 'Farketmez')).upper().replace('Ö', 'O').replace('ö', 'O').replace('Ğ', 'G').replace('ğ', 'G').strip()

def get_partners(t):
    val = str(t.get('İstenmeyen Partner', '')).strip()
    if not val or val.lower() == 'nan':
        return []
    return [p.strip() for p in val.split(',') if p.strip()]

def generate_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_teachers = pd.DataFrame({
            'Ad Soyad': ['Ahmet Hoca', 'Sarah (Native)', 'Mehmet (Danışman)', 'Ayşe Hoca', 'Can Hoca', 'John (Native)'],
            'Rol': ['Destek', 'Native', 'Danışman', 'Ek Görevli', 'Destek', 'Native'],
            'Hedef Ders Sayısı': [4, 4, 3, 2, 4, 4],
            'Tercih (Sabah/Öğle)': ['Sabah', 'Farketmez', 'Sabah', 'Öğle', 'Farketmez', 'Öğle'],
            'Yasaklı Günler': ['Cuma', 'Çarşamba', '', 'Pazartesi,Salı', '', ''],
            'Sabit Sınıf': ['', '', 'A1.01', '', '', ''],
            'Yetkinlik (Seviyeler)': ['A1,A2,B1', 'Hepsi', 'A1,A2', 'B1,B2', 'Hepsi', 'B1,B2'],
            'İstenmeyen Partner': ['', '', 'Ayşe Hoca', 'Mehmet (Danışman)', '', '']
        })
        df_teachers.to_excel(writer, sheet_name='Ogretmenler', index=False)
    return output.getvalue()

# --- BAŞLIK VE YAN PANEL ---
st.title("🛡️ Hazırlık Ders Programı Yönetim Kokpiti (V70)")
st.caption("Google OR-Tools CP-SAT Motoru • Hot Module Destekli • Öğretmen ve Sınıf Çift Çizelgeli")

with st.sidebar:
    st.header("📁 Veri Girişi")
    uploaded_file = st.file_uploader("Öğretmen Excel Dosyası", type=["xlsx"])
    st.download_button("📥 Kılavuzlu Şablonu İndir", generate_template(), "ogretmen_listesi_sablon.xlsx", use_container_width=True)
    st.markdown("---")
    st.markdown("### 📌 Hızlı İpuçları")
    st.info("""
    - **Hot Module:** Sabahçı grupları tek tıkla belirler; kalanları otomatik öğleye atar.
    - **Öğretmen Çizelgesi:** Çözüm sonrası hem sınıf hem bireysel öğretmen takvimi oluşturulur.
    """)

# --- SEKME DÜZENİ (TABS) ---
tab_config, tab_precheck, tab_results = st.tabs([
    "⚙️ 1. Yapılandırma & Kurallar",
    "📊 2. Kapasite & Ön Kontrol",
    "📅 3. Program & Çözüm Ekranı"
])

# ==============================================================================
# SEKME 1: YAPILANDIRMA & KURALLAR
# ==============================================================================
with tab_config:
    st.subheader("🔥 Vardiya Stratejisi (Hot Module)")
    
    col_hm1, col_hm2 = st.columns([1, 2])
    with col_hm1:
        use_hot_module = st.toggle("Hot Module Modunu Kullan", value=True, 
                                   help="Açıkken seçilen seviyeler SABAH olur, kalan tüm seviyeler otomatik ÖĞLE grubuna geçer.")
    
    all_levels = ["A1", "A2", "B1", "B2", "PreFaculty"]
    level_shifts = {}

    if use_hot_module:
        with col_hm2:
            hot_morning_levels = st.multiselect(
                "☀️ Sabah Grubu Olacak Seviyeleri Seçin:",
                options=all_levels,
                default=["A1", "A2"],
                help="Seçilen seviyeler Sabah, seçilmeyenler ise Öğle vardiyasına ayarlanır."
            )
            for lvl in all_levels:
                level_shifts[lvl] = "Sabah" if lvl in hot_morning_levels else "Öğle"
        
        # Seçim Özeti Rozetleri
        morning_badge = ", ".join([l for l, s in level_shifts.items() if s == "Sabah"]) or "Yok"
        afternoon_badge = ", ".join([l for l, s in level_shifts.items() if s == "Öğle"]) or "Yok"
        st.success(f"**Vardiya Planı:** ☀️ **Sabah:** [{morning_badge}] | 🌙 **Öğle:** [{afternoon_badge}]")
    else:
        st.info("ℹ️ Hot Module kapalı. Seviyelerin vardiyalarını aşağıdan manüel olarak belirleyebilirsiniz.")
        cols_manual = st.columns(5)
        for i, lvl in enumerate(all_levels):
            with cols_manual[i]:
                default_val = "Sabah" if lvl in ["A1", "A2"] else "Öğle"
                level_shifts[lvl] = st.selectbox(f"{lvl} Zamanı", ["Sabah", "Öğle"], index=0 if default_val == "Sabah" else 1, key=f"shift_{lvl}")

    st.markdown("---")
    st.subheader("🏫 Şube Sayıları")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: count_a1 = st.number_input("A1 Şube Sayısı", 0, 20, 4)
    with c2: count_a2 = st.number_input("A2 Şube Sayısı", 0, 20, 4)
    with c3: count_b1 = st.number_input("B1 Şube Sayısı", 0, 20, 3)
    with c4: count_b2 = st.number_input("B2 Şube Sayısı", 0, 20, 2)
    with c5: count_pre = st.number_input("PreFaculty Şube", 0, 10, 0)

    st.markdown("---")
    st.subheader("⚙️ Çözücü Kuralları & Toleranslar")
    col_rule1, col_rule2 = st.columns(2)
    with col_rule1:
        max_teachers_per_class = st.slider("Sınıf Başına Max Farklı Öğretmen", 1, 6, 3, 
                                           help="Bir sınıfa hafta boyunca en fazla kaç farklı öğretmen girebilir?")
        allow_empty_slots = st.checkbox("Sıkışınca Boş Ders Bırak", value=True, 
                                        help="Kapatılırsa okulun her dersi için hoca bulunması zorunlu olur (Açık varsa çözüm bulunamaz).")
    with col_rule2:
        strict_forbidden_days = st.checkbox("Yasaklı Günleri Kesin Kural Yap (Hard Constraint)", value=False, 
                                           help="Açıkken hocaya yasaklı gününe ASLA ders yazılmaz. Kapalıyken ağır ceza puanıyla sistem esneyebilir.")
        allow_native_advisor = st.checkbox("Native Hocalar Danışman Olabilir mi?", value=False)

# Sınıfları Üreten Fonksiyon
def build_classes():
    class_list = []
    config = [
        (count_a1, "A1"), (count_a2, "A2"), (count_b1, "B1"),
        (count_b2, "B2"), (count_pre, "PreFaculty")
    ]
    for count, lvl in config:
        time_code = 0 if level_shifts[lvl] == "Sabah" else 1
        for i in range(1, count + 1):
            class_list.append({
                "Sınıf Adı": f"{lvl}.{i:02d}",
                "Seviye": lvl,
                "Zaman Kodu": time_code,
                "Zaman": "Sabah" if time_code == 0 else "Öğle"
            })
    return pd.DataFrame(class_list)

df_classes = build_classes()
classes_list = df_classes.to_dict('records')

# ==============================================================================
# SEKME 2: KAPASİTE & ÖN KONTROL
# ==============================================================================
with tab_precheck:
    st.subheader("📊 Kapasite & Talep Simülasyonu")
    
    if not uploaded_file:
        st.warning("⚠️ Lütfen sol menüden bir öğretmen Excel dosyası yükleyin.")
    else:
        df_teachers = pd.read_excel(uploaded_file, sheet_name='Ogretmenler').fillna("")
        if 'Hedef Ders Sayısı' not in df_teachers.columns and 'Hedef Gün Sayısı' in df_teachers.columns:
            df_teachers.rename(columns={'Hedef Gün Sayısı': 'Hedef Ders Sayısı'}, inplace=True)
            
        teachers_list = df_teachers.to_dict('records')

        # İhtiyaçlar
        morning_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 0])
        afternoon_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 1])
        total_needs = morning_needs + afternoon_needs

        # Kapasiteler
        morning_cap, afternoon_cap, flex_cap = 0, 0, 0
        base_targets = []
        for t in teachers_list:
            forb_cnt = len(str(t.get('Yasaklı Günler', '')).split(',')) if str(t.get('Yasaklı Günler', '')).strip() else 0
            t_cap = min(int(t.get('Hedef Ders Sayısı', 0)), 5 - forb_cnt)
            base_targets.append(t_cap)

            pref = get_pref(t)
            if 'SABAH' in pref: morning_cap += t_cap
            elif 'OGLE' in pref: afternoon_cap += t_cap
            else: flex_cap += t_cap

        total_teacher_cap = sum(base_targets)
        excess = total_teacher_cap - total_needs

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Toplam Ders İhtiyacı", f"{total_needs} Saat")
        m2.metric("Toplam Hoca Kapasitesi", f"{total_teacher_cap} Saat", delta=f"{excess} Saat Fark")
        m3.metric("Sabah Grubu (İhtiyaç/Kapasite)", f"{morning_needs} / {morning_cap} (+{flex_cap})")
        m4.metric("Öğle Grubu (İhtiyaç/Kapasite)", f"{afternoon_needs} / {afternoon_cap} (+{flex_cap})")

        st.markdown("---")
        st.subheader("🔍 Kural & Veri Bütünlüğü Doğrulaması")

        # Mantıksal Analiz
        errors, warnings = [], []
        fixed_classes = []
        teacher_names = [str(t.get('Ad Soyad', '')).strip() for t in teachers_list]

        for t in teachers_list:
            t_name = str(t.get('Ad Soyad', '')).strip()
            role = get_role(t)
            f_class = str(t.get('Sabit Sınıf', '')).strip()

            if f_class:
                if not allow_native_advisor and "NATIVE" in role:
                    errors.append(f"🛑 **{t_name}:** Native öğretmene sabit sınıf verilmiş (Genel Ayarlardan izin verilebilir).")
                if "EK GÖREVL" in role:
                    errors.append(f"🛑 **{t_name}:** Ek Görevli öğretmene sabit sınıf verilemez.")
                if not any(c['Sınıf Adı'] == f_class for c in classes_list):
                    errors.append(f"❌ **{t_name}:** Atandığı '{f_class}' sınıfı şube listesinde mevcut değil.")
                else:
                    fixed_classes.append(f_class)

            for p in get_partners(t):
                if p not in teacher_names:
                    warnings.append(f"⚠️ **{t_name}:** İstenmeyen partner '{p}' öğretmen listesinde bulunamadı.")

        dupe_fixed = [item for item, cnt in collections.Counter(fixed_classes).items() if cnt > 1]
        if dupe_fixed:
            errors.append(f"❌ **Çakışma:** {', '.join(dupe_fixed)} sınıfına birden fazla öğretmen sabitlenmiş!")

        if errors:
            for e in errors: st.error(e)
        else:
            st.success("✅ Veri doğrulaması başarılı. Kritik bir kural çelişkisi bulunamadı.")

        if warnings:
            for w in warnings: st.warning(w)

        # Kırpma Simülasyonu
        adjusted_targets = list(base_targets)
        if excess > 0:
            st.info(f"ℹ️ **Dengeleme Notu:** {excess} saat fazla hoca kapasitesi bulunmaktadır. Ek Görevli ve Destek hocalarından adil kırpma yapılacaktır.")
            ek_idx = [i for i, t in enumerate(teachers_list) if 'EK GÖREVL' in get_role(t)]
            destek_idx = [i for i, t in enumerate(teachers_list) if 'DESTEK' in get_role(t)]
            other_idx = [i for i, t in enumerate(teachers_list) if i not in ek_idx and i not in destek_idx and 'NATIVE' not in get_role(t)]

            exc = excess
            while exc > 0:
                trimmed = False
                for group in [ek_idx, destek_idx, other_idx]:
                    for i in group:
                        if exc == 0: break
                        min_req = 1 if str(teachers_list[i].get('Sabit Sınıf', '')).strip() else 0
                        if adjusted_targets[i] > min_req:
                            adjusted_targets[i] -= 1
                            exc -= 1
                            trimmed = True
                    if exc == 0 or trimmed: break
                if not trimmed: break
        elif excess < 0:
            if not allow_empty_slots:
                st.error("🛑 Kapasite yetersiz ve 'Boş Ders Bırak' ayarı KAPALI. Bu durumda çözüm üretilemez.")
            else:
                st.warning(f"⚠️ {abs(excess)} ders saati öğretmen yetersizliği nedeniyle boş kalacaktır.")

# ==============================================================================
# SEKME 3: PROGRAM & ÇÖZÜM EKRANI
# ==============================================================================
with tab_results:
    st.subheader("🚀 Optimizasyon ve Çizelge Çıktıları")

    if not uploaded_file:
        st.info("Program oluşturmak için lütfen sol menüden dosyanızı yükleyin.")
    else:
        btn_solve = st.button("⚡ Programı Optimize Et ve Oluştur", type="primary", use_container_width=True)

        if btn_solve:
            if errors:
                st.error("Lütfen Sekme 2'deki kritik hataları düzelttikten sonra tekrar deneyin.")
            else:
                with st.spinner("Google OR-Tools CP-SAT motoru çalışıyor (Tüm kısıtlar taranıyor)..."):
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
                            teacher_in_class[(t, c)] = model.NewBoolVar(f't_in_c_{t}_{c}')
                            for d in days:
                                for s in sessions:
                                    x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                    # Sınıfta Olma Durumu (teacher_in_class)
                    for t in range(len(teachers_list)):
                        for c in range(len(classes_list)):
                            model.AddMaxEquality(teacher_in_class[(t, c)], [x[(t, c, d, s)] for d in days for s in sessions])

                    # --- KESİN KURALLAR (HARD) ---
                    # 1. Sınıf Başına Max Hoca
                    for c in range(len(classes_list)):
                        model.Add(sum(teacher_in_class[(t, c)] for t in range(len(teachers_list))) <= max_teachers_per_class)

                    # 2. İstenmeyen Partner
                    name_to_idx = {str(t['Ad Soyad']).strip(): i for i, t in enumerate(teachers_list)}
                    partner_pairs = set()
                    for t_idx, t in enumerate(teachers_list):
                        for p in get_partners(t):
                            if p in name_to_idx:
                                p_idx = name_to_idx[p]
                                pair = tuple(sorted([t_idx, p_idx]))
                                if pair not in partner_pairs and t_idx != p_idx:
                                    partner_pairs.add(pair)
                                    for c in range(len(classes_list)):
                                        model.Add(teacher_in_class[(t_idx, c)] + teacher_in_class[(p_idx, c)] <= 1)

                    # 3. Hoca Çakışması
                    for t in range(len(teachers_list)):
                        for d in days:
                            for s in sessions:
                                model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

                    # 4. Sınıf Oturumu & Boş Ders Kuralı
                    for c_idx, c_data in enumerate(classes_list):
                        req_s = c_data['Zaman Kodu']
                        other_s = 1 - req_s
                        for d in days:
                            if c_data['Seviye'] == "PreFaculty" and d >= 3:
                                model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) == 0)
                            else:
                                if allow_empty_slots:
                                    model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) <= 1)
                                else:
                                    model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) == 1)
                            model.Add(sum(x[(t, c_idx, d, other_s)] for t in range(len(teachers_list))) == 0)

                    # 5. Seviye Yetkinliği
                    for t_idx, t in enumerate(teachers_list):
                        allowed = str(t.get('Yetkinlik (Seviyeler)', '')).strip()
                        if allowed == "": allowed = "Hepsi"
                        if "Hepsi" not in allowed:
                            for c_idx, c in enumerate(classes_list):
                                if c['Seviye'] not in allowed:
                                    for d in days:
                                        for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                                    model.Add(advisor_var[(t_idx, c_idx)] == 0)

                    # 6. Danışmanlık Kuralları
                    for c in range(len(classes_list)):
                        model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) <= 1)
                    for t in range(len(teachers_list)):
                        model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

                    for t_idx, t in enumerate(teachers_list):
                        fixed_c = str(t.get('Sabit Sınıf', '')).strip()
                        if fixed_c:
                            f_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == fixed_c), None)
                            if f_idx is not None: model.Add(advisor_var[(t_idx, f_idx)] == 1)

                        role = get_role(t)
                        if 'EK GÖREVL' in role:
                            for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                        if not allow_native_advisor and 'NATIVE' in role:
                            for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
                        if 'NATIVE' in role:
                            for c_idx, c_data in enumerate(classes_list):
                                if c_data['Seviye'] == 'A1':
                                    for d in days:
                                        for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                        if 'EK GÖREVL' in role:
                            for c_idx in range(len(classes_list)):
                                model.Add(sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions) <= 1)

                    # 7. Tavan Saat Sınırı
                    for t_idx, t in enumerate(teachers_list):
                        model.Add(sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions) <= adjusted_targets[t_idx])

                    # 8. Yasaklı Gün Kesin Kuralı
                    if strict_forbidden_days:
                        for t_idx, t in enumerate(teachers_list):
                            forb = str(t.get('Yasaklı Günler', ''))
                            for d_idx, d_name in enumerate(day_names):
                                if d_name in forb:
                                    for c in range(len(classes_list)):
                                        for s in sessions: model.Add(x[(t_idx, c, d_idx, s)] == 0)

                    # --- AMAC FONKSİYONU VE YUMUŞAK KURALLAR ---
                    objective = []
                    # Taban doldurma ödülü
                    objective.append(sum(x.values()) * 1000000)

                    # Native Dağılımı ve Şelalesi
                    for c_idx, c_data in enumerate(classes_list):
                        natives_in_c = []
                        for t_idx, t in enumerate(teachers_list):
                            if 'NATIVE' in get_role(t):
                                is_p = teacher_in_class[(t_idx, c_idx)]
                                natives_in_c.append(is_p)
                                lvl = c_data['Seviye']
                                score = 100000 if lvl == "B2" else (50000 if lvl == "B1" else (10000 if lvl == "A2" else 1000))
                                objective.append(is_p * score)
                        if natives_in_c:
                            model.Add(sum(natives_in_c) <= 1)

                    # Danışmanlık ve Pazartesi
                    for t_idx, t_data in enumerate(teachers_list):
                        forb_days = str(t_data.get('Yasaklı Günler', ''))
                        for c_idx, c_data in enumerate(classes_list):
                            is_adv = advisor_var[(t_idx, c_idx)]
                            req_s = c_data['Zaman Kodu']
                            if "Pazartesi" not in forb_days:
                                pzt_v = x[(t_idx, c_idx, 0, req_s)]
                                adv_pzt = model.NewBoolVar(f'ap_{t_idx}_{c_idx}')
                                model.AddImplication(adv_pzt, pzt_v)
                                model.AddImplication(adv_pzt, is_adv)
                                objective.append(adv_pzt * 500000)

                            if c_data['Seviye'] != "PreFaculty":
                                cnt_days = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                                is_2p = model.NewBoolVar(f'i2_{t_idx}_{c_idx}')
                                model.Add(cnt_days >= 2).OnlyEnforceIf(is_2p)
                                model.Add(cnt_days <= 1).OnlyEnforceIf(is_2p.Not())
                                adv_2d = model.NewBoolVar(f'a2_{t_idx}_{c_idx}')
                                model.AddImplication(adv_2d, is_2p)
                                model.AddImplication(adv_2d, is_adv)
                                objective.append(adv_2d * 200000)

                    # Vardiya Tercih Cezaları
                    for t_idx, t in enumerate(teachers_list):
                        pref = get_pref(t)
                        if 'SABAH' in pref:
                            for c in range(len(classes_list)):
                                for d in days: objective.append(x[(t_idx, c, d, 1)] * -100000)
                        elif 'OGLE' in pref:
                            for c in range(len(classes_list)):
                                for d in days: objective.append(x[(t_idx, c, d, 0)] * -100000)

                    # Çift Vardiya (Split Shift) Cezası
                    for t_idx, t in enumerate(teachers_list):
                        for d in days:
                            is_m = model.NewBoolVar(f'm_{t_idx}_{d}')
                            is_a = model.NewBoolVar(f'a_{t_idx}_{d}')
                            model.AddMaxEquality(is_m, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                            model.AddMaxEquality(is_a, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                            dbl = model.NewBoolVar(f'dbl_{t_idx}_{d}')
                            model.Add(is_m + is_a - 1 <= dbl)
                            objective.append(dbl * -500000)

                    # Yasaklı Gün Esnek Cezası
                    if not strict_forbidden_days:
                        for t_idx, t in enumerate(teachers_list):
                            forb = str(t.get('Yasaklı Günler', ''))
                            for d_idx, d_name in enumerate(day_names):
                                if d_name in forb:
                                    for c in range(len(classes_list)):
                                        for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -5000000)

                    # Çözüm Başlat
                    model.Maximize(sum(objective))
                    solver = cp_model.CpSolver()
                    solver.parameters.max_time_in_seconds = 90.0
                    solver.parameters.num_search_workers = 8

                    status = solver.Solve(model)

                    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                        # Sonuçları Çıkar
                        class_schedule = []
                        teacher_schedule = {t['Ad Soyad']: {d: "-" for d in day_names} for t in teachers_list}
                        violations = []
                        native_names = [t['Ad Soyad'] for t in teachers_list if 'NATIVE' in get_role(t)]

                        for c_idx, c in enumerate(classes_list):
                            c_name = c['Sınıf Adı']
                            s_req = c['Zaman Kodu']
                            adv_name = "Atanamadı"
                            for t_idx in range(len(teachers_list)):
                                if solver.Value(advisor_var[(t_idx, c_idx)]) == 1:
                                    adv_name = teachers_list[t_idx]['Ad Soyad']
                                    break

                            row = {
                                "Sınıf": c_name, "Seviye": c['Seviye'], 
                                "Danışman": adv_name, "Vardiya": "Sabah" if s_req == 0 else "Öğle"
                            }

                            for d_idx, d_name in enumerate(day_names):
                                val = "🔴 BOŞ"
                                if c['Seviye'] == "PreFaculty" and d_idx >= 3:
                                    val = "⛔ KAPALI"
                                else:
                                    for t_idx, t in enumerate(teachers_list):
                                        if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                            t_name = t['Ad Soyad']
                                            val = t_name
                                            teacher_schedule[t_name][d_name] = f"{c_name} ({row['Vardiya']})"

                                            # İhlal dedektifleri
                                            pref = get_pref(t)
                                            if 'SABAH' in pref and s_req == 1:
                                                violations.append({"Hoca": t_name, "Tür": "Ters Vardiya (Tercih: Sabah)", "Detay": c_name})
                                            elif 'OGLE' in pref and s_req == 0:
                                                violations.append({"Hoca": t_name, "Tür": "Ters Vardiya (Tercih: Öğle)", "Detay": c_name})
                                            if d_name in str(t.get('Yasaklı Günler', '')):
                                                violations.append({"Hoca": t_name, "Tür": f"Yasaklı Günde Ders ({d_name})", "Detay": c_name})
                                            break

                                if val == "🔴 BOŞ":
                                    violations.append({"Hoca": "-", "Tür": f"Boş Ders ({d_name})", "Detay": c_name})

                                row[d_name] = val
                            class_schedule.append(row)

                        # Öğretmen İstatistikleri
                        stats_list = []
                        for t_idx, t in enumerate(teachers_list):
                            t_name = t['Ad Soyad']
                            assigned = sum(solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions)
                            stats_list.append({
                                "Ad Soyad": t_name,
                                "Rol": t.get('Rol', ''),
                                "İlk Hedef": int(t.get('Hedef Ders Sayısı', 0)),
                                "Düzeltilmiş Hedef": adjusted_targets[t_idx],
                                "Atanan Ders": assigned,
                                "Durum": "Kusursuz" if assigned == adjusted_targets[t_idx] else f"{adjusted_targets[t_idx] - assigned} Ders Eksik"
                            })

                        # Öğretmen Programı DataFrame
                        teacher_rows = []
                        for s_row in stats_list:
                            t_name = s_row['Ad Soyad']
                            t_dict = dict(s_row)
                            t_dict.update(teacher_schedule[t_name])
                            teacher_rows.append(t_dict)

                        st.session_state['solution_output'] = {
                            "df_classes": pd.DataFrame(class_schedule),
                            "df_teachers": pd.DataFrame(teacher_rows),
                            "df_stats": pd.DataFrame(stats_list),
                            "df_violations": pd.DataFrame(violations).drop_duplicates() if violations else pd.DataFrame(),
                            "native_names": native_names
                        }
                    else:
                        st.error("❌ Model çözülemedi (Infeasible). 'Sınıf Başına Max Hoca' sayısını artırın veya 'Boş Ders Bırak'ı işaretleyin.")

    # --- ÇÖZÜM GÖRSELLEŞTİRME VE İNDİRME ---
    if st.session_state['solution_output'] is not None:
        st.balloons()
        sol = st.session_state['solution_output']

        st.success("✅ Program başarıyla optimize edildi!")

        # İhlal Bildirimleri (Expander)
        if not sol['df_violations'].empty:
            with st.expander(f"⚠️ Dikkat: {len(sol['df_violations'])} Noktada Esneme / Tercih Dışı Durum Oluştu", expanded=True):
                st.dataframe(sol['df_violations'], use_container_width=True)

        # Görünüm Değiştirici
        view_mode = st.radio(
            "📋 Çizelge Görünüm Modu:",
            ["🏫 Sınıf Bazlı Program", "👨‍🏫 Öğretmen Bazlı Program", "📊 Öğretmen İstatistikleri"],
            horizontal=True
        )

        if view_mode == "🏫 Sınıf Bazlı Program":
            st.dataframe(sol['df_classes'], use_container_width=True)
        elif view_mode == "👨‍🏫 Öğretmen Bazlı Program":
            # Öğretmen filtreleme kutusu
            selected_teacher = st.selectbox("Belirli Bir Öğretmeni İzole Et:", ["Tümü"] + list(sol['df_teachers']['Ad Soyad']))
            if selected_teacher == "Tümü":
                st.dataframe(sol['df_teachers'], use_container_width=True)
            else:
                st.dataframe(sol['df_teachers'][sol['df_teachers']['Ad Soyad'] == selected_teacher], use_container_width=True)
        else:
            st.dataframe(sol['df_stats'], use_container_width=True)

        # --- ÇOK SEKMELİ EXCEL ÇIKTISI ---
        output_res = io.BytesIO()
        with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
            sol['df_classes'].to_excel(writer, index=False, sheet_name="Sinif_Programi")
            sol['df_teachers'].to_excel(writer, index=False, sheet_name="Ogretmen_Programi")
            sol['df_stats'].to_excel(writer, index=False, sheet_name="Istatistikler")
            if not sol['df_violations'].empty:
                sol['df_violations'].to_excel(writer, index=False, sheet_name="Ihlal_Raporu")

            wb = writer.book
            ws_c = writer.sheets['Sinif_Programi']
            ws_t = writer.sheets['Ogretmen_Programi']

            base_fmt = {'border': 1, 'align': 'center', 'valign': 'vcenter'}
            fmt_def = wb.add_format(base_fmt)
            fmt_a1 = wb.add_format(dict(base_fmt, bg_color='#FFD700', font_color='black', bold=True))
            fmt_a2 = wb.add_format(dict(base_fmt, bg_color='#FFA500', font_color='white', bold=True))
            fmt_b1 = wb.add_format(dict(base_fmt, bg_color='#800000', font_color='white', bold=True))
            fmt_b2 = wb.add_format(dict(base_fmt, bg_color='#006400', font_color='white', bold=True))
            fmt_pre = wb.add_format(dict(base_fmt, bg_color='#604878', font_color='white', bold=True))
            fmt_blue = wb.add_format(dict(base_fmt, bg_color='#ADD8E6'))

            ws_c.set_column('A:B', 12)
            ws_c.set_column('C:C', 20)
            ws_c.set_column('E:I', 15)

            for r, row in sol['df_classes'].iterrows():
                e_r = r + 1
                lvl = str(row['Seviye'])
                c_fmt = fmt_a1 if lvl == "A1" else (fmt_a2 if lvl == "A2" else (fmt_b1 if lvl == "B1" else (fmt_b2 if lvl == "B2" else fmt_pre)))
                ws_c.write(e_r, 0, row['Sınıf'], c_fmt)
                ws_c.write(e_r, 1, row['Seviye'], c_fmt)
                ws_c.write(e_r, 2, row['Danışman'], fmt_def)
                ws_c.write(e_r, 3, row['Vardiya'], fmt_def)
                for c_i in range(4, 9):
                    val = row.iloc[c_i]
                    ws_c.write(e_r, c_i, val, fmt_blue if val in sol['native_names'] else fmt_def)

            ws_t.set_column('A:B', 18)
            ws_t.set_column('G:K', 16)

        st.download_button(
            "📥 Çok Sekmeli Kurumsal Excel'i İndir (.xlsx)",
            output_res.getvalue(),
            "hazirlik_ders_programi_v70.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
