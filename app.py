import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import collections

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Hazırlık Ders Programı Koordinatör Kokpiti (V76)",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SESSION STATE BAŞLATMA ---
if 'solution_output' not in st.session_state:
    st.session_state['solution_output'] = None
if 'pinned_assignments' not in st.session_state:
    st.session_state['pinned_assignments'] = []
if 'saved_scenarios' not in st.session_state:
    st.session_state['saved_scenarios'] = {}

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
st.title("🛡️ Hazırlık Ders Programı Koordinatör Kokpiti (V76)")
st.caption("Google OR-Tools CP-SAT Motoru • Asenkron Ders Modülü • Kriz Dedektifli ve Manuel Pinlemeli")

with st.sidebar:
    st.header("📁 Veri Girişi")
    uploaded_file = st.file_uploader("Öğretmen Excel Dosyası", type=["xlsx"])
    st.download_button("📥 Kılavuzlu Şablonu İndir", generate_template(), "ogretmen_listesi_sablon.xlsx", use_container_width=True)
    st.markdown("---")
    st.markdown("### 📌 Aktif Pinlemeler")
    if st.session_state['pinned_assignments']:
        st.write(f"Toplam **{len(st.session_state['pinned_assignments'])}** hücre kilitli.")
        if st.button("🗑️ Tüm Pinleri Temizle", use_container_width=True):
            st.session_state['pinned_assignments'] = []
            st.rerun()
    else:
        st.caption("Henüz kilitlenmiş hücre yok.")

# --- SEKME DÜZENİ (TABS) ---
tab_config, tab_precheck, tab_results, tab_whatif = st.tabs([
    "⚙️ 1. Yapılandırma & Kurallar",
    "📊 2. Kapasite & Ön Kontrol",
    "📅 3. Program & Çözüm Ekranı",
    "⚖️ 4. Senaryo Kıyaslama (What-If)"
])

all_levels = ["A1", "A2", "B1", "B2", "PreFaculty"]
level_shifts = {}

# ==============================================================================
# SEKME 1: YAPILANDIRMA & KURALLAR
# ==============================================================================
with tab_config:
    st.subheader("🔥 Vardiya Stratejisi (Hot Module)")
    col_hm1, col_hm2 = st.columns([1, 2])
    with col_hm1:
        use_hot_module = st.toggle("Hot Module Modunu Kullan", value=True, 
                                   help="Açıkken seçilen seviyeler SABAH olur, kalan tüm seviyeler otomatik ÖĞLE grubuna geçer.")
    
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
        
        morning_badge = ", ".join([l for l, s in level_shifts.items() if s == "Sabah"]) or "Yok"
        afternoon_badge = ", ".join([l for l, s in level_shifts.items() if s == "Öğle"]) or "Yok"
        st.success(f"**Vardiya Planı:** ☀️ **Sabah:** [{morning_badge}] | 🌙 **Öğle:** [{afternoon_badge}]")
    else:
        st.info("ℹ️ Hot Module kapalı. Seviyelerin vardiyalarını aşağıdan manüel belirleyebilirsiniz.")
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
    st.subheader("⚙️ Çözücü Kuralları & Asenkron Ders Modülü")
    col_rule1, col_rule2 = st.columns(2)
    with col_rule1:
        max_teachers_per_class = st.slider("Sınıf Başına Max Farklı Öğretmen (Canlı Ders)", 1, 6, 3, 
                                           help="Bir sınıfa hafta boyunca en fazla kaç farklı canlı ders öğretmeni girebilir?")
        allow_empty_slots = st.checkbox("Sıkışınca Boş Ders Bırak", value=True, 
                                        help="Kapatılırsa okulun her dersi için hoca bulunması zorunlu olur.")
        enable_asynch = st.checkbox(
            "💻 Asenkron Ders Slotu Aç (Yalnızca B1 ve B2)",
            value=True,
            help="İşaretlendiğinde her B1 ve B2 şubesine 1 adet Asenkron ders hoca slotu açılır. Bu slot, atanan hocanın haftalık kotasından 1 ders gününe karşılık gelir."
        )
    with col_rule2:
        strict_forbidden_days = st.checkbox("Yasaklı Günleri Kesin Kural Yap (Hard)", value=False, 
                                           help="Açıkken hocaya yasaklı gününe ASLA ders yazılmaz.")
        allow_native_advisor = st.checkbox("Native Hocalar Danışman Olabilir mi?", value=False)

    # 1A: MANUEL HÜCRE PİNLEME (KİLİTLEME) ARAYÜZÜ
    st.markdown("---")
    with st.expander("📌 Manuel Hücre Kilitleme (Pinning Yöneticisi)", expanded=False):
        st.info("Belirli bir hocanın belirli bir sınıfa, güne veya Asenkron slota kesin olarak atanmasını sabitleyebilirsiniz.")
        
        temp_classes = []
        cfg = [(count_a1, "A1"), (count_a2, "A2"), (count_b1, "B1"), (count_b2, "B2"), (count_pre, "PreFaculty")]
        for cnt, lvl in cfg:
            for idx in range(1, cnt + 1):
                temp_classes.append(f"{lvl}.{idx:02d}")
        
        if uploaded_file and temp_classes:
            df_t_preview = pd.read_excel(uploaded_file, sheet_name='Ogretmenler').fillna("")
            t_names = list(df_t_preview['Ad Soyad'].unique())
            
            p_col1, p_col2, p_col3, p_col4, p_col5 = st.columns([2, 1.5, 1.5, 1.5, 1])
            with p_col1: pin_teacher = st.selectbox("Öğretmen:", t_names, key="pin_t")
            with p_col2: pin_class = st.selectbox("Sınıf:", temp_classes, key="pin_c")
            with p_col3: pin_day = st.selectbox("Gün (Canlı için):", ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "-"], key="pin_d")
            with p_col4: 
                sess_opts = ["Sabah", "Öğle", "Asenkron"] if enable_asynch else ["Sabah", "Öğle"]
                pin_session = st.selectbox("Vardiya / Tür:", sess_opts, key="pin_s")
            with p_col5:
                st.write("")
                st.write("")
                if st.button("📌 Kilitle"):
                    new_pin = {"teacher": pin_teacher, "class": pin_class, "day": pin_day, "session": pin_session}
                    if new_pin not in st.session_state['pinned_assignments']:
                        st.session_state['pinned_assignments'].append(new_pin)
                        st.success("Kilitlendi!")
                        st.rerun()

            if st.session_state['pinned_assignments']:
                st.write("##### Mevcut Kilitli Atamalar:")
                st.dataframe(pd.DataFrame(st.session_state['pinned_assignments']), use_container_width=True)
        else:
            st.caption("Önce sol menüden Excel yükleyin ve şube sayılarını belirleyin.")

# Sınıfları Oluşturma
def build_classes():
    class_list = []
    config = [
        (count_a1, "A1"), (count_a2, "A2"), (count_b1, "B1"),
        (count_b2, "B2"), (count_pre, "PreFaculty")
    ]
    for count, lvl in config:
        time_code = 0 if level_shifts.get(lvl, "Sabah") == "Sabah" else 1
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

        morning_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 0])
        afternoon_needs = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list if c['Zaman Kodu'] == 1])
        asynch_needs = (count_b1 + count_b2) if enable_asynch else 0
        total_needs = morning_needs + afternoon_needs + asynch_needs

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
        m1.metric("Toplam Ders İhtiyacı", f"{total_needs} Saat", help=f"Canlı Dersler: {morning_needs + afternoon_needs}, Asenkron: {asynch_needs}")
        m2.metric("Toplam Hoca Kapasitesi", f"{total_teacher_cap} Saat", delta=f"{excess} Saat Fark")
        m3.metric("Sabah Grubu (Canlı)", f"{morning_needs} / {morning_cap} (+{flex_cap})")
        m4.metric("Öğle Grubu (Canlı)", f"{afternoon_needs} / {afternoon_cap} (+{flex_cap})")

        if enable_asynch:
            st.info(f"💻 **Asenkron Ders Durumu:** B1 ve B2 seviyesindeki {count_b1 + count_b2} şube için toplam **{asynch_needs} ders günü** asenkron kontenjanı hesaplamaya dahil edildi.")

        # Kırpma Simülasyonu
        adjusted_targets = list(base_targets)
        if excess > 0:
            st.info(f"ℹ️ {excess} saat fazla hoca kapasitesi var. Ek Görevli ve Destek hocalarından adil kırpma yapılacaktır.")
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

# ==============================================================================
# 3A: AKILLI KRİZ DEDEKTİFİ
# ==============================================================================
def run_diagnostic_engine(teachers_list, classes_list, adjusted_targets, max_teachers_per_class, strict_forbidden_days, enable_asynch):
    diag_model = cp_model.CpModel()
    days = range(5)
    sessions = range(2)

    x = {}
    teacher_in_class = {}
    for t in range(len(teachers_list)):
        for c in range(len(classes_list)):
            teacher_in_class[(t, c)] = diag_model.NewBoolVar(f'd_tinclass_{t}_{c}')
            for d in days:
                for s in sessions:
                    x[(t, c, d, s)] = diag_model.NewBoolVar(f'd_x_{t}_{c}_{d}_{s}')

    for t in range(len(teachers_list)):
        for c in range(len(classes_list)):
            diag_model.AddMaxEquality(teacher_in_class[(t, c)], [x[(t, c, d, s)] for d in days for s in sessions])

    slacks = []

    # 1. Boş Ders Gevşetmesi
    for c_idx, c_data in enumerate(classes_list):
        req_s = c_data['Zaman Kodu']
        for d in days:
            if c_data['Seviye'] == "PreFaculty" and d >= 3:
                continue
            slack_slot = diag_model.NewBoolVar(f'slack_slot_{c_idx}_{d}')
            diag_model.Add(sum(x[(t, c_idx, d, req_s)] for t in range(len(teachers_list))) + slack_slot >= 1)
            slacks.append(slack_slot * 1000)

    # 2. Sınıf Başına Max Hoca Gevşetmesi
    slack_max_t = {}
    for c in range(len(classes_list)):
        slack_max_t[c] = diag_model.NewIntVar(0, 10, f'slack_maxt_{c}')
        diag_model.Add(sum(teacher_in_class[(t, c)] for t in range(len(teachers_list))) <= max_teachers_per_class + slack_max_t[c])
        slacks.append(slack_max_t[c] * 500)

    # 3. İstenmeyen Partner
    name_to_idx = {str(t['Ad Soyad']).strip(): i for i, t in enumerate(teachers_list)}
    slack_partners = {}
    processed_pairs = set()
    for t_idx, t in enumerate(teachers_list):
        for p in get_partners(t):
            if p in name_to_idx:
                p_idx = name_to_idx[p]
                pair = tuple(sorted([t_idx, p_idx]))
                if pair not in processed_pairs and t_idx != p_idx:
                    processed_pairs.add(pair)
                    for c in range(len(classes_list)):
                        sp = diag_model.NewBoolVar(f'sp_{t_idx}_{p_idx}_{c}')
                        diag_model.Add(teacher_in_class[(t_idx, c)] + teacher_in_class[(p_idx, c)] <= 1 + sp)
                        slack_partners[(t_idx, p_idx, c)] = sp
                        slacks.append(sp * 300)

    # 4. Hoca Çakışması
    for t in range(len(teachers_list)):
        for d in days:
            for s in sessions:
                diag_model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

    # 5. Seviye Yetkinliği
    slack_comp = {}
    for t_idx, t in enumerate(teachers_list):
        allowed = str(t.get('Yetkinlik (Seviyeler)', '')).strip()
        if allowed != "" and "Hepsi" not in allowed:
            for c_idx, c in enumerate(classes_list):
                if c['Seviye'] not in allowed:
                    s_c = diag_model.NewBoolVar(f's_comp_{t_idx}_{c_idx}')
                    slack_comp[(t_idx, c_idx)] = s_c
                    for d in days:
                        for s in sessions:
                            diag_model.Add(x[(t_idx, c_idx, d, s)] <= s_c)
                    slacks.append(s_c * 400)

    # 6. Asenkron Gevşetmesi (Varsa)
    slack_asynch = {}
    if enable_asynch:
        for c_idx, c in enumerate(classes_list):
            if c['Seviye'] in ['B1', 'B2']:
                s_as = diag_model.NewBoolVar(f's_asynch_{c_idx}')
                slack_asynch[c_idx] = s_as
                slacks.append(s_as * 800)

    for t_idx, t in enumerate(teachers_list):
        diag_model.Add(sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions) <= adjusted_targets[t_idx])

    diag_model.Minimize(sum(slacks))
    diag_solver = cp_model.CpSolver()
    diag_solver.parameters.max_time_in_seconds = 15.0
    status = diag_solver.Solve(diag_model)

    diagnostics = []
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for c in range(len(classes_list)):
            extra_hoca = diag_solver.Value(slack_max_t[c])
            if extra_hoca > 0:
                diagnostics.append(f"🏫 **Sınıf Başına Max Hoca Kısıtı:** '{classes_list[c]['Sınıf Adı']}' sınıfı için {max_teachers_per_class} hoca yetersiz kalıyor (En az {max_teachers_per_class + extra_hoca} hoca girmeli).")

        for (t1, t2, c), sp in slack_partners.items():
            if diag_solver.Value(sp) == 1:
                diagnostics.append(f"👥 **İstenmeyen Partner Çatışması:** '{teachers_list[t1]['Ad Soyad']}' ile '{teachers_list[t2]['Ad Soyad']}', '{classes_list[c]['Sınıf Adı']}' sınıfına mecburen birlikte yazılmak zorunda kalıyor.")

        for (t_idx, c_idx), sc in slack_comp.items():
            if diag_solver.Value(sc) == 1:
                diagnostics.append(f"🎓 **Yetkinlik Yetersizliği:** '{teachers_list[t_idx]['Ad Soyad']}', yetkinliği olmadığı halde '{classes_list[c_idx]['Sınıf Adı']}' sınıfına atanmak zorunda kalıyor.")

    if not diagnostics:
        diagnostics.append("Genel kapasite yetersizliği, Asenkron kota darlığı veya aşırı kısıtlayıcı Yasaklı Gün / Pinleme çakışması tespit edildi.")

    return diagnostics

# ==============================================================================
# SEKME 3: PROGRAM & ÇÖZÜM EKRANI
# ==============================================================================
with tab_results:
    st.subheader("🚀 Optimizasyon ve Çözüm Motoru")

    if not uploaded_file:
        st.info("Program oluşturmak için lütfen sol menüden dosyanızı yükleyin.")
    else:
        col_prof1, col_prof2 = st.columns([2, 1])
        with col_prof1:
            variant_profile = st.selectbox(
                "🎯 Optimizasyon Profili (Çözüm Stratejisi):",
                [
                    "Plan A: Dengeli / Kurumsal (Önerilen)",
                    "Plan B: Öğretmen Odaklı (Maksimum Hoca Memnuniyeti)",
                    "Plan C: Pedagoji & Danışman Odaklı (Sınıf-Öğrenci Bağı)"
                ],
                help="Her profil kural ağırlıklarını farklı önceliklendirir."
            )
        with col_prof2:
            st.write("")
            st.write("")
            btn_solve = st.button("⚡ Programı Optimize Et ve Oluştur", type="primary", use_container_width=True)

        if btn_solve:
            with st.spinner(f"'{variant_profile}' stratejisiyle OR-Tools CP-SAT motoru çalışıyor..."):
                model = cp_model.CpModel()
                days = range(5)
                day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
                sessions = range(2)

                x = {}
                advisor_var = {}
                teacher_in_class = {}
                asynch_var = {}

                # Canlı Ders Değişkenleri
                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        advisor_var[(t, c)] = model.NewBoolVar(f'adv_{t}_{c}')
                        teacher_in_class[(t, c)] = model.NewBoolVar(f't_in_c_{t}_{c}')
                        for d in days:
                            for s in sessions:
                                x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

                for t in range(len(teachers_list)):
                    for c in range(len(classes_list)):
                        model.AddMaxEquality(teacher_in_class[(t, c)], [x[(t, c, d, s)] for d in days for s in sessions])

                # Asenkron Ders Değişkenleri (Yalnızca B1 ve B2)
                b1_b2_indices = [i for i, c in enumerate(classes_list) if c['Seviye'] in ['B1', 'B2']]
                if enable_asynch:
                    for t in range(len(teachers_list)):
                        for c_idx in b1_b2_indices:
                            asynch_var[(t, c_idx)] = model.NewBoolVar(f'asynch_{t}_{c_idx}')

                # --- 1A: PINLENMİŞ ATAMALARIN MODELDE ZORLANMASI ---
                name_to_idx = {str(t['Ad Soyad']).strip(): i for i, t in enumerate(teachers_list)}
                for pin in st.session_state['pinned_assignments']:
                    p_t = name_to_idx.get(pin['teacher'])
                    p_c = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == pin['class']), None)
                    if p_t is not None and p_c is not None:
                        if pin['session'] == 'Asenkron' and enable_asynch:
                            if (p_t, p_c) in asynch_var:
                                model.Add(asynch_var[(p_t, p_c)] == 1)
                        else:
                            p_d = day_names.index(pin['day']) if pin['day'] in day_names else None
                            p_s = 0 if pin['session'] == 'Sabah' else 1
                            if p_d is not None:
                                model.Add(x[(p_t, p_c, p_d, p_s)] == 1)

                # --- HARD CONSTRAINTS ---
                # Sınıf Başına Max Hoca (Canlı dersler için)
                for c in range(len(classes_list)):
                    model.Add(sum(teacher_in_class[(t, c)] for t in range(len(teachers_list))) <= max_teachers_per_class)

                # İstenmeyen Partner
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

                # Hoca Çakışması (Aynı anda 1 sınıfta)
                for t in range(len(teachers_list)):
                    for d in days:
                        for s in sessions:
                            model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

                # Sınıf Oturumu & Boş Ders Kuralı
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

                # Seviye Yetkinliği
                for t_idx, t in enumerate(teachers_list):
                    allowed = str(t.get('Yetkinlik (Seviyeler)', '')).strip()
                    if allowed != "" and "Hepsi" not in allowed:
                        for c_idx, c in enumerate(classes_list):
                            if c['Seviye'] not in allowed:
                                for d in days:
                                    for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                                model.Add(advisor_var[(t_idx, c_idx)] == 0)
                                if enable_asynch and (t_idx, c_idx) in asynch_var:
                                    model.Add(asynch_var[(t_idx, c_idx)] == 0)

                # Danışmanlık ve Roller
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

                # --- ASENKRON DERS KISITLARI (B1 ve B2 İÇİN) ---
                if enable_asynch:
                    for c_idx in b1_b2_indices:
                        if allow_empty_slots:
                            model.Add(sum(asynch_var[(t, c_idx)] for t in range(len(teachers_list))) <= 1)
                        else:
                            model.Add(sum(asynch_var[(t, c_idx)] for t in range(len(teachers_list))) == 1)

                    # Tavan Saat: Canlı Dersler + Asenkron Ders == adjusted_target
                    for t_idx, t in enumerate(teachers_list):
                        live_total = sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions)
                        asynch_total = sum(asynch_var[(t_idx, c)] for c in b1_b2_indices)
                        # Asenkron ders 1 ders gününe tekabül eder:
                        model.Add(live_total + asynch_total <= adjusted_targets[t_idx])
                else:
                    for t_idx, t in enumerate(teachers_list):
                        model.Add(sum(x[(t_idx, c, d, s)] for c in range(len(classes_list)) for d in days for s in sessions) <= adjusted_targets[t_idx])

                # Yasaklı Gün Kesin Kuralı
                if strict_forbidden_days:
                    for t_idx, t in enumerate(teachers_list):
                        forb = str(t.get('Yasaklı Günler', ''))
                        for d_idx, d_name in enumerate(day_names):
                            if d_name in forb:
                                for c in range(len(classes_list)):
                                    for s in sessions: model.Add(x[(t_idx, c, d_idx, s)] == 0)

                # --- PROFİLE GÖRE DİNAMİK AĞIRLIKLANDIRMA ---
                objective = []
                base_slot_reward = 1000000
                objective.append(sum(x.values()) * base_slot_reward)

                # Asenkron slotların doldurulması teşviki
                if enable_asynch:
                    for (t, c), a_var in asynch_var.items():
                        objective.append(a_var * base_slot_reward)

                if "Plan B: Öğretmen Odaklı" in variant_profile:
                    w_shift_penalty = -500000
                    w_split_penalty = -1000000
                    w_adv_mon = 200000
                    w_adv_multi = 100000
                    w_native_b2 = 50000
                elif "Plan C: Pedagoji" in variant_profile:
                    w_shift_penalty = -50000
                    w_split_penalty = -300000
                    w_adv_mon = 1000000
                    w_adv_multi = 500000
                    w_native_b2 = 300000
                else:
                    w_shift_penalty = -150000
                    w_split_penalty = -500000
                    w_adv_mon = 500000
                    w_adv_multi = 200000
                    w_native_b2 = 100000

                # Native Dağılımı ve Şelalesi
                for c_idx, c_data in enumerate(classes_list):
                    natives_in_c = []
                    for t_idx, t in enumerate(teachers_list):
                        if 'NATIVE' in get_role(t):
                            is_p = teacher_in_class[(t_idx, c_idx)]
                            natives_in_c.append(is_p)
                            lvl = c_data['Seviye']
                            score = w_native_b2 if lvl == "B2" else (w_native_b2 // 2 if lvl == "B1" else 10000)
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
                            objective.append(adv_pzt * w_adv_mon)

                        if c_data['Seviye'] != "PreFaculty":
                            cnt_days = sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions)
                            is_2p = model.NewBoolVar(f'i2_{t_idx}_{c_idx}')
                            model.Add(cnt_days >= 2).OnlyEnforceIf(is_2p)
                            model.Add(cnt_days <= 1).OnlyEnforceIf(is_2p.Not())
                            adv_2d = model.NewBoolVar(f'a2_{t_idx}_{c_idx}')
                            model.AddImplication(adv_2d, is_2p)
                            model.AddImplication(adv_2d, is_adv)
                            objective.append(adv_2d * w_adv_multi)

                # Vardiya Tercih Cezaları
                for t_idx, t in enumerate(teachers_list):
                    pref = get_pref(t)
                    if 'SABAH' in pref:
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 1)] * w_shift_penalty)
                    elif 'OGLE' in pref:
                        for c in range(len(classes_list)):
                            for d in days: objective.append(x[(t_idx, c, d, 0)] * w_shift_penalty)

                # Çift Vardiya (Split Shift) Cezası
                for t_idx, t in enumerate(teachers_list):
                    for d in days:
                        is_m = model.NewBoolVar(f'm_{t_idx}_{d}')
                        is_a = model.NewBoolVar(f'a_{t_idx}_{d}')
                        model.AddMaxEquality(is_m, [x[(t_idx, c, d, 0)] for c in range(len(classes_list))])
                        model.AddMaxEquality(is_a, [x[(t_idx, c, d, 1)] for c in range(len(classes_list))])
                        dbl = model.NewBoolVar(f'dbl_{t_idx}_{d}')
                        model.Add(is_m + is_a - 1 <= dbl)
                        objective.append(dbl * w_split_penalty)

                # Yasaklı Gün Esnek Cezası
                if not strict_forbidden_days:
                    for t_idx, t in enumerate(teachers_list):
                        forb = str(t.get('Yasaklı Günler', ''))
                        for d_idx, d_name in enumerate(day_names):
                            if d_name in forb:
                                for c in range(len(classes_list)):
                                    for s in sessions: objective.append(x[(t_idx, c, d_idx, s)] * -5000000)

                # Modeli Çöz
                model.Maximize(sum(objective))
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 90.0
                solver.parameters.num_search_workers = 8

                status = solver.Solve(model)

                if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    class_schedule = []
                    teacher_schedule = {t['Ad Soyad']: {d: "-" for d in day_names} for t in teachers_list}
                    teacher_asynch_assigned = {t['Ad Soyad']: [] for t in teachers_list}
                    violations = []
                    native_names = [t['Ad Soyad'] for t in teachers_list if 'NATIVE' in get_role(t)]

                    total_assigned_slots = 0
                    split_shift_count = 0
                    wrong_shift_count = 0
                    advisor_mon_success = 0

                    for c_idx, c in enumerate(classes_list):
                        c_name = c['Sınıf Adı']
                        s_req = c['Zaman Kodu']
                        adv_name = "Atanamadı"
                        for t_idx in range(len(teachers_list)):
                            if solver.Value(advisor_var[(t_idx, c_idx)]) == 1:
                                adv_name = teachers_list[t_idx]['Ad Soyad']
                                break

                        row = {"Sınıf": c_name, "Seviye": c['Seviye'], "Danışman": adv_name, "Vardiya": "Sabah" if s_req == 0 else "Öğle"}

                        # Canlı Günler
                        for d_idx, d_name in enumerate(day_names):
                            val = "🔴 BOŞ"
                            if c['Seviye'] == "PreFaculty" and d_idx >= 3:
                                val = "⛔ KAPALI"
                            else:
                                for t_idx, t in enumerate(teachers_list):
                                    if solver.Value(x[(t_idx, c_idx, d_idx, s_req)]) == 1:
                                        t_name = t['Ad Soyad']
                                        val = t_name
                                        total_assigned_slots += 1
                                        teacher_schedule[t_name][d_name] = f"{c_name} ({row['Vardiya']})"

                                        if d_idx == 0 and t_name == adv_name:
                                            advisor_mon_success += 1

                                        pref = get_pref(t)
                                        if 'SABAH' in pref and s_req == 1:
                                            wrong_shift_count += 1
                                            violations.append({"Hoca": t_name, "Tür": "Ters Vardiya", "Detay": c_name})
                                        elif 'OGLE' in pref and s_req == 0:
                                            wrong_shift_count += 1
                                            violations.append({"Hoca": t_name, "Tür": "Ters Vardiya", "Detay": c_name})
                                        break

                            if val == "🔴 BOŞ":
                                violations.append({"Hoca": "-", "Tür": f"Boş Canlı Ders ({d_name})", "Detay": c_name})

                            row[d_name] = val

                        # Asenkron Slot Değerlendirmesi
                        if enable_asynch:
                            if c['Seviye'] in ['B1', 'B2']:
                                asynch_teacher = "🔴 BOŞ"
                                for t_idx in range(len(teachers_list)):
                                    if solver.Value(asynch_var[(t_idx, c_idx)]) == 1:
                                        asynch_teacher = teachers_list[t_idx]['Ad Soyad']
                                        total_assigned_slots += 1
                                        teacher_asynch_assigned[asynch_teacher].append(c_name)
                                        break
                                row["Asenkron"] = asynch_teacher
                                if asynch_teacher == "🔴 BOŞ":
                                    violations.append({"Hoca": "-", "Tür": "Boş Asenkron Slot", "Detay": c_name})
                            else:
                                row["Asenkron"] = "—"

                        class_schedule.append(row)

                    for t_idx, t in enumerate(teachers_list):
                        for d in days:
                            m_cnt = sum(solver.Value(x[(t_idx, c, d, 0)]) for c in range(len(classes_list)))
                            a_cnt = sum(solver.Value(x[(t_idx, c, d, 1)]) for c in range(len(classes_list)))
                            if m_cnt > 0 and a_cnt > 0:
                                split_shift_count += 1

                    stats_list = []
                    for t_idx, t in enumerate(teachers_list):
                        t_name = t['Ad Soyad']
                        assigned_live = sum(solver.Value(x[(t_idx, c, d, s)]) for c in range(len(classes_list)) for d in days for s in sessions)
                        assigned_as = len(teacher_asynch_assigned[t_name])
                        total_t = assigned_live + assigned_as
                        target_t = adjusted_targets[t_idx]

                        stats_list.append({
                            "Ad Soyad": t_name,
                            "Rol": t.get('Rol', ''),
                            "Hedef": target_t,
                            "Canlı Ders": assigned_live,
                            "Asenkron": assigned_as,
                            "Toplam": total_t,
                            "Durum": "Kusursuz" if total_t == target_t else f"{target_t - total_t} Eksik"
                        })

                    teacher_rows = []
                    for s_row in stats_list:
                        t_name = s_row['Ad Soyad']
                        t_dict = dict(s_row)
                        t_dict.update(teacher_schedule[t_name])
                        if enable_asynch:
                            t_dict["Asenkron Sınıf"] = ", ".join(teacher_asynch_assigned[t_name]) if teacher_asynch_assigned[t_name] else "—"
                        teacher_rows.append(t_dict)

                    total_needed_slots = sum([3 if c['Seviye'] == 'PreFaculty' else 5 for c in classes_list]) + ((count_b1 + count_b2) if enable_asynch else 0)
                    fill_rate = round((total_assigned_slots / total_needed_slots) * 100, 1) if total_needed_slots > 0 else 100
                    adv_mon_rate = round((advisor_mon_success / len(classes_list)) * 100, 1) if classes_list else 100

                    st.session_state['solution_output'] = {
                        "profile_name": variant_profile,
                        "df_classes": pd.DataFrame(class_schedule),
                        "df_teachers": pd.DataFrame(teacher_rows),
                        "df_stats": pd.DataFrame(stats_list),
                        "df_violations": pd.DataFrame(violations).drop_duplicates() if violations else pd.DataFrame(),
                        "native_names": native_names,
                        "enable_asynch": enable_asynch,
                        "metrics": {
                            "Profil": variant_profile,
                            "Ders Doluluk (%)": fill_rate,
                            "Boş Kalan Slot": total_needed_slots - total_assigned_slots,
                            "Ters Vardiya Sayısı": wrong_shift_count,
                            "Çift Vardiya Sayısı": split_shift_count,
                            "Danışman Pazartesi Uyum (%)": adv_mon_rate
                        }
                    }
                else:
                    st.error("❌ Model mevcut kısıtlarla çözülemedi (Infeasible).")
                    with st.spinner("🩺 Akıllı Kriz Dedektifi çalıştırılıyor... Kilitlenmenin kök nedeni taranıyor..."):
                        diags = run_diagnostic_engine(teachers_list, classes_list, adjusted_targets, max_teachers_per_class, strict_forbidden_days, enable_asynch)
                        st.warning("### 🩺 Kriz Dedektifi Teşhis Raporu:")
                        st.info("Aşağıdaki çakışmalar sistemin kilitlenmesine neden oldu. Lütfen bu noktaları esnetmeyi deneyin:")
                        for d in diags:
                            st.markdown(f"- {d}")

    # --- ÇÖZÜM GÖRSELLEŞTİRME VE İNDİRME ---
    if st.session_state['solution_output'] is not None:
        st.balloons()
        sol = st.session_state['solution_output']

        st.success(f"✅ Program başarıyla oluşturuldu! (**Kullanılan Strateji:** {sol['profile_name']})")

        # Senaryo Kaydetme
        sc_col1, sc_col2 = st.columns([3, 1])
        with sc_col1:
            scenario_name_input = st.text_input("Bu Çözümü Senaryo Olarak Kaydet:", value=f"Senaryo_{len(st.session_state['saved_scenarios'])+1} ({sol['profile_name'][:12]})")
        with sc_col2:
            st.write("")
            st.write("")
            if st.button("💾 Senaryoyu Kaydet", use_container_width=True):
                st.session_state['saved_scenarios'][scenario_name_input] = sol['metrics']
                st.toast(f"'{scenario_name_input}' senaryosu 4. Sekmeye kaydedildi!")

        if not sol['df_violations'].empty:
            with st.expander(f"⚠️ {len(sol['df_violations'])} Noktada Esneme / Tercih Dışı Durum Oluştu", expanded=False):
                st.dataframe(sol['df_violations'], use_container_width=True)

        view_mode = st.radio("📋 Çizelge Görünüm Modu:", ["🏫 Sınıf Bazlı Program", "👨‍🏫 Öğretmen Bazlı Program", "📊 Öğretmen İstatistikleri"], horizontal=True)

        if view_mode == "🏫 Sınıf Bazlı Program":
            st.dataframe(sol['df_classes'], use_container_width=True)
        elif view_mode == "👨‍🏫 Öğretmen Bazlı Program":
            selected_teacher = st.selectbox("Belirli Bir Öğretmeni İzole Et:", ["Tümü"] + list(sol['df_teachers']['Ad Soyad']))
            if selected_teacher == "Tümü":
                st.dataframe(sol['df_teachers'], use_container_width=True)
            else:
                st.dataframe(sol['df_teachers'][sol['df_teachers']['Ad Soyad'] == selected_teacher], use_container_width=True)
        else:
            st.dataframe(sol['df_stats'], use_container_width=True)

        # Excel İndirme
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
            fmt_asynch = wb.add_format(dict(base_fmt, bg_color='#E6E6FA', bold=True))

            ws_c.set_column('A:B', 12)
            ws_c.set_column('C:C', 20)
            ws_c.set_column('E:I', 15)
            if sol['enable_asynch']:
                ws_c.set_column('J:J', 18)

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

                if sol['enable_asynch']:
                    as_val = row.get('Asenkron', '—')
                    ws_c.write(e_r, 9, as_val, fmt_asynch if as_val not in ['—', '🔴 BOŞ'] else fmt_def)

            ws_t.set_column('A:B', 18)
            ws_t.set_column('G:L', 16)

        st.download_button(
            "📥 Çok Sekmeli Kurumsal Excel'i İndir (.xlsx)",
            output_res.getvalue(),
            "hazirlik_ders_programi_v76.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ==============================================================================
# SEKME 4: WHAT-IF SENARYO KIYASLAMA EKRANI
# ==============================================================================
with tab_whatif:
    st.subheader("⚖️ 'What-If' Senaryo & Profil Kıyaslama")
    st.write("Farklı kural kombinasyonları veya çözüm profilleriyle ürettiğiniz programları burada yan yana kıyaslayabilirsiniz.")

    if not st.session_state['saved_scenarios']:
        st.info("💡 Henüz kayıtlı bir senaryo yok. Sekme 3'te program oluşturduktan sonra **'💾 Senaryoyu Kaydet'** butonuna basarak buraya ekleyebilirsiniz.")
    else:
        df_scenarios = pd.DataFrame(st.session_state['saved_scenarios']).T
        st.dataframe(df_scenarios, use_container_width=True)

        if st.button("🗑️ Kayıtlı Senaryoları Temizle"):
            st.session_state['saved_scenarios'] = {}
            st.rerun()
