def analyze_data(teachers, classes, allow_native_advisor):
    warnings = []
    errors = []

    for t in teachers:
        role = str(t['Rol']).upper()
        fixed_class = str(t['Sabit Sınıf']).strip()
        forbidden_str = str(t['Yasaklı Günler'])

        if not allow_native_advisor and "NATIVE" in role and fixed_class:
             errors.append(f"🛑 **{t['Ad Soyad']}**: Native hocaya sabit sınıf verilmesi engellendi.")

        if fixed_class:
            target_class = next((c for c in classes if c['Sınıf Adı'] == fixed_class), None)
            if not target_class:
                errors.append(f"❌ **{t['Ad Soyad']}**: Atandığı '{fixed_class}' sınıfı sistemde yok.")

            if "Pazartesi" in forbidden_str:
                warnings.append(f"⚠️ **Uyarı ({t['Ad Soyad']}):** '{fixed_class}' danışmanı ama Pazartesi yasaklı. Sistem Pazartesi ders yazamayabilir.")

    return errors, warnings
