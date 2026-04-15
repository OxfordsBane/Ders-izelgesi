import pandas as pd
from ortools.sat.python import cp_model

def create_automated_classes():
    class_list = []
    config = [
        (4, "A1", 0),
        (4, "A2", 0),
        (4, "B1", 0),
        (2, "B2", 0),
        (0, "PreFaculty", 0),
    ]
    for count, lvl, time_code in config:
        for i in range(1, count + 1):
            class_name = f"{lvl}.{i:02d}"
            class_list.append({"Sınıf Adı": class_name, "Seviye": lvl, "Zaman Kodu": time_code})
    return pd.DataFrame(class_list)

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

df_classes = create_automated_classes()
teachers_list = df_teachers.to_dict('records')
classes_list = df_classes.to_dict('records')

def run_solver(skip_constraint=None):
    model = cp_model.CpModel()
    days = range(5)
    sessions = range(2)

    x = {}
    advisor_var = {}

    for t in range(len(teachers_list)):
        for c in range(len(classes_list)):
            advisor_var[(t, c)] = model.NewBoolVar(f'adv_{t}_{c}')
            for d in days:
                for s in sessions:
                    x[(t, c, d, s)] = model.NewBoolVar(f'x_{t}_{c}_{d}_{s}')

    # 1. Fiziksel Çakışma
    for t in range(len(teachers_list)):
        for d in days:
            for s in sessions:
                model.Add(sum(x[(t, c, d, s)] for c in range(len(classes_list))) <= 1)

    # 2. Sınıf Kapasitesi (includes PreFaculty)
    for c_idx, c_data in enumerate(classes_list):
        req_session = c_data['Zaman Kodu']
        other_session = 1 - req_session
        for d in days:
            if skip_constraint == 'PreFaculty' and c_data['Seviye'] == "PreFaculty" and d >= 3:
                # Disabling PreFaculty closure hard constraint
                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
            elif c_data['Seviye'] == "PreFaculty" and d >= 3:
                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) == 0)
            else:
                model.Add(sum(x[(t, c_idx, d, req_session)] for t in range(len(teachers_list))) <= 1)
            model.Add(sum(x[(t, c_idx, d, other_session)] for t in range(len(teachers_list))) == 0)

    # 3. Yetkinlik (Eligibility)
    if skip_constraint != 'Eligibility':
        for t_idx, t in enumerate(teachers_list):
            allowed = str(t['Yetkinlik (Seviyeler)'])
            if "Hepsi" not in allowed:
                for c_idx, c in enumerate(classes_list):
                    if c['Seviye'] not in allowed:
                        for d in days:
                            for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)
                        model.Add(advisor_var[(t_idx, c_idx)] == 0)

    # 4. Danışman Tekilliği (Uniqueness)
    if skip_constraint != 'Uniqueness':
        for c in range(len(classes_list)):
            model.Add(sum(advisor_var[(t, c)] for t in range(len(teachers_list))) == 1)
        for t in range(len(teachers_list)):
            model.Add(sum(advisor_var[(t, c)] for c in range(len(classes_list))) <= 1)

    # 5. Sabit Sınıf (Fixed class assignments)
    if skip_constraint != 'Fixed class':
        for t_idx, t in enumerate(teachers_list):
            if t['Sabit Sınıf']:
                fixed_c_idx = next((i for i, c in enumerate(classes_list) if c['Sınıf Adı'] == str(t['Sabit Sınıf'])), None)
                if fixed_c_idx is not None:
                    model.Add(advisor_var[(t_idx, fixed_c_idx)] == 1)

    # Rol Kısıtları (Ek Görevli ve Native Danışman)
    allow_native_advisor = False
    for t_idx, t in enumerate(teachers_list):
        if 'Ek Görevli' in str(t['Rol']):
            for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)
        if not allow_native_advisor and 'Native' in str(t['Rol']):
            for c in range(len(classes_list)): model.Add(advisor_var[(t_idx, c)] == 0)

    # Native A1 Yasağı
    for t_idx, t in enumerate(teachers_list):
        if 'Native' in str(t['Rol']):
            for c_idx, c_data in enumerate(classes_list):
                if c_data['Seviye'] == 'A1':
                    for d in days:
                        for s in sessions: model.Add(x[(t_idx, c_idx, d, s)] == 0)

    # Ek Görevli Gezici
    for t_idx, t in enumerate(teachers_list):
        if 'Ek Görevli' in str(t['Rol']):
            for c_idx in range(len(classes_list)):
                model.Add(sum(x[(t_idx, c_idx, d, s)] for d in days for s in sessions) <= 1)

    solver = cp_model.CpSolver()
    status = solver.Solve(model)
    if status == cp_model.OPTIMAL:
        return "OPTIMAL"
    elif status == cp_model.FEASIBLE:
        return "FEASIBLE"
    elif status == cp_model.INFEASIBLE:
        return "INFEASIBLE"
    else:
        return "UNKNOWN"

if __name__ == "__main__":
    print("Testing constraints to find the cause of INFEASIBLE status:")
    print(f"Base Configuration (All rules ON): {run_solver()}")
    print(f"Disabling 'PreFaculty day closures': {run_solver('PreFaculty')}")
    print(f"Disabling 'Eligibility': {run_solver('Eligibility')}")
    print(f"Disabling 'Uniqueness': {run_solver('Uniqueness')}")
    print(f"Disabling 'Fixed class assignments': {run_solver('Fixed class')}")
