import pytest
from app import generate_schedule
from ortools.sat.python import cp_model

def create_mock_teacher(name="Hoca", role="Destek", target=5, preferences="", forbidden="", fixed_class="", allowed_levels="Hepsi"):
    return {
        'Ad Soyad': name,
        'Rol': role,
        'Hedef Ders Sayısı': target,
        'Tercih (Sabah/Öğle)': preferences,
        'Yasaklı Günler': forbidden,
        'Sabit Sınıf': fixed_class,
        'Yetkinlik (Seviyeler)': allowed_levels,
        'İstenmeyen Partner': ''
    }

def create_mock_class(name="A1.01", level="A1", session=0): # 0: Sabah, 1: Öğle
    return {
        'Sınıf Adı': name,
        'Seviye': level,
        'Zaman Kodu': session
    }

def run_solver(teachers, classes, allow_native_advisor=False, reduce_mode=False):
    status, solver, x, advisor_var, adjusted_targets, days, day_names, sessions = generate_schedule(
        teachers, classes, allow_native_advisor, reduce_mode
    )
    return status, solver, x, advisor_var, days, sessions

def test_fiziksel_cakisma():
    t = [
        create_mock_teacher(name="Hoca1", target=10),
        create_mock_teacher(name="Dummy1", target=1),
        create_mock_teacher(name="Dummy2", target=1)
    ]
    c = [
        create_mock_class(name="C1", session=0),
        create_mock_class(name="C2", session=0)
    ]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for d in days:
        for s in sessions:
            assigned = sum(1 for c_idx in range(len(c)) if solver.Value(x[(0, c_idx, d, s)]) == 1)
            assert assigned <= 1

def test_sinif_kapasitesi():
    t = [
        create_mock_teacher(name="H1", target=5),
        create_mock_teacher(name="H2", target=5)
    ]
    c = [create_mock_class(name="C1", session=1)]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for d in days:
        morning_assigned = sum(1 for t_idx in range(len(t)) if solver.Value(x[(t_idx, 0, d, 0)]) == 1)
        assert morning_assigned == 0

        afternoon_assigned = sum(1 for t_idx in range(len(t)) if solver.Value(x[(t_idx, 0, d, 1)]) == 1)
        assert afternoon_assigned <= 1

def test_prefaculty_kapama():
    t = [create_mock_teacher(target=10)]
    c = [create_mock_class(level="PreFaculty", session=0)]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for d in [3, 4]:
        assigned = sum(1 for t_idx in range(len(t)) if solver.Value(x[(0, 0, d, 0)]) == 1)
        assert assigned == 0

def test_yetkinlik():
    t = [
        create_mock_teacher(name="H1", allowed_levels="A1,A2"),
        create_mock_teacher(name="Dummy1", allowed_levels="Hepsi") # To act as advisor and teacher
    ]
    c = [create_mock_class(level="B1", session=0)]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)
    assert solver.Value(adv[(0, 0)]) == 0
    for d in days:
        for s in sessions:
            assert solver.Value(x[(0, 0, d, s)]) == 0

def test_danisman_tekilligi():
    t = [create_mock_teacher("H1"), create_mock_teacher("H2")]
    c = [create_mock_class("C1"), create_mock_class("C2")]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for c_idx in range(len(c)):
        assigned = sum(solver.Value(adv[(t_idx, c_idx)]) for t_idx in range(len(t)))
        assert assigned == 1

    for t_idx in range(len(t)):
        assigned = sum(solver.Value(adv[(t_idx, c_idx)]) for c_idx in range(len(c)))
        assert assigned <= 1

def test_sabit_sinif():
    t = [create_mock_teacher("H1", fixed_class="C1"), create_mock_teacher("H2")]
    c = [create_mock_class("C1", session=0), create_mock_class("C2", session=1)]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)
    assert solver.Value(adv[(0, 0)]) == 1  # H1 -> C1

def test_rol_kisitlari():
    t = [
        create_mock_teacher("H1", role="Ek Görevli"),
        create_mock_teacher("H2", role="Native"),
        create_mock_teacher("Dummy1"),
        create_mock_teacher("Dummy2")
    ]
    c = [create_mock_class("C1"), create_mock_class("C2")]
    status, solver, x, adv, days, sessions = run_solver(t, c, allow_native_advisor=False)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for c_idx in range(len(c)):
        assert solver.Value(adv[(0, c_idx)]) == 0
        assert solver.Value(adv[(1, c_idx)]) == 0

def test_native_a1_yasagi():
    t = [
        create_mock_teacher("H1", role="Native"),
        create_mock_teacher("Dummy1")
    ]
    c = [create_mock_class(level="A1")]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    for d in days:
        for s in sessions:
            assert solver.Value(x[(0, 0, d, s)]) == 0

def test_ek_gorevli_gezici():
    t = [
        create_mock_teacher("H1", role="Ek Görevli", target=5),
        create_mock_teacher("Dummy1")
    ]
    c = [create_mock_class("C1")]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    total_assigned = sum(solver.Value(x[(0, 0, d, s)]) for d in days for s in sessions)
    assert total_assigned <= 1


def test_pazartesi_kurali():
    # Danışman Pazartesi günleri dersinde olmaya TEŞVİK edilir (Soft constraint).
    # H1 C1'in danışmanı olsun ve Pazartesi (d=0) ders versin istiyoruz.
    t = [create_mock_teacher("H1", fixed_class="C1", target=5)]
    c = [create_mock_class("C1", session=0)]

    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    # H1, C1'in danışmanı ve Pazartesi oradaydı (eğer yasaklı değilse ve hedef ders sayısı yeterliyse)
    # Burada tek hoca, tek sınıf ve hedefi 5, o yüzden her gün girecek ve pazartesi orada olması optimum.
    assert solver.Value(adv[(0, 0)]) == 1
    assert solver.Value(x[(0, 0, 0, 0)]) == 1 # d=0 (Pazartesi), s=0 (Sabah)

def test_3_gun_kurali():
    # Danışmanın sınıfına en az 3 gün girmesi teşvik edilir.
    t = [
        create_mock_teacher("H1", fixed_class="C1", target=5),
        create_mock_teacher("Dummy1", target=5) # Dummy
    ]
    c = [create_mock_class("C1", session=0), create_mock_class("C2", session=0)]

    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    # H1 C1'in danışmanı
    assert solver.Value(adv[(0, 0)]) == 1

    # H1'in C1'deki ders sayısı >= 3 olmalı (teşvik edildiği ve engeli olmadığı için)
    assigned_days = sum(solver.Value(x[(0, 0, d, 0)]) for d in days)
    assert assigned_days >= 3

def test_tek_vardiya_kurali():
    # Aynı gün hem sabah hem öğle ders verilmesi cezalandırılır. (Tek vardiya kuralı)
    # Bir öğretmeni hem sabah hem öğle sınıflarına ihtiyaç duyacak şekilde atamaya çalışalım ama soft constraint bunu önlemeye çalışacaktır.
    # Ancak başka çare yoksa girebilir.
    # Burada yeterli öğretmen verirsek, kimse çift vardiya yapmaz.
    t = [
        create_mock_teacher("H1", target=10),
        create_mock_teacher("H2", target=10)
    ]
    c = [
        create_mock_class("C1", session=0), # Sabah
        create_mock_class("C2", session=1)  # Öğle
    ]
    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    # Her gün için H1'in çift vardiya yapıp yapmadığına bakalım
    for t_idx in range(len(t)):
        for d in days:
            morning = sum(solver.Value(x[(t_idx, c_idx, d, 0)]) for c_idx in range(len(c)))
            afternoon = sum(solver.Value(x[(t_idx, c_idx, d, 1)]) for c_idx in range(len(c)))
            assert morning + afternoon <= 1 # Kimse çift vardiya yapmamalı çünkü başkası var

def test_native_siniri():
    # Native hoca bir sınıfa haftada en fazla 1 kez girer. (Soft constraint)
    # Burada da tek Native var, başka şans yoksa birden fazla girer ama ceza alır.
    # Dummy öğretmen varsa, solver cezadan kaçınmak için Native'i max 1 kez atar.
    t = [
        create_mock_teacher("H1", role="Native", target=5),
        create_mock_teacher("Dummy1", target=5)
    ]
    c = [create_mock_class("C1", level="A2", session=0)]

    status, solver, x, adv, days, sessions = run_solver(t, c)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    # H1 (Native) C1'e en fazla 1 kere girmeli
    assigned_days = sum(solver.Value(x[(0, 0, d, s)]) for d in days for s in sessions)
    assert assigned_days <= 1
