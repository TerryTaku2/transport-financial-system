#!/usr/bin/env python3
"""One-off local dev helper: populate a realistic demo dataset across every
module (fleet, drivers, franchise, store, finance) so the app and its
reports/dashboard have something meaningful to show on a fresh local
database. Safe to re-run — guarded by a sentinel vehicle registration, so a
second run is a no-op instead of creating duplicates.

Also removes a handful of stray rows (test vehicles XYZ999/EXP001/SOON001/
OK001, franchise vehicle FR001, and an accompanying audit log entry / admin
user) that were created by ad-hoc manual testing against this same local
database before this script existed — see the app's git history around the
insurance-classification and franchise-deposits features. Harmless if those
rows are already gone.

Run with: python seed_demo_data.py
"""
import random
from datetime import date, timedelta

from app import (
    app, db, touch_sync_fields,
    User, Vehicle, VehicleDocument, Driver, Route, DailyLog, FuelLog,
    MaintenanceLog, MaintenanceSchedule, Loan, LoanPayment, Payable,
    Receivable, CapitalContribution, OwnerDrawing, ExpenseCategory, Expense,
    DriverDeposit, FranchiseVehicle, FranchiseDailyIncome, FranchiseWeeklyIncome,
    SparePart, StorePurchase, StoreSale, AuditLog,
)

SENTINEL_REG = 'ZWD-101'


def cleanup_stray_test_rows():
    """Remove specific rows created by earlier manual testing against this
    same local DB, identified by name — never a blanket wipe."""
    FranchiseDailyIncome.query.filter(
        FranchiseDailyIncome.vehicle_id.in_(
            db.session.query(FranchiseVehicle.id).filter_by(number_plate='FR001'))).delete(
        synchronize_session=False)
    FranchiseWeeklyIncome.query.filter(
        FranchiseWeeklyIncome.vehicle_id.in_(
            db.session.query(FranchiseVehicle.id).filter_by(number_plate='FR001'))).delete(
        synchronize_session=False)
    FranchiseVehicle.query.filter_by(number_plate='FR001').delete()

    stray_vehicle_ids = [v.id for v in Vehicle.query.filter(
        Vehicle.registration.in_(['XYZ999', 'EXP001', 'SOON001', 'OK001'])).all()]
    if stray_vehicle_ids:
        AuditLog.query.filter(AuditLog.table_name == 'vehicles',
                               AuditLog.record_id.in_(stray_vehicle_ids)).delete(synchronize_session=False)
        Vehicle.query.filter(Vehicle.id.in_(stray_vehicle_ids)).delete(synchronize_session=False)

    db.session.commit()


def seed():
    if Vehicle.query.filter_by(registration=SENTINEL_REG).first():
        print('Demo data already present (found ZWD-101) — skipping seed.')
        return

    admin = User.query.filter_by(username='admin').first()
    admin_id = admin.id if admin else None
    today = date.today()

    # ── Vehicles ────────────────────────────────────────────────
    vehicles_spec = [
        dict(registration='ZWD-101', make='Toyota', model='Quantum', year=2019,
             acquisition_cost=18000, status='active', fuel_type='diesel', daily_target=120,
             insurance_provider='Old Mutual', insurance_type='Comprehensive',
             insurance_policy_number='OM-4471', insurance_expiry=today + timedelta(days=200)),
        dict(registration='ZWD-102', make='Toyota', model='Hiace', year=2020,
             acquisition_cost=16500, status='active', fuel_type='diesel', daily_target=110,
             insurance_provider='Zimnat', insurance_type='Third Party, Fire & Theft',
             insurance_policy_number='ZN-8823', insurance_expiry=today + timedelta(days=15)),
        dict(registration='ZWD-103', make='Nissan', model='Caravan', year=2018,
             acquisition_cost=14000, status='active', fuel_type='diesel', daily_target=100,
             insurance_provider='NicozDiamond', insurance_type='Passenger Liability',
             insurance_policy_number='ND-1290', insurance_expiry=today - timedelta(days=10)),
        dict(registration='ZWD-104', make='Toyota', model='Quantum', year=2021,
             acquisition_cost=21000, status='maintenance', fuel_type='diesel', daily_target=120,
             insurance_provider=None, insurance_type=None,
             insurance_policy_number=None, insurance_expiry=None),
        dict(registration='ZWD-105', make='Isuzu', model='NPR Bus', year=2017,
             acquisition_cost=26000, status='active', fuel_type='diesel', daily_target=150,
             insurance_provider='Old Mutual', insurance_type='Comprehensive',
             insurance_policy_number='OM-5510', insurance_expiry=today + timedelta(days=90)),
    ]
    vehicles = []
    for spec in vehicles_spec:
        v = Vehicle(**spec)
        touch_sync_fields(v)
        db.session.add(v)
        vehicles.append(v)
    db.session.flush()

    for v in vehicles:
        for doc_type, days_from_today in [('fitness', 45), ('permit', -5), ('license', 20)]:
            doc = VehicleDocument(
                vehicle_id=v.id, doc_type=doc_type,
                reference_number=f'{doc_type[:3].upper()}-{v.id}{random.randint(100, 999)}',
                issue_date=today - timedelta(days=200),
                expiry_date=today + timedelta(days=days_from_today),
            )
            touch_sync_fields(doc)
            db.session.add(doc)

    schedule = MaintenanceSchedule(
        vehicle_id=vehicles[0].id, description='Full service', interval_days=90,
        last_done_date=today - timedelta(days=75), next_due_date=today + timedelta(days=15),
        status='active')
    touch_sync_fields(schedule)
    db.session.add(schedule)

    # ── Drivers & conductors ───────────────────────────────────
    drivers_spec = [
        dict(name='Tendai Moyo', role='driver', license_number='DL-10001', phone='0771000001', commission_rate=0.15),
        dict(name='Farai Ncube', role='driver', license_number='DL-10002', phone='0771000002', commission_rate=0.15),
        dict(name='Blessing Chikwanha', role='driver', license_number='DL-10003', phone='0771000003', commission_rate=0.15),
        dict(name='Rutendo Gumbo', role='conductor', license_number=None, phone='0771000004', commission_rate=0.10),
        dict(name='Simbarashe Dube', role='conductor', license_number=None, phone='0771000005', commission_rate=0.10),
    ]
    drivers = []
    for spec in drivers_spec:
        d = Driver(status='active', **spec)
        touch_sync_fields(d)
        db.session.add(d)
        drivers.append(d)
    db.session.flush()
    active_drivers = [d for d in drivers if d.role == 'driver']
    active_conductors = [d for d in drivers if d.role == 'conductor']

    # ── Routes ──────────────────────────────────────────────────
    routes_spec = [
        dict(name='Harare - Chitungwiza', start_point='Harare CBD', end_point='Chitungwiza', distance_km=25, fare_rate=1.0),
        dict(name='Harare - Norton', start_point='Harare CBD', end_point='Norton', distance_km=40, fare_rate=1.5),
        dict(name='Harare - Ruwa', start_point='Harare CBD', end_point='Ruwa', distance_km=22, fare_rate=1.0),
    ]
    routes = []
    for spec in routes_spec:
        r = Route(status='active', **spec)
        touch_sync_fields(r)
        db.session.add(r)
        routes.append(r)
    db.session.flush()

    # ── Daily transactions, fuel & maintenance logs (last 30 days) ─
    operating_vehicles = [v for v in vehicles if v.status == 'active']
    for days_ago in range(29, -1, -1):
        d = today - timedelta(days=days_ago)
        for v in operating_vehicles:
            if random.random() < 0.15:
                continue  # the odd vehicle-off-road day, for realism
            driver = random.choice(active_drivers)
            conductor = random.choice(active_conductors)
            route = random.choice(routes)
            trips = random.randint(4, 9)
            revenue = round(trips * random.uniform(35, 65), 2)
            log = DailyLog(
                vehicle_id=v.id, driver_id=driver.id, conductor_id=conductor.id,
                route_id=route.id, log_date=d, trips_completed=trips,
                gross_revenue=revenue, garnish=0.0, created_by=admin_id)
            touch_sync_fields(log)
            db.session.add(log)

            if random.random() < 0.3:
                liters = round(random.uniform(15, 40), 1)
                cost_per_liter = 1.5
                fuel = FuelLog(
                    vehicle_id=v.id, log_date=d, liters=liters,
                    cost_per_liter=cost_per_liter, total_cost=round(liters * cost_per_liter, 2),
                    created_by=admin_id)
                touch_sync_fields(fuel)
                db.session.add(fuel)

        if days_ago % 9 == 0:
            v = random.choice(vehicles)
            parts_cost = round(random.uniform(30, 250), 2)
            labor_cost = round(random.uniform(20, 100), 2)
            maint = MaintenanceLog(
                vehicle_id=v.id, log_date=d, description='Routine service / repair',
                parts_cost=parts_cost, labor_cost=labor_cost,
                total_cost=parts_cost + labor_cost, mechanic='Fleet Workshop',
                created_by=admin_id)
            touch_sync_fields(maint)
            db.session.add(maint)
    db.session.flush()

    # ── Expenses, capital, drawings, deposits ──────────────────
    admin_cat = ExpenseCategory.query.filter_by(name='Admin', parent_id=None).first()
    wages_cat = ExpenseCategory.query.filter_by(name='Wages', parent_id=None).first()
    garage_cat = ExpenseCategory.query.filter_by(name='Garage Fee', parent_id=None).first()

    for i, (cat, amount, vid) in enumerate([
        (admin_cat, 150.0, None), (wages_cat, 600.0, None),
        (garage_cat, 100.0, vehicles[0].id), (garage_cat, 100.0, vehicles[1].id),
    ]):
        exp = Expense(category_id=cat.id, vehicle_id=vid, expense_date=today - timedelta(days=i * 3),
                      description=f'{cat.name} — demo entry', amount=amount, created_by=admin_id)
        touch_sync_fields(exp)
        db.session.add(exp)

    cap = CapitalContribution(contributor='Owner', amount=5000.0,
                               contribution_date=today - timedelta(days=10), created_by=admin_id)
    touch_sync_fields(cap)
    db.session.add(cap)

    draw = OwnerDrawing(amount=800.0, drawing_date=today - timedelta(days=4), created_by=admin_id)
    touch_sync_fields(draw)
    db.session.add(draw)

    for days_ago in (1, 4, 8):
        dep = DriverDeposit(deposit_date=today - timedelta(days=days_ago),
                             amount=round(random.uniform(200, 500), 2), created_by=admin_id)
        touch_sync_fields(dep)
        db.session.add(dep)

    # ── Payables / receivables / loan ──────────────────────────
    payables_spec = [
        ('AutoParts Zimbabwe', 320.0, 'unpaid'),
        ('Fuel Wholesalers', 540.0, 'unpaid'),
        ('Tyre City', 210.0, 'paid'),
    ]
    for supplier, amount, status in payables_spec:
        p = Payable(supplier_name=supplier, description='Demo invoice', amount=amount,
                    invoice_date=today - timedelta(days=15), due_date=today + timedelta(days=15),
                    status=status, paid_date=today - timedelta(days=2) if status == 'paid' else None,
                    created_by=admin_id)
        touch_sync_fields(p)
        db.session.add(p)

    receivables_spec = [
        ('Corporate Charter Client', 450.0, 'outstanding'),
        ('School Contract', 300.0, 'outstanding'),
        ('Wedding Charter', 180.0, 'collected'),
    ]
    for client, amount, status in receivables_spec:
        rcv = Receivable(client_name=client, description='Demo receivable', amount=amount,
                         invoice_date=today - timedelta(days=12), due_date=today + timedelta(days=18),
                         status=status, collected_date=today - timedelta(days=1) if status == 'collected' else None,
                         created_by=admin_id)
        touch_sync_fields(rcv)
        db.session.add(rcv)

    loan = Loan(lender='Local Bank', principal=10000.0, interest_rate=12.0,
                start_date=today - timedelta(days=180), term_months=24, status='active',
                created_by=admin_id)
    touch_sync_fields(loan)
    db.session.add(loan)
    db.session.flush()
    payment = LoanPayment(loan_id=loan.id, payment_date=today - timedelta(days=20), amount=500.0,
                          created_by=admin_id)
    touch_sync_fields(payment)
    db.session.add(payment)

    # ── Spares store ────────────────────────────────────────────
    parts_spec = [
        ('Brake Pads (Set)', 'BP-100', 'set', 12.0, 40, 20, 5),
        ('Engine Oil 5L', 'EO-500', 'unit', 18.0, 30, 15, 4),
        ('Wheel Bearing', 'WB-220', 'unit', 9.0, 35, 10, 3),
    ]
    parts = []
    for name, part_number, unit, cost_price, markup, qty, reorder in parts_spec:
        part = SparePart(name=name, part_number=part_number, unit=unit, cost_price=cost_price,
                         markup_percent=markup, quantity_on_hand=qty, reorder_level=reorder,
                         status='active', created_by=admin_id)
        touch_sync_fields(part)
        db.session.add(part)
        parts.append(part)
    db.session.flush()

    for part in parts:
        purchase = StorePurchase(part_id=part.id, purchase_date=today - timedelta(days=6),
                                 quantity=10, unit_cost=part.cost_price,
                                 total_cost=round(10 * part.cost_price, 2),
                                 supplier='Parts Wholesaler', created_by=admin_id)
        touch_sync_fields(purchase)
        db.session.add(purchase)

    unit_price = round(parts[0].cost_price * (1 + parts[0].markup_percent / 100), 2)
    sale_ext = StoreSale(part_id=parts[0].id, sale_date=today - timedelta(days=2), quantity=2,
                         unit_cost=parts[0].cost_price, unit_price=unit_price,
                         total_amount=round(unit_price * 2, 2), customer_name='Walk-in Customer',
                         created_by=admin_id)
    touch_sync_fields(sale_ext)
    db.session.add(sale_ext)

    unit_price2 = round(parts[1].cost_price * (1 + parts[1].markup_percent / 100), 2)
    sale_internal = StoreSale(part_id=parts[1].id, vehicle_id=vehicles[0].id,
                              sale_date=today - timedelta(days=3), quantity=1,
                              unit_cost=parts[1].cost_price, unit_price=unit_price2,
                              total_amount=unit_price2, created_by=admin_id)
    touch_sync_fields(sale_internal)
    db.session.add(sale_internal)

    # ── Franchise ───────────────────────────────────────────────
    franchise_spec = [
        ('FRX-201', 'Alice Mutasa', 15.0, None),
        ('FRX-202', 'Brian Karimu', None, 80.0),
    ]
    fvs = []
    for plate, name, daily_fee, weekly_fee in franchise_spec:
        fv = FranchiseVehicle(number_plate=plate, franchisee_name=name, status='active',
                              daily_fee=daily_fee, weekly_fee=weekly_fee, amount_owed=0)
        touch_sync_fields(fv)
        db.session.add(fv)
        fvs.append(fv)
    db.session.flush()

    for days_ago in (1, 2, 3, 4, 5):
        d = today - timedelta(days=days_ago)
        income = round(random.uniform(60, 100), 2)
        exp_fines = round(random.uniform(0, 10), 2)
        expenditure = round(random.uniform(5, 20), 2)
        deposited = round(income - exp_fines - expenditure - random.uniform(0, 8), 2)
        entry = FranchiseDailyIncome(
            entry_date=d, vehicle_id=fvs[0].id, income=income, exp_traffic_fines=exp_fines,
            exp_workshop=expenditure, deposited=max(deposited, 0), created_by=admin_id)
        touch_sync_fields(entry)
        db.session.add(entry)

    weekly = FranchiseWeeklyIncome(
        week_start=today - timedelta(days=today.weekday()), vehicle_id=fvs[1].id,
        income=560.0, exp_wages=40.0, exp_facilitation_fees=20.0, deposited=480.0,
        created_by=admin_id)
    touch_sync_fields(weekly)
    db.session.add(weekly)

    db.session.commit()
    print('Demo data seeded successfully.')
    print(f'  Vehicles: {Vehicle.query.count()}  Drivers: {Driver.query.count()}  Routes: {Route.query.count()}')
    print(f'  Daily transactions: {DailyLog.query.count()}  Fuel logs: {FuelLog.query.count()}  Maintenance logs: {MaintenanceLog.query.count()}')
    print(f'  Franchise vehicles: {FranchiseVehicle.query.count()}')


if __name__ == '__main__':
    with app.app_context():
        cleanup_stray_test_rows()
        seed()
