#!/usr/bin/env python3
"""
Transport Fleet & Finance Management System
T-Tech Solutions | June 2026
"""

import os
import csv
import difflib
import io
import json
import re
import secrets
from datetime import datetime, date, timedelta, timezone
from functools import wraps

import openpyxl
from dotenv import load_dotenv
from flask import (Flask, render_template, request, redirect, url_for,
                   flash, jsonify, make_response)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func

load_dotenv()

# ─────────────────────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────────────────────
IS_PRODUCTION = os.environ.get('FLASK_ENV', 'production').lower() == 'production'

secret_key = os.environ.get('SECRET_KEY')
if not secret_key:
    if IS_PRODUCTION:
        raise RuntimeError(
            'SECRET_KEY environment variable must be set in production. '
            'Set FLASK_ENV=development for local work with an auto-generated key.'
        )
    secret_key = secrets.token_hex(32)
    print('WARNING: SECRET_KEY not set — using a random development key '
          '(sessions will not persist across restarts).')

app = Flask(__name__)
app.config.update(
    SECRET_KEY=secret_key,
    SQLALCHEMY_DATABASE_URI=os.environ.get('DATABASE_URL', 'sqlite:///transport_erp.db'),
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    COMMISSION_DRIVER_RATE=0.15,
    COMMISSION_CONDUCTOR_RATE=0.08,
    VEHICLE_USEFUL_LIFE_YEARS=5,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=IS_PRODUCTION,
)

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'warning'

csrf = CSRFProtect(app)
limiter = Limiter(get_remote_address, app=app, storage_uri='memory://',
                  default_limits=[])


# ─────────────────────────────────────────────────────────────
# Models
# ─────────────────────────────────────────────────────────────
class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='manager')
    is_active = db.Column(db.Boolean, default=True)
    permissions = db.Column(db.Text, default='[]')
    driver_id = db.Column(db.Integer, db.ForeignKey('drivers.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    linked_driver = db.relationship('Driver', foreign_keys=[driver_id])

    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)

    def has_permission(self, perm):
        if self.role == 'admin':
            return True
        try:
            return perm in json.loads(self.permissions or '[]')
        except (json.JSONDecodeError, TypeError):
            return False

    def get_permissions(self):
        try:
            return json.loads(self.permissions or '[]')
        except (json.JSONDecodeError, TypeError):
            return []


class Vehicle(db.Model):
    __tablename__ = 'vehicles'
    id = db.Column(db.Integer, primary_key=True)
    registration = db.Column(db.String(20), unique=True, nullable=False)
    make = db.Column(db.String(50), nullable=False)
    model = db.Column(db.String(50), nullable=False)
    year = db.Column(db.Integer, nullable=False)
    acquisition_cost = db.Column(db.Float, default=0.0)
    status = db.Column(db.String(20), default='active')
    fuel_type = db.Column(db.String(10), default='diesel')
    # Expected daily fare for this vehicle, set by the admin. Null means no
    # target is tracked for this vehicle — it's excluded from shortfall
    # flagging (see report_shortfalls / DailyLog.garnish).
    daily_target = db.Column(db.Float, nullable=True)
    # Insurance is tracked as its own first-class field (not a generic
    # VehicleDocument) since every vehicle needs exactly one current policy
    # and it's the compliance item admins check most often — it gets its
    # own alerting the same way documents do (see insurance_status below).
    insurance_provider = db.Column(db.String(100))
    insurance_policy_number = db.Column(db.String(100))
    insurance_expiry = db.Column(db.Date, nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    documents = db.relationship('VehicleDocument', backref='vehicle',
                                lazy=True, cascade='all, delete-orphan')
    daily_logs = db.relationship('DailyLog', backref='vehicle', lazy=True)
    fuel_logs = db.relationship('FuelLog', backref='vehicle', lazy=True)
    maintenance_logs = db.relationship('MaintenanceLog', backref='vehicle', lazy=True)

    @property
    def total_revenue(self):
        return sum(l.gross_revenue for l in self.daily_logs)

    @property
    def total_fuel_liters(self):
        return sum(l.liters for l in self.fuel_logs)

    @property
    def total_maintenance_cost(self):
        return sum(l.total_cost for l in self.maintenance_logs)

    @property
    def insurance_days_to_expiry(self):
        return (self.insurance_expiry - date.today()).days if self.insurance_expiry else None

    @property
    def insurance_status(self):
        if not self.insurance_expiry:
            return 'none'
        d = self.insurance_days_to_expiry
        if d < 0:
            return 'expired'
        if d <= 30:
            return 'warning'
        return 'valid'


class VehicleDocument(db.Model):
    __tablename__ = 'vehicle_documents'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    doc_type = db.Column(db.String(50), nullable=False)
    reference_number = db.Column(db.String(100))
    issue_date = db.Column(db.Date)
    expiry_date = db.Column(db.Date, nullable=False)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def days_to_expiry(self):
        return (self.expiry_date - date.today()).days

    @property
    def status(self):
        d = self.days_to_expiry
        if d < 0:
            return 'expired'
        if d <= 30:
            return 'warning'
        return 'valid'


class Driver(db.Model):
    __tablename__ = 'drivers'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    # Only drivers need a license — conductors don't drive, so this is
    # optional and only enforced as required at the form level for role='driver'.
    license_number = db.Column(db.String(50), unique=True, nullable=True)
    phone = db.Column(db.String(20))
    role = db.Column(db.String(20), default='driver')
    commission_rate = db.Column(db.Float)
    status = db.Column(db.String(20), default='active')
    # For a conductor, the driver they normally work under — informational,
    # and used to auto-select the conductor when logging a trip for that driver.
    paired_driver_id = db.Column(db.Integer, db.ForeignKey('drivers.id'), nullable=True)
    # The vehicle this driver/conductor is normally assigned to — informational,
    # and used to auto-select the driver/conductor when logging a trip for that vehicle.
    assigned_vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=True)
    next_of_kin_name = db.Column(db.String(100))
    next_of_kin_phone = db.Column(db.String(20))
    next_of_kin_relationship = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    paired_driver = db.relationship('Driver', remote_side=[id], backref='paired_conductors')
    assigned_vehicle = db.relationship('Vehicle', backref='assigned_crew')
    driven_logs = db.relationship('DailyLog', foreign_keys='DailyLog.driver_id',
                                  backref='driver', lazy=True)
    conducted_logs = db.relationship('DailyLog', foreign_keys='DailyLog.conductor_id',
                                     backref='conductor', lazy=True)


class Route(db.Model):
    __tablename__ = 'routes'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    start_point = db.Column(db.String(100), nullable=False)
    end_point = db.Column(db.String(100), nullable=False)
    distance_km = db.Column(db.Float)
    fare_rate = db.Column(db.Float, nullable=False)
    status = db.Column(db.String(20), default='active')
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    logs = db.relationship('DailyLog', backref='route', lazy=True)


class DailyLog(db.Model):
    __tablename__ = 'daily_logs'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    # Nullable so an activity/fare can be logged before a driver is known or
    # assigned — see driver_ledger_add. Garnish still requires a driver since
    # it nets against a specific person's commission.
    driver_id = db.Column(db.Integer, db.ForeignKey('drivers.id'), nullable=True)
    conductor_id = db.Column(db.Integer, db.ForeignKey('drivers.id'))
    # Nullable because Daily Transactions (the vehicle ledger) doesn't
    # collect a route per entry — it mirrors a plain date/driver/fare/
    # diesel/mileage logbook. Left over from a route-per-trip form that
    # has since been merged into this page; route_profitability reports
    # simply see None for rows entered here.
    route_id = db.Column(db.Integer, db.ForeignKey('routes.id'), nullable=True)
    log_date = db.Column(db.Date, nullable=False, default=date.today)
    trips_completed = db.Column(db.Integer, default=0)
    gross_revenue = db.Column(db.Float, nullable=False, default=0.0)
    # Garnished from the driver/conductor's commission for this day — typically
    # because they fell short of the admin-set revenue target and the admin
    # decided not to pay out the full percentage. See reason_for_shortfall.
    # Netted against commission in the payroll report.
    garnish = db.Column(db.Float, nullable=False, default=0.0)
    reason_for_shortfall = db.Column(db.Text)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    creator = db.relationship('User', foreign_keys=[created_by])
    updater = db.relationship('User', foreign_keys=[updated_by])


class FuelLog(db.Model):
    __tablename__ = 'fuel_logs'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    log_date = db.Column(db.Date, nullable=False, default=date.today)
    liters = db.Column(db.Float, nullable=False)
    # Fuel cost is no longer tracked by the app (liters/distance only) — these
    # two columns stay on the model, defaulted to 0, purely so existing
    # production databases (no migration tool in this project) don't reject
    # new inserts on their pre-existing NOT NULL constraint.
    cost_per_liter = db.Column(db.Float, nullable=False, default=0.0)
    total_cost = db.Column(db.Float, nullable=False, default=0.0)
    odometer = db.Column(db.Float)
    supplier = db.Column(db.String(100))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    creator = db.relationship('User', foreign_keys=[created_by])


class MaintenanceLog(db.Model):
    __tablename__ = 'maintenance_logs'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    log_date = db.Column(db.Date, nullable=False, default=date.today)
    description = db.Column(db.Text, nullable=False)
    parts_cost = db.Column(db.Float, default=0.0)
    labor_cost = db.Column(db.Float, default=0.0)
    total_cost = db.Column(db.Float, nullable=False)
    mechanic = db.Column(db.String(100))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    creator = db.relationship('User', foreign_keys=[created_by])


class AuditLog(db.Model):
    __tablename__ = 'audit_logs'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    action = db.Column(db.String(20), nullable=False)
    table_name = db.Column(db.String(50))
    record_id = db.Column(db.Integer)
    description = db.Column(db.Text)
    ip_address = db.Column(db.String(50))
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    user = db.relationship('User')


class ImportBatch(db.Model):
    """One row per committed file import (ledger or franchise workbook). This
    is the structured audit trail for imports — separate from the generic
    AuditLog above — because it needs to carry the quarantined error rows
    and link to the records it created, so a bad import can be identified
    and reversed via ImportBatchRecord."""
    __tablename__ = 'import_batches'
    id = db.Column(db.Integer, primary_key=True)
    target_type = db.Column(db.String(30), nullable=False)   # 'ledger' | 'franchise_workbook'
    filename = db.Column(db.String(255), nullable=False)
    total_rows = db.Column(db.Integer, default=0)
    rows_imported = db.Column(db.Integer, default=0)
    rows_failed = db.Column(db.Integer, default=0)
    error_rows = db.Column(db.Text)   # JSON list of {**original_row_columns, 'System_Error': msg}
    status = db.Column(db.String(15), nullable=False, default='committed')  # 'committed' | 'reverted'
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    reverted_at = db.Column(db.DateTime)
    reverted_by = db.Column(db.Integer, db.ForeignKey('users.id'))

    creator = db.relationship('User', foreign_keys=[created_by])
    reverter = db.relationship('User', foreign_keys=[reverted_by])
    records = db.relationship('ImportBatchRecord', backref='batch', lazy=True,
                               cascade='all, delete-orphan')

    @property
    def error_row_list(self):
        return json.loads(self.error_rows) if self.error_rows else []


class ImportBatchRecord(db.Model):
    """Links an ImportBatch to the actual row it created, so a revert can
    delete exactly those rows instead of guessing by timestamp."""
    __tablename__ = 'import_batch_records'
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.Integer, db.ForeignKey('import_batches.id'), nullable=False)
    target_table = db.Column(db.String(30), nullable=False)   # 'daily_logs' | 'fuel_logs' | 'franchise_vehicles'
    record_id = db.Column(db.Integer, nullable=False)


# ─────────────────────────────────────────────────────────────
# Finance: loans, payables, receivables, capital, expenses, budget
# ─────────────────────────────────────────────────────────────
class Loan(db.Model):
    __tablename__ = 'loans'
    id = db.Column(db.Integer, primary_key=True)
    lender = db.Column(db.String(100), nullable=False)
    principal = db.Column(db.Float, nullable=False)
    interest_rate = db.Column(db.Float, default=0.0)
    start_date = db.Column(db.Date, nullable=False)
    term_months = db.Column(db.Integer)
    status = db.Column(db.String(20), default='active')
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    payments = db.relationship('LoanPayment', backref='loan', lazy=True,
                               cascade='all, delete-orphan')

    @property
    def total_repaid(self):
        return sum(p.amount for p in self.payments)

    @property
    def outstanding_balance(self):
        return self.principal - self.total_repaid


class LoanPayment(db.Model):
    __tablename__ = 'loan_payments'
    id = db.Column(db.Integer, primary_key=True)
    loan_id = db.Column(db.Integer, db.ForeignKey('loans.id'), nullable=False)
    payment_date = db.Column(db.Date, nullable=False, default=date.today)
    amount = db.Column(db.Float, nullable=False)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class Payable(db.Model):
    """Accounts payable — amounts owed to suppliers OUTSIDE the fuel/maintenance
    logs (e.g. an insurance premium invoice), tracked on an accrual basis:
    the expense is recognized when the payable is created, not when it's paid."""
    __tablename__ = 'payables'
    id = db.Column(db.Integer, primary_key=True)
    supplier_name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    amount = db.Column(db.Float, nullable=False)
    invoice_date = db.Column(db.Date, nullable=False, default=date.today)
    due_date = db.Column(db.Date)
    status = db.Column(db.String(20), default='unpaid')
    paid_date = db.Column(db.Date)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class Receivable(db.Model):
    """Accounts receivable — revenue earned but not yet collected (e.g. an
    invoiced corporate/charter client), tracked on an accrual basis: revenue
    is recognized when the receivable is created, not when it's collected."""
    __tablename__ = 'receivables'
    id = db.Column(db.Integer, primary_key=True)
    client_name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    amount = db.Column(db.Float, nullable=False)
    invoice_date = db.Column(db.Date, nullable=False, default=date.today)
    due_date = db.Column(db.Date)
    status = db.Column(db.String(20), default='outstanding')
    collected_date = db.Column(db.Date)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class CommissionPayment(db.Model):
    """An actual cash payout of driver/conductor commission. Kept separate
    from the accrued commission figure (computed live from revenue, as the
    payroll report already does) so the two can be reconciled — accrued
    minus paid is the outstanding commission liability."""
    __tablename__ = 'commission_payments'
    id = db.Column(db.Integer, primary_key=True)
    driver_id = db.Column(db.Integer, db.ForeignKey('drivers.id'), nullable=False)
    payment_date = db.Column(db.Date, nullable=False, default=date.today)
    amount = db.Column(db.Float, nullable=False)
    period_start = db.Column(db.Date)
    period_end = db.Column(db.Date)
    method = db.Column(db.String(30))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    driver = db.relationship('Driver')


class CapitalContribution(db.Model):
    __tablename__ = 'capital_contributions'
    id = db.Column(db.Integer, primary_key=True)
    contributor = db.Column(db.String(100), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    contribution_date = db.Column(db.Date, nullable=False, default=date.today)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class OwnerDrawing(db.Model):
    __tablename__ = 'owner_drawings'
    id = db.Column(db.Integer, primary_key=True)
    amount = db.Column(db.Float, nullable=False)
    drawing_date = db.Column(db.Date, nullable=False, default=date.today)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


class ExpenseCategory(db.Model):
    """Two-level expense classification: a top-level heading (parent_id is
    None, e.g. "Maintenance") with optional sub-headings under it (e.g.
    "Engine Oil"). An Expense can be tagged to either a heading or a
    sub-heading — the hierarchy is for organization, not enforcement."""
    __tablename__ = 'expense_categories'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(60), nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey('expense_categories.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    parent = db.relationship('ExpenseCategory', remote_side=[id], backref='children')

    @property
    def display_name(self):
        return f'{self.parent.name} — {self.name}' if self.parent else self.name


class Expense(db.Model):
    __tablename__ = 'expenses'
    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.Integer, db.ForeignKey('expense_categories.id'), nullable=False)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=True)
    expense_date = db.Column(db.Date, nullable=False, default=date.today)
    description = db.Column(db.Text)
    amount = db.Column(db.Float, nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    category = db.relationship('ExpenseCategory', backref='expenses')
    vehicle = db.relationship('Vehicle')


class Budget(db.Model):
    __tablename__ = 'budgets'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(60), nullable=False)
    month = db.Column(db.Date, nullable=False)
    amount = db.Column(db.Float, nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))


# ─────────────────────────────────────────────────────────────
# Franchise Income: a revenue stream separate from vehicle operations.
# Daily-fee and weekly-fee franchisees are two independent obligations, so
# they get two independent entities — never a shared row — each with its
# own income, expenditure breakdown, and cash reconciliation.
# ─────────────────────────────────────────────────────────────
class FranchiseDailyIncome(db.Model):
    """One reconciliation record per calendar date per franchise vehicle for
    daily franchise fee collections — a date can have several entries, one
    per vehicle, since each vehicle's standalone fee is reconciled
    independently. vehicle_id may be null for a whole-franchise entry not
    attributable to one vehicle (e.g. a historical combined-total figure).
    Cash-in-hand, total expenditure, and variance are derived rather than
    stored, so they can never drift out of sync with the entered figures."""
    __tablename__ = 'franchise_daily_income'
    __table_args__ = (db.UniqueConstraint('entry_date', 'vehicle_id', name='uq_franchise_daily_income_date_vehicle'),)
    id = db.Column(db.Integer, primary_key=True)
    entry_date = db.Column(db.Date, nullable=False)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('franchise_vehicles.id'), nullable=True)

    income = db.Column(db.Float, nullable=False, default=0)
    exp_traffic_fines = db.Column(db.Float, nullable=False, default=0)
    exp_facilitation_fees = db.Column(db.Float, nullable=False, default=0)
    exp_workshop = db.Column(db.Float, nullable=False, default=0)
    exp_wages = db.Column(db.Float, nullable=False, default=0)
    other_expenditure = db.Column(db.Float, nullable=False, default=0)

    deposited = db.Column(db.Float, nullable=False, default=0)
    description = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def total_expenditure(self):
        return (self.exp_traffic_fines or 0) + (self.exp_facilitation_fees or 0) \
            + (self.exp_workshop or 0) + (self.exp_wages or 0) + (self.other_expenditure or 0)

    @property
    def cash_in_hand(self):
        return (self.income or 0) - self.total_expenditure

    @property
    def variance(self):
        return (self.deposited or 0) - self.cash_in_hand

    vehicle = db.relationship('FranchiseVehicle')


class FranchiseWeeklyIncome(db.Model):
    """Same shape as FranchiseDailyIncome, but one record per week
    (week_start = that week's Monday) per franchise vehicle for the
    separate weekly franchise fee — kept as its own entity rather than
    extra columns on the same date row, since a vehicle's daily and weekly
    dues are independent obligations, not two figures on one entry.
    vehicle_id may be null for a whole-franchise entry, same as above."""
    __tablename__ = 'franchise_weekly_income'
    __table_args__ = (db.UniqueConstraint('week_start', 'vehicle_id', name='uq_franchise_weekly_income_week_vehicle'),)
    id = db.Column(db.Integer, primary_key=True)
    week_start = db.Column(db.Date, nullable=False)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('franchise_vehicles.id'), nullable=True)

    income = db.Column(db.Float, nullable=False, default=0)
    exp_traffic_fines = db.Column(db.Float, nullable=False, default=0)
    exp_facilitation_fees = db.Column(db.Float, nullable=False, default=0)
    exp_workshop = db.Column(db.Float, nullable=False, default=0)
    exp_wages = db.Column(db.Float, nullable=False, default=0)
    other_expenditure = db.Column(db.Float, nullable=False, default=0)

    deposited = db.Column(db.Float, nullable=False, default=0)
    description = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def total_expenditure(self):
        return (self.exp_traffic_fines or 0) + (self.exp_facilitation_fees or 0) \
            + (self.exp_workshop or 0) + (self.exp_wages or 0) + (self.other_expenditure or 0)

    @property
    def cash_in_hand(self):
        return (self.income or 0) - self.total_expenditure

    @property
    def variance(self):
        return (self.deposited or 0) - self.cash_in_hand

    vehicle = db.relationship('FranchiseVehicle')


class FranchiseVehicle(db.Model):
    """A third-party vehicle paying to operate under the franchise — distinct
    from the company's own fleet (Vehicle above), which is operated directly
    rather than franchised out."""
    __tablename__ = 'franchise_vehicles'
    id = db.Column(db.Integer, primary_key=True)
    number_plate = db.Column(db.String(20), nullable=False, unique=True)
    franchisee_name = db.Column(db.String(100), nullable=False)
    status = db.Column(db.String(20), default='active')
    # The standalone fee this vehicle agreed to pay (e.g. $10/day) — kept on
    # the vehicle rather than derived, since it's a negotiated figure, not
    # something computed from collections.
    agreed_amount = db.Column(db.Float, nullable=False, default=0)
    # Running arrears balance — tracked manually rather than computed,
    # since "how much is owed" depends on a payment schedule this system
    # doesn't model (which days were actually due).
    amount_owed = db.Column(db.Float, nullable=False, default=0)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    collections = db.relationship('FranchiseCollection', backref='vehicle', lazy=True,
                                  order_by='FranchiseCollection.entry_date.desc()')

    @property
    def total_collected(self):
        return sum(c.amount for c in self.collections)

    @property
    def total_expense(self):
        return sum(c.expense for c in self.collections)

    @property
    def net_collected(self):
        return self.total_collected - self.total_expense


class FranchiseCollection(db.Model):
    """What one franchise vehicle paid on one date — the per-vehicle detail
    behind a day's franchise income total. Kept separate from the
    daily/weekly income entities (rather than forcing a sum-to-match)
    because in practice the two are reconciled by a person, not guaranteed
    to tie out line-for-line.

    frequency lives here rather than on FranchiseVehicle because the same
    vehicle can owe both a daily due and a separate weekly franchise fee —
    it's a property of the payment, not a fixed plan per vehicle.

    expense is that same vehicle's own cost for that day/week (fuel,
    fines, etc. it bears itself) — kept alongside amount so each
    franchisee's collection entry nets out to what it actually owes,
    without touching the company-wide franchise income expense figures."""
    __tablename__ = 'franchise_collections'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('franchise_vehicles.id'), nullable=False)
    entry_date = db.Column(db.Date, nullable=False)
    frequency = db.Column(db.String(10), nullable=False, default='daily')  # 'daily' or 'weekly'
    amount = db.Column(db.Float, nullable=False)
    expense = db.Column(db.Float, nullable=False, default=0)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def net(self):
        return self.amount - (self.expense or 0)


class MaintenanceSchedule(db.Model):
    __tablename__ = 'maintenance_schedules'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    description = db.Column(db.String(150), nullable=False)
    interval_days = db.Column(db.Integer)
    interval_km = db.Column(db.Float)
    last_done_date = db.Column(db.Date)
    last_done_odometer = db.Column(db.Float)
    next_due_date = db.Column(db.Date)
    next_due_odometer = db.Column(db.Float)
    status = db.Column(db.String(20), default='active')
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    vehicle = db.relationship('Vehicle')


# ─────────────────────────────────────────────────────────────
# Spares Store: parts inventory, purchases & marked-up sales
# ─────────────────────────────────────────────────────────────
class SparePart(db.Model):
    __tablename__ = 'spare_parts'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    part_number = db.Column(db.String(50))
    unit = db.Column(db.String(20), nullable=False, default='pc')
    # Running weighted-average unit cost, updated on every purchase — not
    # editable directly, it's a derived stock-valuation figure.
    cost_price = db.Column(db.Float, nullable=False, default=0.0)
    markup_percent = db.Column(db.Float, nullable=False, default=0.0)
    quantity_on_hand = db.Column(db.Integer, nullable=False, default=0)
    reorder_level = db.Column(db.Integer, nullable=False, default=0)
    status = db.Column(db.String(20), default='active')
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    purchases = db.relationship('StorePurchase', backref='part', lazy=True,
                                cascade='all, delete-orphan')
    sales = db.relationship('StoreSale', backref='part', lazy=True,
                            cascade='all, delete-orphan')

    @property
    def selling_price(self):
        return round(self.cost_price * (1 + self.markup_percent / 100), 2)

    @property
    def stock_value(self):
        return self.cost_price * self.quantity_on_hand

    @property
    def low_stock(self):
        return self.quantity_on_hand <= self.reorder_level


class StorePurchase(db.Model):
    """A restock — adds quantity to the part and rolls the new cost into
    the part's weighted-average cost_price."""
    __tablename__ = 'store_purchases'
    id = db.Column(db.Integer, primary_key=True)
    part_id = db.Column(db.Integer, db.ForeignKey('spare_parts.id'), nullable=False)
    purchase_date = db.Column(db.Date, nullable=False, default=date.today)
    quantity = db.Column(db.Integer, nullable=False)
    unit_cost = db.Column(db.Float, nullable=False)
    total_cost = db.Column(db.Float, nullable=False)
    supplier = db.Column(db.String(100))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    creator = db.relationship('User', foreign_keys=[created_by])


class StoreSale(db.Model):
    """A sale at the part's marked-up selling price. Cost and price are
    snapshotted at sale time so historical profit doesn't shift when the
    part's cost or markup later changes."""
    __tablename__ = 'store_sales'
    id = db.Column(db.Integer, primary_key=True)
    part_id = db.Column(db.Integer, db.ForeignKey('spare_parts.id'), nullable=False)
    # Set when this sale is to one of the fleet's own vehicles rather than an
    # outside customer — the sale amount then also counts as an expense on
    # that vehicle's income statement (see vehicle_income_totals). Mutually
    # exclusive with customer_name in practice, not enforced at the DB level.
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=True)
    sale_date = db.Column(db.Date, nullable=False, default=date.today)
    quantity = db.Column(db.Integer, nullable=False)
    unit_cost = db.Column(db.Float, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    total_amount = db.Column(db.Float, nullable=False)
    customer_name = db.Column(db.String(100))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    creator = db.relationship('User', foreign_keys=[created_by])
    vehicle = db.relationship('Vehicle')

    @property
    def profit(self):
        return (self.unit_price - self.unit_cost) * self.quantity

    @property
    def customer_display(self):
        return self.vehicle.registration if self.vehicle else (self.customer_name or None)


# ─────────────────────────────────────────────────────────────
# Auth helpers
# ─────────────────────────────────────────────────────────────
@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


PERMISSIONS = {
    'dashboard':    'Dashboard — overview stats, charts & recent activity',
    'crew_portal':  'Crew Portal — log daily income & view team performance leaderboard',
    'vehicles':     'Vehicles — view, add & edit vehicles',
    'drivers':      'Crew — view, add & edit drivers/conductors',
    'routes':       'Routes — view, add & edit routes',
    'daily_logs':   'Daily Transactions — view, record & edit vehicle transactions',
    'fuel_logs':    'Fuel Logs — view & record fuel entries',
    'maintenance':  'Maintenance — view & record maintenance logs',
    'reports':      'Finance & Reports — income statement, payroll, CSV exports',
    'compliance':   'Compliance — vehicle documents & expiry tracker',
    'finance':      'Finance Ledger — loans, payables, receivables, capital, expenses, budget',
    'store':        'Spares Store — parts inventory, purchases & marked-up sales',
    'franchise':    'Franchise Income — collection reconciliation entry & franchise P&L statements',
}

PERMISSION_REDIRECTS = [
    ('dashboard',   'dashboard'),
    ('crew_portal', 'crew_leaderboard'),
    ('vehicles',    'vehicles'),
    ('drivers',     'drivers'),
    ('routes',      'routes_list'),
    ('daily_logs',  'driver_ledger'),
    ('fuel_logs',   'fuel_logs'),
    ('maintenance', 'maintenance_logs'),
    ('reports',     'report_income'),
    ('compliance',  'compliance'),
    ('finance',     'loans_list'),
    ('store',       'store_parts'),
    ('franchise',   'franchise_daily_income_list'),
]


def first_permitted_url(user):
    if user.role == 'admin':
        return url_for('dashboard')
    for perm, endpoint in PERMISSION_REDIRECTS:
        if user.has_permission(perm):
            return url_for(endpoint)
    return url_for('no_access')


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            flash('Admin access required.', 'danger')
            return redirect(first_permitted_url(current_user))
        return f(*args, **kwargs)
    return decorated


def permission_required(perm):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if not current_user.has_permission(perm):
                flash('You do not have permission to access that section.', 'danger')
                return redirect(first_permitted_url(current_user))
            return f(*args, **kwargs)
        return decorated
    return decorator


def permission_required_any(*perms):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if not any(current_user.has_permission(p) for p in perms):
                flash('You do not have permission to access that section.', 'danger')
                return redirect(first_permitted_url(current_user))
            return f(*args, **kwargs)
        return decorated
    return decorator


def log_audit(action, table_name=None, record_id=None, description=None):
    entry = AuditLog(
        user_id=current_user.id if current_user.is_authenticated else None,
        action=action,
        table_name=table_name,
        record_id=record_id,
        description=description,
        ip_address=request.remote_addr,
    )
    db.session.add(entry)


def save_import_batch(target_type, filename, total_rows, imported, error_rows, created_records):
    """Record a committed file import as a structured ImportBatch (unlike
    log_audit's free-text trail, this carries the quarantined error rows and
    links to every record it created via ImportBatchRecord, so the import
    can be found later, its errors re-downloaded, and it can be reversed).
    Call right before db.session.commit() so the batch and the data it
    describes land in the same transaction."""
    batch = ImportBatch(
        target_type=target_type, filename=filename, total_rows=total_rows,
        rows_imported=imported, rows_failed=len(error_rows),
        error_rows=json.dumps(error_rows) if error_rows else None,
        created_by=current_user.id,
    )
    db.session.add(batch)
    db.session.flush()
    for table_name, record_id in created_records:
        db.session.add(ImportBatchRecord(batch_id=batch.id, target_table=table_name, record_id=record_id))
    return batch


def parse_date(s):
    try:
        return datetime.strptime(s, '%Y-%m-%d').date() if s else None
    except ValueError:
        raise ValueError(f'"{s}" is not a valid date (expected YYYY-MM-DD).')


def parse_import_date(value):
    """Parse a date cell from an uploaded CSV/Excel row. Excel cells come
    through openpyxl as datetime objects already; CSV cells are plain
    strings, tried against the common formats a logbook might use."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        raise ValueError('Date is required.')
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f'"{s}" is not a recognized date (use YYYY-MM-DD).')


def parse_import_number(value, label):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        raise ValueError(f'{label} "{s}" is not a number.')


MAX_LEDGER_IMPORT_ROWS = 2000

# Canonical ledger import fields and the header synonyms an uploaded file might
# use for each — used to auto-map an arbitrary file's columns onto the shape
# the importer expects, instead of requiring an exact header match.
CANONICAL_LEDGER_FIELDS = [
    ('date', 'Date', ['date', 'log date', 'trip date', 'day']),
    ('driver', 'Driver', ['driver', 'driver name', 'crew', 'operator']),
    ('fare', 'Fare', ['fare', 'revenue', 'income', 'gross revenue', 'collection',
                      'collections', 'takings']),
    ('diesel_cost', 'Diesel (USD)', ['diesel cost', 'diesel usd', 'diesel amount',
                                     'petrol cost', 'petrol usd', 'petrol amount',
                                     'fuel cost', 'fuel usd', 'fuel amount',
                                     'diesel', 'petrol', 'fuel']),
    ('mileage', 'Mileage', ['mileage', 'odometer', 'odo', 'distance reading', 'km reading']),
]

# The raw-header key each canonical field must land on so the existing
# row-validation loop (which looks for these exact keys) keeps working unchanged.
CANONICAL_TO_ROW_KEY = {
    'date': 'date', 'driver': 'driver', 'fare': 'fare',
    'diesel_cost': 'diesel cost', 'mileage': 'mileage',
}

def _find_header_row(raw_rows, max_scan=10):
    """Locate the header row within the first few rows of an uploaded sheet,
    skipping any title/blank rows some logbooks put above the real columns
    (e.g. a "JULY 2026 DAILY TRANSACTIONS" banner row)."""
    date_synonyms = next(syns for key, _label, syns in CANONICAL_LEDGER_FIELDS if key == 'date')
    for idx, row in enumerate(raw_rows[:max_scan]):
        cells = [str(c).strip().lower() for c in row if c not in (None, '')]
        if len(cells) < 2:
            continue
        if any(c in date_synonyms or difflib.get_close_matches(c, date_synonyms, n=1, cutoff=0.75)
               for c in cells):
            return idx
    return 0


def _json_safe_cell(cell):
    # Excel date cells come through openpyxl as datetime/date objects,
    # which aren't JSON-serializable — the row data gets round-tripped
    # through a hidden form field, so normalize to the 'YYYY-MM-DD'
    # string parse_import_date already accepts.
    if isinstance(cell, datetime):
        return cell.date().isoformat()
    if isinstance(cell, date):
        return cell.isoformat()
    return cell


def _rows_from_raw(raw_rows):
    """Turn a sheet/CSV's raw rows into (headers, rows), locating the real
    header row (skipping any title/blank rows above it) and keying each row
    by its stripped, lowercased header."""
    header_idx = _find_header_row(raw_rows)
    headers = [str(h).strip().lower() if h is not None else '' for h in raw_rows[header_idx]]
    rows = [{h: _json_safe_cell(c) for h, c in zip(headers, r)}
            for r in raw_rows[header_idx + 1:] if any(c not in (None, '') for c in r)]
    return [h for h in headers if h], rows


def _select_sheet_for_vehicle(wb, vehicle):
    """Pick the right worksheet for a single-vehicle import. A single-sheet
    file is unambiguous, but a multi-sheet workbook (e.g. a fleet-wide master
    logbook) must NOT silently fall back to whichever sheet happened to be
    "active" when it was last saved — that would import a different vehicle's
    transactions under the one currently selected. Only proceed if a sheet's
    name matches the selected vehicle's registration; otherwise raise so the
    upload is rejected instead of misattributed."""
    if vehicle is None or len(wb.sheetnames) == 1:
        return wb.active
    target = _normalize_registration(vehicle.registration)
    for name in wb.sheetnames:
        if _normalize_registration(name) == target:
            return wb[name]
    raise ValueError(
        f'This workbook has {len(wb.sheetnames)} sheets and none is named "{vehicle.registration}" — '
        'to avoid importing the wrong vehicle\'s data, upload a file with just one sheet, or use '
        '"Import Fleet Workbook" below to import every vehicle\'s sheet at once.')


def read_uploaded_table(file, vehicle=None):
    """Parse an uploaded CSV/XLSX file into (headers, rows). Each row is a dict
    keyed by its stripped, lowercased header. Raises ValueError on an
    unsupported or unreadable file."""
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    try:
        if ext == 'csv':
            reader = csv.reader(io.StringIO(file.read().decode('utf-8-sig')))
            raw_rows = list(reader)
        elif ext in ('xlsx', 'xlsm'):
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = _select_sheet_for_vehicle(wb, vehicle)
            raw_rows = [list(r) for r in ws.iter_rows(values_only=True)]
        else:
            raise ValueError('Unsupported file type — upload a .csv or .xlsx file.')
        if not raw_rows:
            raise ValueError('The file is empty.')
    except ValueError:
        raise
    except Exception:
        raise ValueError('Could not read that file. Make sure it is a valid CSV or Excel export.')

    headers, rows = _rows_from_raw(raw_rows)
    if len(rows) > MAX_LEDGER_IMPORT_ROWS:
        raise ValueError(f'That file has {len(rows)} rows — the importer previews at most '
                         f'{MAX_LEDGER_IMPORT_ROWS} at a time. Split it into smaller files.')
    return headers, rows


def read_uploaded_workbook_sheets(file):
    """Parse a multi-sheet XLSX/XLSM workbook into {sheet_name: (headers, rows)}
    for every sheet that has data — lets a fleet-wide workbook (one tab per
    vehicle, like a company's master logbook) be imported in a single pass.
    Raises ValueError on an unsupported or unreadable file."""
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext not in ('xlsx', 'xlsm'):
        raise ValueError('Upload a .xlsx or .xlsm workbook — one sheet per vehicle.')
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        raise ValueError('Could not read that file. Make sure it is a valid Excel workbook.')

    sheets = {}
    for ws in wb.worksheets:
        raw_rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not raw_rows:
            continue
        headers, rows = _rows_from_raw(raw_rows)
        if rows:
            sheets[ws.title] = (headers, rows[:MAX_LEDGER_IMPORT_ROWS])
    return sheets


def _normalize_registration(name):
    return re.sub(r'\s+', ' ', str(name or '')).strip().upper()


def auto_map_columns(headers, fields=None):
    """Guess which uploaded header feeds each canonical field (ledger fields by
    default; pass a different fields list for another import shape), by exact
    synonym match first and fuzzy match second. Returns {field_key: header_or_None}."""
    fields = fields if fields is not None else CANONICAL_LEDGER_FIELDS
    mapping = {}
    claimed = set()
    for field_key, _label, synonyms in fields:
        match = next((h for h in headers if h in synonyms and h not in claimed), None)
        if not match:
            candidates = [h for h in headers if h not in claimed]
            # word-boundary substring match first, e.g. "daily fare" contains "fare"
            match = next((h for h in candidates
                          if any(re.search(rf'\b{re.escape(s)}\b', h) for s in synonyms)), None)
        if not match:
            close = difflib.get_close_matches(synonyms[0], candidates, n=1, cutoff=0.6)
            if not close:
                for syn in synonyms[1:]:
                    close = difflib.get_close_matches(syn, candidates, n=1, cutoff=0.6)
                    if close:
                        break
            match = close[0] if close else None
        mapping[field_key] = match
        if match:
            claimed.add(match)
    return mapping


def apply_column_mapping(headers, raw_rows, mapping, row_key_map=None):
    """Rebuild each raw row (keyed by uploaded header) into the row shape an
    import loop expects (keyed by canonical field name), per `mapping`
    ({field_key: header_or_None}). row_key_map defaults to the ledger's."""
    row_key_map = row_key_map if row_key_map is not None else CANONICAL_TO_ROW_KEY
    mapped_rows = []
    for raw in raw_rows:
        row = {}
        for field_key, header in mapping.items():
            row_key = row_key_map[field_key]
            row[row_key] = raw.get(header) if header else None
        mapped_rows.append(row)
    return mapped_rows


def import_ledger_rows(file_rows, vehicle, auto_register_drivers=False):
    """Validate and persist already-mapped ledger rows (keyed by 'date',
    'driver', 'fare', 'diesel cost', 'mileage') as DailyLog/FuelLog entries
    for `vehicle`. Returns (imported_count, error_messages, created_driver_names);
    does not commit — the caller decides when to commit/rollback.

    If auto_register_drivers is True, a fare row naming a driver that isn't
    on file registers a new active Driver instead of erroring — used by the
    fleet-wide bulk import, where dozens of real driver names showing up for
    the first time is the expected case, not a data-entry mistake worth
    blocking on. The single-vehicle import leaves this off, since there a
    human already reviewed the file and can add the driver deliberately.

    created_records is a list of (target_table, record_id) tuples for every
    DailyLog/FuelLog row created — used to build an ImportBatchRecord trail
    so the import can be identified and reversed later. error_rows mirrors
    errors but keeps the original row data plus a 'System_Error' column, for
    a downloadable quarantine CSV."""
    driver_by_name = {d.name.strip().lower(): d for d in
                       Driver.query.filter_by(role='driver', status='active').all()}
    created_drivers = []
    created_records = []
    last_driver = None  # carries forward across blank-driver rows, see below

    imported = 0
    errors = []
    error_rows = []
    for i, raw_row in enumerate(file_rows, start=2):  # row 1 is the header
        row = {(k or '').strip().lower(): v for k, v in raw_row.items()}
        try:
            date_raw = row.get('date')
            if date_raw in (None, ''):
                continue
            log_date = parse_import_date(date_raw)

            driver_name = str(row.get('driver') or '').strip()
            driver = driver_by_name.get(driver_name.lower()) if driver_name else None

            cost_key = next((k for k in ('diesel cost', 'petrol cost', 'fuel cost') if k in row), 'diesel cost')
            fare = parse_import_number(row.get('fare'), 'Fare')
            diesel_cost = parse_import_number(row.get(cost_key), 'Diesel (USD)')
            mileage = parse_import_number(row.get('mileage'), 'Mileage')

            # Real-world logbooks often put a note ("GARAGE", "ARRESTED", "DRIVER
            # NOT FEELING WELL") in the driver column on no-fare days — only
            # require a recognized driver when there's actual revenue to attribute.
            if fare and not driver:
                if not driver_name and last_driver is not None:
                    # Some logbooks only write the driver's name once and leave
                    # it blank on the following days to mean "same as above".
                    driver = last_driver
                elif driver_name and auto_register_drivers:
                    driver = Driver(name=driver_name, role='driver', status='active')
                    db.session.add(driver)
                    db.session.flush()
                    driver_by_name[driver_name.lower()] = driver
                    created_drivers.append(driver_name)
                else:
                    raise ValueError(f'unknown driver "{driver_name}".' if driver_name else 'fare needs a driver name.')
            if fare is None and diesel_cost is None and mileage is None:
                continue

            if driver_name and driver:
                last_driver = driver

            if fare:
                conductor = driver.paired_conductors[0] if driver.paired_conductors else None
                log = DailyLog(
                    vehicle_id=vehicle.id, driver_id=driver.id,
                    conductor_id=conductor.id if conductor else None,
                    log_date=log_date, gross_revenue=fare, created_by=current_user.id,
                )
                db.session.add(log)
                db.session.flush()
                created_records.append(('daily_logs', log.id))
            if diesel_cost is not None:
                fuel = FuelLog(
                    vehicle_id=vehicle.id, log_date=log_date, liters=0,
                    total_cost=diesel_cost, odometer=mileage, created_by=current_user.id,
                )
                db.session.add(fuel)
                db.session.flush()
                created_records.append(('fuel_logs', fuel.id))
            elif mileage is not None:
                fuel = FuelLog(
                    vehicle_id=vehicle.id, log_date=log_date, liters=0,
                    odometer=mileage, created_by=current_user.id,
                )
                db.session.add(fuel)
                db.session.flush()
                created_records.append(('fuel_logs', fuel.id))
            imported += 1
        except ValueError as e:
            errors.append(f'Row {i}: {e}')
            error_rows.append({**row, 'System_Error': str(e)})

    return imported, errors, error_rows, created_drivers, created_records


# ─────────────────────────────────────────────────────────────
# Franchise Workbook Import — reads a whole multi-sheet Excel workbook (like
# the franchise's own monthly file) and finds every vehicle x date
# collection grid in it, by recognizing each table's own headers rather
# than requiring one flat header row per sheet.
# ─────────────────────────────────────────────────────────────
MAX_WORKBOOK_IMPORT_CELLS = 500000

# Row labels from the franchise's own "Weekly Analysis" recap block — seen
# while hunting for vehicle-collection grids, these confirm we've wandered
# into that summary table rather than a real per-vehicle grid, so it's
# skipped rather than mis-imported as vehicles named "Net Profit".
_FRANCHISE_ANALYSIS_LABELS = {
    'daily franchise', 'weekly franchise', 'total income', 'less expenses',
    'bridge', 'morning tickets', 'monday payments', 'tuesday payments',
    'net profit', 'total net profit', 'day', 'date', 'total',
}


def _vehicle_matrix_blocks(ws):
    """Find every vehicle x date collection grid in a worksheet: a row of
    ≥3 consecutive dates starting at column B, above rows whose column A
    holds a vehicle plate/name and whose other columns hold amounts paid on
    the matching date. Frequency ('daily' vs 'weekly') is read from the
    nearest section label above the grid (e.g. "FRANCHISE DAILY
    COLLECTIONS"), falling back to the sheet title, then 'daily'."""
    blocks = []
    for row in ws.iter_rows():
        date_cols = []
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=row[0].row, column=c).value
            if isinstance(v, (datetime, date)):
                date_cols.append(c)
            elif date_cols:
                break
        if len(date_cols) < 3:
            continue
        # Confirm this is really a vehicle grid, not the "Weekly Analysis"
        # recap block (which also has a row of ~7 dates above summary labels).
        next_a = ws.cell(row=row[0].row + 1, column=1).value
        if next_a is None or str(next_a).strip().lower() in _FRANCHISE_ANALYSIS_LABELS:
            continue

        label = ''
        for lookback in range(1, 15):
            v = ws.cell(row=row[0].row - lookback, column=1).value
            if isinstance(v, str) and v.strip():
                label = v.strip().upper()
                break
        if 'DAILY' in label:
            frequency = 'daily'
        elif 'WEEKLY' in label:
            frequency = 'weekly'
        else:
            title = ws.cell(row=1, column=1).value
            frequency = 'weekly' if isinstance(title, str) and 'WEEKLY' in title.upper() else 'daily'

        blocks.append(dict(date_row=row[0].row, date_cols=date_cols,
                           data_start=row[0].row + 1, frequency=frequency))
    return blocks


_PLATE_RE = re.compile(r'^([A-Z]{2,3}\s?\d{3,5})\s*(.*)$')


def _split_plate_name(raw):
    raw = (raw or '').strip()
    m = _PLATE_RE.match(raw.upper())
    if not m:
        return raw.upper(), (raw.upper().title() or '(unnamed)')
    plate = re.sub(r'\s+', ' ', m.group(1).strip())
    name = m.group(2).strip().title() if m.group(2).strip() else '(unnamed)'
    return plate, name


def _extract_vehicle_matrix_block(ws, block):
    """Turn one vehicle grid into rows of {plate, name, date, amount, frequency}."""
    date_cols = {c: (ws.cell(row=block['date_row'], column=c).value.date()
                      if isinstance(ws.cell(row=block['date_row'], column=c).value, datetime)
                      else ws.cell(row=block['date_row'], column=c).value)
                 for c in block['date_cols']}
    rows = []
    r = block['data_start']
    blank_streak = 0
    while blank_streak < 3:
        plate_raw = ws.cell(row=r, column=1).value
        if plate_raw in (None, ''):
            blank_streak += 1
            r += 1
            continue
        if str(plate_raw).strip().lower() in _FRANCHISE_ANALYSIS_LABELS:
            break  # ran into the next section (e.g. a "TOTAL INCOME" row)
        blank_streak = 0
        plate, name = _split_plate_name(str(plate_raw))
        for c, d in date_cols.items():
            amt = ws.cell(row=r, column=c).value
            if isinstance(amt, (int, float)) and amt:
                rows.append(dict(plate=plate, name=name, date=d.isoformat(),
                                  amount=float(amt), frequency=block['frequency']))
        r += 1
    return rows


def import_franchise_workbook(file, created_by=None):
    """Read every sheet of an uploaded workbook, auto-detect every
    vehicle-collection table in it by header text (not sheet name or
    position), and sort each into FranchiseVehicle / FranchiseCollection.
    Returns a summary dict; does not commit — the caller decides when to
    commit/rollback."""
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        raise ValueError('Could not read that file. Make sure it is a valid Excel workbook.')

    matrix_rows = []
    cell_budget = MAX_WORKBOOK_IMPORT_CELLS
    for ws in wb.worksheets:
        for block in _vehicle_matrix_blocks(ws):
            matrix_rows.extend(_extract_vehicle_matrix_block(ws, block))
        cell_budget -= ws.max_row * ws.max_column
        if cell_budget < 0:
            raise ValueError('That workbook is too large to import in one pass — split it into smaller files.')

    # A table can legitimately repeat across sheets (the source file's own
    # "Weekly Analysis" tab re-states some collection rows already on the
    # monthly sheet) — first occurrence wins rather than erroring.
    seen_matrix = set()
    deduped_matrix = []
    for row in matrix_rows:
        key = (row['plate'], row['date'], row['frequency'])
        if key in seen_matrix:
            continue
        seen_matrix.add(key)
        deduped_matrix.append(row)

    vehicle_by_plate = {v.number_plate: v for v in FranchiseVehicle.query.all()}
    vehicles_created, collections_created, collections_skipped = 0, 0, 0
    created_records = []
    for row in deduped_matrix:
        vehicle = vehicle_by_plate.get(row['plate'])
        if not vehicle:
            vehicle = FranchiseVehicle(number_plate=row['plate'], franchisee_name=row['name'], status='active')
            db.session.add(vehicle)
            db.session.flush()
            vehicle_by_plate[row['plate']] = vehicle
            created_records.append(('franchise_vehicles', vehicle.id))
            vehicles_created += 1
        entry_date = parse_import_date(row['date'])
        if FranchiseCollection.query.filter_by(
                vehicle_id=vehicle.id, entry_date=entry_date, frequency=row['frequency']).first():
            collections_skipped += 1
            continue
        collection = FranchiseCollection(
            vehicle_id=vehicle.id, entry_date=entry_date, frequency=row['frequency'],
            amount=row['amount'], created_by=created_by,
        )
        db.session.add(collection)
        db.session.flush()
        created_records.append(('franchise_collections', collection.id))
        collections_created += 1

    return dict(
        vehicles_created=vehicles_created, collections_created=collections_created,
        collections_skipped=collections_skipped, created_records=created_records,
        total_rows=len(matrix_rows),
    )


def form_float(form, field, label=None, required=True, default=None, min_value=None):
    label = label or field.replace('_', ' ').capitalize()
    raw = (form.get(field) or '').strip()
    if not raw:
        if required:
            raise ValueError(f'{label} is required.')
        return default
    try:
        value = float(raw)
    except ValueError:
        raise ValueError(f'{label} must be a number.')
    if min_value is not None and value < min_value:
        raise ValueError(f'{label} cannot be less than {min_value}.')
    return value


def form_int(form, field, label=None, required=True, default=None, min_value=None):
    label = label or field.replace('_', ' ').capitalize()
    raw = (form.get(field) or '').strip()
    if not raw:
        if required:
            raise ValueError(f'{label} is required.')
        return default
    try:
        value = int(raw)
    except ValueError:
        raise ValueError(f'{label} must be a whole number.')
    if min_value is not None and value < min_value:
        raise ValueError(f'{label} cannot be less than {min_value}.')
    return value


def check_unique(model, field_name, value, label=None, exclude_id=None):
    """Raise a friendly ValueError if another row already has this value,
    instead of letting the DB's UNIQUE constraint crash with a 500."""
    label = label or field_name.replace('_', ' ').capitalize()
    q = model.query.filter(getattr(model, field_name) == value)
    if exclude_id is not None:
        q = q.filter(model.id != exclude_id)
    if q.first():
        raise ValueError(f'{label} "{value}" is already in use.')


def handle_form_errors(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if request.method == 'POST':
            try:
                return f(*args, **kwargs)
            except KeyError as e:
                db.session.rollback()
                flash(f'Missing required field: {e}', 'danger')
                return redirect(request.url)
            except ValueError as e:
                db.session.rollback()
                flash(str(e), 'danger')
                return redirect(request.url)
        return f(*args, **kwargs)
    return decorated


def query_date_range(default_from=None, default_to=None):
    """Read date_from/date_to from the query string for report pages.
    Falls back to defaults on missing/invalid input and auto-swaps a
    reversed range, so any date (or combination) can be requested
    without raising — the caller always gets a usable (from, to) pair."""
    today = date.today()
    default_from = default_from or today.replace(day=1)
    default_to = default_to or today

    from_str = request.args.get('date_from', '').strip()
    to_str = request.args.get('date_to', '').strip()

    try:
        df = parse_date(from_str) if from_str else default_from
    except ValueError:
        flash(f'"{from_str}" is not a valid start date — showing {default_from} instead.', 'warning')
        df = default_from

    try:
        dt = parse_date(to_str) if to_str else default_to
    except ValueError:
        flash(f'"{to_str}" is not a valid end date — showing {default_to} instead.', 'warning')
        dt = default_to

    if df > dt:
        df, dt = dt, df
        flash('Start date was after end date — the range was swapped.', 'warning')

    return df, dt


def query_single_date(param='as_of', default=None):
    """Read a single date query param (e.g. 'as at' a point in time),
    falling back to a default on missing/invalid input instead of raising."""
    default = default or date.today()
    raw = request.args.get(param, '').strip()
    if not raw:
        return default
    try:
        return parse_date(raw)
    except ValueError:
        flash(f'"{raw}" is not a valid date — showing {default} instead.', 'warning')
        return default


def compute_commission_accrued(as_of):
    """Total driver/conductor commission earned (accrued) on all revenue
    up to as_of, using the same per-driver rate logic as the payroll report,
    but over the driver's entire history rather than one filter period."""
    dr_rate = app.config['COMMISSION_DRIVER_RATE']
    co_rate = app.config['COMMISSION_CONDUCTOR_RATE']
    total = 0.0
    for d in Driver.query.all():
        driven = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
            DailyLog.driver_id == d.id, DailyLog.log_date <= as_of).scalar() or 0
        conducted = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
            DailyLog.conductor_id == d.id, DailyLog.log_date <= as_of).scalar() or 0
        garnish_driven = db.session.query(func.sum(DailyLog.garnish)).filter(
            DailyLog.driver_id == d.id, DailyLog.log_date <= as_of).scalar() or 0
        garnish_conducted = db.session.query(func.sum(DailyLog.garnish)).filter(
            DailyLog.conductor_id == d.id, DailyLog.log_date <= as_of).scalar() or 0
        rate = d.commission_rate if d.commission_rate is not None else (
            dr_rate if d.role == 'driver' else co_rate)
        total += max(driven + conducted - garnish_driven - garnish_conducted, 0) * rate
    return total


def compute_financial_position(as_of):
    """Simplified statement of financial position as at a given date.

    This is built on real records (loans, payables, receivables, capital
    contributions, owner drawings, commission payments, expenses) rather
    than a single inferred "cash" figure, but it's still an approximation,
    not bookkeeping-grade — there's no real cash/bank reconciliation.
    Key modelling choices:
      - Vehicle purchases are a pure asset swap (cash down, fixed asset up)
        — they are NOT assumed to be capital-funded. If vehicles were
        bought before any Capital Contribution or Loan was recorded to
        explain the cash, Cash will legitimately show negative here. That's
        the system being honest about a genuine gap in what's recorded —
        fix it by adding an opening Capital Contribution or Loan entry.
      - Vehicles are depreciated straight-line over VEHICLE_USEFUL_LIFE_YEARS
        from when each was added to the fleet (no acquisition-date field
        exists separately from that).
      - Commission is accrued on all revenue earned (matching the payroll
        report's calculation), not just what's been paid out — the
        difference sits as a Commission Payable liability.
      - Payables/Receivables are accrual-based: recognized as an
        expense/revenue when created, not when settled.
      - Loan repayments are treated as pure principal reduction (interest
        is not separately expensed) — a deliberate simplification.
    """
    useful_life = app.config['VEHICLE_USEFUL_LIFE_YEARS']

    vehicle_rows = []
    total_vehicle_cost = 0.0
    total_accum_dep = 0.0
    for v in Vehicle.query.order_by(Vehicle.registration).all():
        acquired = v.created_at.date()
        if acquired > as_of:
            continue
        years_in_service = (as_of - acquired).days / 365.25
        accum_dep = min(v.acquisition_cost, v.acquisition_cost * years_in_service / useful_life) \
            if useful_life else 0.0
        nbv = v.acquisition_cost - accum_dep
        vehicle_rows.append({
            'vehicle': v,
            'cost': v.acquisition_cost,
            'accumulated_depreciation': accum_dep,
            'net_book_value': nbv,
            'years_in_service': years_in_service,
        })
        total_vehicle_cost += v.acquisition_cost
        total_accum_dep += accum_dep

    total_nbv = total_vehicle_cost - total_accum_dep

    total_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date <= as_of).scalar() or 0
    total_maintenance = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date <= as_of).scalar() or 0
    total_expenses = db.session.query(func.sum(Expense.amount)).filter(
        Expense.expense_date <= as_of).scalar() or 0

    commission_accrued = compute_commission_accrued(as_of)
    commission_paid = db.session.query(func.sum(CommissionPayment.amount)).filter(
        CommissionPayment.payment_date <= as_of).scalar() or 0
    commission_payable = commission_accrued - commission_paid

    loan_proceeds = db.session.query(func.sum(Loan.principal)).filter(
        Loan.start_date <= as_of).scalar() or 0
    loan_repayments = db.session.query(func.sum(LoanPayment.amount)).filter(
        LoanPayment.payment_date <= as_of).scalar() or 0
    loans_outstanding = loan_proceeds - loan_repayments

    payables_total = db.session.query(func.sum(Payable.amount)).filter(
        Payable.invoice_date <= as_of).scalar() or 0
    payables_paid = db.session.query(func.sum(Payable.amount)).filter(
        Payable.status == 'paid', Payable.paid_date.isnot(None), Payable.paid_date <= as_of).scalar() or 0
    payables_outstanding = payables_total - payables_paid

    receivables_total = db.session.query(func.sum(Receivable.amount)).filter(
        Receivable.invoice_date <= as_of).scalar() or 0
    receivables_collected = db.session.query(func.sum(Receivable.amount)).filter(
        Receivable.status == 'collected', Receivable.collected_date.isnot(None),
        Receivable.collected_date <= as_of).scalar() or 0
    receivables_outstanding = receivables_total - receivables_collected

    capital_contributions = db.session.query(func.sum(CapitalContribution.amount)).filter(
        CapitalContribution.contribution_date <= as_of).scalar() or 0
    owner_drawings = db.session.query(func.sum(OwnerDrawing.amount)).filter(
        OwnerDrawing.drawing_date <= as_of).scalar() or 0

    cash_and_equivalents = (
        total_revenue - total_maintenance - total_expenses
        - commission_paid - total_vehicle_cost
        + loan_proceeds - loan_repayments
        + capital_contributions - owner_drawings
        + receivables_collected - payables_paid
    )

    retained_earnings = (
        (total_revenue + receivables_total)
        - (total_maintenance + total_expenses + payables_total)
        - total_accum_dep - commission_accrued
    )
    owners_capital = capital_contributions - owner_drawings

    total_assets = total_nbv + cash_and_equivalents + receivables_outstanding
    total_liabilities = loans_outstanding + payables_outstanding + commission_payable
    total_equity = owners_capital + retained_earnings

    return {
        'as_of': as_of,
        'useful_life': useful_life,
        'vehicle_rows': vehicle_rows,
        'total_cost': total_vehicle_cost,
        'total_accum_dep': total_accum_dep,
        'total_nbv': total_nbv,
        'cash_and_equivalents': cash_and_equivalents,
        'receivables_outstanding': receivables_outstanding,
        'total_assets': total_assets,
        'loans_outstanding': loans_outstanding,
        'payables_outstanding': payables_outstanding,
        'commission_payable': commission_payable,
        'total_liabilities': total_liabilities,
        'owners_capital': owners_capital,
        'retained_earnings': retained_earnings,
        'total_equity': total_equity,
    }


# ─────────────────────────────────────────────────────────────
# Auth routes
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('dashboard') if current_user.is_authenticated else url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
@limiter.limit('10 per minute')
def login():
    if current_user.is_authenticated:
        return redirect(first_permitted_url(current_user))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password) and user.is_active:
            login_user(user)
            log_audit('LOGIN', description=f'User {username} logged in')
            db.session.commit()
            return redirect(first_permitted_url(user))
        flash('Invalid username or password.', 'danger')
    return render_template('auth/login.html')


@app.route('/logout')
@login_required
def logout():
    log_audit('LOGOUT', description=f'User {current_user.username} logged out')
    db.session.commit()
    logout_user()
    flash('You have been signed out.', 'info')
    return redirect(url_for('login'))


@app.route('/account/change-password', methods=['GET', 'POST'])
@login_required
def change_password():
    if request.method == 'POST':
        current_pw = request.form.get('current_password', '')
        new_pw = request.form.get('new_password', '')
        confirm_pw = request.form.get('confirm_password', '')
        if not current_user.check_password(current_pw):
            flash('Current password is incorrect.', 'danger')
        elif len(new_pw) < 6:
            flash('New password must be at least 6 characters.', 'danger')
        elif new_pw != confirm_pw:
            flash('New passwords do not match.', 'danger')
        else:
            current_user.set_password(new_pw)
            log_audit('UPDATE', 'users', current_user.id, f'{current_user.username} changed their password')
            db.session.commit()
            flash('Password changed successfully.', 'success')
            return redirect(url_for('dashboard'))
    return render_template('auth/change_password.html')


# ─────────────────────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────────────────────
@app.route('/no-access')
@login_required
def no_access():
    return render_template('auth/no_access.html')


@app.route('/dashboard')
@login_required
@permission_required('dashboard')
def dashboard():
    today = date.today()
    month_start = today.replace(day=1)

    today_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date == today).scalar() or 0

    month_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date >= month_start).scalar() or 0

    month_maintenance = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date >= month_start).scalar() or 0

    month_expenses = month_maintenance
    month_profit = month_revenue - month_expenses

    active_vehicles = Vehicle.query.filter_by(status='active').count()
    active_drivers = Driver.query.filter_by(status='active').count()

    expiry_threshold = today + timedelta(days=30)
    expiring_docs = VehicleDocument.query.filter(
        VehicleDocument.expiry_date.between(today, expiry_threshold)).count()
    expired_docs = VehicleDocument.query.filter(
        VehicleDocument.expiry_date < today).count()
    expiring_docs += Vehicle.query.filter(
        Vehicle.insurance_expiry.between(today, expiry_threshold)).count()
    expired_docs += Vehicle.query.filter(Vehicle.insurance_expiry < today).count()

    recent_logs = DailyLog.query.order_by(DailyLog.log_date.desc()).limit(6).all()

    rev_chart = []
    for i in range(6, -1, -1):
        d = today - timedelta(days=i)
        rev = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
            DailyLog.log_date == d).scalar() or 0
        rev_chart.append({'date': d.strftime('%d %b'), 'revenue': float(rev)})

    return render_template('dashboard.html',
        today_revenue=today_revenue, month_revenue=month_revenue,
        month_expenses=month_expenses, month_profit=month_profit,
        active_vehicles=active_vehicles, active_drivers=active_drivers,
        expiring_docs=expiring_docs, expired_docs=expired_docs,
        recent_logs=recent_logs, revenue_chart=json.dumps(rev_chart),
        today=today)


# ─────────────────────────────────────────────────────────────
# Vehicles
# ─────────────────────────────────────────────────────────────
@app.route('/vehicles')
@login_required
@permission_required('vehicles')
def vehicles():
    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('vehicles/index.html', vehicles=all_vehicles)


@app.route('/vehicles/add', methods=['GET', 'POST'])
@login_required
@permission_required('vehicles')
@handle_form_errors
def vehicle_add():
    if request.method == 'POST':
        registration = request.form['registration'].upper().strip()
        check_unique(Vehicle, 'registration', registration)
        fuel_type = request.form.get('fuel_type', 'diesel')
        if fuel_type not in ('diesel', 'petrol'):
            raise ValueError('Fuel type must be Diesel or Petrol.')
        v = Vehicle(
            registration=registration,
            make=request.form['make'].strip(),
            model=request.form['model'].strip(),
            year=form_int(request.form, 'year', min_value=1980),
            acquisition_cost=form_float(request.form, 'acquisition_cost', required=False, default=0, min_value=0),
            status=request.form.get('status', 'active'),
            fuel_type=fuel_type,
            daily_target=form_float(request.form, 'daily_target', required=False, min_value=0),
            insurance_provider=request.form.get('insurance_provider', '').strip() or None,
            insurance_policy_number=request.form.get('insurance_policy_number', '').strip() or None,
            insurance_expiry=parse_date(request.form.get('insurance_expiry')),
        )
        db.session.add(v)
        db.session.flush()
        log_audit('CREATE', 'vehicles', v.id, f'Added vehicle {v.registration}')
        db.session.commit()
        flash(f'Vehicle {v.registration} registered successfully.', 'success')
        return redirect(url_for('vehicles'))
    return render_template('vehicles/form.html', vehicle=None, action='Register')


@app.route('/vehicles/<int:vid>')
@login_required
@permission_required('vehicles')
def vehicle_detail(vid):
    v = Vehicle.query.get_or_404(vid)
    today = date.today()
    recent_logs = DailyLog.query.filter_by(vehicle_id=vid).order_by(DailyLog.log_date.desc()).limit(10).all()
    recent_fuel = FuelLog.query.filter_by(vehicle_id=vid).order_by(FuelLog.log_date.desc()).limit(5).all()
    recent_maint = MaintenanceLog.query.filter_by(vehicle_id=vid).order_by(MaintenanceLog.log_date.desc()).limit(5).all()
    return render_template('vehicles/detail.html', vehicle=v, today=today,
                           recent_logs=recent_logs, recent_fuel=recent_fuel,
                           recent_maint=recent_maint)


@app.route('/vehicles/<int:vid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('vehicles')
@handle_form_errors
def vehicle_edit(vid):
    v = Vehicle.query.get_or_404(vid)
    if request.method == 'POST':
        registration = request.form['registration'].upper().strip()
        check_unique(Vehicle, 'registration', registration, exclude_id=v.id)
        v.registration = registration
        v.make = request.form['make'].strip()
        v.model = request.form['model'].strip()
        v.year = form_int(request.form, 'year', min_value=1980)
        v.acquisition_cost = form_float(request.form, 'acquisition_cost', required=False, default=0, min_value=0)
        v.status = request.form.get('status', 'active')
        fuel_type = request.form.get('fuel_type', 'diesel')
        if fuel_type not in ('diesel', 'petrol'):
            raise ValueError('Fuel type must be Diesel or Petrol.')
        v.fuel_type = fuel_type
        v.daily_target = form_float(request.form, 'daily_target', required=False, min_value=0)
        v.insurance_provider = request.form.get('insurance_provider', '').strip() or None
        v.insurance_policy_number = request.form.get('insurance_policy_number', '').strip() or None
        v.insurance_expiry = parse_date(request.form.get('insurance_expiry'))
        log_audit('UPDATE', 'vehicles', v.id, f'Updated vehicle {v.registration}')
        db.session.commit()
        flash(f'Vehicle {v.registration} updated.', 'success')
        return redirect(url_for('vehicle_detail', vid=vid))
    return render_template('vehicles/form.html', vehicle=v, action='Edit')


@app.route('/vehicles/<int:vid>/delete', methods=['POST'])
@login_required
@admin_required
def vehicle_delete(vid):
    v = Vehicle.query.get_or_404(vid)
    reg = v.registration
    log_audit('DELETE', 'vehicles', vid, f'Deleted vehicle {reg}')
    db.session.delete(v)
    db.session.commit()
    flash(f'Vehicle {reg} removed.', 'warning')
    return redirect(url_for('vehicles'))


@app.route('/vehicles/<int:vehicle_id>/documents/add', methods=['GET', 'POST'])
@login_required
@permission_required('vehicles')
@handle_form_errors
def document_add(vehicle_id):
    vehicle = Vehicle.query.get_or_404(vehicle_id)
    if request.method == 'POST':
        doc = VehicleDocument(
            vehicle_id=vehicle_id,
            doc_type=request.form['doc_type'],
            reference_number=request.form.get('reference_number', '').strip(),
            issue_date=parse_date(request.form.get('issue_date')),
            expiry_date=parse_date(request.form['expiry_date']),
            notes=request.form.get('notes', '').strip(),
        )
        db.session.add(doc)
        log_audit('CREATE', 'vehicle_documents', None,
                  f'Added {doc.doc_type} for {vehicle.registration}')
        db.session.commit()
        flash('Document added successfully.', 'success')
        return redirect(url_for('vehicle_detail', vid=vehicle_id))
    return render_template('vehicles/document_form.html', vehicle=vehicle)


@app.route('/documents/<int:did>/delete', methods=['POST'])
@login_required
@permission_required('vehicles')
def document_delete(did):
    doc = VehicleDocument.query.get_or_404(did)
    vid = doc.vehicle_id
    log_audit('DELETE', 'vehicle_documents', did, f'Deleted {doc.doc_type} document')
    db.session.delete(doc)
    db.session.commit()
    flash('Document deleted.', 'warning')
    return redirect(url_for('vehicle_detail', vid=vid))


# ─────────────────────────────────────────────────────────────
# Drivers
# ─────────────────────────────────────────────────────────────
@app.route('/drivers')
@login_required
@permission_required('drivers')
def drivers():
    all_drivers = Driver.query.order_by(Driver.name).all()
    return render_template('drivers/index.html', drivers=all_drivers)


@app.route('/drivers/add', methods=['GET', 'POST'])
@login_required
@permission_required('drivers')
@handle_form_errors
def driver_add():
    if request.method == 'POST':
        rate_input = form_float(request.form, 'commission_rate', required=False, min_value=0)
        role = request.form.get('role', 'driver')
        license_number = request.form.get('license_number', '').strip() or None
        if license_number:
            check_unique(Driver, 'license_number', license_number, label='License number')
        paired_driver_id = form_int(request.form, 'paired_driver_id', required=False) if role == 'conductor' else None
        d = Driver(
            name=request.form['name'].strip(),
            license_number=license_number,
            phone=request.form.get('phone', '').strip(),
            role=role,
            commission_rate=rate_input / 100 if rate_input is not None else None,
            status=request.form.get('status', 'active'),
            paired_driver_id=paired_driver_id,
            assigned_vehicle_id=form_int(request.form, 'assigned_vehicle_id', required=False),
            next_of_kin_name=request.form.get('next_of_kin_name', '').strip(),
            next_of_kin_phone=request.form.get('next_of_kin_phone', '').strip(),
            next_of_kin_relationship=request.form.get('next_of_kin_relationship', '').strip(),
        )
        db.session.add(d)
        db.session.flush()
        log_audit('CREATE', 'drivers', d.id, f'Added driver {d.name}')
        db.session.commit()
        flash(f'Driver {d.name} registered.', 'success')
        return redirect(url_for('drivers'))
    eligible_drivers = Driver.query.filter_by(role='driver', status='active').order_by(Driver.name).all()
    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('drivers/form.html', driver=None, action='Register',
                           eligible_drivers=eligible_drivers, vehicles=all_vehicles)


@app.route('/drivers/<int:did>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('drivers')
@handle_form_errors
def driver_edit(did):
    d = Driver.query.get_or_404(did)
    if request.method == 'POST':
        rate_input = form_float(request.form, 'commission_rate', required=False, min_value=0)
        role = request.form.get('role', 'driver')
        license_number = request.form.get('license_number', '').strip() or None
        if license_number:
            check_unique(Driver, 'license_number', license_number, label='License number', exclude_id=d.id)
        paired_driver_id = form_int(request.form, 'paired_driver_id', required=False) if role == 'conductor' else None
        if paired_driver_id == d.id:
            raise ValueError('A conductor cannot be paired with themself.')
        d.name = request.form['name'].strip()
        d.license_number = license_number
        d.phone = request.form.get('phone', '').strip()
        d.role = role
        d.commission_rate = rate_input / 100 if rate_input is not None else None
        d.status = request.form.get('status', 'active')
        d.paired_driver_id = paired_driver_id
        d.assigned_vehicle_id = form_int(request.form, 'assigned_vehicle_id', required=False)
        d.next_of_kin_name = request.form.get('next_of_kin_name', '').strip()
        d.next_of_kin_phone = request.form.get('next_of_kin_phone', '').strip()
        d.next_of_kin_relationship = request.form.get('next_of_kin_relationship', '').strip()
        log_audit('UPDATE', 'drivers', d.id, f'Updated driver {d.name}')
        db.session.commit()
        flash(f'Driver {d.name} updated.', 'success')
        return redirect(url_for('drivers'))
    eligible_drivers = Driver.query.filter_by(role='driver', status='active').order_by(Driver.name).all()
    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('drivers/form.html', driver=d, action='Edit',
                           eligible_drivers=eligible_drivers, vehicles=all_vehicles)


@app.route('/drivers/<int:did>/delete', methods=['POST'])
@login_required
@admin_required
def driver_delete(did):
    d = Driver.query.get_or_404(did)
    name = d.name
    log_audit('DELETE', 'drivers', did, f'Deleted driver {name}')
    db.session.delete(d)
    db.session.commit()
    flash(f'Driver {name} removed.', 'warning')
    return redirect(url_for('drivers'))


# ─────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────
@app.route('/routes')
@login_required
@permission_required('routes')
def routes_list():
    all_routes = Route.query.order_by(Route.name).all()
    return render_template('routes/index.html', routes=all_routes)


@app.route('/routes/add', methods=['GET', 'POST'])
@login_required
@permission_required('routes')
@handle_form_errors
def route_add():
    if request.method == 'POST':
        r = Route(
            name=request.form['name'].strip(),
            start_point=request.form['start_point'].strip(),
            end_point=request.form['end_point'].strip(),
            distance_km=form_float(request.form, 'distance_km', required=False, min_value=0),
            fare_rate=form_float(request.form, 'fare_rate', min_value=0),
            status=request.form.get('status', 'active'),
        )
        db.session.add(r)
        db.session.flush()
        log_audit('CREATE', 'routes', r.id, f'Added route {r.name}')
        db.session.commit()
        flash(f'Route "{r.name}" added.', 'success')
        return redirect(url_for('routes_list'))
    return render_template('routes/form.html', route=None, action='Add')


@app.route('/routes/<int:rid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('routes')
@handle_form_errors
def route_edit(rid):
    r = Route.query.get_or_404(rid)
    if request.method == 'POST':
        r.name = request.form['name'].strip()
        r.start_point = request.form['start_point'].strip()
        r.end_point = request.form['end_point'].strip()
        r.distance_km = form_float(request.form, 'distance_km', required=False, min_value=0)
        r.fare_rate = form_float(request.form, 'fare_rate', min_value=0)
        r.status = request.form.get('status', 'active')
        log_audit('UPDATE', 'routes', r.id, f'Updated route {r.name}')
        db.session.commit()
        flash(f'Route "{r.name}" updated.', 'success')
        return redirect(url_for('routes_list'))
    return render_template('routes/form.html', route=r, action='Edit')


@app.route('/routes/<int:rid>/delete', methods=['POST'])
@login_required
@admin_required
def route_delete(rid):
    r = Route.query.get_or_404(rid)
    name = r.name
    log_audit('DELETE', 'routes', rid, f'Deleted route {name}')
    db.session.delete(r)
    db.session.commit()
    flash(f'Route "{name}" deleted.', 'warning')
    return redirect(url_for('routes_list'))


# ─────────────────────────────────────────────────────────────
# Daily Transactions — edit/delete a single vehicle/date entry from the
# Vehicle Ledger below. This replaces the old standalone Daily Logs
# CRUD pages, which duplicated the same DailyLog data behind a second,
# heavier form — everything now lives on one page: /logs/ledger.
# ─────────────────────────────────────────────────────────────
@app.route('/logs/ledger/<int:vehicle_id>/<log_date_str>/edit', methods=['GET', 'POST'])
@login_required
@permission_required_any('daily_logs', 'crew_portal')
@handle_form_errors
def ledger_entry_edit(vehicle_id, log_date_str):
    vehicle = Vehicle.query.get_or_404(vehicle_id)
    try:
        log_date = parse_date(log_date_str)
    except ValueError:
        flash(f'"{log_date_str}" is not a valid date.', 'danger')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id))

    period = request.values.get('period', 'month')
    daily_logs_for_date = DailyLog.query.filter_by(
        vehicle_id=vehicle_id, log_date=log_date).order_by(DailyLog.id).all()
    fuel_logs_for_date = FuelLog.query.filter_by(
        vehicle_id=vehicle_id, log_date=log_date).order_by(FuelLog.id).all()
    if len(daily_logs_for_date) > 1 or len(fuel_logs_for_date) > 1:
        flash('This day has more than one entry for this vehicle and can\'t be edited '
              'as a single row — delete it and re-enter instead.', 'warning')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))

    log = daily_logs_for_date[0] if daily_logs_for_date else None
    fuel = fuel_logs_for_date[0] if fuel_logs_for_date else None

    if request.method == 'POST':
        driver_id = form_int(request.form, 'driver_id', required=False)
        fare = form_float(request.form, 'fare', required=False, min_value=0)
        garnish = form_float(request.form, 'garnish', required=False, min_value=0)
        reason_for_shortfall = request.form.get('reason_for_shortfall', '').strip() or None
        diesel_cost = form_float(request.form, 'diesel_cost', required=False, min_value=0)
        mileage = form_float(request.form, 'mileage', required=False, min_value=0)

        if log is not None or fare is not None or garnish is not None:
            if fare is None:
                raise ValueError('Fare is required for this entry.')
        if garnish is not None and not driver_id:
            raise ValueError('Select a driver to record a garnish against.')

        if fare is None and diesel_cost is None and mileage is None and garnish is None:
            raise ValueError('Enter at least a fare, diesel cost, mileage reading, or garnish.')

        if fare is not None:
            driver = Driver.query.get(driver_id)
            conductor = driver.paired_conductors[0] if driver and driver.paired_conductors else None
            if log is None:
                log = DailyLog(vehicle_id=vehicle_id, log_date=log_date, created_by=current_user.id)
                db.session.add(log)
            log.driver_id = driver_id
            log.conductor_id = conductor.id if conductor else None
            log.gross_revenue = fare
            log.garnish = garnish or 0.0
            log.reason_for_shortfall = reason_for_shortfall
            log.updated_by = current_user.id
            log.updated_at = datetime.now(timezone.utc)

        if diesel_cost is not None or mileage is not None:
            if fuel is None:
                fuel = FuelLog(vehicle_id=vehicle_id, log_date=log_date, liters=0, created_by=current_user.id)
                db.session.add(fuel)
            fuel.total_cost = diesel_cost or 0
            fuel.odometer = mileage
        elif fuel is not None:
            db.session.delete(fuel)

        log_audit('UPDATE', 'daily_logs', log.id if log else None,
                  f'Edited ledger entry for {vehicle.registration} on {log_date}')
        db.session.commit()
        flash('Entry updated.', 'success')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))

    all_drivers = Driver.query.filter_by(role='driver', status='active').order_by(Driver.name).all()
    return render_template('logs/ledger_entry_form.html', vehicle=vehicle, log=log, fuel=fuel,
                           log_date=log_date, drivers=all_drivers, period=period)


@app.route('/logs/ledger/<int:vehicle_id>/<log_date_str>/delete', methods=['POST'])
@login_required
@admin_required
def ledger_entry_delete(vehicle_id, log_date_str):
    vehicle = Vehicle.query.get_or_404(vehicle_id)
    try:
        log_date = parse_date(log_date_str)
    except ValueError:
        flash(f'"{log_date_str}" is not a valid date.', 'danger')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id))

    period = request.form.get('period', 'month')
    DailyLog.query.filter_by(vehicle_id=vehicle_id, log_date=log_date).delete()
    FuelLog.query.filter_by(vehicle_id=vehicle_id, log_date=log_date).delete()
    log_audit('DELETE', 'daily_logs', None, f'Deleted ledger entry for {vehicle.registration} on {log_date}')
    db.session.commit()
    flash('Entry deleted.', 'warning')
    return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))


# ─────────────────────────────────────────────────────────────
# Vehicle Ledger — one running sheet per vehicle (date, driver, fare,
# diesel, mileage), matching how the fleet's paper/Excel logbooks are
# kept — one sheet per vehicle, driver rotating day to day. Posts to
# the same DailyLog/FuelLog tables the rest of the system uses.
# Replaces the old Crew Portal "Log Income" form. Filterable by
# day/week/month. Diesel is captured as a USD amount, not liters — crew
# report what they spent on fuel, not a metered liter reading, so these
# FuelLog rows carry liters=0 and are skipped by the Fuel Efficiency
# report (which needs liters) rather than showing a false 0 L/100km.
# ─────────────────────────────────────────────────────────────
def vehicle_ledger_rows(vehicle_id, df=None, dt=None, daily_target=None):
    """Merge DailyLog (fare, any driver) and FuelLog (diesel cost/mileage)
    for this vehicle by date, with distance computed the same way the Fuel
    Efficiency report does — delta from the previous odometer reading.
    If df/dt are given, only rows in that range are returned, but the
    distance baseline still uses the last odometer reading before df so
    the first visible row isn't wrongly shown as having no distance.
    If daily_target is given, each day with a driver entry is flagged
    against it (see 'shortfall' on each row) — the same target used by
    the Revenue Shortfalls report, surfaced here at entry time too."""
    daily_q = DailyLog.query.filter_by(vehicle_id=vehicle_id)
    fuel_q = FuelLog.query.filter_by(vehicle_id=vehicle_id)
    if df:
        daily_q = daily_q.filter(DailyLog.log_date >= df)
        fuel_q = fuel_q.filter(FuelLog.log_date >= df)
    if dt:
        daily_q = daily_q.filter(DailyLog.log_date <= dt)
        fuel_q = fuel_q.filter(FuelLog.log_date <= dt)

    daily_by_date = {}
    for log in daily_q.order_by(DailyLog.log_date).all():
        daily_by_date.setdefault(log.log_date, []).append(log)

    fuel_by_date = {}
    for log in fuel_q.order_by(FuelLog.log_date).all():
        fuel_by_date.setdefault(log.log_date, []).append(log)

    prev_odometer = None
    if df:
        baseline = FuelLog.query.filter(
            FuelLog.vehicle_id == vehicle_id, FuelLog.log_date < df,
            FuelLog.odometer.isnot(None)).order_by(FuelLog.log_date.desc()).first()
        prev_odometer = baseline.odometer if baseline else None

    all_dates = sorted(set(daily_by_date) | set(fuel_by_date))
    rows = []
    total_fare = 0.0
    total_diesel_cost = 0.0
    total_garnish = 0.0
    for d in all_dates:
        daily_logs = daily_by_date.get(d, [])
        fare = sum(l.gross_revenue for l in daily_logs)
        garnish = sum(l.garnish for l in daily_logs)
        reason_for_shortfall = '; '.join(n for n in (l.reason_for_shortfall for l in daily_logs) if n) or None
        driver_names = ', '.join(sorted({l.driver.name for l in daily_logs if l.driver})) or None
        fuel_logs = fuel_by_date.get(d, [])
        diesel_cost = sum(f.total_cost for f in fuel_logs)
        odometer = max((f.odometer for f in fuel_logs if f.odometer is not None), default=None)

        distance = None
        if odometer is not None and prev_odometer is not None:
            distance = odometer - prev_odometer
        if odometer is not None:
            prev_odometer = odometer

        total_fare += fare
        total_diesel_cost += diesel_cost
        total_garnish += garnish
        # Only flag a day that actually has a driver/fare entry — a
        # fuel-only row shouldn't read as a missed revenue target.
        shortfall = None
        if daily_target and daily_logs and fare < daily_target:
            shortfall = daily_target - fare
        rows.append({
            'date': d, 'driver_names': driver_names, 'fare': fare,
            'garnish': garnish, 'reason_for_shortfall': reason_for_shortfall,
            'shortfall': shortfall,
            'diesel_cost': diesel_cost,
            'odometer': odometer, 'distance': distance,
            # More than one entry on the same day (e.g. a driver change
            # mid-shift) can't be represented/edited as a single row —
            # the UI falls back to "delete all, re-enter" for these.
            'multiple': len(daily_logs) > 1 or len(fuel_logs) > 1,
        })
    return rows, total_fare, total_diesel_cost, total_garnish


def resolve_ledger_period(period, today):
    if period == 'today':
        df = dt = today
    elif period == 'week':
        df, dt = today - timedelta(days=today.weekday()), today
    elif period == 'all':
        df = dt = None
    else:
        period = 'month'
        df, dt = today.replace(day=1), today
    return period, df, dt


@app.route('/logs/ledger')
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger():
    today = date.today()

    period, df, dt = resolve_ledger_period(request.args.get('period', 'month'), today)
    date_from_str = df.strftime('%Y-%m-%d') if df else ''
    date_to_str = dt.strftime('%Y-%m-%d') if dt else ''

    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    vehicle_id = request.args.get('vehicle_id', '')
    vehicle = Vehicle.query.get(vehicle_id) if vehicle_id else (all_vehicles[0] if all_vehicles else None)

    rows, total_fare, total_diesel_cost, total_garnish = [], 0.0, 0.0, 0.0
    latest_odometer = None
    if vehicle:
        rows, total_fare, total_diesel_cost, total_garnish = vehicle_ledger_rows(
            vehicle.id, df, dt, daily_target=vehicle.daily_target)
        latest_fuel = FuelLog.query.filter(
            FuelLog.vehicle_id == vehicle.id, FuelLog.odometer.isnot(None)
        ).order_by(FuelLog.log_date.desc(), FuelLog.id.desc()).first()
        latest_odometer = latest_fuel.odometer if latest_fuel else None

    all_drivers = Driver.query.filter_by(role='driver', status='active').order_by(Driver.name).all()
    return render_template('logs/ledger.html', vehicles=all_vehicles, vehicle=vehicle,
        drivers=all_drivers,
        rows=rows, total_fare=total_fare, total_diesel_cost=total_diesel_cost,
        total_garnish=total_garnish,
        period=period, date_from=date_from_str, date_to=date_to_str,
        today=today.strftime('%Y-%m-%d'), latest_odometer=latest_odometer)


@app.route('/logs/ledger/add', methods=['POST'])
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger_add():
    # This is a POST-only route with no GET counterpart at the same URL, so
    # (unlike the other forms in the app) errors can't redirect back to
    # request.url — that would GET this same POST-only URL and 405. Errors
    # are handled locally here and always redirect to the GET ledger page.
    period = request.form.get('period', 'month')
    vehicle_id = form_int(request.form, 'vehicle_id', required=False)
    try:
        vehicle = Vehicle.query.get(vehicle_id) if vehicle_id else None
        if not vehicle:
            raise ValueError('Select a vehicle.')

        log_date = parse_date(request.form['log_date'])
        driver_id = form_int(request.form, 'driver_id', required=False)
        fare = form_float(request.form, 'fare', required=False, min_value=0)
        garnish = form_float(request.form, 'garnish', required=False, min_value=0)
        reason_for_shortfall = request.form.get('reason_for_shortfall', '').strip() or None
        diesel_cost = form_float(request.form, 'diesel_cost', required=False, min_value=0)
        mileage = form_float(request.form, 'mileage', required=False, min_value=0)

        if garnish is not None and not driver_id:
            raise ValueError('Select a driver to record a garnish against.')
        if driver_id and fare is None and garnish is None:
            raise ValueError('Fare or a garnish is required when a driver is selected.')

        if fare is None and diesel_cost is None and mileage is None and garnish is None:
            raise ValueError('Enter at least a fare, diesel cost, mileage reading, or garnish.')

        if fare is not None or garnish is not None:
            driver = Driver.query.get(driver_id)
            conductor = driver.paired_conductors[0] if driver and driver.paired_conductors else None
            daily = DailyLog(
                vehicle_id=vehicle_id, driver_id=driver_id, conductor_id=conductor.id if conductor else None,
                log_date=log_date, gross_revenue=fare or 0.0,
                garnish=garnish or 0.0, reason_for_shortfall=reason_for_shortfall,
                created_by=current_user.id,
            )
            db.session.add(daily)
            log_audit('CREATE', 'daily_logs', None,
                       f'Ledger entry for {vehicle.registration} on {log_date}: fare {fare or 0.0}' +
                       (f', garnish {garnish} ({reason_for_shortfall})' if garnish else ''))

        if diesel_cost is not None:
            fuel = FuelLog(
                vehicle_id=vehicle_id, log_date=log_date, liters=0,
                total_cost=diesel_cost, odometer=mileage, created_by=current_user.id,
            )
            db.session.add(fuel)
            log_audit('CREATE', 'fuel_logs', None, f'Ledger entry for {vehicle.registration} on {log_date}: diesel ${diesel_cost}')
        elif mileage is not None:
            fuel = FuelLog(
                vehicle_id=vehicle_id, log_date=log_date, liters=0, odometer=mileage, created_by=current_user.id,
            )
            db.session.add(fuel)

        db.session.commit()
        flash('Ledger entry recorded.', 'success')

        # Auto-flag a shortfall right at entry time, not just later on the
        # Revenue Shortfalls report — checks the day's total fare (this entry
        # plus any others already logged for the same vehicle/date).
        if fare is not None and vehicle.daily_target:
            day_total = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
                DailyLog.vehicle_id == vehicle_id, DailyLog.log_date == log_date).scalar() or 0
            if day_total < vehicle.daily_target:
                gap = vehicle.daily_target - day_total
                flash(f'Flagged: {vehicle.registration} is ${gap:,.2f} short of its ${vehicle.daily_target:,.2f} '
                      f'daily target on {log_date} — see Revenue Shortfalls to garnish.', 'warning')
    except KeyError as e:
        db.session.rollback()
        flash(f'Missing required field: {e}', 'danger')
    except ValueError as e:
        db.session.rollback()
        flash(str(e), 'danger')

    return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))


@app.route('/logs/ledger/export')
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger_export():
    vehicle_id = request.args.get('vehicle_id', '')
    vehicle = Vehicle.query.get(vehicle_id) if vehicle_id else None
    if not vehicle:
        flash('Select a vehicle to export.', 'danger')
        return redirect(url_for('driver_ledger'))

    period, df, dt = resolve_ledger_period(request.args.get('period', 'month'), date.today())
    rows, _, _, _ = vehicle_ledger_rows(vehicle.id, df, dt)

    fuel_label = vehicle.fuel_type.capitalize()
    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['Date', 'Driver', 'Fare', f'{fuel_label} (USD)', 'Mileage', 'Distance',
                'Garnish', 'Reason for Shortfall'])
    for row in rows:
        w.writerow([
            row['date'], row['driver_names'] or '',
            f"{row['fare']:.2f}" if row['fare'] else '',
            f"{row['diesel_cost']:.2f}" if row['diesel_cost'] else '',
            row['odometer'] if row['odometer'] is not None else '',
            row['distance'] if row['distance'] is not None else '',
            f"{row['garnish']:.2f}" if row['garnish'] else '',
            row['reason_for_shortfall'] or '',
        ])
    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    safe_reg = vehicle.registration.replace(' ', '_')
    resp.headers['Content-Disposition'] = f'attachment; filename={safe_reg}_ledger_{period}_{date.today()}.csv'
    return resp


def _resolve_ledger_import_vehicle():
    period = request.form.get('period', 'month')
    vehicle_id = form_int(request.form, 'vehicle_id', required=False)
    vehicle = Vehicle.query.get(vehicle_id) if vehicle_id else None
    return vehicle, vehicle_id, period


@app.route('/logs/ledger/import/preview', methods=['POST'])
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger_import_preview():
    vehicle, vehicle_id, period = _resolve_ledger_import_vehicle()
    if not vehicle:
        flash('Select a vehicle before importing.', 'danger')
        return redirect(url_for('driver_ledger', period=period))

    file = request.files.get('file')
    if file and file.filename:
        filename = file.filename
        try:
            headers, raw_rows = read_uploaded_table(file, vehicle=vehicle)
        except ValueError as e:
            flash(str(e), 'danger')
            return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))
        mapping = auto_map_columns(headers)
    else:
        # Re-preview after the user adjusted the mapping — the file itself
        # isn't resubmitted, the previously parsed rows travel via raw_data.
        try:
            filename = request.form.get('filename', 'uploaded file')
            payload = json.loads(request.form.get('raw_data') or '{}')
            headers, raw_rows = payload.get('headers') or [], payload.get('rows') or []
            if not headers or not raw_rows:
                raise ValueError('Choose a CSV or Excel file to import.')
            mapping = {field_key: (request.form.get(f'map_{field_key}') or None)
                       for field_key, _label, _syn in CANONICAL_LEDGER_FIELDS}
        except (ValueError, json.JSONDecodeError, TypeError):
            flash('Choose a CSV or Excel file to import — the previous preview session expired.', 'danger')
            return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))

    if not raw_rows:
        flash('That file has no data rows to import — it only has a header row. '
              'Add rows with a Date and Fare/Diesel (USD)/Mileage, then re-import.', 'warning')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))

    preview_rows = apply_column_mapping(headers, raw_rows[:10], mapping)
    return render_template('logs/ledger_import_preview.html',
                           vehicle=vehicle, period=period, filename=filename,
                           headers=headers, mapping=mapping, fields=CANONICAL_LEDGER_FIELDS,
                           preview_rows=preview_rows, row_count=len(raw_rows),
                           raw_data=json.dumps({'headers': headers, 'rows': raw_rows}))


@app.route('/logs/ledger/import/confirm', methods=['POST'])
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger_import_confirm():
    vehicle, vehicle_id, period = _resolve_ledger_import_vehicle()
    if not vehicle:
        flash('Select a vehicle before importing.', 'danger')
        return redirect(url_for('driver_ledger', period=period))

    filename = request.form.get('filename', 'uploaded file')
    try:
        payload = json.loads(request.form.get('raw_data') or '{}')
        headers, raw_rows = payload.get('headers') or [], payload.get('rows') or []
        if not raw_rows:
            raise ValueError('empty')
    except (ValueError, json.JSONDecodeError, TypeError):
        flash('That preview session expired — please choose the file again.', 'danger')
        return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))

    mapping = {field_key: (request.form.get(f'map_{field_key}') or None)
               for field_key, _label, _syn in CANONICAL_LEDGER_FIELDS}
    file_rows = apply_column_mapping(headers, raw_rows, mapping)
    imported, errors, error_rows, created_drivers, created_records = import_ledger_rows(
        file_rows, vehicle, auto_register_drivers=True)

    if imported or error_rows:
        # Commit even when imported == 0: a batch made only of failed rows
        # still needs to persist so its quarantine CSV can be downloaded.
        save_import_batch('ledger', filename, len(raw_rows), imported, error_rows, created_records)
        if imported:
            log_audit('CREATE', 'daily_logs', None,
                       f'Imported {imported} ledger row(s) for {vehicle.registration} from {filename}')
        for driver in Driver.query.filter(Driver.name.in_(created_drivers)).all():
            log_audit('CREATE', 'drivers', driver.id,
                       f'Auto-registered driver "{driver.name}" from ledger import ({filename}) — '
                       f'not on file, added because a fare row named them.')
        db.session.commit()
    else:
        db.session.rollback()

    if imported:
        flash(f'Imported {imported} row(s) for {vehicle.registration}.', 'success')
    if created_drivers:
        flash(f'Auto-registered new driver(s): {", ".join(created_drivers)}.', 'success')
    if errors:
        shown = errors[:10]
        more = f' (+{len(errors) - 10} more)' if len(errors) > 10 else ''
        flash('Skipped rows — ' + '; '.join(shown) + more, 'warning')
    if not imported and not errors:
        flash('No rows found to import.', 'warning')

    return redirect(url_for('driver_ledger', vehicle_id=vehicle_id, period=period))


@app.route('/logs/ledger/import/bulk', methods=['POST'])
@login_required
@permission_required_any('daily_logs', 'crew_portal')
def driver_ledger_import_bulk():
    """Import a fleet-wide workbook with one sheet per vehicle (sheet name
    matched against each vehicle's registration) in a single pass. Unlike the
    single-vehicle import there's no manual column-mapping step — with dozens
    of sheets that isn't practical, so each sheet is auto-mapped the same way
    and the results page reports what happened per vehicle."""
    period = request.form.get('period', 'month')
    file = request.files.get('file')
    if not file or not file.filename:
        flash('Choose an Excel workbook to import.', 'danger')
        return redirect(url_for('driver_ledger', period=period))

    try:
        sheets = read_uploaded_workbook_sheets(file)
    except ValueError as e:
        flash(str(e), 'danger')
        return redirect(url_for('driver_ledger', period=period))

    if not sheets:
        flash('That workbook has no sheets with data rows to import.', 'warning')
        return redirect(url_for('driver_ledger', period=period))

    auto_register = request.form.get('auto_register') == '1'
    vehicles_by_reg = {_normalize_registration(v.registration): v for v in Vehicle.query.all()}

    results = []
    for sheet_name, (headers, rows) in sheets.items():
        savepoint = db.session.begin_nested()
        try:
            mapping = auto_map_columns(headers)
            vehicle = vehicles_by_reg.get(_normalize_registration(sheet_name))
            vehicle_created = False

            # Distinguish an actual per-vehicle ledger tab from a fleet-wide
            # summary sheet like "DAILY TOTAL INCOME" before ever auto-registering
            # a "vehicle" for it. A summary sheet also has a Date column and its
            # "INCOME" header even fuzzy-matches the Fare synonym list, so Fare
            # alone isn't a safe signal — a real per-vehicle logbook always
            # names a Driver or logs Mileage; a financial rollup never does.
            looks_like_ledger = bool(mapping.get('date')) and bool(
                mapping.get('driver') or mapping.get('mileage'))

            if not vehicle and looks_like_ledger and auto_register:
                registration = re.sub(r'\s+', ' ', sheet_name).strip().upper()
                vehicle = Vehicle(registration=registration, make='Unknown', model='Unknown',
                                   year=date.today().year, status='active', fuel_type='diesel')
                db.session.add(vehicle)
                db.session.flush()
                vehicles_by_reg[_normalize_registration(registration)] = vehicle
                vehicle_created = True
                log_audit('CREATE', 'vehicles', vehicle.id,
                          f'Auto-registered vehicle {registration} from fleet workbook '
                          f'import (sheet "{sheet_name}") — make/model/year are placeholders.')

            if not vehicle:
                reason = ('No vehicle matches this sheet name.' if looks_like_ledger else
                          "Doesn't look like a per-vehicle ledger sheet.")
                results.append({'sheet': sheet_name, 'vehicle': None, 'mapping': {},
                                 'imported': 0, 'errors': [], 'created_drivers': [],
                                 'vehicle_created': False, 'skip_reason': reason})
                savepoint.commit()
                continue

            if not mapping.get('date'):
                results.append({'sheet': sheet_name, 'vehicle': vehicle, 'mapping': mapping,
                                 'imported': 0, 'errors': [], 'created_drivers': [],
                                 'vehicle_created': vehicle_created,
                                 'skip_reason': 'No Date column detected.'})
                savepoint.commit()
                continue

            mapped_rows = apply_column_mapping(headers, rows, mapping)
            imported, errors, _error_rows, created_drivers, _created_records = import_ledger_rows(
                mapped_rows, vehicle, auto_register_drivers=auto_register)
            if imported:
                log_audit('CREATE', 'daily_logs', None,
                           f'Imported {imported} ledger row(s) for {vehicle.registration} '
                           f'from {file.filename} (sheet "{sheet_name}")')
            for driver in Driver.query.filter(Driver.name.in_(created_drivers)).all():
                log_audit('CREATE', 'drivers', driver.id,
                           f'Auto-registered driver "{driver.name}" from fleet workbook import '
                           f'({file.filename}, sheet "{sheet_name}") — not on file, added because '
                           f'a fare row named them.')
            results.append({'sheet': sheet_name, 'vehicle': vehicle, 'mapping': mapping,
                             'imported': imported, 'errors': errors,
                             'created_drivers': created_drivers,
                             'vehicle_created': vehicle_created, 'skip_reason': None})
            savepoint.commit()
        except Exception as e:
            savepoint.rollback()
            results.append({'sheet': sheet_name, 'vehicle': None, 'mapping': {},
                             'imported': 0, 'errors': [], 'created_drivers': [],
                             'vehicle_created': False, 'skip_reason': f'Unexpected error: {e}'})

    total_imported = sum(r['imported'] for r in results)
    total_registered = (sum(1 for r in results if r['vehicle_created']) +
                        sum(len(r['created_drivers']) for r in results))
    if total_imported or total_registered:
        db.session.commit()
    else:
        db.session.rollback()

    return render_template('logs/ledger_bulk_import_result.html',
                           filename=file.filename, results=results,
                           total_imported=total_imported, period=period,
                           fields=CANONICAL_LEDGER_FIELDS)


# ─────────────────────────────────────────────────────────────
# Fuel Logs
# ─────────────────────────────────────────────────────────────
@app.route('/logs/fuel')
@login_required
@permission_required('fuel_logs')
def fuel_logs():
    page = request.args.get('page', 1, type=int)
    vehicle_id = request.args.get('vehicle_id', '')
    q = FuelLog.query
    if vehicle_id:
        q = q.filter(FuelLog.vehicle_id == vehicle_id)
    logs = q.order_by(FuelLog.log_date.desc()).paginate(page=page, per_page=20)
    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('logs/fuel/index.html', logs=logs, vehicles=all_vehicles,
                           vehicle_id=vehicle_id)


@app.route('/logs/fuel/add', methods=['GET', 'POST'])
@login_required
@permission_required('fuel_logs')
@handle_form_errors
def fuel_log_add():
    if request.method == 'POST':
        liters = form_float(request.form, 'liters', min_value=0)
        log = FuelLog(
            vehicle_id=form_int(request.form, 'vehicle_id'),
            log_date=parse_date(request.form['log_date']),
            liters=liters,
            odometer=form_float(request.form, 'odometer', required=False, min_value=0),
            supplier=request.form.get('supplier', '').strip(),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(log)
        db.session.flush()
        log_audit('CREATE', 'fuel_logs', log.id,
                  f'Fuel log for {log.vehicle.registration}: {liters}L')
        db.session.commit()
        flash('Fuel log recorded.', 'success')
        return redirect(url_for('fuel_logs'))

    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('logs/fuel/form.html', vehicles=all_vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/logs/fuel/<int:lid>/delete', methods=['POST'])
@login_required
@admin_required
def fuel_log_delete(lid):
    log = FuelLog.query.get_or_404(lid)
    log_audit('DELETE', 'fuel_logs', lid, f'Deleted fuel log {lid}')
    db.session.delete(log)
    db.session.commit()
    flash('Fuel log deleted.', 'warning')
    return redirect(url_for('fuel_logs'))


# ─────────────────────────────────────────────────────────────
# Maintenance Logs
# ─────────────────────────────────────────────────────────────
@app.route('/logs/maintenance')
@login_required
@permission_required('maintenance')
def maintenance_logs():
    page = request.args.get('page', 1, type=int)
    vehicle_id = request.args.get('vehicle_id', '')
    q = MaintenanceLog.query
    if vehicle_id:
        q = q.filter(MaintenanceLog.vehicle_id == vehicle_id)
    logs = q.order_by(MaintenanceLog.log_date.desc()).paginate(page=page, per_page=20)
    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('logs/maintenance/index.html', logs=logs, vehicles=all_vehicles,
                           vehicle_id=vehicle_id)


@app.route('/logs/maintenance/add', methods=['GET', 'POST'])
@login_required
@permission_required('maintenance')
@handle_form_errors
def maintenance_log_add():
    if request.method == 'POST':
        parts = form_float(request.form, 'parts_cost', required=False, default=0, min_value=0)
        labor = form_float(request.form, 'labor_cost', required=False, default=0, min_value=0)
        log = MaintenanceLog(
            vehicle_id=form_int(request.form, 'vehicle_id'),
            log_date=parse_date(request.form['log_date']),
            description=request.form['description'].strip(),
            parts_cost=parts,
            labor_cost=labor,
            total_cost=parts + labor,
            mechanic=request.form.get('mechanic', '').strip(),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(log)
        db.session.flush()
        log_audit('CREATE', 'maintenance_logs', log.id,
                  f'Maintenance for {log.vehicle.registration}')
        db.session.commit()
        flash('Maintenance log recorded.', 'success')
        return redirect(url_for('maintenance_logs'))

    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('logs/maintenance/form.html', vehicles=all_vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/logs/maintenance/<int:lid>/delete', methods=['POST'])
@login_required
@admin_required
def maintenance_log_delete(lid):
    log = MaintenanceLog.query.get_or_404(lid)
    log_audit('DELETE', 'maintenance_logs', lid, f'Deleted maintenance log {lid}')
    db.session.delete(log)
    db.session.commit()
    flash('Maintenance log deleted.', 'warning')
    return redirect(url_for('maintenance_logs'))


def latest_odometer(vehicle_id):
    latest_fuel = FuelLog.query.filter_by(vehicle_id=vehicle_id).filter(
        FuelLog.odometer.isnot(None)).order_by(FuelLog.log_date.desc()).first()
    return latest_fuel.odometer if latest_fuel else None


def schedule_status(sched):
    today = date.today()
    odometer = latest_odometer(sched.vehicle_id)
    due_by_date = sched.next_due_date is not None and sched.next_due_date <= today
    due_by_km = (sched.next_due_odometer is not None and odometer is not None
                 and odometer >= sched.next_due_odometer)
    if due_by_date or due_by_km:
        return 'overdue'
    soon_by_date = sched.next_due_date is not None and (sched.next_due_date - today).days <= 14
    soon_by_km = (sched.next_due_odometer is not None and odometer is not None
                  and sched.next_due_odometer - odometer <= 500)
    if soon_by_date or soon_by_km:
        return 'due_soon'
    return 'ok'


# ─────────────────────────────────────────────────────────────
# Preventive Maintenance Scheduling
# ─────────────────────────────────────────────────────────────
@app.route('/maintenance/schedules')
@login_required
@permission_required('maintenance')
def maintenance_schedules():
    schedules = MaintenanceSchedule.query.filter_by(status='active').join(Vehicle).order_by(
        Vehicle.registration).all()
    rows = [{'schedule': s, 'status': schedule_status(s), 'odometer': latest_odometer(s.vehicle_id)}
            for s in schedules]
    rows.sort(key=lambda r: {'overdue': 0, 'due_soon': 1, 'ok': 2}[r['status']])
    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('maintenance/schedules.html', rows=rows, vehicles=all_vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/maintenance/schedules/add', methods=['GET', 'POST'])
@login_required
@permission_required('maintenance')
@handle_form_errors
def maintenance_schedule_add():
    if request.method == 'POST':
        interval_days = form_int(request.form, 'interval_days', required=False, min_value=1)
        interval_km = form_float(request.form, 'interval_km', required=False, min_value=1)
        last_done_date = parse_date(request.form.get('last_done_date')) or date.today()
        last_done_odometer = form_float(request.form, 'last_done_odometer', required=False, min_value=0)

        sched = MaintenanceSchedule(
            vehicle_id=form_int(request.form, 'vehicle_id'),
            description=request.form['description'].strip(),
            interval_days=interval_days,
            interval_km=interval_km,
            last_done_date=last_done_date,
            last_done_odometer=last_done_odometer,
            next_due_date=(last_done_date + timedelta(days=interval_days)) if interval_days else None,
            next_due_odometer=(last_done_odometer + interval_km)
                if (interval_km and last_done_odometer is not None) else None,
        )
        db.session.add(sched)
        db.session.flush()
        log_audit('CREATE', 'maintenance_schedules', sched.id, f'Added maintenance schedule: {sched.description}')
        db.session.commit()
        flash('Maintenance schedule added.', 'success')
        return redirect(url_for('maintenance_schedules'))
    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('maintenance/schedule_form.html', vehicles=all_vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/maintenance/schedules/<int:sid>/done', methods=['POST'])
@login_required
@permission_required('maintenance')
@handle_form_errors
def maintenance_schedule_done(sid):
    sched = MaintenanceSchedule.query.get_or_404(sid)
    done_date = parse_date(request.form.get('done_date')) or date.today()
    done_odometer = form_float(request.form, 'done_odometer', required=False, min_value=0)

    sched.last_done_date = done_date
    sched.last_done_odometer = done_odometer if done_odometer is not None else sched.last_done_odometer
    sched.next_due_date = (done_date + timedelta(days=sched.interval_days)) if sched.interval_days else None
    sched.next_due_odometer = (sched.last_done_odometer + sched.interval_km) \
        if (sched.interval_km and sched.last_done_odometer is not None) else None

    log_audit('UPDATE', 'maintenance_schedules', sid, f'Marked "{sched.description}" as done')
    db.session.commit()
    flash('Marked as done — next due date recalculated.', 'success')
    return redirect(url_for('maintenance_schedules'))


@app.route('/maintenance/schedules/<int:sid>/delete', methods=['POST'])
@login_required
@admin_required
def maintenance_schedule_delete(sid):
    sched = MaintenanceSchedule.query.get_or_404(sid)
    log_audit('DELETE', 'maintenance_schedules', sid, f'Deleted maintenance schedule: {sched.description}')
    db.session.delete(sched)
    db.session.commit()
    flash('Maintenance schedule deleted.', 'warning')
    return redirect(url_for('maintenance_schedules'))


# ─────────────────────────────────────────────────────────────
# Spares Store: parts inventory, purchases & marked-up sales
# ─────────────────────────────────────────────────────────────
@app.route('/store/parts')
@login_required
@permission_required('store')
def store_parts():
    parts = SparePart.query.order_by(SparePart.name).all()
    total_stock_value = sum(p.stock_value for p in parts)
    low_stock_count = sum(1 for p in parts if p.status == 'active' and p.low_stock)
    month_start = date.today().replace(day=1)
    month_profit = sum(s.profit for s in
                       StoreSale.query.filter(StoreSale.sale_date >= month_start).all())
    return render_template('store/parts.html', parts=parts,
                           total_stock_value=total_stock_value,
                           low_stock_count=low_stock_count, month_profit=month_profit)


@app.route('/store/parts/add', methods=['GET', 'POST'])
@login_required
@permission_required('store')
@handle_form_errors
def store_part_add():
    if request.method == 'POST':
        part = SparePart(
            name=request.form['name'].strip(),
            part_number=request.form.get('part_number', '').strip(),
            unit=request.form.get('unit', 'pc').strip() or 'pc',
            markup_percent=form_float(request.form, 'markup_percent', required=False,
                                      default=0, min_value=0),
            reorder_level=form_int(request.form, 'reorder_level', required=False,
                                   default=0, min_value=0),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(part)
        db.session.flush()
        log_audit('CREATE', 'spare_parts', part.id, f'Added spare part: {part.name}')
        db.session.commit()
        flash('Spare part added. Record a purchase to bring in stock.', 'success')
        return redirect(url_for('store_parts'))
    return render_template('store/part_form.html', part=None)


@app.route('/store/parts/<int:pid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('store')
@handle_form_errors
def store_part_edit(pid):
    part = SparePart.query.get_or_404(pid)
    if request.method == 'POST':
        part.name = request.form['name'].strip()
        part.part_number = request.form.get('part_number', '').strip()
        part.unit = request.form.get('unit', 'pc').strip() or 'pc'
        part.markup_percent = form_float(request.form, 'markup_percent', required=False,
                                         default=0, min_value=0)
        part.reorder_level = form_int(request.form, 'reorder_level', required=False,
                                      default=0, min_value=0)
        part.status = request.form.get('status', 'active')
        part.notes = request.form.get('notes', '').strip()
        log_audit('UPDATE', 'spare_parts', part.id, f'Updated spare part: {part.name}')
        db.session.commit()
        flash('Spare part updated.', 'success')
        return redirect(url_for('store_parts'))
    return render_template('store/part_form.html', part=part)


@app.route('/store/parts/<int:pid>/delete', methods=['POST'])
@login_required
@admin_required
def store_part_delete(pid):
    part = SparePart.query.get_or_404(pid)
    log_audit('DELETE', 'spare_parts', pid, f'Deleted spare part: {part.name}')
    db.session.delete(part)
    db.session.commit()
    flash('Spare part deleted.', 'warning')
    return redirect(url_for('store_parts'))


@app.route('/store/purchases')
@login_required
@permission_required('store')
def store_purchases():
    page = request.args.get('page', 1, type=int)
    part_id = request.args.get('part_id', '')
    q = StorePurchase.query
    if part_id:
        q = q.filter(StorePurchase.part_id == part_id)
    purchases = q.order_by(StorePurchase.purchase_date.desc()).paginate(page=page, per_page=20)
    all_parts = SparePart.query.order_by(SparePart.name).all()
    return render_template('store/purchases.html', purchases=purchases, parts=all_parts,
                           part_id=part_id)


@app.route('/store/purchases/add', methods=['GET', 'POST'])
@login_required
@permission_required('store')
@handle_form_errors
def store_purchase_add():
    if request.method == 'POST':
        part = SparePart.query.get_or_404(form_int(request.form, 'part_id'))
        quantity = form_int(request.form, 'quantity', min_value=1)
        unit_cost = form_float(request.form, 'unit_cost', min_value=0)

        purchase = StorePurchase(
            part_id=part.id,
            purchase_date=parse_date(request.form['purchase_date']),
            quantity=quantity,
            unit_cost=unit_cost,
            total_cost=quantity * unit_cost,
            supplier=request.form.get('supplier', '').strip(),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        new_total_qty = part.quantity_on_hand + quantity
        part.cost_price = ((part.quantity_on_hand * part.cost_price) +
                           (quantity * unit_cost)) / new_total_qty
        part.quantity_on_hand = new_total_qty

        db.session.add(purchase)
        db.session.flush()
        log_audit('CREATE', 'store_purchases', purchase.id,
                  f'Purchased {quantity} x {part.name} @ {unit_cost}')
        db.session.commit()
        flash('Purchase recorded and stock updated.', 'success')
        return redirect(url_for('store_purchases'))

    all_parts = SparePart.query.filter_by(status='active').order_by(SparePart.name).all()
    return render_template('store/purchase_form.html', parts=all_parts,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/store/purchases/<int:pid>/delete', methods=['POST'])
@login_required
@admin_required
def store_purchase_delete(pid):
    purchase = StorePurchase.query.get_or_404(pid)
    part = purchase.part
    part.quantity_on_hand = max(0, part.quantity_on_hand - purchase.quantity)
    log_audit('DELETE', 'store_purchases', pid,
              f'Deleted purchase of {purchase.quantity} x {part.name}')
    db.session.delete(purchase)
    db.session.commit()
    flash('Purchase deleted and stock reduced accordingly. Note: this does not '
         'recompute historical average cost.', 'warning')
    return redirect(url_for('store_purchases'))


@app.route('/store/sales')
@login_required
@permission_required('store')
def store_sales():
    page = request.args.get('page', 1, type=int)
    part_id = request.args.get('part_id', '')
    q = StoreSale.query
    if part_id:
        q = q.filter(StoreSale.part_id == part_id)
    sales = q.order_by(StoreSale.sale_date.desc()).paginate(page=page, per_page=20)
    all_parts = SparePart.query.order_by(SparePart.name).all()
    return render_template('store/sales.html', sales=sales, parts=all_parts, part_id=part_id)


@app.route('/store/sales/add', methods=['GET', 'POST'])
@login_required
@permission_required('store')
@handle_form_errors
def store_sale_add():
    if request.method == 'POST':
        part = SparePart.query.get_or_404(form_int(request.form, 'part_id'))
        quantity = form_int(request.form, 'quantity', min_value=1)
        if quantity > part.quantity_on_hand:
            raise ValueError(f'Only {part.quantity_on_hand} {part.unit}(s) of {part.name} in stock.')
        unit_price = form_float(request.form, 'unit_price', required=False,
                                default=part.selling_price, min_value=0)
        vehicle_id = form_int(request.form, 'vehicle_id', required=False)

        sale = StoreSale(
            part_id=part.id,
            vehicle_id=vehicle_id,
            sale_date=parse_date(request.form['sale_date']),
            quantity=quantity,
            unit_cost=part.cost_price,
            unit_price=unit_price,
            total_amount=quantity * unit_price,
            customer_name=request.form.get('customer_name', '').strip() if not vehicle_id else None,
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        part.quantity_on_hand -= quantity

        db.session.add(sale)
        db.session.flush()
        log_audit('CREATE', 'store_sales', sale.id,
                  f'Sold {quantity} x {part.name} @ {unit_price}' +
                  (f' to vehicle {sale.vehicle.registration} (booked as an expense on that vehicle)'
                   if sale.vehicle else ''))
        db.session.commit()
        flash('Sale recorded and stock updated.', 'success')
        return redirect(url_for('store_sales'))

    all_parts = SparePart.query.filter(SparePart.status == 'active',
                                       SparePart.quantity_on_hand > 0).order_by(SparePart.name).all()
    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    return render_template('store/sale_form.html', parts=all_parts, vehicles=all_vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/store/sales/<int:sid>/delete', methods=['POST'])
@login_required
@admin_required
def store_sale_delete(sid):
    sale = StoreSale.query.get_or_404(sid)
    part = sale.part
    part.quantity_on_hand += sale.quantity
    log_audit('DELETE', 'store_sales', sid, f'Deleted sale of {sale.quantity} x {part.name}')
    db.session.delete(sale)
    db.session.commit()
    flash('Sale deleted and stock restored.', 'warning')
    return redirect(url_for('store_sales'))


@app.route('/store/trading-account')
@login_required
@permission_required('store')
def store_trading_account():
    """Standalone Trading & Profit or Loss Account for the Spares Store,
    run as its own cost centre. Revenue is ALL sales — including internal
    sales to company vehicles at the same marked-up price a walk-in
    customer pays — because from the store's side that's a real sale.
    That same amount also lands as a maintenance expense on the buying
    vehicle's income statement (see vehicle_income_totals): the store
    recognizes revenue, the vehicle recognizes a cost — two sides of one
    internal transfer, not a double-count within either statement.
    Cost of sales is summed from each sale's snapshotted unit_cost, which
    is more accurate than an opening/closing-stock estimate here since
    cost_price is a live weighted average, not a point-in-time figure."""
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    sales = StoreSale.query.filter(StoreSale.sale_date.between(df, dt)).all()

    sales_revenue = sum(s.total_amount for s in sales)
    cost_of_sales = sum(s.unit_cost * s.quantity for s in sales)
    gross_profit = sales_revenue - cost_of_sales
    gross_margin = (gross_profit / sales_revenue * 100) if sales_revenue else 0

    external = [s for s in sales if not s.vehicle_id]
    internal = [s for s in sales if s.vehicle_id]
    external_sales = sum(s.total_amount for s in external)
    external_cogs = sum(s.unit_cost * s.quantity for s in external)
    internal_sales = sum(s.total_amount for s in internal)
    internal_cogs = sum(s.unit_cost * s.quantity for s in internal)

    purchases_total = db.session.query(func.sum(StorePurchase.total_cost)).filter(
        StorePurchase.purchase_date.between(df, dt)).scalar() or 0
    closing_stock_value = sum(p.stock_value for p in SparePart.query.all())

    by_part = {}
    for s in sales:
        row = by_part.setdefault(s.part_id,
            {'part': s.part, 'quantity': 0, 'sales': 0.0, 'cost_of_sales': 0.0})
        row['quantity'] += s.quantity
        row['sales'] += s.total_amount
        row['cost_of_sales'] += s.unit_cost * s.quantity
    part_breakdown = list(by_part.values())
    for row in part_breakdown:
        row['gross_profit'] = row['sales'] - row['cost_of_sales']
        row['margin'] = (row['gross_profit'] / row['sales'] * 100) if row['sales'] else 0
    part_breakdown.sort(key=lambda r: r['gross_profit'], reverse=True)

    return render_template('store/trading_account.html',
        sales_revenue=sales_revenue, cost_of_sales=cost_of_sales,
        gross_profit=gross_profit, gross_margin=gross_margin,
        external_sales=external_sales, external_cogs=external_cogs,
        external_profit=external_sales - external_cogs,
        internal_sales=internal_sales, internal_cogs=internal_cogs,
        internal_profit=internal_sales - internal_cogs,
        purchases_total=purchases_total, closing_stock_value=closing_stock_value,
        part_breakdown=part_breakdown,
        date_from=date_from_str, date_to=date_to_str)


# ─────────────────────────────────────────────────────────────
# Reports
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# Crew Portal
# ─────────────────────────────────────────────────────────────
# Income logging moved to the Driver Ledger (see driver_ledger / driver_ledger_add
# above) — it replaced this page so fare, diesel and mileage are captured together.


@app.route('/crew/leaderboard')
@login_required
@permission_required('crew_portal')
def crew_leaderboard():
    today = date.today()
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    dr_rate = app.config['COMMISSION_DRIVER_RATE']
    co_rate = app.config['COMMISSION_CONDUCTOR_RATE']

    rows = []
    for d in Driver.query.filter_by(status='active').order_by(Driver.name).all():
        driven = db.session.query(
            func.sum(DailyLog.gross_revenue),
            func.sum(DailyLog.trips_completed),
            func.count(DailyLog.id),
        ).filter(DailyLog.driver_id == d.id,
                 DailyLog.log_date.between(df, dt)).first()
        conducted = db.session.query(
            func.sum(DailyLog.gross_revenue),
            func.sum(DailyLog.trips_completed),
            func.count(DailyLog.id),
        ).filter(DailyLog.conductor_id == d.id,
                 DailyLog.log_date.between(df, dt)).first()

        total_rev   = (driven[0] or 0) + (conducted[0] or 0)
        total_trips = (driven[1] or 0) + (conducted[1] or 0)
        days_worked = (driven[2] or 0) + (conducted[2] or 0)
        if days_worked == 0:
            continue
        rate = d.commission_rate if d.commission_rate is not None else (
            dr_rate if d.role == 'driver' else co_rate)
        rows.append({
            'driver': d,
            'total_revenue': total_rev,
            'total_trips': total_trips,
            'days_worked': days_worked,
            'avg_per_day': total_rev / days_worked if days_worked else 0,
            'commission': total_rev * rate,
            'rate_pct': rate * 100,
        })

    rows.sort(key=lambda r: r['total_revenue'], reverse=True)
    for i, r in enumerate(rows):
        r['rank'] = i + 1

    my_driver = current_user.linked_driver
    my_row = next((r for r in rows if my_driver and r['driver'].id == my_driver.id), None)

    return render_template('crew/leaderboard.html',
                           rows=rows, my_row=my_row,
                           date_from=date_from_str, date_to=date_to_str, today=today)


def vehicle_income_totals(df, dt, vehicle_id=None):
    """Revenue/maintenance/expense/spares totals for one vehicle (or the
    whole fleet if vehicle_id is None) over [df, dt]. Only expenses (and
    store sales) explicitly tagged to a vehicle count toward that vehicle's
    statement — general overhead (untagged expenses) only appears in the
    consolidated total. Spares sold from the in-house store to a company
    vehicle are booked as an expense on that vehicle here, at the same
    marked-up price the store charges any other customer — see StoreSale."""
    rev_q = db.session.query(func.sum(DailyLog.gross_revenue)).filter(DailyLog.log_date.between(df, dt))
    maint_q = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(MaintenanceLog.log_date.between(df, dt))
    exp_q = db.session.query(func.sum(Expense.amount)).filter(Expense.expense_date.between(df, dt))
    spares_q = db.session.query(func.sum(StoreSale.total_amount)).filter(StoreSale.sale_date.between(df, dt))

    if vehicle_id:
        rev_q = rev_q.filter(DailyLog.vehicle_id == vehicle_id)
        maint_q = maint_q.filter(MaintenanceLog.vehicle_id == vehicle_id)
        exp_q = exp_q.filter(Expense.vehicle_id == vehicle_id)
        spares_q = spares_q.filter(StoreSale.vehicle_id == vehicle_id)
    else:
        exp_q = exp_q.filter(Expense.vehicle_id.is_(None))
        spares_q = spares_q.filter(StoreSale.vehicle_id.isnot(None))

    revenue = rev_q.scalar() or 0
    maintenance = maint_q.scalar() or 0
    expenses = exp_q.scalar() or 0
    spares = spares_q.scalar() or 0
    return revenue, maintenance, expenses, spares


def expense_breakdown_by_category(df, dt, vehicle_id=None):
    """Group Expense amounts by heading/sub-heading for the period and scope,
    for the itemized breakdown on the income statement. Consolidated scope
    (vehicle_id=None) includes both vehicle-tagged and general expenses,
    matching the consolidated statement's total; per-vehicle scope only
    includes that vehicle's tagged expenses."""
    q = db.session.query(Expense.category_id, func.sum(Expense.amount)).filter(
        Expense.expense_date.between(df, dt))
    if vehicle_id:
        q = q.filter(Expense.vehicle_id == vehicle_id)
    totals_by_cat = dict(q.group_by(Expense.category_id).all())

    rows = []
    for h in ExpenseCategory.query.filter_by(parent_id=None).order_by(ExpenseCategory.name).all():
        h_direct = totals_by_cat.get(h.id, 0)
        children = []
        for c in sorted(h.children, key=lambda x: x.name):
            amt = totals_by_cat.get(c.id, 0)
            if amt:
                children.append({'name': c.name, 'amount': amt})
        heading_total = h_direct + sum(c['amount'] for c in children)
        if heading_total:
            rows.append({'name': h.name, 'direct': h_direct, 'children': children, 'total': heading_total})
    rows.sort(key=lambda r: r['total'], reverse=True)
    return rows


def statement_expense_line_items(df, dt, vehicle_id=None):
    """The actual source rows behind each Statement Summary category —
    MaintenanceLog entries and vehicle-tagged StoreSale spares under
    Maintenance, Expense rows grouped by their top-level heading (or
    'Other' for a custom heading) everywhere else. Powers the income
    statement's click-to-drill-down UI so a category total isn't a dead end."""
    standard_names = ('Maintenance', 'Wages', 'Traffic Fines', 'Insurance', 'Admin')
    items = {name: [] for name in standard_names}
    items['Other'] = []

    maint_q = MaintenanceLog.query.filter(MaintenanceLog.log_date.between(df, dt))
    if vehicle_id:
        maint_q = maint_q.filter(MaintenanceLog.vehicle_id == vehicle_id)
    for m in maint_q.all():
        items['Maintenance'].append({
            'date': m.log_date, 'source': 'Maintenance Log',
            'description': m.description or '—', 'vehicle': m.vehicle.registration,
            'amount': m.total_cost,
        })

    spares_q = StoreSale.query.filter(StoreSale.sale_date.between(df, dt))
    if vehicle_id:
        spares_q = spares_q.filter(StoreSale.vehicle_id == vehicle_id)
    else:
        spares_q = spares_q.filter(StoreSale.vehicle_id.isnot(None))
    for s in spares_q.all():
        items['Maintenance'].append({
            'date': s.sale_date, 'source': 'Store Sale',
            'description': s.part.name, 'vehicle': s.vehicle.registration if s.vehicle else '—',
            'amount': s.total_amount,
        })

    heading_name_by_cat_id = {}
    for h in ExpenseCategory.query.filter_by(parent_id=None).all():
        heading_name_by_cat_id[h.id] = h.name
        for c in h.children:
            heading_name_by_cat_id[c.id] = h.name

    exp_q = Expense.query.filter(Expense.expense_date.between(df, dt))
    if vehicle_id:
        exp_q = exp_q.filter(Expense.vehicle_id == vehicle_id)
    for e in exp_q.all():
        heading = heading_name_by_cat_id.get(e.category_id)
        bucket = heading if heading in items else 'Other'
        items[bucket].append({
            'date': e.expense_date, 'source': e.category.display_name,
            'description': e.description or '—',
            'vehicle': e.vehicle.registration if e.vehicle else '—',
            'amount': e.amount,
        })

    for rows in items.values():
        rows.sort(key=lambda r: r['date'], reverse=True)
    return items


@app.route('/reports/income')
@login_required
@permission_required('reports')
def report_income():
    vehicle_id = request.args.get('vehicle_id', '')
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    if vehicle_id:
        # Per-vehicle statement: only costs directly attributable to this vehicle.
        gross_revenue, maintenance_cost, vehicle_expenses, spares_cost = vehicle_income_totals(df, dt, vehicle_id)
        general_expenses = 0
    else:
        # Consolidated statement: fleet-wide maintenance plus ALL expenses
        # (both vehicle-tagged and general overhead), plus spares sold from
        # the store to any company vehicle.
        gross_revenue, maintenance_cost, general_expenses, spares_cost = vehicle_income_totals(df, dt, None)
        vehicle_expenses = db.session.query(func.sum(Expense.amount)).filter(
            Expense.expense_date.between(df, dt), Expense.vehicle_id.isnot(None)).scalar() or 0

    total_expenses = maintenance_cost + vehicle_expenses + general_expenses + spares_cost
    net_profit = gross_revenue - total_expenses
    profit_margin = (net_profit / gross_revenue * 100) if gross_revenue else 0

    vehicle_breakdown = []
    for v in Vehicle.query.order_by(Vehicle.registration).all():
        v_rev, v_maint, v_exp, v_spares = vehicle_income_totals(df, dt, v.id)
        if v_rev == 0 and v_maint == 0 and v_exp == 0 and v_spares == 0:
            continue
        v_total_cost = v_maint + v_exp + v_spares
        v_net = v_rev - v_total_cost
        vehicle_breakdown.append({
            # Spares bought from the in-house store for this vehicle count as
            # a maintenance cost, alongside logged maintenance-job spend.
            'vehicle': v, 'revenue': v_rev, 'maintenance': v_maint + v_spares,
            'expenses': v_exp, 'net_profit': v_net,
            'margin': (v_net / v_rev * 100) if v_rev else 0,
        })
    vehicle_breakdown.sort(key=lambda r: r['net_profit'], reverse=True)
    expense_breakdown = expense_breakdown_by_category(df, dt, vehicle_id or None)

    # Statement Summary classifies expenses under five fixed headings
    # (matching create_default_expense_categories) rather than the
    # tagged/untagged split — Maintenance folds in logged maintenance-job
    # costs, spares sold from the store to a company vehicle, and any
    # Expense rows booked to the Maintenance category, all together.
    # Anything booked under a custom heading a user has added falls into
    # "Other" so it still counts toward Total Operating Expenses.
    statement_category_names = ('Maintenance', 'Wages', 'Traffic Fines', 'Insurance', 'Admin')
    category_totals = {name: 0.0 for name in statement_category_names}
    other_expenses = 0.0
    for row in expense_breakdown:
        if row['name'] in category_totals:
            category_totals[row['name']] = row['total']
        else:
            other_expenses += row['total']
    category_totals['Maintenance'] += maintenance_cost + spares_cost

    statement_expenses = [(name, category_totals[name]) for name in statement_category_names]
    if other_expenses:
        statement_expenses.append(('Other', other_expenses))
    statement_expense_items = statement_expense_line_items(df, dt, vehicle_id or None)

    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('reports/income.html',
        gross_revenue=gross_revenue,
        maintenance_cost=maintenance_cost, vehicle_expenses=vehicle_expenses,
        general_expenses=general_expenses, total_expenses=total_expenses,
        statement_expenses=statement_expenses, statement_expense_items=statement_expense_items,
        net_profit=net_profit, profit_margin=profit_margin,
        vehicle_breakdown=vehicle_breakdown, expense_breakdown=expense_breakdown,
        vehicles=all_vehicles,
        date_from=date_from_str, date_to=date_to_str, vehicle_id=vehicle_id)


@app.route('/reports/payroll')
@login_required
@permission_required('reports')
def report_payroll():
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    dr_rate = app.config['COMMISSION_DRIVER_RATE']
    co_rate = app.config['COMMISSION_CONDUCTOR_RATE']

    earnings = []
    for d in Driver.query.filter_by(status='active').order_by(Driver.name).all():
        driven = db.session.query(func.sum(DailyLog.gross_revenue),
                                  func.count(DailyLog.id)).filter(
            DailyLog.driver_id == d.id,
            DailyLog.log_date.between(df, dt)).first()
        conducted = db.session.query(func.sum(DailyLog.gross_revenue),
                                     func.count(DailyLog.id)).filter(
            DailyLog.conductor_id == d.id,
            DailyLog.log_date.between(df, dt)).first()
        garnish_driven = db.session.query(func.sum(DailyLog.garnish)).filter(
            DailyLog.driver_id == d.id,
            DailyLog.log_date.between(df, dt)).scalar() or 0
        garnish_conducted = db.session.query(func.sum(DailyLog.garnish)).filter(
            DailyLog.conductor_id == d.id,
            DailyLog.log_date.between(df, dt)).scalar() or 0

        rev = (driven[0] or 0) + (conducted[0] or 0)
        days = (driven[1] or 0) + (conducted[1] or 0)
        garnish = garnish_driven + garnish_conducted
        if days == 0 and not garnish:
            continue
        rate = d.commission_rate if d.commission_rate is not None else (
            dr_rate if d.role == 'driver' else co_rate)
        # Garnish is netted off revenue before commission is calculated, not
        # deducted from the commission afterwards — the commission percentage
        # applies to what the crew member actually brought in after garnish.
        commission = max(rev - garnish, 0) * rate
        paid = db.session.query(func.sum(CommissionPayment.amount)).filter(
            CommissionPayment.driver_id == d.id,
            CommissionPayment.payment_date.between(df, dt)).scalar() or 0
        earnings.append({
            'driver': d,
            'total_revenue': rev,
            'days_worked': days,
            'rate_pct': rate * 100,
            'commission': commission,
            'garnish': garnish,
            'paid': paid,
            'outstanding': commission - paid,
            'conductors': [],
        })

    total_commissions = sum(e['commission'] for e in earnings)
    total_garnish = sum(e['garnish'] for e in earnings)
    total_paid = sum(e['paid'] for e in earnings)
    total_outstanding = sum(e['outstanding'] for e in earnings)

    # Nest each conductor's row under their paired driver so payroll reads
    # crew-by-crew (driver + the conductor who rode with them) instead of one
    # flat alphabetical list mixing roles. A conductor only nests if their
    # paired driver also has an earnings row this period — otherwise (no
    # pairing set, or the paired driver didn't earn anything) it stays as
    # its own top-level row, same as before. If a driver has no conductor
    # to nest, a placeholder row still prints under them labeled "Conductor",
    # with commission projected off the driver's own revenue at the standard
    # conductor rate — there's no CommissionPayment target without a real
    # person, so it's informational (no Pay action), not an amount owed.
    grouped, nested_ids = [], set()
    for e in earnings:
        if e['driver'].role == 'conductor':
            continue
        for ce in earnings:
            if ce['driver'].role == 'conductor' and ce['driver'].paired_driver_id == e['driver'].id:
                e['conductors'].append(ce)
                nested_ids.add(ce['driver'].id)
        if not e['conductors']:
            placeholder_commission = max(e['total_revenue'] - e['garnish'], 0) * co_rate
            e['conductors'].append({
                'driver': None, 'is_placeholder': True,
                'total_revenue': e['total_revenue'], 'days_worked': e['days_worked'],
                'rate_pct': co_rate * 100,
                'commission': placeholder_commission, 'garnish': e['garnish'], 'paid': 0,
                'outstanding': placeholder_commission,
            })
        grouped.append(e)
    grouped.extend(e for e in earnings if e['driver'].role == 'conductor' and e['driver'].id not in nested_ids)
    earnings = grouped

    return render_template('reports/payroll.html',
        earnings=earnings, total_commissions=total_commissions,
        total_garnish=total_garnish,
        total_paid=total_paid, total_outstanding=total_outstanding,
        date_from=date_from_str, date_to=date_to_str)


@app.route('/reports/shortfalls')
@login_required
@permission_required('reports')
def report_shortfalls():
    """Flags every vehicle/day where actual fare fell below that vehicle's
    admin-set daily_target — vehicles with no target set are skipped
    entirely. Each flagged day shows how much garnish (if any) has already
    been applied against the shortfall, so the admin can see at a glance
    what's still unresolved and act on it inline."""
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    rows = []
    targeted_vehicles = Vehicle.query.filter(
        Vehicle.daily_target.isnot(None), Vehicle.daily_target > 0
    ).order_by(Vehicle.registration).all()
    for v in targeted_vehicles:
        logs = DailyLog.query.filter(
            DailyLog.vehicle_id == v.id, DailyLog.log_date.between(df, dt)
        ).all()
        by_date = {}
        for log in logs:
            by_date.setdefault(log.log_date, []).append(log)
        for d, day_logs in by_date.items():
            fare = sum(l.gross_revenue for l in day_logs)
            if fare >= v.daily_target:
                continue
            garnish = sum(l.garnish for l in day_logs)
            reasons = '; '.join(n for n in (l.reason_for_shortfall for l in day_logs) if n) or None
            drivers = sorted({l.driver for l in day_logs if l.driver}, key=lambda dr: dr.name)
            shortfall = v.daily_target - fare
            rows.append({
                'vehicle': v, 'date': d, 'drivers': drivers,
                'target': v.daily_target, 'fare': fare, 'shortfall': shortfall,
                'garnish': garnish, 'remaining': shortfall - garnish,
                'reason_for_shortfall': reasons,
            })

    rows.sort(key=lambda r: r['date'], reverse=True)
    total_shortfall = sum(r['shortfall'] for r in rows)
    total_garnish = sum(r['garnish'] for r in rows)
    total_remaining = sum(max(r['remaining'], 0) for r in rows)
    pending_count = sum(1 for r in rows if r['remaining'] > 0)

    return render_template('reports/shortfalls.html', rows=rows,
        total_shortfall=total_shortfall, total_garnish=total_garnish,
        total_remaining=total_remaining, pending_count=pending_count,
        date_from=date_from_str, date_to=date_to_str)


@app.route('/finance/commission-payments/add', methods=['POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def commission_payment_add():
    payment = CommissionPayment(
        driver_id=form_int(request.form, 'driver_id'),
        payment_date=parse_date(request.form['payment_date']),
        amount=form_float(request.form, 'amount', min_value=0),
        period_start=parse_date(request.form.get('period_start')),
        period_end=parse_date(request.form.get('period_end')),
        method=request.form.get('method', '').strip(),
        notes=request.form.get('notes', '').strip(),
        created_by=current_user.id,
    )
    db.session.add(payment)
    db.session.flush()
    log_audit('CREATE', 'commission_payments', payment.id,
              f'Commission payment of {payment.amount} to driver #{payment.driver_id}')
    db.session.commit()
    flash('Commission payment recorded.', 'success')
    return redirect(request.referrer or url_for('report_payroll'))


@app.route('/finance/commission-payments/<int:pid>/delete', methods=['POST'])
@login_required
@admin_required
def commission_payment_delete(pid):
    payment = CommissionPayment.query.get_or_404(pid)
    log_audit('DELETE', 'commission_payments', pid, f'Deleted commission payment of {payment.amount}')
    db.session.delete(payment)
    db.session.commit()
    flash('Commission payment deleted.', 'warning')
    return redirect(request.referrer or url_for('report_payroll'))


@app.route('/reports/cash-flow')
@login_required
@permission_required('reports')
def report_cash_flow():
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    def in_range(col):
        return col.between(df, dt)

    operating_in = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        in_range(DailyLog.log_date)).scalar() or 0
    receivables_in = db.session.query(func.sum(Receivable.amount)).filter(
        Receivable.status == 'collected', in_range(Receivable.collected_date)).scalar() or 0

    maint_out = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        in_range(MaintenanceLog.log_date)).scalar() or 0
    expenses_out = db.session.query(func.sum(Expense.amount)).filter(
        in_range(Expense.expense_date)).scalar() or 0
    commission_out = db.session.query(func.sum(CommissionPayment.amount)).filter(
        in_range(CommissionPayment.payment_date)).scalar() or 0
    payables_out = db.session.query(func.sum(Payable.amount)).filter(
        Payable.status == 'paid', in_range(Payable.paid_date)).scalar() or 0

    net_operating = operating_in + receivables_in - maint_out - expenses_out - commission_out - payables_out

    vehicles_bought = [v for v in Vehicle.query.all() if df <= v.created_at.date() <= dt]
    investing_out = sum(v.acquisition_cost for v in vehicles_bought)
    net_investing = -investing_out

    loan_proceeds_in = db.session.query(func.sum(Loan.principal)).filter(
        in_range(Loan.start_date)).scalar() or 0
    loan_repay_out = db.session.query(func.sum(LoanPayment.amount)).filter(
        in_range(LoanPayment.payment_date)).scalar() or 0
    capital_in = db.session.query(func.sum(CapitalContribution.amount)).filter(
        in_range(CapitalContribution.contribution_date)).scalar() or 0
    drawings_out = db.session.query(func.sum(OwnerDrawing.amount)).filter(
        in_range(OwnerDrawing.drawing_date)).scalar() or 0
    net_financing = loan_proceeds_in - loan_repay_out + capital_in - drawings_out

    net_change = net_operating + net_investing + net_financing

    opening_cash = compute_financial_position(df - timedelta(days=1))['cash_and_equivalents']
    closing_cash = compute_financial_position(dt)['cash_and_equivalents']

    return render_template('reports/cash_flow.html',
        date_from=date_from_str, date_to=date_to_str,
        operating_in=operating_in, receivables_in=receivables_in,
        maint_out=maint_out, expenses_out=expenses_out,
        commission_out=commission_out, payables_out=payables_out, net_operating=net_operating,
        investing_out=investing_out, net_investing=net_investing, vehicles_bought=vehicles_bought,
        loan_proceeds_in=loan_proceeds_in, loan_repay_out=loan_repay_out,
        capital_in=capital_in, drawings_out=drawings_out, net_financing=net_financing,
        net_change=net_change, opening_cash=opening_cash, closing_cash=closing_cash)


@app.route('/reports/budget')
@login_required
@permission_required('reports')
def report_budget():
    today = date.today()
    month_str = request.args.get('month', today.strftime('%Y-%m'))
    try:
        month_start = datetime.strptime(month_str, '%Y-%m').date().replace(day=1)
    except ValueError:
        flash(f'"{month_str}" is not a valid month — showing {today.strftime("%Y-%m")} instead.', 'warning')
        month_start = today.replace(day=1)
        month_str = month_start.strftime('%Y-%m')
    month_end = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

    # Expense category labels are prefixed to avoid colliding with the
    # fixed Revenue/Maintenance keys below — an admin could otherwise
    # name a heading "Maintenance" (as in the worked example) and silently
    # shadow the MaintenanceLog-derived figure.
    actuals = {
        'Revenue': db.session.query(func.sum(DailyLog.gross_revenue)).filter(
            DailyLog.log_date.between(month_start, month_end)).scalar() or 0,
        'Maintenance': db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
            MaintenanceLog.log_date.between(month_start, month_end)).scalar() or 0,
    }
    for cat in ExpenseCategory.query.all():
        actuals[f'Expense: {cat.display_name}'] = db.session.query(func.sum(Expense.amount)).filter(
            Expense.category_id == cat.id,
            Expense.expense_date.between(month_start, month_end)).scalar() or 0

    budgets = {b.category: b.amount for b in Budget.query.filter_by(month=month_start).all()}
    all_categories = sorted(set(list(actuals.keys()) + list(budgets.keys())))
    rows = []
    for cat in all_categories:
        budget_amt = budgets.get(cat, 0)
        actual_amt = actuals.get(cat, 0)
        rows.append({'category': cat, 'budget': budget_amt, 'actual': actual_amt,
                     'variance': actual_amt - budget_amt})

    categories_available = ['Revenue', 'Maintenance'] + \
        [f'Expense: {c.display_name}' for c in ExpenseCategory.query.all()]
    return render_template('reports/budget.html', rows=rows, month=month_str,
        month_label=month_start.strftime('%B %Y'), categories=categories_available)


@app.route('/reports/budget/set', methods=['POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def budget_set():
    month_str = request.form.get('month', '')
    month_start = datetime.strptime(month_str, '%Y-%m').date().replace(day=1)
    category = request.form['category'].strip()
    amount = form_float(request.form, 'amount', min_value=0)

    existing = Budget.query.filter_by(category=category, month=month_start).first()
    if existing:
        existing.amount = amount
    else:
        db.session.add(Budget(category=category, month=month_start, amount=amount, created_by=current_user.id))
    log_audit('UPDATE', 'budgets', None, f'Set budget for {category} in {month_str}: {amount}')
    db.session.commit()
    flash(f'Budget for {category} set.', 'success')
    return redirect(url_for('report_budget', month=month_str))


@app.route('/reports/fuel-efficiency')
@login_required
@permission_required('reports')
def report_fuel_efficiency():
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    rows = []
    for v in Vehicle.query.order_by(Vehicle.registration).all():
        logs = FuelLog.query.filter(
            FuelLog.vehicle_id == v.id, FuelLog.log_date.between(df, dt),
            FuelLog.odometer.isnot(None)).order_by(FuelLog.odometer).all()
        if len(logs) < 2:
            continue
        segments = []
        for prev, curr in zip(logs, logs[1:]):
            distance = curr.odometer - prev.odometer
            if distance <= 0:
                continue
            # Skip fill-ups with no liters recorded (e.g. a diesel entry logged
            # only as a cash amount, or a mileage-only reading) — we genuinely
            # don't know how much fuel covered this distance, so it can't be
            # counted rather than being shown as an implausible 0 L/100km.
            if not curr.liters:
                continue
            # Fuel burned to cover this segment is the fill-up that ends it.
            l_per_100km = (curr.liters / distance) * 100
            segments.append({'from_date': prev.log_date, 'to_date': curr.log_date,
                             'distance': distance, 'liters': curr.liters,
                             'l_per_100km': l_per_100km})
        if not segments:
            continue
        total_distance = sum(s['distance'] for s in segments)
        total_liters = sum(s['liters'] for s in segments)
        # Aggregate consumption over the whole period (distance-weighted), which
        # is more accurate than averaging each segment's ratio equally.
        overall_l_per_100km = (total_liters / total_distance) * 100 if total_distance else 0
        km_per_liter = (total_distance / total_liters) if total_liters else 0
        rows.append({'vehicle': v, 'segments': segments,
                     'avg_l_per_100km': overall_l_per_100km,
                     'total_distance': total_distance, 'total_liters': total_liters,
                     'km_per_liter': km_per_liter})

    # Fleet-wide figures aggregated across all measured distance/fuel.
    fleet_distance = sum(r['total_distance'] for r in rows)
    fleet_liters = sum(r['total_liters'] for r in rows)
    fleet_avg = (fleet_liters / fleet_distance) * 100 if fleet_distance else 0
    return render_template('reports/fuel_efficiency.html', rows=rows, fleet_avg=fleet_avg,
        fleet_distance=fleet_distance, fleet_liters=fleet_liters,
        date_from=date_from_str, date_to=date_to_str)


@app.route('/reports/distance-travelled')
@login_required
@permission_required('reports')
def report_distance_travelled():
    d = query_single_date('date')
    date_str = d.strftime('%Y-%m-%d')

    rows = []
    for v in Vehicle.query.order_by(Vehicle.registration).all():
        # Odometer reading logged for this exact date (max, in case of more
        # than one fuel/mileage entry that day) — same basis as the Vehicle
        # Ledger and Fuel Efficiency report.
        odometer = db.session.query(func.max(FuelLog.odometer)).filter(
            FuelLog.vehicle_id == v.id, FuelLog.log_date == d,
            FuelLog.odometer.isnot(None)).scalar()

        distance = prev_odometer = prev_date = None
        if odometer is not None:
            prev = FuelLog.query.filter(
                FuelLog.vehicle_id == v.id, FuelLog.log_date < d,
                FuelLog.odometer.isnot(None)).order_by(FuelLog.log_date.desc()).first()
            if prev:
                prev_odometer, prev_date = prev.odometer, prev.log_date
                distance = odometer - prev_odometer

        rows.append({'vehicle': v, 'odometer': odometer,
                     'prev_odometer': prev_odometer, 'prev_date': prev_date,
                     'distance': distance})

    fleet_distance = sum(r['distance'] for r in rows if r['distance'] is not None)
    vehicles_reporting = sum(1 for r in rows if r['distance'] is not None)
    return render_template('reports/distance_travelled.html', rows=rows,
        date_str=date_str, fleet_distance=fleet_distance,
        vehicles_reporting=vehicles_reporting, fleet_size=len(rows))


@app.route('/reports/route-profitability')
@login_required
@permission_required('reports')
def report_route_profitability():
    df, dt = query_date_range()
    date_from_str, date_to_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    total_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date.between(df, dt)).scalar() or 0
    total_maintenance = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date.between(df, dt)).scalar() or 0
    total_costs = total_maintenance

    route_data = db.session.query(
        Route.id, Route.name, Route.start_point, Route.end_point,
        func.sum(DailyLog.gross_revenue).label('revenue'),
        func.sum(DailyLog.trips_completed).label('trips'),
        func.count(DailyLog.id).label('log_days'),
    ).join(DailyLog, Route.id == DailyLog.route_id).filter(
        DailyLog.log_date.between(df, dt)
    ).group_by(Route.id).all()

    rows = []
    for r in route_data:
        revenue = r.revenue or 0
        allocated_cost = (revenue / total_revenue * total_costs) if total_revenue else 0
        rows.append({
            'route': r, 'revenue': revenue, 'trips': r.trips or 0, 'log_days': r.log_days,
            'allocated_cost': allocated_cost, 'net_profit': revenue - allocated_cost,
        })

    # Entries with no route (e.g. from the Vehicle Ledger, which doesn't ask
    # for one) would otherwise be silently dropped by the INNER JOIN above,
    # leaving the per-route rows short of the fleet-wide total shown.
    unrouted = db.session.query(
        func.sum(DailyLog.gross_revenue).label('revenue'),
        func.sum(DailyLog.trips_completed).label('trips'),
        func.count(DailyLog.id).label('log_days'),
    ).filter(DailyLog.route_id.is_(None), DailyLog.log_date.between(df, dt)).first()
    if unrouted and unrouted.log_days:
        revenue = unrouted.revenue or 0
        allocated_cost = (revenue / total_revenue * total_costs) if total_revenue else 0
        rows.append({
            'route': None, 'revenue': revenue, 'trips': unrouted.trips or 0, 'log_days': unrouted.log_days,
            'allocated_cost': allocated_cost, 'net_profit': revenue - allocated_cost,
        })

    rows.sort(key=lambda x: x['net_profit'], reverse=True)

    return render_template('reports/route_profitability.html', rows=rows,
        total_revenue=total_revenue, total_costs=total_costs,
        date_from=date_from_str, date_to=date_to_str)


@app.route('/reports/financial-position')
@login_required
@permission_required('reports')
def report_financial_position():
    as_of = query_single_date('as_of')
    fp = compute_financial_position(as_of)
    return render_template('reports/financial_position.html',
        as_of_str=as_of.strftime('%Y-%m-%d'), **fp)


# ─────────────────────────────────────────────────────────────
# CSV Exports
# ─────────────────────────────────────────────────────────────
@app.route('/reports/export/daily-logs')
@login_required
@permission_required('reports')
def export_daily_logs():
    df = request.args.get('date_from', '')
    dt = request.args.get('date_to', '')
    q = DailyLog.query.order_by(DailyLog.log_date.desc())
    try:
        if df:
            q = q.filter(DailyLog.log_date >= parse_date(df))
        if dt:
            q = q.filter(DailyLog.log_date <= parse_date(dt))
    except ValueError as e:
        flash(str(e), 'danger')
        return redirect(url_for('driver_ledger'))

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['Date', 'Vehicle', 'Driver', 'Conductor', 'Route',
                'Trips', 'Gross Revenue (USD)', 'Garnish', 'Reason for Shortfall',
                'Entered By', 'Notes'])
    for log in q.all():
        w.writerow([log.log_date, log.vehicle.registration, log.driver.name if log.driver else '',
                    log.conductor.name if log.conductor else '',
                    log.route.name if log.route else '', log.trips_completed,
                    f'{log.gross_revenue:.2f}',
                    f'{log.garnish:.2f}' if log.garnish else '',
                    log.reason_for_shortfall or '',
                    log.creator.username if log.creator else '',
                    log.notes or ''])
    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = f'attachment; filename=daily_transactions_{date.today()}.csv'
    return resp


@app.route('/reports/export/income')
@login_required
@permission_required('reports')
def export_income():
    vehicle_id = request.args.get('vehicle_id', '')
    df, dt = query_date_range()
    df_str, dt_str = df.strftime('%Y-%m-%d'), dt.strftime('%Y-%m-%d')

    daily_q = DailyLog.query.filter(DailyLog.log_date.between(df, dt))
    fuel_q = FuelLog.query.filter(FuelLog.log_date.between(df, dt))
    maint_q = MaintenanceLog.query.filter(MaintenanceLog.log_date.between(df, dt))
    exp_q = Expense.query.filter(Expense.expense_date.between(df, dt))
    spares_q = StoreSale.query.filter(StoreSale.sale_date.between(df, dt))

    vehicle_label = 'Consolidated (fleet-wide)'
    if vehicle_id:
        v = Vehicle.query.get(vehicle_id)
        vehicle_label = f'{v.registration} — {v.make} {v.model}' if v else f'Vehicle #{vehicle_id}'
        daily_q = daily_q.filter(DailyLog.vehicle_id == vehicle_id)
        fuel_q = fuel_q.filter(FuelLog.vehicle_id == vehicle_id)
        maint_q = maint_q.filter(MaintenanceLog.vehicle_id == vehicle_id)
        exp_q = exp_q.filter(Expense.vehicle_id == vehicle_id)
        spares_q = spares_q.filter(StoreSale.vehicle_id == vehicle_id)
    else:
        spares_q = spares_q.filter(StoreSale.vehicle_id.isnot(None))

    daily = daily_q.order_by(DailyLog.log_date).all()
    fuel = fuel_q.order_by(FuelLog.log_date).all()
    maintenance = maint_q.order_by(MaintenanceLog.log_date).all()
    expenses = exp_q.order_by(Expense.expense_date).all()
    spares = spares_q.order_by(StoreSale.sale_date).all()

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['TRANSPORT FLEET INCOME STATEMENT (ZIMRA COMPLIANT)'])
    w.writerow([f'Scope: {vehicle_label}'])
    w.writerow([f'Period: {df_str} to {dt_str}'])
    w.writerow([f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")} by {current_user.username}'])
    w.writerow([])

    w.writerow(['REVENUE'])
    w.writerow(['Date', 'Vehicle', 'Route', 'Trips', 'Gross Revenue (USD)'])
    total_rev = 0
    for l in daily:
        w.writerow([l.log_date, l.vehicle.registration, l.route.name if l.route else '',
                    l.trips_completed, f'{l.gross_revenue:.2f}'])
        total_rev += l.gross_revenue
    w.writerow(['', '', '', 'TOTAL REVENUE', f'{total_rev:.2f}'])
    w.writerow([])

    w.writerow(['FUEL CONSUMPTION (not a cost — tracked in liters only)'])
    w.writerow(['Date', 'Vehicle', 'Liters', 'Supplier'])
    total_fuel_liters = 0
    for f in fuel:
        w.writerow([f.log_date, f.vehicle.registration, f.liters, f.supplier or ''])
        total_fuel_liters += f.liters
    w.writerow(['', '', 'TOTAL LITERS', f'{total_fuel_liters:.1f}'])
    w.writerow([])

    w.writerow(['MAINTENANCE EXPENSES'])
    w.writerow(['Date', 'Vehicle', 'Description', 'Parts (USD)', 'Labor (USD)', 'Total (USD)'])
    total_maint = 0
    for m in maintenance:
        w.writerow([m.log_date, m.vehicle.registration, m.description,
                    f'{m.parts_cost:.2f}', f'{m.labor_cost:.2f}', f'{m.total_cost:.2f}'])
        total_maint += m.total_cost
    w.writerow(['', '', '', '', 'TOTAL MAINTENANCE', f'{total_maint:.2f}'])
    w.writerow([])

    w.writerow(['OTHER EXPENSES'])
    w.writerow(['Date', 'Category', 'Vehicle', 'Description', 'Amount (USD)'])
    total_exp = 0
    for e in expenses:
        w.writerow([e.expense_date, e.category.display_name, e.vehicle.registration if e.vehicle else '(general)',
                    e.description or '', f'{e.amount:.2f}'])
        total_exp += e.amount
    w.writerow(['', '', '', 'TOTAL OTHER EXPENSES', f'{total_exp:.2f}'])
    w.writerow([])

    w.writerow(['SPARES SOLD TO COMPANY VEHICLES (booked as a maintenance expense on that vehicle)'])
    w.writerow(['Date', 'Part', 'Vehicle', 'Qty', 'Unit Price (USD)', 'Total (USD)'])
    total_spares = 0
    for s in spares:
        w.writerow([s.sale_date, s.part.name, s.vehicle.registration if s.vehicle else '',
                    s.quantity, f'{s.unit_price:.2f}', f'{s.total_amount:.2f}'])
        total_spares += s.total_amount
    w.writerow(['', '', '', '', 'TOTAL SPARES', f'{total_spares:.2f}'])
    w.writerow([])

    w.writerow(['NET PROFIT', f'{total_rev - total_maint - total_exp - total_spares:.2f}'])

    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    scope_suffix = f'_vehicle{vehicle_id}' if vehicle_id else '_consolidated'
    resp.headers['Content-Disposition'] = f'attachment; filename=income{scope_suffix}_{df_str}_to_{dt_str}.csv'
    return resp


@app.route('/reports/export/financial-position')
@login_required
@permission_required('reports')
def export_financial_position():
    as_of = query_single_date('as_of')
    fp = compute_financial_position(as_of)

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['STATEMENT OF FINANCIAL POSITION (SIMPLIFIED — SEE NOTES)'])
    w.writerow([f'As at: {as_of}'])
    w.writerow([f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")} by {current_user.username}'])
    w.writerow([])

    w.writerow(['ASSETS'])
    w.writerow(['Non-Current Assets — Vehicles'])
    w.writerow(['Vehicle', 'Cost (USD)', 'Accumulated Depreciation (USD)', 'Net Book Value (USD)'])
    for row in fp['vehicle_rows']:
        w.writerow([row['vehicle'].registration, f"{row['cost']:.2f}",
                    f"{row['accumulated_depreciation']:.2f}", f"{row['net_book_value']:.2f}"])
    w.writerow(['', '', 'TOTAL VEHICLES (NBV)', f"{fp['total_nbv']:.2f}"])
    w.writerow([])
    w.writerow(['Current Assets'])
    w.writerow(['Cash & Cash Equivalents', f"{fp['cash_and_equivalents']:.2f}"])
    w.writerow(['Receivables Outstanding', f"{fp['receivables_outstanding']:.2f}"])
    w.writerow([])
    w.writerow(['TOTAL ASSETS', f"{fp['total_assets']:.2f}"])
    w.writerow([])

    w.writerow(['LIABILITIES'])
    w.writerow(['Loans Outstanding', f"{fp['loans_outstanding']:.2f}"])
    w.writerow(['Payables Outstanding', f"{fp['payables_outstanding']:.2f}"])
    w.writerow(['Commission Payable (accrued, unpaid)', f"{fp['commission_payable']:.2f}"])
    w.writerow(['TOTAL LIABILITIES', f"{fp['total_liabilities']:.2f}"])
    w.writerow([])

    w.writerow(['EQUITY'])
    w.writerow(["Owner's Capital", f"{fp['owners_capital']:.2f}"])
    w.writerow(['Retained Earnings', f"{fp['retained_earnings']:.2f}"])
    w.writerow(['TOTAL EQUITY', f"{fp['total_equity']:.2f}"])
    w.writerow([])
    w.writerow(['TOTAL LIABILITIES + EQUITY', f"{fp['total_liabilities'] + fp['total_equity']:.2f}"])
    w.writerow([])
    w.writerow(['NOTE: Simplified statement. Vehicle purchases are a pure asset swap, not assumed '
                 'capital-funded — negative cash means a vehicle was bought before a Capital '
                 'Contribution or Loan was recorded to explain the funding. Vehicles are '
                 f"straight-line depreciated over {fp['useful_life']} years from the date each was "
                 'added to the fleet. Commission is accrued on all revenue earned, not just what has '
                 'been paid. Loan repayments are treated as pure principal reduction.'])

    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = f'attachment; filename=financial_position_{as_of}.csv'
    return resp


@app.route('/reports/export/distance-travelled')
@login_required
@permission_required('reports')
def export_distance_travelled():
    d = query_single_date('date')

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow([f'DISTANCE TRAVELLED — {d}'])
    w.writerow([])
    w.writerow(['Vehicle', 'Previous Reading Date', 'Previous Odometer (km)',
                'Odometer on Date (km)', 'Distance Travelled (km)'])
    for v in Vehicle.query.order_by(Vehicle.registration).all():
        odometer = db.session.query(func.max(FuelLog.odometer)).filter(
            FuelLog.vehicle_id == v.id, FuelLog.log_date == d,
            FuelLog.odometer.isnot(None)).scalar()
        prev_odometer = prev_date = distance = None
        if odometer is not None:
            prev = FuelLog.query.filter(
                FuelLog.vehicle_id == v.id, FuelLog.log_date < d,
                FuelLog.odometer.isnot(None)).order_by(FuelLog.log_date.desc()).first()
            if prev:
                prev_odometer, prev_date = prev.odometer, prev.log_date
                distance = odometer - prev_odometer
        w.writerow([v.registration, prev_date or '', prev_odometer or '',
                    odometer if odometer is not None else '',
                    f'{distance:.0f}' if distance is not None else ''])

    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = f'attachment; filename=distance_travelled_{d}.csv'
    return resp


# ─────────────────────────────────────────────────────────────
# Finance Ledger: Loans
# ─────────────────────────────────────────────────────────────
@app.route('/finance/loans')
@login_required
@permission_required('finance')
def loans_list():
    all_loans = Loan.query.order_by(Loan.start_date.desc()).all()
    return render_template('finance/loans.html', loans=all_loans, today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/loans/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def loan_add():
    if request.method == 'POST':
        loan = Loan(
            lender=request.form['lender'].strip(),
            principal=form_float(request.form, 'principal', min_value=0),
            interest_rate=form_float(request.form, 'interest_rate', required=False, default=0, min_value=0),
            start_date=parse_date(request.form['start_date']),
            term_months=form_int(request.form, 'term_months', required=False),
            status='active',
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(loan)
        db.session.flush()
        log_audit('CREATE', 'loans', loan.id, f'Added loan from {loan.lender} for {loan.principal}')
        db.session.commit()
        flash('Loan recorded.', 'success')
        return redirect(url_for('loans_list'))
    return render_template('finance/loan_form.html', today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/loans/<int:lid>/delete', methods=['POST'])
@login_required
@admin_required
def loan_delete(lid):
    loan = Loan.query.get_or_404(lid)
    log_audit('DELETE', 'loans', lid, f'Deleted loan from {loan.lender}')
    db.session.delete(loan)
    db.session.commit()
    flash('Loan deleted.', 'warning')
    return redirect(url_for('loans_list'))


@app.route('/finance/loans/<int:lid>/payment', methods=['POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def loan_payment_add(lid):
    loan = Loan.query.get_or_404(lid)
    payment = LoanPayment(
        loan_id=lid,
        payment_date=parse_date(request.form['payment_date']),
        amount=form_float(request.form, 'amount', min_value=0),
        notes=request.form.get('notes', '').strip(),
        created_by=current_user.id,
    )
    db.session.add(payment)
    log_audit('CREATE', 'loan_payments', None, f'Repayment of {payment.amount} on loan from {loan.lender}')
    db.session.commit()
    flash('Loan repayment recorded.', 'success')
    return redirect(url_for('loans_list'))


@app.route('/finance/loan-payments/<int:pid>/delete', methods=['POST'])
@login_required
@admin_required
def loan_payment_delete(pid):
    payment = LoanPayment.query.get_or_404(pid)
    log_audit('DELETE', 'loan_payments', pid, f'Deleted loan repayment of {payment.amount}')
    db.session.delete(payment)
    db.session.commit()
    flash('Loan repayment deleted.', 'warning')
    return redirect(url_for('loans_list'))


# ─────────────────────────────────────────────────────────────
# Finance Ledger: Payables (Accounts Payable)
# ─────────────────────────────────────────────────────────────
@app.route('/finance/payables')
@login_required
@permission_required('finance')
def payables_list():
    all_payables = Payable.query.order_by(Payable.invoice_date.desc()).all()
    return render_template('finance/payables.html', payables=all_payables)


@app.route('/finance/payables/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def payable_add():
    if request.method == 'POST':
        p = Payable(
            supplier_name=request.form['supplier_name'].strip(),
            description=request.form.get('description', '').strip(),
            amount=form_float(request.form, 'amount', min_value=0),
            invoice_date=parse_date(request.form['invoice_date']),
            due_date=parse_date(request.form.get('due_date')),
            created_by=current_user.id,
        )
        db.session.add(p)
        db.session.flush()
        log_audit('CREATE', 'payables', p.id, f'Payable to {p.supplier_name}: {p.amount}')
        db.session.commit()
        flash('Payable recorded.', 'success')
        return redirect(url_for('payables_list'))
    return render_template('finance/payable_form.html', today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/payables/<int:pid>/mark-paid', methods=['POST'])
@login_required
@permission_required('finance')
def payable_mark_paid(pid):
    p = Payable.query.get_or_404(pid)
    p.status = 'paid'
    p.paid_date = date.today()
    log_audit('UPDATE', 'payables', pid, f'Marked payable to {p.supplier_name} as paid')
    db.session.commit()
    flash('Payable marked as paid.', 'success')
    return redirect(url_for('payables_list'))


@app.route('/finance/payables/<int:pid>/delete', methods=['POST'])
@login_required
@admin_required
def payable_delete(pid):
    p = Payable.query.get_or_404(pid)
    log_audit('DELETE', 'payables', pid, f'Deleted payable to {p.supplier_name}')
    db.session.delete(p)
    db.session.commit()
    flash('Payable deleted.', 'warning')
    return redirect(url_for('payables_list'))


# ─────────────────────────────────────────────────────────────
# Finance Ledger: Receivables (Accounts Receivable)
# ─────────────────────────────────────────────────────────────
@app.route('/finance/receivables')
@login_required
@permission_required('finance')
def receivables_list():
    all_receivables = Receivable.query.order_by(Receivable.invoice_date.desc()).all()
    return render_template('finance/receivables.html', receivables=all_receivables)


@app.route('/finance/receivables/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def receivable_add():
    if request.method == 'POST':
        r = Receivable(
            client_name=request.form['client_name'].strip(),
            description=request.form.get('description', '').strip(),
            amount=form_float(request.form, 'amount', min_value=0),
            invoice_date=parse_date(request.form['invoice_date']),
            due_date=parse_date(request.form.get('due_date')),
            created_by=current_user.id,
        )
        db.session.add(r)
        db.session.flush()
        log_audit('CREATE', 'receivables', r.id, f'Receivable from {r.client_name}: {r.amount}')
        db.session.commit()
        flash('Receivable recorded.', 'success')
        return redirect(url_for('receivables_list'))
    return render_template('finance/receivable_form.html', today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/receivables/<int:rid>/mark-collected', methods=['POST'])
@login_required
@permission_required('finance')
def receivable_mark_collected(rid):
    r = Receivable.query.get_or_404(rid)
    r.status = 'collected'
    r.collected_date = date.today()
    log_audit('UPDATE', 'receivables', rid, f'Marked receivable from {r.client_name} as collected')
    db.session.commit()
    flash('Receivable marked as collected.', 'success')
    return redirect(url_for('receivables_list'))


@app.route('/finance/receivables/<int:rid>/delete', methods=['POST'])
@login_required
@admin_required
def receivable_delete(rid):
    r = Receivable.query.get_or_404(rid)
    log_audit('DELETE', 'receivables', rid, f'Deleted receivable from {r.client_name}')
    db.session.delete(r)
    db.session.commit()
    flash('Receivable deleted.', 'warning')
    return redirect(url_for('receivables_list'))


# ─────────────────────────────────────────────────────────────
# Finance Ledger: Owner's Capital & Drawings
# ─────────────────────────────────────────────────────────────
@app.route('/finance/capital')
@login_required
@permission_required('finance')
def capital_list():
    contributions = CapitalContribution.query.order_by(CapitalContribution.contribution_date.desc()).all()
    drawings = OwnerDrawing.query.order_by(OwnerDrawing.drawing_date.desc()).all()
    return render_template('finance/capital.html', contributions=contributions, drawings=drawings)


@app.route('/finance/capital/contributions/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def capital_contribution_add():
    if request.method == 'POST':
        c = CapitalContribution(
            contributor=request.form['contributor'].strip(),
            amount=form_float(request.form, 'amount', min_value=0),
            contribution_date=parse_date(request.form['contribution_date']),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(c)
        db.session.flush()
        log_audit('CREATE', 'capital_contributions', c.id, f'Capital contribution from {c.contributor}: {c.amount}')
        db.session.commit()
        flash('Capital contribution recorded.', 'success')
        return redirect(url_for('capital_list'))
    return render_template('finance/capital_contribution_form.html', today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/capital/contributions/<int:cid>/delete', methods=['POST'])
@login_required
@admin_required
def capital_contribution_delete(cid):
    c = CapitalContribution.query.get_or_404(cid)
    log_audit('DELETE', 'capital_contributions', cid, f'Deleted capital contribution from {c.contributor}')
    db.session.delete(c)
    db.session.commit()
    flash('Capital contribution deleted.', 'warning')
    return redirect(url_for('capital_list'))


@app.route('/finance/capital/drawings/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def owner_drawing_add():
    if request.method == 'POST':
        d = OwnerDrawing(
            amount=form_float(request.form, 'amount', min_value=0),
            drawing_date=parse_date(request.form['drawing_date']),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(d)
        db.session.flush()
        log_audit('CREATE', 'owner_drawings', d.id, f'Owner drawing: {d.amount}')
        db.session.commit()
        flash('Owner drawing recorded.', 'success')
        return redirect(url_for('capital_list'))
    return render_template('finance/owner_drawing_form.html', today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/capital/drawings/<int:did>/delete', methods=['POST'])
@login_required
@admin_required
def owner_drawing_delete(did):
    d = OwnerDrawing.query.get_or_404(did)
    log_audit('DELETE', 'owner_drawings', did, f'Deleted owner drawing of {d.amount}')
    db.session.delete(d)
    db.session.commit()
    flash('Owner drawing deleted.', 'warning')
    return redirect(url_for('capital_list'))


# ─────────────────────────────────────────────────────────────
# Finance Ledger: Expense Categories & Expenses
# ─────────────────────────────────────────────────────────────
@app.route('/finance/expenses')
@login_required
@permission_required('finance')
def expenses_list():
    page = request.args.get('page', 1, type=int)
    expenses = Expense.query.order_by(Expense.expense_date.desc()).paginate(page=page, per_page=20)
    headings = ExpenseCategory.query.filter_by(parent_id=None).order_by(ExpenseCategory.name).all()
    return render_template('finance/expenses.html', expenses=expenses, headings=headings)


@app.route('/finance/expenses/add', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def expense_add():
    if request.method == 'POST':
        e = Expense(
            category_id=form_int(request.form, 'category_id'),
            vehicle_id=form_int(request.form, 'vehicle_id', required=False),
            expense_date=parse_date(request.form['expense_date']),
            description=request.form.get('description', '').strip(),
            amount=form_float(request.form, 'amount', min_value=0),
            created_by=current_user.id,
        )
        db.session.add(e)
        db.session.flush()
        log_audit('CREATE', 'expenses', e.id, f'Expense of {e.amount}')
        db.session.commit()
        flash('Expense recorded.', 'success')
        return redirect(url_for('expenses_list'))
    headings = ExpenseCategory.query.filter_by(parent_id=None).order_by(ExpenseCategory.name).all()
    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    selected_category_id = request.args.get('new_category_id', '')
    return render_template('finance/expense_form.html', headings=headings, vehicles=all_vehicles,
                           selected_category_id=selected_category_id,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/finance/expenses/<int:eid>/delete', methods=['POST'])
@login_required
@admin_required
def expense_delete(eid):
    e = Expense.query.get_or_404(eid)
    log_audit('DELETE', 'expenses', eid, f'Deleted expense of {e.amount}')
    db.session.delete(e)
    db.session.commit()
    flash('Expense deleted.', 'warning')
    return redirect(url_for('expenses_list'))


@app.route('/finance/expense-categories/add', methods=['POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def expense_category_add():
    name = request.form.get('name', '').strip()
    if not name:
        raise ValueError('Category name is required.')
    parent_id = form_int(request.form, 'parent_id', required=False)
    parent = ExpenseCategory.query.get(parent_id) if parent_id else None
    if parent_id and not parent:
        raise ValueError('Selected heading does not exist.')

    redirect_to = request.form.get('redirect_to') or request.referrer or url_for('expenses_list')

    existing = ExpenseCategory.query.filter_by(name=name, parent_id=parent_id).first()
    if existing:
        flash(f'"{name}" already exists under {parent.name if parent else "top-level headings"}.', 'warning')
    else:
        new_cat = ExpenseCategory(name=name, parent_id=parent_id)
        db.session.add(new_cat)
        db.session.flush()
        log_audit('CREATE', 'expense_categories', new_cat.id, f'Added expense category {new_cat.display_name}')
        db.session.commit()
        flash(f'"{new_cat.display_name}" added.', 'success')
        sep = '&' if '?' in redirect_to else '?'
        redirect_to = f'{redirect_to}{sep}new_category_id={new_cat.id}'
    return redirect(redirect_to)


@app.route('/finance/expense-categories/<int:cid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('finance')
@handle_form_errors
def expense_category_edit(cid):
    cat = ExpenseCategory.query.get_or_404(cid)
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            raise ValueError('Category name is required.')
        parent_id = form_int(request.form, 'parent_id', required=False)
        if parent_id == cat.id:
            raise ValueError('A category cannot be its own heading.')
        if parent_id and any(child.id == parent_id for child in cat.children):
            raise ValueError('Cannot move a heading under one of its own sub-headings.')
        parent = ExpenseCategory.query.get(parent_id) if parent_id else None
        if parent_id and not parent:
            raise ValueError('Selected heading does not exist.')
        if parent_id and cat.children:
            raise ValueError('This category has sub-headings under it, so it cannot also become a sub-heading itself.')

        old_label = cat.display_name
        cat.name = name
        cat.parent_id = parent_id
        log_audit('UPDATE', 'expense_categories', cid, f'Renamed expense category {old_label} to {cat.display_name}')
        db.session.commit()
        flash(f'"{cat.display_name}" updated.', 'success')
        return redirect(url_for('expenses_list'))

    headings = ExpenseCategory.query.filter_by(parent_id=None).order_by(ExpenseCategory.name).all()
    return render_template('finance/expense_category_form.html', category=cat, headings=headings)


@app.route('/finance/expense-categories/<int:cid>/delete', methods=['POST'])
@login_required
@permission_required('finance')
def expense_category_delete(cid):
    cat = ExpenseCategory.query.get_or_404(cid)
    if cat.children:
        flash(f'Cannot delete "{cat.name}" — it has sub-headings under it. Delete those first.', 'danger')
        return redirect(request.referrer or url_for('expenses_list'))
    if Expense.query.filter_by(category_id=cid).first():
        flash(f'Cannot delete "{cat.display_name}" — expenses are recorded against it.', 'danger')
        return redirect(request.referrer or url_for('expenses_list'))
    name = cat.display_name
    log_audit('DELETE', 'expense_categories', cid, f'Deleted expense category {name}')
    db.session.delete(cat)
    db.session.commit()
    flash(f'"{name}" deleted.', 'warning')
    return redirect(request.referrer or url_for('expenses_list'))


# ─────────────────────────────────────────────────────────────
# Franchise Income — daily and weekly franchise fee collections, kept as
# two fully independent entities (see FranchiseDailyIncome/
# FranchiseWeeklyIncome above), each reconciled against its own income,
# expenditure and cash deposited.
# ─────────────────────────────────────────────────────────────
FRANCHISE_INCOME_EXPENSE_FIELDS = (
    ('exp_traffic_fines', 'Traffic Fines'),
    ('exp_facilitation_fees', 'Facilitation Fees'),
    ('exp_workshop', 'Workshop'),
    ('exp_wages', 'Wages'),
)


@app.route('/franchise/daily-income')
@login_required
@permission_required('franchise')
def franchise_daily_income_list():
    page = request.args.get('page', 1, type=int)
    entries = FranchiseDailyIncome.query.outerjoin(FranchiseVehicle).order_by(
        FranchiseDailyIncome.entry_date.desc(), FranchiseVehicle.number_plate).paginate(page=page, per_page=20)
    return render_template('franchise/daily_income_list.html', entries=entries)


@app.route('/franchise/daily-income/add', methods=['GET', 'POST'])
@login_required
@permission_required('franchise')
@handle_form_errors
def franchise_daily_income_add():
    if request.method == 'POST':
        entry_date = parse_date(request.form['entry_date'])
        vehicle_id = form_int(request.form, 'vehicle_id', required=False)
        vehicle = FranchiseVehicle.query.get(vehicle_id) if vehicle_id else None
        if vehicle_id and not vehicle:
            raise ValueError('Select a valid franchise vehicle.')
        label = vehicle.number_plate if vehicle else 'the whole franchise'
        if FranchiseDailyIncome.query.filter_by(entry_date=entry_date, vehicle_id=vehicle.id if vehicle else None).first():
            raise ValueError(f'A daily income entry for {label} on {entry_date} already exists — delete it first to re-enter.')
        amounts = {f: form_float(request.form, f, label=lbl, required=False, default=0)
                   for f, lbl in FRANCHISE_INCOME_EXPENSE_FIELDS}
        entry = FranchiseDailyIncome(
            entry_date=entry_date,
            vehicle_id=vehicle.id if vehicle else None,
            income=form_float(request.form, 'income', required=False, default=0),
            other_expenditure=form_float(request.form, 'other_expenditure', required=False, default=0),
            deposited=form_float(request.form, 'deposited', required=False, default=0),
            description=request.form.get('description', '').strip(),
            created_by=current_user.id,
            **amounts,
        )
        db.session.add(entry)
        db.session.flush()
        log_audit('CREATE', 'franchise_daily_income', entry.id,
                  f'Daily franchise income for {label} on {entry_date}: income {entry.income}, '
                  f'expenditure {entry.total_expenditure}, deposited {entry.deposited}')
        db.session.commit()
        flash('Daily franchise income recorded.', 'success')
        return redirect(url_for('franchise_daily_income_list'))
    vehicles = FranchiseVehicle.query.filter_by(status='active').order_by(FranchiseVehicle.franchisee_name).all()
    return render_template('franchise/daily_income_form.html', vehicles=vehicles, today=date.today().strftime('%Y-%m-%d'))


@app.route('/franchise/daily-income/<int:eid>/delete', methods=['POST'])
@login_required
@admin_required
def franchise_daily_income_delete(eid):
    entry = FranchiseDailyIncome.query.get_or_404(eid)
    log_audit('DELETE', 'franchise_daily_income', eid, f'Deleted daily franchise income entry for {entry.entry_date}')
    db.session.delete(entry)
    db.session.commit()
    flash('Daily franchise income entry deleted.', 'warning')
    return redirect(url_for('franchise_daily_income_list'))


@app.route('/franchise/weekly-income')
@login_required
@permission_required('franchise')
def franchise_weekly_income_list():
    page = request.args.get('page', 1, type=int)
    entries = FranchiseWeeklyIncome.query.outerjoin(FranchiseVehicle).order_by(
        FranchiseWeeklyIncome.week_start.desc(), FranchiseVehicle.number_plate).paginate(page=page, per_page=20)
    return render_template('franchise/weekly_income_list.html', entries=entries)


@app.route('/franchise/weekly-income/add', methods=['GET', 'POST'])
@login_required
@permission_required('franchise')
@handle_form_errors
def franchise_weekly_income_add():
    if request.method == 'POST':
        raw_date = parse_date(request.form['week_start'])
        week_start = raw_date - timedelta(days=raw_date.weekday())  # normalize to that week's Monday
        vehicle_id = form_int(request.form, 'vehicle_id', required=False)
        vehicle = FranchiseVehicle.query.get(vehicle_id) if vehicle_id else None
        if vehicle_id and not vehicle:
            raise ValueError('Select a valid franchise vehicle.')
        label = vehicle.number_plate if vehicle else 'the whole franchise'
        if FranchiseWeeklyIncome.query.filter_by(week_start=week_start, vehicle_id=vehicle.id if vehicle else None).first():
            raise ValueError(f'A weekly income entry for {label} for the week of {week_start} already exists — delete it first to re-enter.')
        amounts = {f: form_float(request.form, f, label=lbl, required=False, default=0)
                   for f, lbl in FRANCHISE_INCOME_EXPENSE_FIELDS}
        entry = FranchiseWeeklyIncome(
            week_start=week_start,
            vehicle_id=vehicle.id if vehicle else None,
            income=form_float(request.form, 'income', required=False, default=0),
            other_expenditure=form_float(request.form, 'other_expenditure', required=False, default=0),
            deposited=form_float(request.form, 'deposited', required=False, default=0),
            description=request.form.get('description', '').strip(),
            created_by=current_user.id,
            **amounts,
        )
        db.session.add(entry)
        db.session.flush()
        log_audit('CREATE', 'franchise_weekly_income', entry.id,
                  f'Weekly franchise income for {label} for week of {week_start}: income {entry.income}, '
                  f'expenditure {entry.total_expenditure}, deposited {entry.deposited}')
        db.session.commit()
        flash('Weekly franchise income recorded.', 'success')
        return redirect(url_for('franchise_weekly_income_list'))
    vehicles = FranchiseVehicle.query.filter_by(status='active').order_by(FranchiseVehicle.franchisee_name).all()
    return render_template('franchise/weekly_income_form.html', vehicles=vehicles, today=date.today().strftime('%Y-%m-%d'))


@app.route('/franchise/weekly-income/<int:eid>/delete', methods=['POST'])
@login_required
@admin_required
def franchise_weekly_income_delete(eid):
    entry = FranchiseWeeklyIncome.query.get_or_404(eid)
    log_audit('DELETE', 'franchise_weekly_income', eid, f'Deleted weekly franchise income entry for week of {entry.week_start}')
    db.session.delete(entry)
    db.session.commit()
    flash('Weekly franchise income entry deleted.', 'warning')
    return redirect(url_for('franchise_weekly_income_list'))


@app.route('/franchise/import/workbook', methods=['POST'])
@login_required
@permission_required('franchise')
def franchise_import_workbook():
    """Upload a whole multi-sheet workbook (like the franchise's own monthly
    file) in one go — every vehicle-collection grid in it is found by its
    own headers and sorted into FranchiseVehicle/FranchiseCollection, no
    column mapping needed and no sheet-naming convention required."""
    file = request.files.get('file')
    if not file or not file.filename:
        flash('Choose an Excel workbook to import.', 'danger')
        return redirect(url_for('franchise_vehicles'))
    try:
        summary = import_franchise_workbook(file, created_by=current_user.id)
    except ValueError as e:
        db.session.rollback()
        flash(str(e), 'danger')
        return redirect(url_for('franchise_vehicles'))

    total_imported = summary['vehicles_created'] + summary['collections_created']
    if total_imported:
        save_import_batch('franchise_workbook', file.filename, summary['total_rows'], total_imported,
                          [], summary['created_records'])
        log_audit('CREATE', 'franchise_workbook', None,
                  f"Imported workbook {file.filename}: {summary['vehicles_created']} new vehicle(s), "
                  f"{summary['collections_created']} collection(s)")
        db.session.commit()
    else:
        db.session.rollback()

    if total_imported:
        flash(f"Imported from {file.filename}: {summary['vehicles_created']} new vehicle(s), "
              f"{summary['collections_created']} collection(s) "
              f"({summary['collections_skipped']} already on file, skipped).", 'success')
    elif summary['total_rows']:
        flash(f"Nothing new to import from {file.filename} — every collection found "
              f"({summary['collections_skipped']}) is already on file.", 'warning')
    else:
        flash('No vehicle-collection tables were recognized in that workbook.', 'warning')

    return redirect(url_for('franchise_vehicles'))


# ─────────────────────────────────────────────────────────────
# Franchise Vehicles — the registry of third-party vehicles paying to
# operate under the franchise, and what each one has paid.
# ─────────────────────────────────────────────────────────────
@app.route('/franchise/vehicles')
@login_required
@permission_required('franchise')
def franchise_vehicles():
    vehicles = FranchiseVehicle.query.order_by(FranchiseVehicle.franchisee_name).all()
    return render_template('franchise/vehicles.html', vehicles=vehicles)


@app.route('/franchise/vehicles/add', methods=['GET', 'POST'])
@login_required
@permission_required('franchise')
@handle_form_errors
def franchise_vehicle_add():
    if request.method == 'POST':
        number_plate = request.form.get('number_plate', '').strip().upper()
        if not number_plate:
            raise ValueError('Number plate is required.')
        check_unique(FranchiseVehicle, 'number_plate', number_plate, label='Number plate')
        vehicle = FranchiseVehicle(
            number_plate=number_plate,
            franchisee_name=request.form.get('franchisee_name', '').strip(),
            status=request.form.get('status', 'active'),
            agreed_amount=form_float(request.form, 'agreed_amount', label='Agreed franchise amount', required=False, default=0, min_value=0),
            amount_owed=form_float(request.form, 'amount_owed', label='Amount owed', required=False, default=0, min_value=0),
            notes=request.form.get('notes', '').strip(),
        )
        if not vehicle.franchisee_name:
            raise ValueError('Franchisee name is required.')
        db.session.add(vehicle)
        db.session.flush()
        log_audit('CREATE', 'franchise_vehicles', vehicle.id,
                  f'Added franchise vehicle {vehicle.number_plate} ({vehicle.franchisee_name})')
        db.session.commit()
        flash('Franchise vehicle added.', 'success')
        return redirect(url_for('franchise_vehicles'))
    return render_template('franchise/vehicle_form.html', vehicle=None, action='Add')


@app.route('/franchise/vehicles/<int:vid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('franchise')
@handle_form_errors
def franchise_vehicle_edit(vid):
    vehicle = FranchiseVehicle.query.get_or_404(vid)
    if request.method == 'POST':
        number_plate = request.form.get('number_plate', '').strip().upper()
        if not number_plate:
            raise ValueError('Number plate is required.')
        check_unique(FranchiseVehicle, 'number_plate', number_plate, label='Number plate', exclude_id=vehicle.id)
        franchisee_name = request.form.get('franchisee_name', '').strip()
        if not franchisee_name:
            raise ValueError('Franchisee name is required.')
        vehicle.number_plate = number_plate
        vehicle.franchisee_name = franchisee_name
        vehicle.status = request.form.get('status', 'active')
        vehicle.agreed_amount = form_float(request.form, 'agreed_amount', label='Agreed franchise amount', required=False, default=0, min_value=0)
        vehicle.amount_owed = form_float(request.form, 'amount_owed', label='Amount owed', required=False, default=0, min_value=0)
        vehicle.notes = request.form.get('notes', '').strip()
        log_audit('UPDATE', 'franchise_vehicles', vehicle.id, f'Updated franchise vehicle {vehicle.number_plate}')
        db.session.commit()
        flash('Franchise vehicle updated.', 'success')
        return redirect(url_for('franchise_vehicles'))
    return render_template('franchise/vehicle_form.html', vehicle=vehicle, action='Edit')


@app.route('/franchise/vehicles/<int:vid>/delete', methods=['POST'])
@login_required
@admin_required
def franchise_vehicle_delete(vid):
    vehicle = FranchiseVehicle.query.get_or_404(vid)
    if FranchiseCollection.query.filter_by(vehicle_id=vid).first():
        flash('Cannot delete this vehicle — it has collection history. Mark it Inactive instead.', 'danger')
        return redirect(url_for('franchise_vehicles'))
    log_audit('DELETE', 'franchise_vehicles', vid, f'Deleted franchise vehicle {vehicle.number_plate}')
    db.session.delete(vehicle)
    db.session.commit()
    flash('Franchise vehicle deleted.', 'warning')
    return redirect(url_for('franchise_vehicles'))


# ─────────────────────────────────────────────────────────────
# Franchise Collections — what each franchise vehicle paid, and when.
# ─────────────────────────────────────────────────────────────
@app.route('/franchise/collections')
@login_required
@permission_required('franchise')
def franchise_collections():
    page = request.args.get('page', 1, type=int)
    vehicle_id = request.args.get('vehicle_id', type=int)
    frequency = request.args.get('frequency', '')
    q = FranchiseCollection.query
    if vehicle_id:
        q = q.filter_by(vehicle_id=vehicle_id)
    if frequency in ('daily', 'weekly'):
        q = q.filter_by(frequency=frequency)
    entries = q.order_by(FranchiseCollection.entry_date.desc()).paginate(page=page, per_page=30)
    vehicles = FranchiseVehicle.query.order_by(FranchiseVehicle.franchisee_name).all()
    return render_template('franchise/collections.html', entries=entries, vehicles=vehicles,
                           vehicle_id=vehicle_id, frequency=frequency)


@app.route('/franchise/collections/add', methods=['GET', 'POST'])
@login_required
@permission_required('franchise')
@handle_form_errors
def franchise_collection_add():
    if request.method == 'POST':
        vehicle_id = form_int(request.form, 'vehicle_id')
        vehicle = FranchiseVehicle.query.get(vehicle_id)
        if not vehicle:
            raise ValueError('Select a valid franchise vehicle.')
        frequency = request.form.get('frequency', '')
        if frequency not in ('daily', 'weekly'):
            raise ValueError('Frequency must be Daily or Weekly.')
        collection = FranchiseCollection(
            vehicle_id=vehicle.id,
            entry_date=parse_date(request.form['entry_date']),
            frequency=frequency,
            amount=form_float(request.form, 'amount', min_value=0),
            expense=form_float(request.form, 'expense', required=False, default=0, min_value=0),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(collection)
        db.session.flush()
        log_audit('CREATE', 'franchise_collections', collection.id,
                  f'{vehicle.number_plate} paid {collection.amount} ({frequency}), expense {collection.expense}, '
                  f'on {collection.entry_date}')
        db.session.commit()
        flash('Collection recorded.', 'success')
        return redirect(url_for('franchise_collections'))
    vehicles = FranchiseVehicle.query.filter_by(status='active').order_by(FranchiseVehicle.franchisee_name).all()
    return render_template('franchise/collection_form.html', vehicles=vehicles,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/franchise/collections/<int:cid>/delete', methods=['POST'])
@login_required
@admin_required
def franchise_collection_delete(cid):
    collection = FranchiseCollection.query.get_or_404(cid)
    log_audit('DELETE', 'franchise_collections', cid,
              f'Deleted collection of {collection.amount} for {collection.vehicle.number_plate} on {collection.entry_date}')
    db.session.delete(collection)
    db.session.commit()
    flash('Collection entry deleted.', 'warning')
    return redirect(url_for('franchise_collections'))


def _collections_by_day(frequency, df, dt):
    """Group FranchiseCollection entries of one frequency into per-date rows
    (vehicle count + total collected) — shared by the Daily and Weekly
    Franchise report pages, which are identical in shape and differ only in
    which frequency they show."""
    entries = FranchiseCollection.query.filter(
        FranchiseCollection.entry_date.between(df, dt), FranchiseCollection.frequency == frequency
    ).order_by(FranchiseCollection.entry_date.desc(), FranchiseCollection.id.desc()).all()

    days = {}
    for c in entries:
        days.setdefault(c.entry_date, []).append(c)
    day_rows = [
        dict(entry_date=d, collections=day_entries, vehicle_count=len(day_entries),
             total=sum(c.amount for c in day_entries), total_expense=sum(c.expense for c in day_entries),
             net=sum(c.net for c in day_entries))
        for d, day_entries in sorted(days.items(), reverse=True)
    ]
    return day_rows, sum(c.amount for c in entries), sum(c.expense for c in entries)


@app.route('/reports/franchise/daily-collections')
@login_required
@permission_required('franchise')
def report_franchise_daily_collections():
    """Daily Franchise — per-date rollup of vehicles paying the daily fee."""
    df, dt = query_date_range()
    day_rows, total, total_expense = _collections_by_day('daily', df, dt)
    return render_template('franchise/daily_collections.html', title='Daily Franchise Collections',
                           frequency='daily', days=day_rows, total=total, total_expense=total_expense,
                           net=total - total_expense,
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


@app.route('/reports/franchise/weekly-collections')
@login_required
@permission_required('franchise')
def report_franchise_weekly_collections():
    """Weekly Franchise — per-date rollup of vehicles paying the weekly fee."""
    df, dt = query_date_range()
    day_rows, total, total_expense = _collections_by_day('weekly', df, dt)
    return render_template('franchise/daily_collections.html', title='Weekly Franchise Collections',
                           frequency='weekly', days=day_rows, total=total, total_expense=total_expense,
                           net=total - total_expense,
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


@app.route('/reports/franchise/analysis')
@login_required
@permission_required('franchise')
def report_franchise_analysis():
    """Franchise Analysis — for each date, Daily Franchise collected vs.
    Weekly Franchise collected vs. their combined total, so the two payment
    populations can be compared day by day."""
    df, dt = query_date_range()
    entries = FranchiseCollection.query.filter(FranchiseCollection.entry_date.between(df, dt)).all()

    days = {}
    for c in entries:
        bucket = days.setdefault(c.entry_date, {'daily': 0.0, 'weekly': 0.0})
        bucket[c.frequency] += c.amount
    day_rows = [
        dict(entry_date=d, weekday=d.strftime('%A'), daily=b['daily'], weekly=b['weekly'],
             total=b['daily'] + b['weekly'])
        for d, b in sorted(days.items())
    ]
    totals = dict(daily=sum(r['daily'] for r in day_rows), weekly=sum(r['weekly'] for r in day_rows),
                  total=sum(r['total'] for r in day_rows))
    return render_template('franchise/analysis.html', title='Franchise Analysis',
                           days=day_rows, totals=totals,
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


def _income_entry_totals(entries):
    """Sum a list of FranchiseDailyIncome or FranchiseWeeklyIncome entries
    (identical shape) into the same fields as a single entry, so a period
    total can be rendered with the same formatting as one row."""
    return dict(
        income=sum(e.income for e in entries),
        exp_traffic_fines=sum(e.exp_traffic_fines for e in entries),
        exp_facilitation_fees=sum(e.exp_facilitation_fees for e in entries),
        exp_workshop=sum(e.exp_workshop for e in entries),
        exp_wages=sum(e.exp_wages for e in entries),
        other_expenditure=sum(e.other_expenditure for e in entries),
        total_expenditure=sum(e.total_expenditure for e in entries),
        cash_in_hand=sum(e.cash_in_hand for e in entries),
        deposited=sum(e.deposited for e in entries),
        variance=sum(e.variance for e in entries),
    )


def _group_income_by_period(entries, period_attr):
    """Group entries (FranchiseDailyIncome or FranchiseWeeklyIncome) by
    their date/week_start, summing every vehicle's figures into one row per
    period — matching the franchise's own paper reconciliation schedule,
    which records a single combined total per date/week rather than a
    per-vehicle breakdown. Returns rows sorted by period, each carrying the
    period value, how many vehicles fed into it, and the summed totals."""
    by_period = {}
    for e in entries:
        by_period.setdefault(getattr(e, period_attr), []).append(e)
    return [
        dict(period=period, vehicle_count=len(period_entries), **_income_entry_totals(period_entries))
        for period, period_entries in sorted(by_period.items())
    ]


@app.route('/reports/franchise/reconciliation')
@login_required
@permission_required('franchise')
def report_franchise_reconciliation():
    """Reconciliation Schedule — daily and weekly franchise income are two
    independent streams, so this page shows two self-contained sections
    (each with its own per-date/week breakdown and subtotal) rather than
    one table with columns from both. Each row is a combined total across
    every vehicle for that date/week, matching the franchise's own paper
    schedule — per-vehicle detail lives on the Daily/Weekly Income list
    pages instead."""
    df, dt = query_date_range()
    daily_entries = FranchiseDailyIncome.query.filter(FranchiseDailyIncome.entry_date.between(df, dt)).all()
    weekly_entries = FranchiseWeeklyIncome.query.filter(FranchiseWeeklyIncome.week_start.between(df, dt)).all()

    daily_rows = _group_income_by_period(daily_entries, 'entry_date')
    weekly_rows = _group_income_by_period(weekly_entries, 'week_start')

    daily_totals = _income_entry_totals(daily_entries)
    weekly_totals = _income_entry_totals(weekly_entries)
    combined_totals = {k: daily_totals[k] + weekly_totals[k] for k in daily_totals}
    combined_totals['net_profit'] = combined_totals['income'] - combined_totals['total_expenditure']

    return render_template('franchise/reconciliation.html', title='Franchise Collection Reconciliation Schedule',
                           daily_rows=daily_rows, daily_totals=daily_totals,
                           weekly_rows=weekly_rows, weekly_totals=weekly_totals,
                           combined_totals=combined_totals,
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


@app.route('/reports/franchise/weekly')
@login_required
@permission_required('franchise')
def report_franchise_weekly():
    """Weekly Analysis — for each Monday-Sunday week in range, roll up that
    week's Daily Income entries and combine them with the standalone Weekly
    Income entry for that week (if any), so daily and weekly figures — kept
    as separate entities — can still be seen combined at the week level."""
    df, dt = query_date_range()
    daily_entries = FranchiseDailyIncome.query.filter(FranchiseDailyIncome.entry_date.between(df, dt)) \
        .order_by(FranchiseDailyIncome.entry_date.asc()).all()
    weekly_entries = FranchiseWeeklyIncome.query.filter(FranchiseWeeklyIncome.week_start.between(df, dt)) \
        .order_by(FranchiseWeeklyIncome.week_start.asc()).all()
    # A week can have several weekly entries — one per franchisee — so group
    # into lists rather than assuming a single entry per week_start.
    weekly_by_week = {}
    for e in weekly_entries:
        weekly_by_week.setdefault(e.week_start, []).append(e)

    daily_by_week = {}
    for e in daily_entries:
        week_start = e.entry_date - timedelta(days=e.entry_date.weekday())
        daily_by_week.setdefault(week_start, []).append(e)

    week_starts = sorted(set(daily_by_week.keys()) | set(weekly_by_week.keys()))
    week_rows = []
    for start in week_starts:
        week_daily_entries = daily_by_week.get(start, [])
        daily_totals = _income_entry_totals(week_daily_entries)
        weekly_totals = _income_entry_totals(weekly_by_week.get(start, []))
        week_rows.append(dict(
            week_start=start, week_end=start + timedelta(days=6), days=len(week_daily_entries),
            daily_income=daily_totals['income'], daily_expenditure=daily_totals['total_expenditure'],
            weekly_income=weekly_totals['income'], weekly_expenditure=weekly_totals['total_expenditure'],
            total_income=daily_totals['income'] + weekly_totals['income'],
            total_expenditure=daily_totals['total_expenditure'] + weekly_totals['total_expenditure'],
            net_profit=(daily_totals['income'] + weekly_totals['income'])
                - (daily_totals['total_expenditure'] + weekly_totals['total_expenditure']),
        ))

    totals = dict(
        daily_income=sum(r['daily_income'] for r in week_rows),
        daily_expenditure=sum(r['daily_expenditure'] for r in week_rows),
        weekly_income=sum(r['weekly_income'] for r in week_rows),
        weekly_expenditure=sum(r['weekly_expenditure'] for r in week_rows),
        total_income=sum(r['total_income'] for r in week_rows),
        total_expenditure=sum(r['total_expenditure'] for r in week_rows),
        net_profit=sum(r['net_profit'] for r in week_rows),
    )
    return render_template('franchise/weekly_analysis.html', title='Franchise Weekly Analysis',
                           weeks=week_rows, totals=totals,
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


@app.route('/reports/franchise/consolidated')
@login_required
@permission_required('franchise')
def report_franchise_consolidated():
    """Consolidated P&L — single summary of income, expenditure by category,
    and the cash reconciliation for the whole period, combining the daily
    and weekly income entities."""
    df, dt = query_date_range()
    daily_entries = FranchiseDailyIncome.query.filter(FranchiseDailyIncome.entry_date.between(df, dt)).all()
    weekly_entries = FranchiseWeeklyIncome.query.filter(FranchiseWeeklyIncome.week_start.between(df, dt)).all()

    daily_totals = _income_entry_totals(daily_entries)
    weekly_totals = _income_entry_totals(weekly_entries)
    totals = {k: daily_totals[k] + weekly_totals[k] for k in daily_totals}
    totals['income_daily'] = daily_totals['income']
    totals['income_weekly'] = weekly_totals['income']
    totals['total_income'] = totals.pop('income')
    totals['net_profit'] = totals['total_income'] - totals['total_expenditure']

    return render_template('franchise/consolidated.html', title='Consolidated Franchise P&L Statement',
                           totals=totals, entry_count=len(daily_entries) + len(weekly_entries),
                           date_from=df.strftime('%Y-%m-%d'), date_to=dt.strftime('%Y-%m-%d'))


# ─────────────────────────────────────────────────────────────
# Compliance
# ─────────────────────────────────────────────────────────────
@app.route('/compliance')
@login_required
@permission_required('compliance')
def compliance():
    today = date.today()
    threshold = today + timedelta(days=30)
    expired = VehicleDocument.query.filter(
        VehicleDocument.expiry_date < today).order_by(VehicleDocument.expiry_date).all()
    expiring = VehicleDocument.query.filter(
        VehicleDocument.expiry_date.between(today, threshold)).order_by(
        VehicleDocument.expiry_date).all()
    valid = VehicleDocument.query.filter(
        VehicleDocument.expiry_date > threshold).order_by(
        VehicleDocument.expiry_date).all()

    # Insurance is a Vehicle field, not a VehicleDocument row, but belongs on
    # the same expiry tracker — wrapped in plain dicts (duck-typed the same
    # shape the template already expects: vehicle/doc_type/expiry_date/
    # days_to_expiry) so it sorts and displays alongside real documents.
    for v in Vehicle.query.filter(Vehicle.insurance_expiry.isnot(None)).all():
        entry = {'vehicle': v, 'doc_type': 'Insurance', 'expiry_date': v.insurance_expiry,
                  'days_to_expiry': v.insurance_days_to_expiry}
        if v.insurance_status == 'expired':
            expired.append(entry)
        elif v.insurance_status == 'warning':
            expiring.append(entry)
        else:
            valid.append(entry)
    expired.sort(key=lambda d: d['expiry_date'] if isinstance(d, dict) else d.expiry_date)
    expiring.sort(key=lambda d: d['expiry_date'] if isinstance(d, dict) else d.expiry_date)
    valid.sort(key=lambda d: d['expiry_date'] if isinstance(d, dict) else d.expiry_date)

    return render_template('compliance/index.html',
        expired=expired, expiring=expiring, valid=valid, today=today)


# ─────────────────────────────────────────────────────────────
# Users (Admin)
# ─────────────────────────────────────────────────────────────
@app.route('/users')
@login_required
@admin_required
def users():
    all_users = User.query.order_by(User.username).all()
    return render_template('users/index.html', users=all_users)


@app.route('/users/add', methods=['GET', 'POST'])
@login_required
@admin_required
@handle_form_errors
def user_add():
    if request.method == 'POST':
        password = request.form.get('password', '')
        if len(password) < 6:
            raise ValueError('Password must be at least 6 characters.')
        username = request.form['username'].strip().lower()
        email = request.form['email'].strip().lower()
        check_unique(User, 'username', username)
        check_unique(User, 'email', email)
        u = User(
            username=username,
            email=email,
            role=request.form.get('role', 'manager'),
        )
        u.set_password(password)
        db.session.add(u)
        db.session.flush()
        log_audit('CREATE', 'users', u.id, f'Created user {u.username}')
        db.session.commit()
        flash(f'User "{u.username}" created.', 'success')
        return redirect(url_for('users'))
    return render_template('users/form.html')


@app.route('/users/<int:uid>/toggle', methods=['POST'])
@login_required
@admin_required
def user_toggle(uid):
    u = User.query.get_or_404(uid)
    if u.id == current_user.id:
        flash('Cannot deactivate your own account.', 'danger')
        return redirect(url_for('users'))
    u.is_active = not u.is_active
    log_audit('UPDATE', 'users', uid,
              f'Set user {u.username} active={u.is_active}')
    db.session.commit()
    flash(f'User {u.username} {"activated" if u.is_active else "deactivated"}.', 'info')
    return redirect(url_for('users'))


@app.route('/users/<int:uid>/permissions', methods=['GET', 'POST'])
@login_required
@admin_required
def user_permissions(uid):
    u = User.query.get_or_404(uid)
    if u.role == 'admin':
        flash('Admin users always have full access — no permission configuration needed.', 'info')
        return redirect(url_for('users'))
    if request.method == 'POST':
        granted = request.form.getlist('permissions')
        u.permissions = json.dumps(granted)
        log_audit('UPDATE', 'users', uid,
                  f'Updated permissions for {u.username}: {", ".join(granted) or "none"}')
        db.session.commit()
        flash(f'Permissions updated for {u.username}.', 'success')
        return redirect(url_for('users'))
    return render_template('users/permissions.html', u=u,
                           all_permissions=PERMISSIONS,
                           current_perms=u.get_permissions())


@app.route('/users/<int:uid>/reset-password', methods=['POST'])
@login_required
@admin_required
def user_reset_password(uid):
    u = User.query.get_or_404(uid)
    new_pw = request.form.get('new_password', '').strip()
    if len(new_pw) < 6:
        flash('Password must be at least 6 characters.', 'danger')
        return redirect(url_for('users'))
    u.set_password(new_pw)
    log_audit('UPDATE', 'users', uid, f'Reset password for {u.username}')
    db.session.commit()
    flash(f'Password reset for {u.username}.', 'success')
    return redirect(url_for('users'))


# ─────────────────────────────────────────────────────────────
# Audit Log
# ─────────────────────────────────────────────────────────────
@app.route('/audit')
@login_required
@admin_required
def audit_log():
    page = request.args.get('page', 1, type=int)
    logs = AuditLog.query.order_by(AuditLog.timestamp.desc()).paginate(page=page, per_page=30)
    return render_template('audit/index.html', logs=logs)


# ─────────────────────────────────────────────────────────────
# Import Batches — audit trail, quarantine error export, revert
# ─────────────────────────────────────────────────────────────
IMPORT_TARGET_PERMISSIONS = {
    'ledger': ('daily_logs', 'crew_portal'),
    'franchise_workbook': ('franchise',),
}

IMPORT_REVERT_MODELS = {
    'daily_logs': DailyLog,
    'fuel_logs': FuelLog,
    'franchise_vehicles': FranchiseVehicle,
    'franchise_collections': FranchiseCollection,
}


@app.route('/imports')
@login_required
@admin_required
def import_history():
    page = request.args.get('page', 1, type=int)
    batches = ImportBatch.query.order_by(ImportBatch.created_at.desc()).paginate(page=page, per_page=30)
    return render_template('imports/history.html', batches=batches)


@app.route('/imports/<int:batch_id>/errors.csv')
@login_required
def import_batch_errors_csv(batch_id):
    batch = ImportBatch.query.get_or_404(batch_id)
    perms = IMPORT_TARGET_PERMISSIONS.get(batch.target_type, ())
    if not any(current_user.has_permission(p) for p in perms):
        flash('You do not have permission to access that import.', 'danger')
        return redirect(first_permitted_url(current_user))

    rows = batch.error_row_list
    out = io.StringIO()
    if rows:
        writer = csv.DictWriter(out, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = (
        f'attachment; filename={batch.target_type}_import_errors_{batch.id}.csv')
    return resp


@app.route('/imports/<int:batch_id>/revert', methods=['POST'])
@login_required
@admin_required
def import_batch_revert(batch_id):
    batch = ImportBatch.query.get_or_404(batch_id)
    if batch.status == 'reverted':
        flash('That import was already reverted.', 'warning')
        return redirect(url_for('import_history'))

    deleted = 0
    for rec in batch.records:
        model = IMPORT_REVERT_MODELS.get(rec.target_table)
        obj = model.query.get(rec.record_id) if model else None
        if obj:
            db.session.delete(obj)
            deleted += 1

    batch.status = 'reverted'
    batch.reverted_at = datetime.now(timezone.utc)
    batch.reverted_by = current_user.id
    log_audit('DELETE', batch.target_type, batch.id,
              f'Reverted import batch #{batch.id} ({batch.filename}) — removed {deleted} record(s)')
    db.session.commit()
    flash(f'Reverted import #{batch.id} — removed {deleted} record(s).', 'success')
    return redirect(url_for('import_history'))


# ─────────────────────────────────────────────────────────────
# API (chart data)
# ─────────────────────────────────────────────────────────────
@app.route('/api/revenue/monthly')
@login_required
@permission_required('dashboard')
def api_revenue_monthly():
    today = date.today()
    data = []
    for i in range(5, -1, -1):
        m = today.month - i
        y = today.year
        while m <= 0:
            m += 12
            y -= 1
        start = date(y, m, 1)
        end = date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)
        rev = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
            DailyLog.log_date >= start, DailyLog.log_date < end).scalar() or 0
        maint = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
            MaintenanceLog.log_date >= start, MaintenanceLog.log_date < end).scalar() or 0
        data.append({
            'month': start.strftime('%b %Y'),
            'revenue': float(rev),
            'expenses': float(maint),
            'profit': float(rev - maint),
        })
    return jsonify(data)


@app.route('/api/vehicles/performance')
@login_required
@permission_required('dashboard')
def api_vehicle_performance():
    today = date.today()
    month_start = today.replace(day=1)
    result = db.session.query(
        Vehicle.registration,
        func.sum(DailyLog.gross_revenue).label('revenue'),
        func.count(DailyLog.id).label('days'),
    ).join(DailyLog, Vehicle.id == DailyLog.vehicle_id).filter(
        DailyLog.log_date >= month_start
    ).group_by(Vehicle.id).all()
    return jsonify([{'vehicle': r.registration, 'revenue': float(r.revenue or 0),
                     'days': r.days} for r in result])


@app.route('/api/expenses/breakdown')
@login_required
@permission_required('dashboard')
def api_expenses_breakdown():
    today = date.today()
    m_start = today.replace(day=1)
    maint = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date >= m_start).scalar() or 0
    return jsonify({'maintenance': float(maint)})


# ─────────────────────────────────────────────────────────────
# WhatsApp Webhook stub
# ─────────────────────────────────────────────────────────────
@app.route('/api/whatsapp/webhook', methods=['POST'])
@csrf.exempt
def whatsapp_webhook():
    # Future: parse Twilio/Vonage WhatsApp messages and create daily logs
    return jsonify({'status': 'received'})


# ─────────────────────────────────────────────────────────────
# Template filters
# ─────────────────────────────────────────────────────────────
@app.template_filter('currency')
def currency_filter(value):
    if value is None:
        return '$0.00'
    return f'${value:,.2f}'


@app.template_filter('pct')
def pct_filter(value):
    return f'{value:.1f}%' if value is not None else '0.0%'


# ─────────────────────────────────────────────────────────────
# Bootstrap DB + first admin
# ─────────────────────────────────────────────────────────────
def migrate_db():
    from sqlalchemy import inspect, text
    inspector = inspect(db.engine)
    user_cols = [c['name'] for c in inspector.get_columns('users')]
    with db.engine.connect() as conn:
        if 'permissions' not in user_cols:
            conn.execute(text("ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT '[]'"))
        if 'driver_id' not in user_cols:
            conn.execute(text("ALTER TABLE users ADD COLUMN driver_id INTEGER REFERENCES drivers(id)"))

        # Depots were removed — drop the leftover columns/table from any DB that has them.
        for table in ('vehicles', 'drivers', 'routes'):
            cols = [c['name'] for c in inspector.get_columns(table)]
            if 'depot_id' in cols:
                conn.execute(text(f"ALTER TABLE {table} DROP COLUMN depot_id"))

        vehicle_cols = [c['name'] for c in inspector.get_columns('vehicles')]
        if 'fuel_type' not in vehicle_cols:
            conn.execute(text("ALTER TABLE vehicles ADD COLUMN fuel_type VARCHAR(10) DEFAULT 'diesel'"))
        if 'daily_target' not in vehicle_cols:
            conn.execute(text("ALTER TABLE vehicles ADD COLUMN daily_target FLOAT"))
        if 'insurance_provider' not in vehicle_cols:
            conn.execute(text("ALTER TABLE vehicles ADD COLUMN insurance_provider VARCHAR(100)"))
        if 'insurance_policy_number' not in vehicle_cols:
            conn.execute(text("ALTER TABLE vehicles ADD COLUMN insurance_policy_number VARCHAR(100)"))
        if 'insurance_expiry' not in vehicle_cols:
            conn.execute(text("ALTER TABLE vehicles ADD COLUMN insurance_expiry DATE"))
        if inspector.has_table('depots'):
            conn.execute(text("DROP TABLE depots"))

        if inspector.has_table('expense_categories'):
            exp_cat_cols = [c['name'] for c in inspector.get_columns('expense_categories')]
            if 'parent_id' not in exp_cat_cols:
                conn.execute(text(
                    "ALTER TABLE expense_categories ADD COLUMN parent_id INTEGER REFERENCES expense_categories(id)"))

        if inspector.has_table('store_sales'):
            store_sale_cols = [c['name'] for c in inspector.get_columns('store_sales')]
            if 'vehicle_id' not in store_sale_cols:
                conn.execute(text(
                    "ALTER TABLE store_sales ADD COLUMN vehicle_id INTEGER REFERENCES vehicles(id)"))

        driver_cols = inspector.get_columns('drivers')
        driver_col_names = [c['name'] for c in driver_cols]
        if 'paired_driver_id' not in driver_col_names:
            conn.execute(text("ALTER TABLE drivers ADD COLUMN paired_driver_id INTEGER REFERENCES drivers(id)"))
        if 'assigned_vehicle_id' not in driver_col_names:
            conn.execute(text("ALTER TABLE drivers ADD COLUMN assigned_vehicle_id INTEGER REFERENCES vehicles(id)"))
        for col in ('next_of_kin_name', 'next_of_kin_phone', 'next_of_kin_relationship'):
            if col not in driver_col_names:
                conn.execute(text(f"ALTER TABLE drivers ADD COLUMN {col} VARCHAR(100)"))

        license_col = next((c for c in driver_cols if c['name'] == 'license_number'), None)
        if license_col is not None and not license_col['nullable']:
            # SQLite can't relax a NOT NULL constraint with ALTER TABLE — rebuild the table.
            # Conductors don't need a license number, so this must become optional.
            conn.execute(text("""
                CREATE TABLE drivers_new (
                    id INTEGER PRIMARY KEY,
                    name VARCHAR(100) NOT NULL,
                    license_number VARCHAR(50) UNIQUE,
                    phone VARCHAR(20),
                    role VARCHAR(20),
                    commission_rate FLOAT,
                    status VARCHAR(20),
                    paired_driver_id INTEGER REFERENCES drivers(id),
                    assigned_vehicle_id INTEGER REFERENCES vehicles(id),
                    next_of_kin_name VARCHAR(100),
                    next_of_kin_phone VARCHAR(100),
                    next_of_kin_relationship VARCHAR(100),
                    created_at DATETIME
                )
            """))
            conn.execute(text("""
                INSERT INTO drivers_new (id, name, license_number, phone, role,
                    commission_rate, status, paired_driver_id, assigned_vehicle_id,
                    next_of_kin_name, next_of_kin_phone, next_of_kin_relationship, created_at)
                SELECT id, name, license_number, phone, role,
                    commission_rate, status, paired_driver_id, assigned_vehicle_id,
                    next_of_kin_name, next_of_kin_phone, next_of_kin_relationship, created_at FROM drivers
            """))
            conn.execute(text("DROP TABLE drivers"))
            conn.execute(text("ALTER TABLE drivers_new RENAME TO drivers"))

        daily_log_cols = inspector.get_columns('daily_logs')
        route_col = next((c for c in daily_log_cols if c['name'] == 'route_id'), None)
        if route_col is not None and not route_col['nullable']:
            # Same story as license_number above — the Vehicle Ledger doesn't
            # track a route per entry, so this must become optional.
            conn.execute(text("""
                CREATE TABLE daily_logs_new (
                    id INTEGER PRIMARY KEY,
                    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id),
                    driver_id INTEGER NOT NULL REFERENCES drivers(id),
                    conductor_id INTEGER REFERENCES drivers(id),
                    route_id INTEGER REFERENCES routes(id),
                    log_date DATE NOT NULL,
                    trips_completed INTEGER,
                    gross_revenue FLOAT NOT NULL,
                    notes TEXT,
                    created_by INTEGER REFERENCES users(id),
                    updated_by INTEGER REFERENCES users(id),
                    created_at DATETIME,
                    updated_at DATETIME
                )
            """))
            conn.execute(text("""
                INSERT INTO daily_logs_new (id, vehicle_id, driver_id, conductor_id, route_id,
                    log_date, trips_completed, gross_revenue, notes, created_by, updated_by,
                    created_at, updated_at)
                SELECT id, vehicle_id, driver_id, conductor_id, route_id,
                    log_date, trips_completed, gross_revenue, notes, created_by, updated_by,
                    created_at, updated_at FROM daily_logs
            """))
            conn.execute(text("DROP TABLE daily_logs"))
            conn.execute(text("ALTER TABLE daily_logs_new RENAME TO daily_logs"))

        daily_log_col_names = [c['name'] for c in inspector.get_columns('daily_logs')]
        # Renamed from staff_deduction/deduction_notes to garnish/reason_for_shortfall
        # to match the "missed revenue target" concept this field represents.
        if 'garnish' not in daily_log_col_names:
            if 'staff_deduction' in daily_log_col_names:
                conn.execute(text("ALTER TABLE daily_logs RENAME COLUMN staff_deduction TO garnish"))
            else:
                conn.execute(text("ALTER TABLE daily_logs ADD COLUMN garnish FLOAT NOT NULL DEFAULT 0.0"))
        if 'reason_for_shortfall' not in daily_log_col_names:
            if 'deduction_notes' in daily_log_col_names:
                conn.execute(text("ALTER TABLE daily_logs RENAME COLUMN deduction_notes TO reason_for_shortfall"))
            else:
                conn.execute(text("ALTER TABLE daily_logs ADD COLUMN reason_for_shortfall TEXT"))

        driver_id_col = next((c for c in inspector.get_columns('daily_logs') if c['name'] == 'driver_id'), None)
        if driver_id_col is not None and not driver_id_col['nullable']:
            # SQLite can't relax a NOT NULL constraint with ALTER TABLE — rebuild the table.
            # Daily Transactions needs to record a fare/activity before a driver is assigned.
            conn.execute(text("""
                CREATE TABLE daily_logs_new2 (
                    id INTEGER PRIMARY KEY,
                    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id),
                    driver_id INTEGER REFERENCES drivers(id),
                    conductor_id INTEGER REFERENCES drivers(id),
                    route_id INTEGER REFERENCES routes(id),
                    log_date DATE NOT NULL,
                    trips_completed INTEGER,
                    gross_revenue FLOAT NOT NULL,
                    garnish FLOAT NOT NULL DEFAULT 0.0,
                    reason_for_shortfall TEXT,
                    notes TEXT,
                    created_by INTEGER REFERENCES users(id),
                    updated_by INTEGER REFERENCES users(id),
                    created_at DATETIME,
                    updated_at DATETIME
                )
            """))
            conn.execute(text("""
                INSERT INTO daily_logs_new2 (id, vehicle_id, driver_id, conductor_id, route_id,
                    log_date, trips_completed, gross_revenue, garnish, reason_for_shortfall, notes,
                    created_by, updated_by, created_at, updated_at)
                SELECT id, vehicle_id, driver_id, conductor_id, route_id,
                    log_date, trips_completed, gross_revenue, garnish, reason_for_shortfall, notes,
                    created_by, updated_by, created_at, updated_at FROM daily_logs
            """))
            conn.execute(text("DROP TABLE daily_logs"))
            conn.execute(text("ALTER TABLE daily_logs_new2 RENAME TO daily_logs"))

        # franchise_income (combined daily+weekly columns on one row) and
        # franchise_weekly_expenses (the old free-form weekly deduction
        # list) were retired in favor of two independent entities,
        # franchise_daily_income and franchise_weekly_income — drop the old
        # tables and let db.create_all() (called again after this function
        # returns) create the new ones.
        if inspector.has_table('franchise_income'):
            conn.execute(text("DROP TABLE franchise_income"))
        if inspector.has_table('franchise_weekly_expenses'):
            conn.execute(text("DROP TABLE franchise_weekly_expenses"))

        # franchise_daily_income / franchise_weekly_income moved from a
        # free-text franchisee_name to a vehicle_id FK on FranchiseVehicle
        # (one entry per date/week per vehicle, since each vehicle pays its
        # own standalone fee), and vehicle_id was then relaxed to nullable
        # so a whole-franchise entry not attributable to one vehicle can
        # coexist with per-vehicle entries — drop and let db.create_all()
        # rebuild with the new column/nullability and unique constraint.
        if inspector.has_table('franchise_daily_income'):
            daily_income_cols = {c['name']: c for c in inspector.get_columns('franchise_daily_income')}
            if 'vehicle_id' not in daily_income_cols or not daily_income_cols['vehicle_id']['nullable']:
                conn.execute(text("DROP TABLE franchise_daily_income"))
        if inspector.has_table('franchise_weekly_income'):
            weekly_income_cols = {c['name']: c for c in inspector.get_columns('franchise_weekly_income')}
            if 'vehicle_id' not in weekly_income_cols or not weekly_income_cols['vehicle_id']['nullable']:
                conn.execute(text("DROP TABLE franchise_weekly_income"))

        if inspector.has_table('franchise_collections'):
            collection_cols = [c['name'] for c in inspector.get_columns('franchise_collections')]
            if 'frequency' not in collection_cols:
                conn.execute(text(
                    "ALTER TABLE franchise_collections ADD COLUMN frequency VARCHAR(10) NOT NULL DEFAULT 'daily'"))
            if 'expense' not in collection_cols:
                conn.execute(text(
                    "ALTER TABLE franchise_collections ADD COLUMN expense FLOAT NOT NULL DEFAULT 0"))

        # franchise_vehicles gained agreed_amount/amount_owed/notes — added
        # in place (not drop-and-recreate) since this table holds real
        # registered vehicles, unlike the income tables above.
        if inspector.has_table('franchise_vehicles'):
            vehicle_cols = [c['name'] for c in inspector.get_columns('franchise_vehicles')]
            if 'agreed_amount' not in vehicle_cols:
                conn.execute(text("ALTER TABLE franchise_vehicles ADD COLUMN agreed_amount FLOAT NOT NULL DEFAULT 0"))
            if 'amount_owed' not in vehicle_cols:
                conn.execute(text("ALTER TABLE franchise_vehicles ADD COLUMN amount_owed FLOAT NOT NULL DEFAULT 0"))
            if 'notes' not in vehicle_cols:
                conn.execute(text("ALTER TABLE franchise_vehicles ADD COLUMN notes TEXT"))

        conn.commit()


def create_default_expense_categories():
    """Vehicle Expenses are classified under exactly five top-level
    headings: Maintenance, Wages, Traffic Fines, Insurance, Admin."""
    headings = ('Maintenance', 'Wages', 'Traffic Fines', 'Insurance', 'Admin')
    for name in headings:
        if not ExpenseCategory.query.filter_by(name=name, parent_id=None).first():
            db.session.add(ExpenseCategory(name=name))
    db.session.flush()

    # Retire the older, differently-named default headings this list
    # replaces. Each is only deleted outright if it holds no expenses and
    # no sub-categories — if it does, that data is folded into the closest
    # matching new heading first, so nothing already booked gets silently
    # dropped or orphaned.
    legacy_fold_into = {
        'Salaries & Wages': 'Wages',
        'Licensing & Permits': 'Admin',
        'Rent & Utilities': 'Admin',
        'Other Overhead': 'Admin',
        'Tax': 'Admin',
    }
    new_headings_by_name = {h.name: h for h in
        ExpenseCategory.query.filter(ExpenseCategory.parent_id.is_(None),
                                     ExpenseCategory.name.in_(headings)).all()}
    for old_name, new_name in legacy_fold_into.items():
        old = ExpenseCategory.query.filter_by(name=old_name, parent_id=None).first()
        if not old:
            continue
        has_data = Expense.query.filter_by(category_id=old.id).first() is not None or old.children
        if has_data:
            new = new_headings_by_name.get(new_name)
            if new:
                Expense.query.filter_by(category_id=old.id).update({'category_id': new.id})
                for child in list(old.children):
                    child.parent_id = new.id
        db.session.delete(old)
    db.session.flush()

    maintenance = ExpenseCategory.query.filter_by(name='Maintenance', parent_id=None).first()
    if not maintenance:
        maintenance = ExpenseCategory(name='Maintenance')
        db.session.add(maintenance)
        db.session.flush()
    for name in ('Brake Pads and Tyres', 'Brake Shoes', 'Tie Rod Ends',
                 'Wheel Bearing', 'Ball Joints', 'Engine Oil'):
        if not ExpenseCategory.query.filter_by(name=name, parent_id=maintenance.id).first():
            db.session.add(ExpenseCategory(name=name, parent_id=maintenance.id))
    db.session.commit()


def create_default_admin():
    if not User.query.filter_by(username='admin').first():
        admin_password = os.environ.get('ADMIN_PASSWORD') or secrets.token_urlsafe(12)
        admin = User(username='admin', email='admin@transport.local', role='admin')
        admin.set_password(admin_password)
        db.session.add(admin)
        db.session.commit()
        print(f'Default admin created — username: admin  password: {admin_password}')
        print('Log in and change this password immediately.')


with app.app_context():
    db.create_all()
    migrate_db()
    db.create_all()  # recreate any tables migrate_db() dropped for a schema change
    create_default_admin()
    create_default_expense_categories()

if __name__ == '__main__':
    debug_mode = not IS_PRODUCTION
    host = os.environ.get('HOST', '127.0.0.1' if IS_PRODUCTION else '0.0.0.0')
    app.run(debug=debug_mode, host=host, port=int(os.environ.get('PORT', 5000)))
