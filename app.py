#!/usr/bin/env python3
"""
Transport Fleet & Finance Management System
T-Tech Solutions | June 2026
"""

import os
import csv
import io
import json
from datetime import datetime, date, timedelta
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for,
                   flash, jsonify, make_response)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func

# ─────────────────────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get('SECRET_KEY', 'trp-erp-dev-secret-2026'),
    SQLALCHEMY_DATABASE_URI='sqlite:///transport_erp.db',
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    COMMISSION_DRIVER_RATE=0.15,
    COMMISSION_CONDUCTOR_RATE=0.10,
)

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Please log in to access this page.'
login_manager.login_message_category = 'warning'


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
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

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
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    documents = db.relationship('VehicleDocument', backref='vehicle',
                                lazy=True, cascade='all, delete-orphan')
    daily_logs = db.relationship('DailyLog', backref='vehicle', lazy=True)
    fuel_logs = db.relationship('FuelLog', backref='vehicle', lazy=True)
    maintenance_logs = db.relationship('MaintenanceLog', backref='vehicle', lazy=True)

    @property
    def total_revenue(self):
        return sum(l.gross_revenue for l in self.daily_logs)

    @property
    def total_fuel_cost(self):
        return sum(l.total_cost for l in self.fuel_logs)

    @property
    def total_maintenance_cost(self):
        return sum(l.total_cost for l in self.maintenance_logs)


class VehicleDocument(db.Model):
    __tablename__ = 'vehicle_documents'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    doc_type = db.Column(db.String(50), nullable=False)
    reference_number = db.Column(db.String(100))
    issue_date = db.Column(db.Date)
    expiry_date = db.Column(db.Date, nullable=False)
    notes = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

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
    license_number = db.Column(db.String(50), unique=True, nullable=False)
    phone = db.Column(db.String(20))
    email = db.Column(db.String(120))
    role = db.Column(db.String(20), default='driver')
    commission_rate = db.Column(db.Float)
    status = db.Column(db.String(20), default='active')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

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
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    logs = db.relationship('DailyLog', backref='route', lazy=True)


class DailyLog(db.Model):
    __tablename__ = 'daily_logs'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    driver_id = db.Column(db.Integer, db.ForeignKey('drivers.id'), nullable=False)
    conductor_id = db.Column(db.Integer, db.ForeignKey('drivers.id'))
    route_id = db.Column(db.Integer, db.ForeignKey('routes.id'), nullable=False)
    log_date = db.Column(db.Date, nullable=False, default=date.today)
    trips_completed = db.Column(db.Integer, default=0)
    gross_revenue = db.Column(db.Float, nullable=False, default=0.0)
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    updated_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    creator = db.relationship('User', foreign_keys=[created_by])
    updater = db.relationship('User', foreign_keys=[updated_by])


class FuelLog(db.Model):
    __tablename__ = 'fuel_logs'
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey('vehicles.id'), nullable=False)
    log_date = db.Column(db.Date, nullable=False, default=date.today)
    liters = db.Column(db.Float, nullable=False)
    cost_per_liter = db.Column(db.Float, nullable=False)
    total_cost = db.Column(db.Float, nullable=False)
    odometer = db.Column(db.Float)
    supplier = db.Column(db.String(100))
    notes = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

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
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

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
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User')


# ─────────────────────────────────────────────────────────────
# Auth helpers
# ─────────────────────────────────────────────────────────────
@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


PERMISSIONS = {
    'vehicles':    'Vehicles — view, add & edit vehicles',
    'drivers':     'Crew — view, add & edit drivers/conductors',
    'routes':      'Routes — view, add & edit routes',
    'daily_logs':  'Daily Logs — view, record & edit trip logs',
    'fuel_logs':   'Fuel Logs — view & record fuel entries',
    'maintenance': 'Maintenance — view & record maintenance logs',
    'reports':     'Finance & Reports — income statement, payroll, CSV exports',
    'compliance':  'Compliance — vehicle documents & expiry tracker',
}


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            flash('Admin access required.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated


def permission_required(perm):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if not current_user.has_permission(perm):
                flash('You do not have permission to access that section.', 'danger')
                return redirect(url_for('dashboard'))
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


def parse_date(s):
    return datetime.strptime(s, '%Y-%m-%d').date() if s else None


# ─────────────────────────────────────────────────────────────
# Auth routes
# ─────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('dashboard') if current_user.is_authenticated else url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password) and user.is_active:
            login_user(user)
            log_audit('LOGIN', description=f'User {username} logged in')
            db.session.commit()
            return redirect(url_for('dashboard'))
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


# ─────────────────────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    today = date.today()
    month_start = today.replace(day=1)

    today_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date == today).scalar() or 0

    month_revenue = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date >= month_start).scalar() or 0

    month_fuel = db.session.query(func.sum(FuelLog.total_cost)).filter(
        FuelLog.log_date >= month_start).scalar() or 0

    month_maintenance = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date >= month_start).scalar() or 0

    month_expenses = month_fuel + month_maintenance
    month_profit = month_revenue - month_expenses

    active_vehicles = Vehicle.query.filter_by(status='active').count()
    active_drivers = Driver.query.filter_by(status='active').count()

    expiry_threshold = today + timedelta(days=30)
    expiring_docs = VehicleDocument.query.filter(
        VehicleDocument.expiry_date.between(today, expiry_threshold)).count()
    expired_docs = VehicleDocument.query.filter(
        VehicleDocument.expiry_date < today).count()

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
def vehicle_add():
    if request.method == 'POST':
        v = Vehicle(
            registration=request.form['registration'].upper().strip(),
            make=request.form['make'].strip(),
            model=request.form['model'].strip(),
            year=int(request.form['year']),
            acquisition_cost=float(request.form.get('acquisition_cost') or 0),
            status=request.form.get('status', 'active'),
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
def vehicle_edit(vid):
    v = Vehicle.query.get_or_404(vid)
    if request.method == 'POST':
        v.registration = request.form['registration'].upper().strip()
        v.make = request.form['make'].strip()
        v.model = request.form['model'].strip()
        v.year = int(request.form['year'])
        v.acquisition_cost = float(request.form.get('acquisition_cost') or 0)
        v.status = request.form.get('status', 'active')
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
def driver_add():
    if request.method == 'POST':
        rate_input = request.form.get('commission_rate', '').strip()
        d = Driver(
            name=request.form['name'].strip(),
            license_number=request.form['license_number'].strip(),
            phone=request.form.get('phone', '').strip(),
            email=request.form.get('email', '').strip(),
            role=request.form.get('role', 'driver'),
            commission_rate=float(rate_input) / 100 if rate_input else None,
            status=request.form.get('status', 'active'),
        )
        db.session.add(d)
        db.session.flush()
        log_audit('CREATE', 'drivers', d.id, f'Added driver {d.name}')
        db.session.commit()
        flash(f'Driver {d.name} registered.', 'success')
        return redirect(url_for('drivers'))
    return render_template('drivers/form.html', driver=None, action='Register')


@app.route('/drivers/<int:did>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('drivers')
def driver_edit(did):
    d = Driver.query.get_or_404(did)
    if request.method == 'POST':
        rate_input = request.form.get('commission_rate', '').strip()
        d.name = request.form['name'].strip()
        d.license_number = request.form['license_number'].strip()
        d.phone = request.form.get('phone', '').strip()
        d.email = request.form.get('email', '').strip()
        d.role = request.form.get('role', 'driver')
        d.commission_rate = float(rate_input) / 100 if rate_input else None
        d.status = request.form.get('status', 'active')
        log_audit('UPDATE', 'drivers', d.id, f'Updated driver {d.name}')
        db.session.commit()
        flash(f'Driver {d.name} updated.', 'success')
        return redirect(url_for('drivers'))
    return render_template('drivers/form.html', driver=d, action='Edit')


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
def route_add():
    if request.method == 'POST':
        r = Route(
            name=request.form['name'].strip(),
            start_point=request.form['start_point'].strip(),
            end_point=request.form['end_point'].strip(),
            distance_km=float(request.form['distance_km']) if request.form.get('distance_km') else None,
            fare_rate=float(request.form['fare_rate']),
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
def route_edit(rid):
    r = Route.query.get_or_404(rid)
    if request.method == 'POST':
        r.name = request.form['name'].strip()
        r.start_point = request.form['start_point'].strip()
        r.end_point = request.form['end_point'].strip()
        r.distance_km = float(request.form['distance_km']) if request.form.get('distance_km') else None
        r.fare_rate = float(request.form['fare_rate'])
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
# Daily Logs
# ─────────────────────────────────────────────────────────────
@app.route('/logs/daily')
@login_required
@permission_required('daily_logs')
def daily_logs():
    page = request.args.get('page', 1, type=int)
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    vehicle_id = request.args.get('vehicle_id', '')

    q = DailyLog.query
    if date_from:
        q = q.filter(DailyLog.log_date >= parse_date(date_from))
    if date_to:
        q = q.filter(DailyLog.log_date <= parse_date(date_to))
    if vehicle_id:
        q = q.filter(DailyLog.vehicle_id == vehicle_id)

    logs = q.order_by(DailyLog.log_date.desc()).paginate(page=page, per_page=20)
    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('logs/daily/index.html', logs=logs, vehicles=all_vehicles,
                           date_from=date_from, date_to=date_to, vehicle_id=vehicle_id)


@app.route('/logs/daily/add', methods=['GET', 'POST'])
@login_required
@permission_required('daily_logs')
def daily_log_add():
    if request.method == 'POST':
        log = DailyLog(
            vehicle_id=int(request.form['vehicle_id']),
            driver_id=int(request.form['driver_id']),
            conductor_id=int(request.form['conductor_id']) if request.form.get('conductor_id') else None,
            route_id=int(request.form['route_id']),
            log_date=parse_date(request.form['log_date']),
            trips_completed=int(request.form.get('trips_completed') or 0),
            gross_revenue=float(request.form['gross_revenue']),
            notes=request.form.get('notes', '').strip(),
            created_by=current_user.id,
        )
        db.session.add(log)
        db.session.flush()
        log_audit('CREATE', 'daily_logs', log.id,
                  f'Daily log for {log.vehicle.registration} on {log.log_date}')
        db.session.commit()
        flash('Daily log recorded.', 'success')
        return redirect(url_for('daily_logs'))

    all_vehicles = Vehicle.query.filter_by(status='active').order_by(Vehicle.registration).all()
    all_drivers = Driver.query.filter_by(status='active').order_by(Driver.name).all()
    all_routes = Route.query.filter_by(status='active').order_by(Route.name).all()
    return render_template('logs/daily/form.html', log=None, action='Record',
                           vehicles=all_vehicles, drivers=all_drivers, routes=all_routes,
                           today=date.today().strftime('%Y-%m-%d'))


@app.route('/logs/daily/<int:lid>/edit', methods=['GET', 'POST'])
@login_required
@permission_required('daily_logs')
def daily_log_edit(lid):
    log = DailyLog.query.get_or_404(lid)
    if request.method == 'POST':
        log.vehicle_id = int(request.form['vehicle_id'])
        log.driver_id = int(request.form['driver_id'])
        log.conductor_id = int(request.form['conductor_id']) if request.form.get('conductor_id') else None
        log.route_id = int(request.form['route_id'])
        log.log_date = parse_date(request.form['log_date'])
        log.trips_completed = int(request.form.get('trips_completed') or 0)
        log.gross_revenue = float(request.form['gross_revenue'])
        log.notes = request.form.get('notes', '').strip()
        log.updated_by = current_user.id
        log.updated_at = datetime.utcnow()
        log_audit('UPDATE', 'daily_logs', lid, f'Updated daily log {lid}')
        db.session.commit()
        flash('Daily log updated.', 'success')
        return redirect(url_for('daily_logs'))

    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    all_drivers = Driver.query.order_by(Driver.name).all()
    all_routes = Route.query.order_by(Route.name).all()
    return render_template('logs/daily/form.html', log=log, action='Edit',
                           vehicles=all_vehicles, drivers=all_drivers, routes=all_routes,
                           today=log.log_date.strftime('%Y-%m-%d'))


@app.route('/logs/daily/<int:lid>/delete', methods=['POST'])
@login_required
@admin_required
def daily_log_delete(lid):
    log = DailyLog.query.get_or_404(lid)
    log_audit('DELETE', 'daily_logs', lid, f'Deleted daily log {lid}')
    db.session.delete(log)
    db.session.commit()
    flash('Daily log deleted.', 'warning')
    return redirect(url_for('daily_logs'))


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
def fuel_log_add():
    if request.method == 'POST':
        liters = float(request.form['liters'])
        cpl = float(request.form['cost_per_liter'])
        log = FuelLog(
            vehicle_id=int(request.form['vehicle_id']),
            log_date=parse_date(request.form['log_date']),
            liters=liters,
            cost_per_liter=cpl,
            total_cost=liters * cpl,
            odometer=float(request.form['odometer']) if request.form.get('odometer') else None,
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
def maintenance_log_add():
    if request.method == 'POST':
        parts = float(request.form.get('parts_cost') or 0)
        labor = float(request.form.get('labor_cost') or 0)
        log = MaintenanceLog(
            vehicle_id=int(request.form['vehicle_id']),
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


# ─────────────────────────────────────────────────────────────
# Reports
# ─────────────────────────────────────────────────────────────
@app.route('/reports/income')
@login_required
@permission_required('reports')
def report_income():
    today = date.today()
    date_from_str = request.args.get('date_from', today.replace(day=1).strftime('%Y-%m-%d'))
    date_to_str = request.args.get('date_to', today.strftime('%Y-%m-%d'))
    vehicle_id = request.args.get('vehicle_id', '')

    df = parse_date(date_from_str)
    dt = parse_date(date_to_str)

    rev_q = db.session.query(func.sum(DailyLog.gross_revenue)).filter(
        DailyLog.log_date.between(df, dt))
    fuel_q = db.session.query(func.sum(FuelLog.total_cost)).filter(
        FuelLog.log_date.between(df, dt))
    maint_q = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date.between(df, dt))

    if vehicle_id:
        rev_q = rev_q.filter(DailyLog.vehicle_id == vehicle_id)
        fuel_q = fuel_q.filter(FuelLog.vehicle_id == vehicle_id)
        maint_q = maint_q.filter(MaintenanceLog.vehicle_id == vehicle_id)

    gross_revenue = rev_q.scalar() or 0
    fuel_cost = fuel_q.scalar() or 0
    maintenance_cost = maint_q.scalar() or 0
    total_expenses = fuel_cost + maintenance_cost
    net_profit = gross_revenue - total_expenses
    profit_margin = (net_profit / gross_revenue * 100) if gross_revenue else 0

    vehicle_breakdown = db.session.query(
        Vehicle.registration,
        Vehicle.make,
        Vehicle.model,
        func.sum(DailyLog.gross_revenue).label('revenue'),
        func.count(DailyLog.id).label('log_days'),
    ).join(DailyLog, Vehicle.id == DailyLog.vehicle_id).filter(
        DailyLog.log_date.between(df, dt)
    ).group_by(Vehicle.id).all()

    all_vehicles = Vehicle.query.order_by(Vehicle.registration).all()
    return render_template('reports/income.html',
        gross_revenue=gross_revenue, fuel_cost=fuel_cost,
        maintenance_cost=maintenance_cost, total_expenses=total_expenses,
        net_profit=net_profit, profit_margin=profit_margin,
        vehicle_breakdown=vehicle_breakdown,
        vehicles=all_vehicles,
        date_from=date_from_str, date_to=date_to_str, vehicle_id=vehicle_id)


@app.route('/reports/payroll')
@login_required
@permission_required('reports')
def report_payroll():
    today = date.today()
    date_from_str = request.args.get('date_from', today.replace(day=1).strftime('%Y-%m-%d'))
    date_to_str = request.args.get('date_to', today.strftime('%Y-%m-%d'))

    df = parse_date(date_from_str)
    dt = parse_date(date_to_str)

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

        rev = (driven[0] or 0) + (conducted[0] or 0)
        days = (driven[1] or 0) + (conducted[1] or 0)
        if days == 0:
            continue
        rate = d.commission_rate if d.commission_rate is not None else (
            dr_rate if d.role == 'driver' else co_rate)
        earnings.append({
            'driver': d,
            'total_revenue': rev,
            'days_worked': days,
            'rate_pct': rate * 100,
            'commission': rev * rate,
        })

    total_commissions = sum(e['commission'] for e in earnings)
    return render_template('reports/payroll.html',
        earnings=earnings, total_commissions=total_commissions,
        date_from=date_from_str, date_to=date_to_str)


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
    if df:
        q = q.filter(DailyLog.log_date >= parse_date(df))
    if dt:
        q = q.filter(DailyLog.log_date <= parse_date(dt))

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['Date', 'Vehicle', 'Driver', 'Conductor', 'Route',
                'Trips', 'Gross Revenue (USD)', 'Entered By', 'Notes'])
    for log in q.all():
        w.writerow([log.log_date, log.vehicle.registration, log.driver.name,
                    log.conductor.name if log.conductor else '',
                    log.route.name, log.trips_completed,
                    f'{log.gross_revenue:.2f}',
                    log.creator.username if log.creator else '',
                    log.notes or ''])
    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = f'attachment; filename=daily_logs_{date.today()}.csv'
    return resp


@app.route('/reports/export/income')
@login_required
@permission_required('reports')
def export_income():
    today = date.today()
    df_str = request.args.get('date_from', today.replace(day=1).strftime('%Y-%m-%d'))
    dt_str = request.args.get('date_to', today.strftime('%Y-%m-%d'))
    df = parse_date(df_str)
    dt = parse_date(dt_str)

    daily = DailyLog.query.filter(DailyLog.log_date.between(df, dt)).order_by(DailyLog.log_date).all()
    fuel = FuelLog.query.filter(FuelLog.log_date.between(df, dt)).order_by(FuelLog.log_date).all()
    maintenance = MaintenanceLog.query.filter(
        MaintenanceLog.log_date.between(df, dt)).order_by(MaintenanceLog.log_date).all()

    out = io.StringIO()
    w = csv.writer(out)
    w.writerow(['TRANSPORT FLEET INCOME STATEMENT (ZIMRA COMPLIANT)'])
    w.writerow([f'Period: {df_str} to {dt_str}'])
    w.writerow([f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M")} by {current_user.username}'])
    w.writerow([])

    w.writerow(['REVENUE'])
    w.writerow(['Date', 'Vehicle', 'Route', 'Trips', 'Gross Revenue (USD)'])
    total_rev = 0
    for l in daily:
        w.writerow([l.log_date, l.vehicle.registration, l.route.name,
                    l.trips_completed, f'{l.gross_revenue:.2f}'])
        total_rev += l.gross_revenue
    w.writerow(['', '', '', 'TOTAL REVENUE', f'{total_rev:.2f}'])
    w.writerow([])

    w.writerow(['FUEL EXPENSES'])
    w.writerow(['Date', 'Vehicle', 'Liters', 'Cost/Liter (USD)', 'Total (USD)', 'Supplier'])
    total_fuel = 0
    for f in fuel:
        w.writerow([f.log_date, f.vehicle.registration, f.liters,
                    f'{f.cost_per_liter:.4f}', f'{f.total_cost:.2f}', f.supplier or ''])
        total_fuel += f.total_cost
    w.writerow(['', '', '', '', 'TOTAL FUEL', f'{total_fuel:.2f}'])
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
    w.writerow(['NET PROFIT', f'{total_rev - total_fuel - total_maint:.2f}'])

    out.seek(0)
    resp = make_response(out.getvalue())
    resp.headers['Content-Type'] = 'text/csv'
    resp.headers['Content-Disposition'] = f'attachment; filename=income_{df_str}_to_{dt_str}.csv'
    return resp


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
def user_add():
    if request.method == 'POST':
        u = User(
            username=request.form['username'].strip().lower(),
            email=request.form['email'].strip().lower(),
            role=request.form.get('role', 'manager'),
        )
        u.set_password(request.form['password'])
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
# API (chart data)
# ─────────────────────────────────────────────────────────────
@app.route('/api/revenue/monthly')
@login_required
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
        fuel = db.session.query(func.sum(FuelLog.total_cost)).filter(
            FuelLog.log_date >= start, FuelLog.log_date < end).scalar() or 0
        maint = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
            MaintenanceLog.log_date >= start, MaintenanceLog.log_date < end).scalar() or 0
        data.append({
            'month': start.strftime('%b %Y'),
            'revenue': float(rev),
            'expenses': float(fuel + maint),
            'profit': float(rev - fuel - maint),
        })
    return jsonify(data)


@app.route('/api/vehicles/performance')
@login_required
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
def api_expenses_breakdown():
    today = date.today()
    m_start = today.replace(day=1)
    fuel = db.session.query(func.sum(FuelLog.total_cost)).filter(
        FuelLog.log_date >= m_start).scalar() or 0
    maint = db.session.query(func.sum(MaintenanceLog.total_cost)).filter(
        MaintenanceLog.log_date >= m_start).scalar() or 0
    return jsonify({'fuel': float(fuel), 'maintenance': float(maint)})


# ─────────────────────────────────────────────────────────────
# WhatsApp Webhook stub
# ─────────────────────────────────────────────────────────────
@app.route('/api/whatsapp/webhook', methods=['POST'])
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
    cols = [c['name'] for c in inspector.get_columns('users')]
    if 'permissions' not in cols:
        with db.engine.connect() as conn:
            conn.execute(text("ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT '[]'"))
            conn.commit()


def create_default_admin():
    if not User.query.filter_by(username='admin').first():
        admin = User(username='admin', email='admin@transport.local', role='admin')
        admin.set_password('admin123')
        db.session.add(admin)
        db.session.commit()
        print('Default admin created — username: admin  password: admin123')


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        migrate_db()
        create_default_admin()
    app.run(debug=True, host='0.0.0.0', port=5000)
