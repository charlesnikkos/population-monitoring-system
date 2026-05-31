import base64
import csv
import datetime
import io
import os
import re
from functools import wraps
from flask import Flask, render_template, request, redirect, session, flash, url_for, send_file, send_from_directory, abort, Response
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import case, func, or_, text
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import qrcode
except ImportError:
    qrcode = None

app = Flask(__name__)
app.config['SECRET_KEY'] = 'barangay_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'uploads')
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls', 'csv', 'png', 'jpg', 'jpeg', 'webp'}

DB = SQLAlchemy(app)

# ---------------- MODELS ----------------

class User(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    username = DB.Column(DB.String(100), unique=True, nullable=False)
    password = DB.Column(DB.String(200), nullable=False)
    role = DB.Column(DB.String(20), nullable=False, default='admin')

    def is_admin(self):
        return self.role == 'admin'

class Citizen(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    fullname = DB.Column(DB.String(200), nullable=False)
    birthdate = DB.Column(DB.String(100), nullable=False)
    place_of_birth = DB.Column(DB.String(200), nullable=False)
    nationality = DB.Column(DB.String(200), nullable=False)
    email = DB.Column(DB.String(200))
    contact_number = DB.Column(DB.String(100), nullable=False)
    age = DB.Column(DB.Integer, nullable=False)
    gender = DB.Column(DB.String(50), nullable=False)
    address = DB.Column(DB.String(200), nullable=False)
    status = DB.Column(DB.String(100), nullable=False)
    occupation = DB.Column(DB.String(200), nullable=False, default='N/A')
    special_status = DB.Column(DB.String(200), nullable=False, default='None')
    purok = DB.Column(DB.String(100), nullable=False)
    photo = DB.Column(DB.String(300), default='default.png')

    @property
    def age_category(self):
        if self.age is None:
            return 'Unknown'
        if self.age < 13:
            return 'Child'
        if self.age < 20:
            return 'Teen'
        if self.age < 60:
            return 'Adult'
        return 'Senior'

class BarangayClearance(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    citizen_id = DB.Column(DB.Integer, DB.ForeignKey('citizen.id'), nullable=False)
    reason = DB.Column(DB.String(200), nullable=False)
    status = DB.Column(DB.String(50), default='Pending')
    date_requested = DB.Column(DB.String(100), nullable=False)
    date_released = DB.Column(DB.String(100))
    citizen = DB.relationship('Citizen', backref='clearances')

class CertificateOfIndigency(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    citizen_id = DB.Column(DB.Integer, DB.ForeignKey('citizen.id'), nullable=False)
    reason = DB.Column(DB.String(200), nullable=False)
    status = DB.Column(DB.String(50), default='Pending')
    date_requested = DB.Column(DB.String(100), nullable=False)
    date_released = DB.Column(DB.String(100))
    citizen = DB.relationship('Citizen', backref='indigency_certs')

class Ayuda(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    citizen_id = DB.Column(DB.Integer, DB.ForeignKey('citizen.id'), nullable=False)
    ayuda_type = DB.Column(DB.String(100), nullable=False)
    amount = DB.Column(DB.Float, nullable=False)
    status = DB.Column(DB.String(50), default='Pending')
    date_requested = DB.Column(DB.String(100), nullable=False)
    date_released = DB.Column(DB.String(100))
    citizen = DB.relationship('Citizen', backref='ayudas')


class LuponCase(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    citizen_id = DB.Column(DB.Integer, DB.ForeignKey('citizen.id'), nullable=False)
    title = DB.Column(DB.String(300), nullable=False)
    article = DB.Column(DB.String(200))
    notes = DB.Column(DB.String(1000))
    date_added = DB.Column(DB.String(100), nullable=False)
    citizen = DB.relationship('Citizen', backref='lupon_cases')

class CertificateRequest(DB.Model):
    id = DB.Column(DB.Integer, primary_key=True)
    citizen_id = DB.Column(DB.Integer, DB.ForeignKey('citizen.id'), nullable=False)
    cert_type = DB.Column(DB.String(200), nullable=False)
    purpose = DB.Column(DB.String(500))
    status = DB.Column(DB.String(50), default='Pending')
    date_requested = DB.Column(DB.String(100), nullable=False)
    date_released = DB.Column(DB.String(100))
    citizen = DB.relationship('Citizen', backref='certificate_requests')

# ---------------- HELPERS ----------------

def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return view(*args, **kwargs)
    return wrapped_view

def admin_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if session.get('role') != 'admin':
            flash('Admin access required.', 'warning')
            return redirect(url_for('dashboard'))
        return view(*args, **kwargs)
    return wrapped_view

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# ---------------- PHOTO HELPER FUNCTIONS ----------------

def get_photo_url(citizen):
    """Get the URL for a citizen's photo with proper error handling"""
    if citizen and hasattr(citizen, 'photo') and citizen.photo and citizen.photo != 'default.png':
        # 1) Check uploaded files directory (static/uploads)
        photo_path = os.path.join(app.config['UPLOAD_FOLDER'], citizen.photo)
        if os.path.exists(photo_path):
            return url_for('uploaded_file', filename=citizen.photo)

        # 2) Fall back to checking the top-level static folder (some images are stored there)
        static_path = os.path.join(app.static_folder, citizen.photo)
        if os.path.exists(static_path):
            return url_for('static', filename=citizen.photo)

    # Default avatar when no photo found
    return url_for('static', filename='default-avatar.png')

def save_uploaded_photo(photo_file, photo_data, fullname, existing_photo='default.png'):
    """Save uploaded photo and return filename"""
    photo_filename = existing_photo
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    # Handle file upload
    if photo_file and photo_file.filename and photo_file.filename != '':
        ext = os.path.splitext(photo_file.filename)[1].lower()
        if ext in {'.png', '.jpg', '.jpeg', '.webp'}:
            safe_name = re.sub(r'[^a-zA-Z0-9]', '_', fullname)
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            photo_filename = f'{safe_name}_{timestamp}{ext}'
            photo_path = os.path.join(app.config['UPLOAD_FOLDER'], photo_filename)
            photo_file.save(photo_path)
            print(f"✅ Photo saved from file: {photo_filename}")
            return photo_filename
    
    # Handle camera capture
    if photo_data and photo_data.startswith('data:image'):
        try:
            header, encoded = photo_data.split(',', 1)
            ext = header.split('/')[1].split(';')[0].lower()
            if ext == 'jpeg':
                ext = 'jpg'
            if ext in {'png', 'jpg', 'jpeg', 'webp'}:
                safe_name = re.sub(r'[^a-zA-Z0-9]', '_', fullname)
                timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                photo_filename = f'{safe_name}_{timestamp}.{ext}'
                photo_path = os.path.join(app.config['UPLOAD_FOLDER'], photo_filename)
                with open(photo_path, 'wb') as f:
                    f.write(base64.b64decode(encoded))
                print(f"✅ Photo saved from camera: {photo_filename}")
                return photo_filename
        except Exception as e:
            print(f"❌ Error saving captured photo: {e}")
    
    return photo_filename

# ---------------- HELPERS ----------------

def split_fullname(fullname):
    parts = fullname.split()
    if not parts:
        return '', '', ''
    if len(parts) == 1:
        return parts[0], '', ''
    if len(parts) == 2:
        return parts[0], '', parts[1]
    return parts[0], ' '.join(parts[1:-1]), parts[-1]

# ---------------- CONTEXT PROCESSORS ----------------

@app.context_processor
def inject_site_assets():
    site_logo_url = None
    site_banner_url = None

    static_logo = os.path.join(app.static_folder, 'logo.png')
    if os.path.exists(static_logo):
        site_logo_url = url_for('static', filename='logo.png')

    static_banner = os.path.join(app.static_folder, 'brgy_taft.jpg')
    if os.path.exists(static_banner):
        site_banner_url = url_for('static', filename='brgy_taft.jpg')

    uploads_dir = os.path.join(app.static_folder, 'uploads')
    if os.path.isdir(uploads_dir):
        for fname in os.listdir(uploads_dir):
            if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.svg')):
                if site_logo_url is None:
                    site_logo_url = url_for('static', filename=f'uploads/{fname}')
                if site_banner_url is None and 'banner' in fname.lower():
                    site_banner_url = url_for('static', filename=f'uploads/{fname}')

    if site_banner_url is None and site_logo_url is not None:
        site_banner_url = site_logo_url

    return dict(site_logo_url=site_logo_url, site_banner_url=site_banner_url)

@app.context_processor
def inject_officials():
    """Inject barangay officials data into all templates"""
    officials = {
        'captain': None,
        'kagawads': [],
        'secretary': None,
        'treasurer': None
    }
    
    try:
        # Get captain (Punong Barangay)
        captain = Citizen.query.filter(Citizen.fullname.ilike('%Gesta%')).first()
        if captain:
            officials['captain'] = captain
        
        # Get kagawads
        kagawad_names = ['Bonono', 'Balo', 'Alvarez', 'Go', 'Parada', 'Boquilon', 'Eviota']
        for name in kagawad_names:
            kagawad = Citizen.query.filter(Citizen.fullname.ilike(f'%{name}%')).first()
            if kagawad:
                officials['kagawads'].append(kagawad)
        
        # Get secretary and treasurer
        secretary = Citizen.query.filter(Citizen.fullname.ilike('%Alvarado%')).first()
        if secretary:
            officials['secretary'] = secretary
        
        treasurer = Citizen.query.filter(Citizen.fullname.ilike('%Martinez%')).first()
        if treasurer:
            officials['treasurer'] = treasurer
    except Exception as e:
        print(f"Error loading officials: {e}")
    
    return dict(officials=officials)

@app.context_processor
def utility_processor():
    def get_photo_url_func(citizen):
        return get_photo_url(citizen)
    return dict(get_photo_url=get_photo_url_func)

# ---------------- ROUTES ----------------

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """Serve uploaded files with proper error handling"""
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
    except:
        abort(404)

@app.route('/')
def home():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            session['user_id'] = user.id
            session['username'] = user.username
            session['role'] = user.role
            return redirect(url_for('dashboard'))

        flash('Invalid username or password.', 'danger')

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    total_population = Citizen.query.count()
    male_population = Citizen.query.filter_by(gender='Male').count()
    female_population = Citizen.query.filter_by(gender='Female').count()
    senior_population = Citizen.query.filter(Citizen.special_status.ilike('%senior%')).count()
    pwd_population = Citizen.query.filter(Citizen.special_status.ilike('%pwd%')).count()
    solo_parents_population = Citizen.query.filter(Citizen.special_status.ilike('%solo parent%')).count()
    indigenous_population = Citizen.query.filter(Citizen.special_status.ilike('%indigenous%')).count()
    fourps_population = Citizen.query.filter(Citizen.special_status.ilike('%4p%')).count()
    kids_population = Citizen.query.filter(Citizen.age < 13).count()
    teens_population = Citizen.query.filter(Citizen.age >= 13, Citizen.age < 20).count()
    adults_population = Citizen.query.filter(Citizen.age >= 20, Citizen.age < 60).count()
    seniors_population = Citizen.query.filter(Citizen.age >= 60).count()

    purok_data = DB.session.query(
        Citizen.purok.label('purok'),
        func.count(Citizen.id).label('total'),
        func.sum(case((Citizen.gender == 'Male', 1), else_=0)).label('male'),
        func.sum(case((Citizen.gender == 'Female', 1), else_=0)).label('female'),
        func.sum(case((Citizen.age < 13, 1), else_=0)).label('kids'),
        func.sum(case(((Citizen.age >= 13) & (Citizen.age < 20), 1), else_=0)).label('teens'),
        func.sum(case(((Citizen.age >= 20) & (Citizen.age < 60), 1), else_=0)).label('adults'),
        func.sum(case((Citizen.age >= 60, 1), else_=0)).label('seniors'),
        func.sum(case((Citizen.special_status.ilike('%pwd%'), 1), else_=0)).label('pwd'),
        func.sum(case((Citizen.special_status.ilike('%solo parent%'), 1), else_=0)).label('solo_parents'),
        func.sum(case((Citizen.special_status.ilike('%indigenous%'), 1), else_=0)).label('indigenous'),
        func.sum(case((Citizen.special_status.ilike('%4p%'), 1), else_=0)).label('fourps')
    ).group_by(Citizen.purok).all()

    certificates_preview = [
        {'title': 'Barangay Residency', 'description': 'Proof of residency for barangay purposes.'},
        {'title': 'Certificate of Residency', 'description': 'Official residency verification document.'},
        {'title': 'Certificate of No Objection', 'description': 'Statement of no objection from barangay.'},
        {'title': 'Certificate of Solo Parent', 'description': 'Support documentation for solo parents.'},
        {'title': 'Barangay Business Clearance', 'description': 'Permit to operate a small business in the barangay.'},
        {'title': 'Good Moral Character', 'description': 'Personal character certificate for school or work.'},
        {'title': 'First Time Job Seeker', 'description': 'Certificate for first-time employment applicants.'},
        {'title': 'Certificate of Cohabitation', 'description': 'Proof of living together as partners.'},
        {'title': 'Cedula / Community Tax Certificate', 'description': 'Local community tax clearance or cedula.'},
        {'title': 'Barangay ID', 'description': 'Official Barangay identity card registration.'}
    ]

    return render_template(
        'dashboard.html',
        total_population=total_population,
        male_population=male_population,
        female_population=female_population,
        senior_population=senior_population,
        pwd_population=pwd_population,
        solo_parents_population=solo_parents_population,
        indigenous_population=indigenous_population,
        fourps_population=fourps_population,
        kids_population=kids_population,
        teens_population=teens_population,
        adults_population=adults_population,
        seniors_population=seniors_population,
        purok_data=purok_data,
        certificates_preview=certificates_preview
    )

@app.route('/certificate_requests')
@login_required
@admin_required
def certificate_requests():
    cert_reqs = CertificateRequest.query.all()
    clearances = BarangayClearance.query.all()
    indigencies = CertificateOfIndigency.query.all()
    ayudas = Ayuda.query.all()

    requests = []
    for req in cert_reqs:
        requests.append({
            'id': req.id,
            'citizen': req.citizen,
            'request_type': req.cert_type,
            'notes': req.purpose or 'N/A',
            'status': req.status,
            'date_requested': req.date_requested,
            'date_released': req.date_released,
            'model': 'certificate',
            'approve_url': url_for('approve_certificate', cert_id=req.id),
            'reject_url': url_for('reject_certificate', cert_id=req.id),
            'release_url': url_for('release_certificate', cert_id=req.id),
            'manage_url': url_for('certificate_requests')
        })
    for req in clearances:
        requests.append({
            'id': req.id,
            'citizen': req.citizen,
            'request_type': 'Barangay Clearance',
            'notes': req.reason or 'N/A',
            'status': req.status,
            'date_requested': req.date_requested,
            'date_released': req.date_released,
            'model': 'clearance',
            'approve_url': url_for('approve_clearance', clearance_id=req.id),
            'reject_url': url_for('reject_clearance', clearance_id=req.id),
            'release_url': url_for('release_clearance', clearance_id=req.id),
            'manage_url': url_for('clearance')
        })
    for req in indigencies:
        requests.append({
            'id': req.id,
            'citizen': req.citizen,
            'request_type': 'Certificate of Indigency',
            'notes': req.reason or 'N/A',
            'status': req.status,
            'date_requested': req.date_requested,
            'date_released': req.date_released,
            'model': 'indigency',
            'approve_url': url_for('approve_indigency', cert_id=req.id),
            'reject_url': url_for('reject_indigency', cert_id=req.id),
            'release_url': url_for('release_indigency', cert_id=req.id),
            'manage_url': url_for('indigency')
        })
    for req in ayudas:
        notes = f"{req.ayuda_type} — ₱{req.amount:,.2f}"
        requests.append({
            'id': req.id,
            'citizen': req.citizen,
            'request_type': 'Ayuda Assistance',
            'notes': notes,
            'status': req.status,
            'date_requested': req.date_requested,
            'date_released': req.date_released,
            'model': 'ayuda',
            'approve_url': url_for('approve_ayuda', ayuda_id=req.id),
            'reject_url': url_for('reject_ayuda', ayuda_id=req.id),
            'release_url': url_for('release_ayuda', ayuda_id=req.id),
            'manage_url': url_for('ayuda')
        })

    requests.sort(key=lambda r: r['date_requested'] or '', reverse=True)

    total_requests = len(requests)
    pending_count = sum(1 for req in requests if req['status'] == 'Pending')
    approved_count = sum(1 for req in requests if req['status'] == 'Approved')
    released_count = sum(1 for req in requests if req['status'] == 'Released')

    return render_template(
        'certificate_requests.html',
        requests=requests,
        total_requests=total_requests,
        pending_count=pending_count,
        approved_count=approved_count,
        released_count=released_count
    )

@app.route('/certificate_request/<int:cert_id>/release', methods=['POST'])
@login_required
@admin_required
def release_certificate(cert_id):
    cert_req = CertificateRequest.query.get_or_404(cert_id)
    date_released = request.form.get('date_released')
    cert_req.status = 'Released'
    cert_req.date_released = date_released
    DB.session.commit()
    flash(f'Certificate released to {cert_req.citizen.fullname}.', 'success')
    return redirect(url_for('certificate_requests'))

@app.route('/certificate_request/<int:cert_id>/approve', methods=['POST'])
@login_required
@admin_required
def approve_certificate(cert_id):
    cert_req = CertificateRequest.query.get_or_404(cert_id)
    cert_req.status = 'Approved'
    DB.session.commit()
    flash(f'Certificate approved for {cert_req.citizen.fullname}.', 'success')
    return redirect(url_for('certificate_requests'))

@app.route('/certificate_request/<int:cert_id>/reject', methods=['POST'])
@login_required
@admin_required
def reject_certificate(cert_id):
    cert_req = CertificateRequest.query.get_or_404(cert_id)
    cert_req.status = 'Rejected'
    DB.session.commit()
    flash(f'Certificate request rejected.', 'info')
    return redirect(url_for('certificate_requests'))

@app.route('/certificates')
@login_required
@admin_required
def certificates():
    def slugify(name):
        return re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')

    pending_count = CertificateRequest.query.filter_by(status='Pending').count()
    approved_count = CertificateRequest.query.filter_by(status='Approved').count()
    released_count = CertificateRequest.query.filter_by(status='Released').count()

    certificates = [
        {'title': 'Barangay Residency', 'icon': 'fa-home', 'description': 'Proof of barangay residency for official requests.'},
        {'title': 'Certificate of Residency', 'icon': 'fa-map-marker-alt', 'description': 'Formal residency verification document.'},
        {'title': 'Certificate of No Objection', 'icon': 'fa-thumbs-up', 'description': 'Statement of no objection from the barangay.'},
        {'title': 'Certificate of Solo Parent', 'icon': 'fa-child', 'description': 'Document for solo parent proof and assistance.'},
        {'title': 'Barangay Business Clearance', 'icon': 'fa-briefcase', 'description': 'Barangay permit for small business operations.'},
        {'title': 'Barangay Clearance', 'icon': 'fa-file-earmark-check', 'description': 'Request barangay clearance through the certificate center.'},
        {'title': 'Certificate of Indigency', 'icon': 'fa-file-earmark-text', 'description': 'Request certificate of indigency through the certificate center.'},
        {'title': 'Ayuda Assistance', 'icon': 'fa-heart', 'description': 'Request ayuda assistance through the certificate center.'},
        {'title': 'Certificate of Good Moral Character', 'icon': 'fa-user-check', 'description': 'Character certificate needed for school or work.'},
        {'title': 'Certificate of First Time Job Seeker', 'icon': 'fa-user-graduate', 'description': 'Certificate for first time employment applicants.'},
        {'title': 'Certificate of Cohabitation', 'icon': 'fa-ring', 'description': 'Proof of living together as partners.'},
        {'title': 'Cedula / Community Tax Certificate (CTC)', 'icon': 'fa-file-invoice-dollar', 'description': 'Local community tax or cedula document.'},
        {'title': 'Barangay ID', 'icon': 'fa-id-card', 'description': 'Barangay identification card registration.'}
    ]

    for c in certificates:
        c['slug'] = slugify(c['title'])

    return render_template('certificates.html', certificates=certificates, pending_count=pending_count, approved_count=approved_count, released_count=released_count)


@app.route('/cases')
@login_required
def cases():
    """Lupon cases list for barangay complaints"""
    cases = [
        {'title': 'Swindling (Estafa)', 'article': 'Art. 315, 4° RPC'},
        {'title': 'Other Forms of Swindling', 'article': 'Art. 316, RPC'},
        {'title': 'Swindling a Minor', 'article': 'Art. 317, RPC'},
        {'title': 'Other Deceits', 'article': 'Art. 318, RPC'},
        {'title': 'Removal, Sale, Pledge of Mortgaged Property', 'article': 'Art. 319, RPC'},
        {'title': 'Municipal Ordinances', 'article': 'Art. 320, RPC'},
        {'title': 'Theft', 'article': 'Art. 321, RPC'},
        {'title': 'Some Forms of Theft', 'article': 'Art. 322, RPC'},
        {'title': 'Altering Boundaries on Landmarks', 'article': 'Art. 323, RPC'},
        {'title': 'Arson of Property of Small Value', 'article': 'Art. 324, RPC'},
        {'title': 'Special Cases of Malicious Mischief', 'article': 'Art. 325, RPC'},
        {'title': 'Other Mischief', 'article': 'Art. 326, RPC'},
        {'title': 'Slight Slander', 'article': 'Art. 327, RPC'},
        {'title': 'Slander by Deed', 'article': 'Art. 328, RPC'},
        {'title': 'Intriguing Against Honor', 'article': 'Art. 329, RPC'},
        {'title': 'Persons Exempt from Criminal Liability', 'article': 'Art. 330, RPC'},
        {'title': 'Alarms and Scandal', 'article': 'Art. 331, RPC'},
        {'title': 'Using False Certificates', 'article': 'Art. 332, RPC'},
        {'title': 'Using Fictitious Name and Concealing True Name', 'article': 'Art. 333, RPC'},
        {'title': 'Physical Injuries Inflicted in a Tumultuous Affray', 'article': 'Art. 334, RPC'},
        {'title': 'Slight Physical Injuries and Maltreatment', 'article': 'Art. 266, RPC'},
        {'title': 'Other Forms of Trespass', 'article': 'Art. 266, RPC'},
        {'title': 'Other Light Threats', 'article': 'Art. 285, RPC'},
        {'title': 'Light Coercions', 'article': 'Art. 287, RPC'},
        {'title': 'Grave Coercions', 'article': 'Art. 286, RPC'},
        {'title': 'Other Similar Coercions', 'article': 'Art. 288, RPC'},
        {'title': 'Formation/Maintenance/Prohibition of Combination of Capital or Labor', 'article': 'Art. 289, RPC'},
        {'title': 'Giving Assistance to Suicide - When not Consummated', 'article': 'Art. 253, RPC'},
        {'title': 'Abortion Practiced by a Physician or Midwife and Dispensing of Abortives', 'article': 'Art. 259, RPC'},
        {'title': 'Responsibility of Participants in Duel (Less Serious)', 'article': 'Art. 259, RPC'},
        {'title': 'Less Serious Physical Injuries', 'article': 'Art. 265, RPC'},
        {'title': 'Inducing Minor to Abandon his Home', 'article': 'Art. 271, RPC'},
        {'title': 'Abandoning a Minor', 'article': 'Art. 276, RPC'},
        {'title': 'Abandonment of Minor by Person Entrusted with his Custody', 'article': 'Art. 277, RPC'},
        {'title': 'Abandonment of Person in Danger', 'article': ''},
        {'title': 'Qualified Trespass to Dwelling', 'article': ''},
        {'title': 'Discovering Secrets through Seizure of Correspondence', 'article': ''},
        {'title': 'False Certificate', 'article': 'Art. 174, RPC'},
        {'title': 'False Testimony', 'article': 'Art. 180, RPC'},
        {'title': 'Illegal use of Uniform or Insignia', 'article': 'Art. 179, RPC'},
        {'title': 'Unlawful Arrest', 'article': 'Art. 269, RPC'},
        {'title': 'Grave Scandal', 'article': 'Art. 200, RPC'},
        {'title': 'Grave Threats No. 2', 'article': 'Art. 282, RPC'},
        {'title': 'Light Threats', 'article': 'Art. 283, RPC'},
        {'title': 'Revealing Secrets with Abuse of Office', 'article': 'Art. 291, RPC'},
        {'title': 'Occupation of Real Property or Usurpation of Real Rights in Property', 'article': 'Art. 312, RPC'},
        {'title': 'Simple Seduction', 'article': 'Art. 338, RPC'},
        {'title': 'Acts of Lasciviousness with Consent', 'article': 'Art. 339, RPC'},
        {'title': 'Threatening to Publish and Offer to Prevent Such Publication for A Compensation', 'article': 'Art. 356, RPC'},
        {'title': 'Directly Incriminating or Imputing to an Innocent Person the Commission of a Crime', 'article': 'Art. 363, RPC'},
    ]

    return render_template('cases.html', cases=cases)


@app.route('/cases/assigned')
@login_required
@admin_required
def assigned_cases():
    cases = LuponCase.query.order_by(LuponCase.date_added.desc()).all()
    unique_citizens = len({c.citizen_id for c in cases})
    return render_template('cases_assigned.html', cases=cases, total_cases=len(cases), unique_citizens=unique_citizens)


@app.route('/cases/add', methods=['GET', 'POST'])
@login_required
@admin_required
def add_case():
    citizens = Citizen.query.order_by(Citizen.fullname).all()
    master_cases = [
        'Swindling (Estafa)', 'Other Forms of Swindling', 'Swindling a Minor', 'Other Deceits',
        'Removal, Sale, Pledge of Mortgaged Property', 'Municipal Ordinances', 'Theft', 'Some Forms of Theft',
        'Altering Boundaries on Landmarks', 'Arson of Property of Small Value', 'Special Cases of Malicious Mischief',
        'Other Mischief', 'Slight Slander', 'Slander by Deed', 'Intriguing Against Honor', 'Persons Exempt from Criminal Liability',
        'Alarms and Scandal', 'Using False Certificates', 'Using Fictitious Name and Concealing True Name',
        'Physical Injuries Inflicted in a Tumultuous Affray', 'Slight Physical Injuries and Maltreatment', 'Other Forms of Trespass',
        'Other Light Threats', 'Light Coercions', 'Grave Coercions', 'Other Similar Coercions',
        'Formation/Maintenance/Prohibition of Combination of Capital or Labor', 'Giving Assistance to Suicide - When not Consummated',
        'Abortion Practiced by a Physician or Midwife and Dispensing of Abortives', 'Responsibility of Participants in Duel (Less Serious)',
        'Less Serious Physical Injuries', 'Inducing Minor to Abandon his Home', 'Abandoning a Minor',
        'Abandonment of Minor by Person Entrusted with his Custody', 'Abandonment of Person in Danger', 'Qualified Trespass to Dwelling',
        'Discovering Secrets through Seizure of Correspondence', 'False Certificate', 'False Testimony', 'Illegal use of Uniform or Insignia',
        'Unlawful Arrest', 'Grave Scandal', 'Grave Threats No. 2', 'Light Threats', 'Revealing Secrets with Abuse of Office',
        'Occupation of Real Property or Usurpation of Real Rights in Property', 'Simple Seduction', 'Acts of Lasciviousness with Consent',
        'Threatening to Publish and Offer to Prevent Such Publication for A Compensation', 'Directly Incriminating or Imputing to an Innocent Person the Commission of a Crime'
    ]

    if request.method == 'POST':
        citizen_id = int(request.form.get('citizen_id'))
        title = request.form.get('title', '').strip()
        article = request.form.get('article', '').strip()
        notes = request.form.get('notes', '').strip()
        if not citizen_id or not title:
            flash('Please select a citizen and provide a case title.', 'warning')
            return render_template('add_case.html', citizens=citizens, master_cases=master_cases)
        from datetime import date
        new_case = LuponCase(citizen_id=citizen_id, title=title, article=article, notes=notes, date_added=str(date.today()))
        DB.session.add(new_case)
        DB.session.commit()
        flash('Case assigned to citizen.', 'success')
        return redirect(url_for('assigned_cases'))

    return render_template('add_case.html', citizens=citizens, master_cases=master_cases)

def _available_certificates():
    return [
        'Barangay Residency',
        'Certificate of Residency',
        'Certificate of No Objection',
        'Certificate of Solo Parent',
        'Barangay Business Clearance',
        'Barangay Clearance',
        'Certificate of Indigency',
        'Ayuda Assistance',
        'Certificate of Good Moral Character',
        'Certificate of First Time Job Seeker',
        'Certificate of Cohabitation',
        'Cedula / Community Tax Certificate (CTC)',
        'Barangay ID'
    ]

def slugify(name):
    return re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')

def slug_to_title(slug):
    for name in _available_certificates():
        if slugify(name) == slug:
            return name
    return None

@app.route('/certificate/request/<cert_slug>', methods=['GET', 'POST'])
@login_required
@admin_required
def request_certificate(cert_slug):
    title = slug_to_title(cert_slug)
    if not title:
        flash('Unknown certificate type.', 'warning')
        return redirect(url_for('certificates'))

    citizens = Citizen.query.order_by(Citizen.fullname).all()

    if request.method == 'POST':
        citizen_id = int(request.form['citizen_id'])
        date_requested = request.form['date_requested']

        if title == 'Barangay Clearance':
            reason = request.form.get('reason', '').strip()
            new_req = BarangayClearance(
                citizen_id=citizen_id,
                reason=reason,
                date_requested=date_requested,
                status='Pending'
            )
        elif title == 'Certificate of Indigency':
            reason = request.form.get('reason', '').strip()
            new_req = CertificateOfIndigency(
                citizen_id=citizen_id,
                reason=reason,
                date_requested=date_requested,
                status='Pending'
            )
        elif title == 'Ayuda Assistance':
            ayuda_type = request.form.get('ayuda_type', '').strip()
            amount = float(request.form.get('amount', '0') or 0)
            new_req = Ayuda(
                citizen_id=citizen_id,
                ayuda_type=ayuda_type,
                amount=amount,
                date_requested=date_requested,
                status='Pending'
            )
        else:
            purpose = request.form.get('purpose', '').strip()
            new_req = CertificateRequest(
                citizen_id=citizen_id,
                cert_type=title,
                purpose=purpose,
                date_requested=date_requested,
                status='Pending'
            )

        DB.session.add(new_req)
        DB.session.commit()

        flash(f'{title} request created.', 'success')
        return redirect(url_for('certificates'))

    from datetime import date
    return render_template('request_certificate.html', title=title, citizens=citizens, today=date.today())

@app.route('/about')
@login_required
def about():
    return render_template('about.html')

@app.route('/citizens')
@login_required
def citizens():
    data = Citizen.query.order_by(Citizen.fullname).all()
    next_action = request.args.get('next')
    next_label = None
    if next_action == 'clearance':
        next_label = 'Barangay Clearance'
    elif next_action == 'indigency':
        next_label = 'Certificate of Indigency'
    elif next_action == 'ayuda':
        next_label = 'Ayuda Assistance'
    return render_template('citizens.html', citizens=data, next_action=next_action, next_label=next_label)

@app.route('/purok/<path:purok_name>')
@login_required
def purok_detail(purok_name):
    search = request.args.get('q', '').strip()
    base_query = Citizen.query.filter_by(purok=purok_name).order_by(Citizen.fullname)
    all_citizens = base_query.all()

    if search:
        search_pattern = f'%{search}%'
        citizens = base_query.filter(
            or_(
                Citizen.fullname.ilike(search_pattern),
                Citizen.address.ilike(search_pattern)
            )
        ).all()
    else:
        citizens = all_citizens

    total_population = len(all_citizens)
    male_population = sum(1 for c in all_citizens if c.gender == 'Male')
    female_population = sum(1 for c in all_citizens if c.gender == 'Female')
    senior_population = sum(1 for c in all_citizens if 'senior' in c.special_status.lower())
    pwd_population = sum(1 for c in all_citizens if 'pwd' in c.special_status.lower())
    solo_parents_population = sum(1 for c in all_citizens if 'solo parent' in c.special_status.lower())
    indigenous_population = sum(1 for c in all_citizens if 'indigenous' in c.special_status.lower())
    fourps_population = sum(1 for c in all_citizens if '4p' in c.special_status.lower())

    kids_population = sum(1 for c in all_citizens if c.age < 13)
    teens_population = sum(1 for c in all_citizens if 13 <= c.age < 20)
    adults_population = sum(1 for c in all_citizens if 20 <= c.age < 60)
    seniors_population = sum(1 for c in all_citizens if c.age >= 60)

    return render_template(
        'purok_detail.html',
        purok_name=purok_name,
        citizens=citizens,
        search=search,
        filtered_count=len(citizens),
        total_population=total_population,
        male_population=male_population,
        female_population=female_population,
        kids_population=kids_population,
        teens_population=teens_population,
        adults_population=adults_population,
        seniors_population=seniors_population,
        senior_population=senior_population,
        pwd_population=pwd_population,
        solo_parents_population=solo_parents_population,
        indigenous_population=indigenous_population,
        fourps_population=fourps_population
    )

@app.route('/citizen/<int:citizen_id>')
@login_required
def citizen_detail(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    qr_code = None
    if qrcode is not None:
        # QR code links to the citizen card page using Flask URL generator
        qr_payload = url_for('citizen_card', citizen_id=citizen_id, _external=True)
        
        qr = qrcode.QRCode(border=4, box_size=8)
        qr.add_data(qr_payload)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buffer = io.BytesIO()
        img.save(buffer, format='PNG')
        qr_code = base64.b64encode(buffer.getvalue()).decode('ascii')

    return render_template('citizen_detail.html', citizen=citizen, qr_code=qr_code)

@app.route('/citizen/<int:citizen_id>/qr_image')
@login_required
def citizen_qr_image(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    if qrcode is None:
        flash('QR code generation not available. Install the `qrcode` Python package.', 'warning')
        return redirect(url_for('citizen_detail', citizen_id=citizen_id))

    # Use request.host_url which includes protocol, host, and port
    base_url = request.host_url.rstrip('/')
    qr_payload = f"{base_url}/citizen/{citizen_id}/card"
    
    qr = qrcode.QRCode(border=4, box_size=10)
    qr.add_data(qr_payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')

@app.route('/citizen/<int:citizen_id>/card')
def citizen_card(citizen_id):
    """Display formatted citizen ID card when QR code is scanned."""
    citizen = Citizen.query.get_or_404(citizen_id)
    return render_template('citizen_card.html', citizen=citizen)


@app.route('/citizen/<int:citizen_id>/download_info')
@login_required
def download_info(citizen_id):
    """Provide a vCard (.vcf) download for the citizen's contact/info."""
    citizen = Citizen.query.get_or_404(citizen_id)

    # Prepare vCard content (simple, widely-compatible)
    # Split fullname into parts for N: field
    parts = citizen.fullname.split()
    first = parts[0] if parts else ''
    last = parts[-1] if len(parts) > 1 else ''
    middle = ' '.join(parts[1:-1]) if len(parts) > 2 else ''

    vcard_lines = [
        'BEGIN:VCARD',
        'VERSION:3.0',
        f'N:{last};{first};{middle};;',
        f'FN:{citizen.fullname}',
    ]

    if citizen.email:
        vcard_lines.append(f'EMAIL;TYPE=INTERNET:{citizen.email}')
    if citizen.contact_number:
        vcard_lines.append(f'TEL;TYPE=CELL:{citizen.contact_number}')
    if citizen.birthdate and citizen.birthdate != 'N/A':
        # try to normalize simple YYYY-MM-DD if possible; otherwise include as-is
        vcard_lines.append(f'BDAY:{citizen.birthdate}')

    # Put address and other notes in NOTE field
    notes = []
    if citizen.address:
        notes.append(f'Address: {citizen.address}')
    if citizen.purok:
        notes.append(f'Purok: {citizen.purok}')
    if citizen.place_of_birth:
        notes.append(f'Place of Birth: {citizen.place_of_birth}')
    if citizen.nationality:
        notes.append(f'Nationality: {citizen.nationality}')

    if notes:
        vcard_lines.append(f'NOTE:{"; ".join(notes)}')

    vcard_lines.append('END:VCARD')

    vcard_content = '\r\n'.join(vcard_lines) + '\r\n'

    filename = re.sub(r'[^a-zA-Z0-9]', '_', citizen.fullname) or f'citizen_{citizen_id}'
    filename = f"{filename}.vcf"

    resp = Response(vcard_content, mimetype='text/vcard')
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp

@app.route('/add_citizen', methods=['GET', 'POST'])
@login_required
@admin_required
def add_citizen():
    if request.method == 'POST':
        try:
            # Get form values
            first_name = request.form.get('first_name', '').strip()
            middle_name = request.form.get('middle_name', '').strip()
            last_name = request.form.get('last_name', '').strip()
            fullname = ' '.join(filter(None, [first_name, middle_name, last_name])).strip()
            birthdate = request.form.get('birthdate', '').strip()
            place_of_birth = request.form.get('place_of_birth', '').strip()
            nationality = request.form.get('nationality', '').strip()
            email = request.form.get('email', '').strip()
            contact_number = request.form.get('contact_number', '').strip()
            age_str = request.form.get('age', '0').strip()
            gender = request.form.get('gender', '').strip()
            address = request.form.get('address', '').strip()
            status = request.form.get('status', '').strip()
            occupation = request.form.get('occupation', 'N/A').strip()
            special_status = request.form.get('special_status', 'None').strip()
            purok = request.form.get('purok', '').strip()
            
            # Handle photo
            photo_file = request.files.get('photo')
            photo_data = request.form.get('captured_photo', '')
            
            # Validate required fields
            if not first_name or not last_name or not gender or not address or not purok or not place_of_birth or not nationality or not contact_number:
                flash('Please fill in all required fields (First Name, Last Name, Birth Place, Nationality, Contact Number, Gender, Address, Purok).', 'warning')
                return render_template('add_citizen.html', first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            
            # Convert age to int
            try:
                age = int(age_str)
                if age < 0 or age > 150:
                    flash('Invalid age. Please enter a valid age (0-150).', 'warning')
                    return render_template('add_citizen.html', first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            except ValueError:
                flash('Invalid age value. Please enter a number.', 'warning')
                return render_template('add_citizen.html', first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            
            # Save photo
            photo_filename = save_uploaded_photo(photo_file, photo_data, fullname)
            
            # Create new citizen
            new_citizen = Citizen(
                fullname=fullname,
                birthdate=birthdate or 'N/A',
                place_of_birth=place_of_birth,
                nationality=nationality,
                email=email,
                contact_number=contact_number,
                age=age,
                gender=gender,
                address=address,
                status=status or 'N/A',
                occupation=occupation or 'N/A',
                special_status=special_status or 'None',
                purok=purok,
                photo=photo_filename
            )
            DB.session.add(new_citizen)
            DB.session.commit()
            flash('Citizen added successfully.', 'success')
            return redirect(url_for('citizens'))
            
        except Exception as e:
            flash(f'Error adding citizen: {str(e)}', 'danger')
            return render_template('add_citizen.html', first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
    
    return render_template('add_citizen.html', first_name='', middle_name='', last_name='', birthdate='', place_of_birth='', nationality='', email='', contact_number='', age='', gender='', address='', status='Single', occupation='N/A', special_status='None', purok='')


@app.route('/citizen/<int:citizen_id>/add_case', methods=['POST'])
@login_required
@admin_required
def add_case_to_citizen(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    title = request.form.get('case_title', '').strip()
    article = request.form.get('case_article', '').strip()
    notes = request.form.get('case_notes', '').strip()
    if not title:
        flash('Please provide a case title.', 'warning')
        return redirect(url_for('citizen_detail', citizen_id=citizen_id))
    from datetime import date
    new_case = LuponCase(
        citizen_id=citizen_id,
        title=title,
        article=article,
        notes=notes,
        date_added=str(date.today())
    )
    DB.session.add(new_case)
    DB.session.commit()
    flash('Case added to citizen record.', 'success')
    return redirect(url_for('citizen_detail', citizen_id=citizen_id))


@app.route('/citizen/<int:citizen_id>/cases/download')
@login_required
def download_citizen_cases(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    cases = LuponCase.query.filter_by(citizen_id=citizen_id).order_by(LuponCase.date_added.desc()).all()

    csv_buf = io.StringIO()
    writer = csv.writer(csv_buf)
    writer.writerow(['Citizen', 'Case Title', 'Article', 'Notes', 'Date Added'])
    for c in cases:
        writer.writerow([citizen.fullname, c.title, c.article or '', c.notes or '', c.date_added or ''])

    csv_bytes = io.BytesIO(csv_buf.getvalue().encode('utf-8-sig'))
    csv_bytes.seek(0)
    filename = f"{re.sub(r'[^a-zA-Z0-9]', '_', citizen.fullname)}_cases.csv"
    return send_file(csv_bytes, mimetype='text/csv', as_attachment=True, download_name=filename)


@app.route('/citizen/<int:citizen_id>/cases/print')
@login_required
def print_citizen_cases(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    cases = LuponCase.query.filter_by(citizen_id=citizen_id).order_by(LuponCase.date_added.desc()).all()
    return render_template('citizen_cases_print.html', citizen=citizen, cases=cases)


@app.route('/case/<int:case_id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_case(case_id):
    case = LuponCase.query.get_or_404(case_id)
    DB.session.delete(case)
    DB.session.commit()
    flash('Case removed.', 'success')
    return redirect(url_for('assigned_cases'))

@app.route('/edit_citizen/<int:citizen_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_citizen(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    if request.method == 'POST':
        try:
            first_name = request.form.get('first_name', '').strip()
            middle_name = request.form.get('middle_name', '').strip()
            last_name = request.form.get('last_name', '').strip()
            fullname = ' '.join(filter(None, [first_name, middle_name, last_name])).strip()
            birthdate = request.form.get('birthdate', '').strip()
            place_of_birth = request.form.get('place_of_birth', '').strip()
            nationality = request.form.get('nationality', '').strip()
            email = request.form.get('email', '').strip()
            contact_number = request.form.get('contact_number', '').strip()
            age_str = request.form.get('age', '0').strip()
            gender = request.form.get('gender', '').strip()
            address = request.form.get('address', '').strip()
            status = request.form.get('status', '').strip()
            occupation = request.form.get('occupation', 'N/A').strip()
            special_status = request.form.get('special_status', 'None').strip()
            purok = request.form.get('purok', '').strip()

            photo_file = request.files.get('photo')
            photo_data = request.form.get('captured_photo', '')
            
            if not first_name or not last_name or not gender or not address or not purok or not place_of_birth or not nationality or not contact_number:
                flash('Please fill in all required fields (First Name, Last Name, Birth Place, Nationality, Contact Number, Gender, Address, Purok).', 'warning')
                return render_template('add_citizen.html', citizen=citizen, action_url=url_for('edit_citizen', citizen_id=citizen.id), first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            
            try:
                age = int(age_str)
                if age < 0 or age > 150:
                    flash('Invalid age. Please enter a valid age (0-150).', 'warning')
                    return render_template('add_citizen.html', citizen=citizen, action_url=url_for('edit_citizen', citizen_id=citizen.id), first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            except ValueError:
                flash('Invalid age value. Please enter a number.', 'warning')
                return render_template('add_citizen.html', citizen=citizen, action_url=url_for('edit_citizen', citizen_id=citizen.id), first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)
            
            # Save photo if new one is provided
            if (photo_file and photo_file.filename) or (photo_data and photo_data.startswith('data:image')):
                photo_filename = save_uploaded_photo(photo_file, photo_data, fullname, citizen.photo)
                citizen.photo = photo_filename
            
            citizen.fullname = fullname
            citizen.birthdate = birthdate or 'N/A'
            citizen.place_of_birth = place_of_birth
            citizen.nationality = nationality
            citizen.email = email
            citizen.contact_number = contact_number
            citizen.age = age
            citizen.gender = gender
            citizen.address = address
            citizen.status = status or 'N/A'
            citizen.occupation = occupation or 'N/A'
            citizen.special_status = special_status or 'None'
            citizen.purok = purok

            DB.session.commit()
            flash('Citizen updated successfully.', 'success')
            return redirect(url_for('citizen_detail', citizen_id=citizen.id))

        except Exception as e:
            flash(f'Error updating citizen: {str(e)}', 'danger')
            return render_template('add_citizen.html', citizen=citizen, action_url=url_for('edit_citizen', citizen_id=citizen.id), first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=birthdate, place_of_birth=place_of_birth, nationality=nationality, email=email, contact_number=contact_number, age=age_str, gender=gender, address=address, status=status, occupation=occupation, special_status=special_status, purok=purok)

    first_name, middle_name, last_name = split_fullname(citizen.fullname)
    return render_template('add_citizen.html', citizen=citizen, action_url=url_for('edit_citizen', citizen_id=citizen.id), first_name=first_name, middle_name=middle_name, last_name=last_name, birthdate=citizen.birthdate if citizen.birthdate != 'N/A' else '', place_of_birth=citizen.place_of_birth, nationality=citizen.nationality, email=citizen.email, contact_number=citizen.contact_number, age=citizen.age, gender=citizen.gender, address=citizen.address, status=citizen.status, occupation=citizen.occupation, special_status=citizen.special_status, purok=citizen.purok)

@app.route('/delete_citizen/<int:citizen_id>')
@login_required
@admin_required
def delete_citizen(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    # Delete photo file if not default
    if citizen.photo and citizen.photo != 'default.png':
        photo_path = os.path.join(app.config['UPLOAD_FOLDER'], citizen.photo)
        if os.path.exists(photo_path):
            try:
                os.remove(photo_path)
            except:
                pass
    DB.session.delete(citizen)
    DB.session.commit()
    flash('Citizen record deleted.', 'success')
    return redirect(url_for('citizens'))

@app.route('/accounts')
@login_required
@admin_required
def accounts():
    users = User.query.order_by(User.username).all()
    return render_template('accounts.html', users=users)

@app.route('/add_account', methods=['GET', 'POST'])
@login_required
@admin_required
def add_account():
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']
        role = request.form.get('role', 'admin')

        if User.query.filter_by(username=username).first():
            flash('Username already exists.', 'warning')
            return render_template('add_account.html')

        new_user = User(
            username=username,
            password=generate_password_hash(password),
            role=role
        )
        DB.session.add(new_user)
        DB.session.commit()
        flash('Account created successfully.', 'success')
        return redirect(url_for('accounts'))

    return render_template('add_account.html')

@app.route('/edit_account/<int:user_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_account(user_id):
    user = User.query.get_or_404(user_id)

    if request.method == 'POST':
        username = request.form['username'].strip()
        role = request.form.get('role', 'staff')
        new_password = request.form.get('password', '').strip()

        if username != user.username and User.query.filter_by(username=username).first():
            flash('Username already exists.', 'warning')
            return render_template('edit_account.html', user=user)

        user.username = username
        user.role = role

        if new_password:
            user.password = generate_password_hash(new_password)

        DB.session.commit()
        flash('Account updated successfully.', 'success')
        return redirect(url_for('accounts'))

    return render_template('edit_account.html', user=user)

# ---------------- IMPORT FUNCTIONS ----------------

def normalize_header(header):
    return str(header or '').strip().lower()

def parse_uploaded_rows(file_path):
    ext = os.path.splitext(file_path)[1].lower().lstrip('.')
    rows = []

    if ext == 'csv':
        with open(file_path, mode='r', encoding='utf-8-sig', newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if any(row.values()):
                    rows.append({normalize_header(k): v for k, v in row.items()})

    elif ext in {'xlsx', 'xls'}:
        if openpyxl is None:
            raise RuntimeError('OpenPyXL is required to import Excel files. Install it with pip install openpyxl.')
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        data = list(sheet.iter_rows(values_only=True))
        if not data:
            return rows
        headers = [normalize_header(cell) for cell in data[0]]
        for row in data[1:]:
            if any(row):
                rows.append({headers[idx]: row[idx] for idx in range(len(headers)) if idx < len(row)})
    else:
        raise RuntimeError('Unsupported file type.')

    return rows

def import_citizens_from_file(file_path):
    field_map = {
        'name': 'fullname',
        'full name': 'fullname',
        'fullname': 'fullname',
        'full_name': 'fullname',
        'age': 'age',
        'gender': 'gender',
        'birthdate': 'birthdate',
        'date of birth': 'birthdate',
        'dob': 'birthdate',
        'address': 'address',
        'status': 'status',
        'civil status': 'status',
        'occupation': 'occupation',
        'job': 'occupation',
        'work': 'occupation',
        'special status': 'special_status',
        'special_status': 'special_status',
        'indigenous': 'special_status',
        'pwd': 'special_status',
        'senior': 'special_status',
        'purok': 'purok',
        'zone': 'purok',
        'barangay': 'purok',
        'nationality': 'nationality',
        'country': 'nationality',
        'contact number': 'contact_number',
        'contact_number': 'contact_number',
        'phone': 'contact_number',
        'telephone': 'contact_number',
        'email': 'email',
        'email address': 'email',
        'email_address': 'email',
        'place of birth': 'place_of_birth',
        'place_of_birth': 'place_of_birth',
        'birthplace': 'place_of_birth',
        'pob': 'place_of_birth'
    }

    inserted = 0
    skipped = 0
    rows = parse_uploaded_rows(file_path)

    for row in rows:
        citizen_data = {}
        for key, value in row.items():
            mapped = field_map.get(key)
            if mapped:
                citizen_data[mapped] = value

        fullname = str(citizen_data.get('fullname', '')).strip()
        age_value = citizen_data.get('age', '')
        gender = str(citizen_data.get('gender', '')).strip()

        if not fullname or not gender or age_value == '':
            skipped += 1
            continue

        try:
            age = int(float(age_value))
        except Exception:
            skipped += 1
            continue

        birthdate = citizen_data.get('birthdate')
        if birthdate is None:
            birthdate = ''
        if isinstance(birthdate, datetime.date):
            birthdate = birthdate.isoformat()
        birthdate = str(birthdate).strip() or 'N/A'

        address = str(citizen_data.get('address', '')).strip() or 'N/A'
        status = str(citizen_data.get('status', '')).strip() or 'N/A'
        occupation = str(citizen_data.get('occupation', '')).strip() or 'N/A'
        special_status = str(citizen_data.get('special_status', '')).strip() or 'None'
        purok = str(citizen_data.get('purok', '')).strip() or 'N/A'
        nationality = str(citizen_data.get('nationality', '')).strip() or 'N/A'
        contact_number = str(citizen_data.get('contact_number', '')).strip() or 'N/A'
        email = str(citizen_data.get('email', '')).strip() or None
        place_of_birth = str(citizen_data.get('place_of_birth', '')).strip() or 'N/A'

        new_citizen = Citizen(
            fullname=fullname,
            birthdate=birthdate,
            age=age,
            gender=gender,
            address=address,
            status=status,
            occupation=occupation,
            special_status=special_status,
            purok=purok,
            nationality=nationality,
            contact_number=contact_number,
            email=email,
            place_of_birth=place_of_birth
        )
        DB.session.add(new_citizen)
        inserted += 1

    DB.session.commit()
    return inserted, skipped

@app.route('/upload_citizens', methods=['GET', 'POST'])
@login_required
@admin_required
def upload_citizens():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('No file selected for upload.', 'warning')
            return redirect(request.url)

        if not allowed_file(file.filename):
            flash('Allowed file types: xlsx, xls, csv.', 'warning')
            return redirect(request.url)

        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        filename = secure_filename(file.filename)
        timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        save_filename = f"{timestamp}_{filename}"
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], save_filename)
        file.save(save_path)

        try:
            inserted, skipped = import_citizens_from_file(save_path)
            flash(f'Imported {inserted} citizen records. Skipped {skipped} invalid rows.', 'success')
        except Exception as e:
            flash(f'Upload failed: {e}', 'danger')
        return redirect(url_for('citizens'))

    return render_template('upload.html')

@app.route('/download_upload_template')
@login_required
@admin_required
def download_upload_template():
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(['Name', 'Birthdate', 'Gender', 'Age', 'Address', 'Status', 'Purok', 'Nationality', 'Contact Number', 'Email Address', 'Place of Birth'])
    writer.writerow(['Juan Dela Cruz', '1990-01-01', 'Male', '34', '123 Mabini St', 'Single', 'Purok 1', 'Filipino', '09123456789', 'juan@email.com', 'Manila'])
    csv_buffer.seek(0)

    return send_file(
        io.BytesIO(csv_buffer.getvalue().encode('utf-8-sig')),
        mimetype='text/csv',
        as_attachment=True,
        download_name='citizen_upload_template.csv'
    )

@app.route('/settings')
@login_required
@admin_required
def settings():
    total_users = User.query.count()
    total_citizens = Citizen.query.count()
    return render_template('settings.html', total_users=total_users, total_citizens=total_citizens)

@app.route('/clearance')
@login_required
@admin_required
def clearance():
    clearances = BarangayClearance.query.all()
    return render_template('clearance.html', clearances=clearances)

@app.route('/clearance/request/<int:citizen_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def request_clearance(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    if request.method == 'POST':
        new_clearance = BarangayClearance(
            citizen_id=citizen_id,
            reason=request.form['reason'],
            date_requested=request.form['date_requested'],
            status='Pending'
        )
        DB.session.add(new_clearance)
        DB.session.commit()
        flash('Clearance request created.', 'success')
        return redirect(url_for('clearance'))
    from datetime import date
    return render_template('request_clearance.html', citizen=citizen, today=date.today())

@app.route('/clearance/release/<int:clearance_id>')
@login_required
@admin_required
def release_clearance(clearance_id):
    clearance = BarangayClearance.query.get_or_404(clearance_id)
    from datetime import date
    clearance.status = 'Released'
    clearance.date_released = str(date.today())
    DB.session.commit()
    flash('Barangay clearance released.', 'success')
    return redirect(url_for('clearance'))


@app.route('/clearance/<int:clearance_id>/approve', methods=['POST'])
@login_required
@admin_required
def approve_clearance(clearance_id):
    clearance = BarangayClearance.query.get_or_404(clearance_id)
    clearance.status = 'Approved'
    DB.session.commit()
    flash('Clearance approved.', 'success')
    return redirect(url_for('certificate_requests'))


@app.route('/clearance/<int:clearance_id>/reject', methods=['POST'])
@login_required
@admin_required
def reject_clearance(clearance_id):
    clearance = BarangayClearance.query.get_or_404(clearance_id)
    clearance.status = 'Rejected'
    DB.session.commit()
    flash('Clearance request rejected.', 'info')
    return redirect(url_for('certificate_requests'))

@app.route('/indigency')
@login_required
@admin_required
def indigency():
    certs = CertificateOfIndigency.query.all()
    return render_template('indigency.html', certs=certs)

@app.route('/indigency/request/<int:citizen_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def request_indigency(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    if request.method == 'POST':
        new_cert = CertificateOfIndigency(
            citizen_id=citizen_id,
            reason=request.form['reason'],
            date_requested=request.form['date_requested'],
            status='Pending'
        )
        DB.session.add(new_cert)
        DB.session.commit()
        flash('Certificate of indigency request created.', 'success')
        return redirect(url_for('indigency'))
    from datetime import date
    return render_template('request_indigency.html', citizen=citizen, today=date.today())

@app.route('/indigency/release/<int:cert_id>')
@login_required
@admin_required
def release_indigency(cert_id):
    cert = CertificateOfIndigency.query.get_or_404(cert_id)
    from datetime import date
    cert.status = 'Released'
    cert.date_released = str(date.today())
    DB.session.commit()
    flash('Certificate of indigency released.', 'success')
    return redirect(url_for('indigency'))


@app.route('/indigency/<int:cert_id>/approve', methods=['POST'])
@login_required
@admin_required
def approve_indigency(cert_id):
    cert = CertificateOfIndigency.query.get_or_404(cert_id)
    cert.status = 'Approved'
    DB.session.commit()
    flash('Indigency certificate approved.', 'success')
    return redirect(url_for('certificate_requests'))


@app.route('/indigency/<int:cert_id>/reject', methods=['POST'])
@login_required
@admin_required
def reject_indigency(cert_id):
    cert = CertificateOfIndigency.query.get_or_404(cert_id)
    cert.status = 'Rejected'
    DB.session.commit()
    flash('Indigency request rejected.', 'info')
    return redirect(url_for('certificate_requests'))

@app.route('/ayuda')
@login_required
@admin_required
def ayuda():
    ayudas = Ayuda.query.all()
    return render_template('ayuda.html', ayudas=ayudas)

@app.route('/ayuda/request/<int:citizen_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def request_ayuda(citizen_id):
    citizen = Citizen.query.get_or_404(citizen_id)
    if request.method == 'POST':
        new_ayuda = Ayuda(
            citizen_id=citizen_id,
            ayuda_type=request.form['ayuda_type'],
            amount=float(request.form['amount']),
            date_requested=request.form['date_requested'],
            status='Pending'
        )
        DB.session.add(new_ayuda)
        DB.session.commit()
        flash('Ayuda request created.', 'success')
        return redirect(url_for('ayuda'))
    from datetime import date
    return render_template('request_ayuda.html', citizen=citizen, today=date.today())

@app.route('/ayuda/release/<int:ayuda_id>')
@login_required
@admin_required
def release_ayuda(ayuda_id):
    ayuda = Ayuda.query.get_or_404(ayuda_id)
    from datetime import date
    ayuda.status = 'Released'
    ayuda.date_released = str(date.today())
    DB.session.commit()
    flash('Ayuda released.', 'success')
    return redirect(url_for('ayuda'))


@app.route('/ayuda/<int:ayuda_id>/approve', methods=['POST'])
@login_required
@admin_required
def approve_ayuda(ayuda_id):
    ayuda = Ayuda.query.get_or_404(ayuda_id)
    ayuda.status = 'Approved'
    DB.session.commit()
    flash('Ayuda approved.', 'success')
    return redirect(url_for('certificate_requests'))


@app.route('/ayuda/<int:ayuda_id>/reject', methods=['POST'])
@login_required
@admin_required
def reject_ayuda(ayuda_id):
    ayuda = Ayuda.query.get_or_404(ayuda_id)
    ayuda.status = 'Rejected'
    DB.session.commit()
    flash('Ayuda request rejected.', 'info')
    return redirect(url_for('certificate_requests'))

# ---------------- MIGRATION FUNCTION ----------------

def ensure_all_columns():
    """Safely add missing columns to existing tables"""
    with DB.engine.connect() as conn:
        # For citizen table
        citizen_exists = conn.execute(text("SELECT name FROM sqlite_master WHERE type='table' AND name='citizen'")).first()
        if citizen_exists:
            columns = [row[1] for row in conn.execute(text("PRAGMA table_info('citizen')"))]
            
            if 'photo' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN photo VARCHAR(300) DEFAULT 'default.png'"))
                print("✅ Added 'photo' column to citizen table")
            
            if 'occupation' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN occupation VARCHAR(200) DEFAULT 'N/A'"))
                print("✅ Added 'occupation' column to citizen table")
            
            if 'special_status' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN special_status VARCHAR(200) DEFAULT 'None'"))
                print("✅ Added 'special_status' column to citizen table")

            if 'place_of_birth' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN place_of_birth VARCHAR(200) DEFAULT 'N/A'"))
                print("✅ Added 'place_of_birth' column to citizen table")

            if 'nationality' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN nationality VARCHAR(200) DEFAULT 'N/A'"))
                print("✅ Added 'nationality' column to citizen table")

            if 'email' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN email VARCHAR(200)"))
                print("✅ Added 'email' column to citizen table")

            if 'contact_number' not in columns:
                conn.execute(text("ALTER TABLE citizen ADD COLUMN contact_number VARCHAR(100) DEFAULT 'N/A'"))
                print("✅ Added 'contact_number' column to citizen table")
        
        # For user table
        user_exists = conn.execute(text("SELECT name FROM sqlite_master WHERE type='table' AND name='user'")).first()
        if user_exists:
            user_columns = [row[1] for row in conn.execute(text("PRAGMA table_info('user')"))]
            if 'role' not in user_columns:
                conn.execute(text("ALTER TABLE user ADD COLUMN role VARCHAR(20) DEFAULT 'admin'"))
                print("✅ Added 'role' column to user table")

# ---------------- CREATE DEFAULT AVATAR ----------------

def create_default_avatar():
    """Create a default avatar image if it doesn't exist"""
    default_avatar_path = os.path.join(app.static_folder, 'default-avatar.png')
    if not os.path.exists(default_avatar_path):
        try:
            from PIL import Image, ImageDraw, ImageFont
            # Create a simple avatar image
            img = Image.new('RGB', (200, 200), color='#667eea')
            draw = ImageDraw.Draw(img)
            draw.ellipse((25, 25, 175, 175), fill='#764ba2')
            draw.text((85, 75), "👤", fill='white', size=80)
            img.save(default_avatar_path)
            print("✅ Created default avatar image")
        except:
            print("⚠️ Could not create default avatar. Install PIL: pip install Pillow")

# ---------------- INIT DB ----------------

if __name__ == '__main__':
    with app.app_context():
        DB.create_all()
        ensure_all_columns()
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        create_default_avatar()

        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(
                username='admin',
                password=generate_password_hash('admin123'),
                role='admin'
            )
            DB.session.add(admin)
            DB.session.commit()
            print("✅ Default admin account created: admin / admin123")

    app.run(host='0.0.0.0', debug=True)