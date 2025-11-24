import os
import sqlite3
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from io import BytesIO
from datetime import datetime
import webbrowser
from threading import Timer

app = Flask(__name__)
app.secret_key = 'transer_segredo_total'
DB_NAME = 'transer.db'

# --- BANCO DE DADOS ---
def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                external_id TEXT NOT NULL UNIQUE,
                document TEXT,
                address TEXT,
                city TEXT,
                email TEXT,
                contract_num TEXT,
                contract_val REAL,
                contract_limit REAL,
                extra_val REAL,
                periodicity TEXT,
                created_at TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE,
                password TEXT,
                role TEXT,
                name TEXT
            )
        ''')
        cursor.execute("SELECT * FROM users WHERE username = 'admin'")
        if not cursor.fetchone():
            cursor.execute("INSERT INTO users (username, password, role, name) VALUES (?, ?, ?, ?)",
                           ('admin', 'admambiental', 'admin', 'Administrador Padrão'))
        conn.commit()

# --- ROTAS ---
@app.route('/')
def index():
    if 'user_id' not in session: return redirect(url_for('login'))
    return render_template('index.html', view='dashboard', user=session.get('user_name'), role=session.get('role'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u, p = request.form['username'], request.form['password']
        with sqlite3.connect(DB_NAME) as conn:
            user = conn.execute("SELECT * FROM users WHERE username = ? AND password = ?", (u, p)).fetchone()
            if user:
                session['user_id'], session['role'], session['user_name'] = user[0], user[3], user[4]
                return redirect(url_for('index'))
            else: flash('Login inválido', 'error')
    return render_template('index.html', view='login')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/clients', methods=['GET', 'POST'])
def clients():
    if 'user_id' not in session: return redirect(url_for('login'))
    if request.method == 'POST':
        d = request.form
        try:
            with sqlite3.connect(DB_NAME) as conn:
                now = datetime.now().strftime("%Y-%m-%d")
                if d.get('client_id'):
                    conn.execute('UPDATE clients SET name=?, external_id=?, document=?, address=?, city=?, email=?, contract_num=?, contract_val=?, contract_limit=?, extra_val=?, periodicity=? WHERE id=?',
                                 (d['name'], d['external_id'], d['document'], d['address'], d['city'], d['email'], d['contract_num'], d['contract_val'], d['contract_limit'], d['extra_val'], d['periodicity'], d['client_id']))
                    flash('Atualizado!', 'success')
                else:
                    conn.execute('INSERT INTO clients (name, external_id, document, address, city, email, contract_num, contract_val, contract_limit, extra_val, periodicity, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                                 (d['name'], d['external_id'], d['document'], d['address'], d['city'], d['email'], d['contract_num'], d['contract_val'], d['contract_limit'], d['extra_val'], d['periodicity'], now))
                    flash('Cadastrado!', 'success')
        except Exception as e: flash(f'Erro: {e}', 'error')
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        return render_template('index.html', view='clients', clients=conn.execute("SELECT * FROM clients ORDER BY name").fetchall(), user=session.get('user_name'), role=session.get('role'))

@app.route('/delete_client/<int:id>')
def delete_client(id):
    if 'user_id' not in session: return redirect(url_for('login'))
    with sqlite3.connect(DB_NAME) as conn: conn.execute("DELETE FROM clients WHERE id=?", (id,))
    return redirect(url_for('clients'))

@app.route('/users', methods=['GET', 'POST'])
def users():
    if 'user_id' not in session or session.get('role') != 'admin': return redirect(url_for('index'))
    if request.method == 'POST':
        d = request.form
        try:
            with sqlite3.connect(DB_NAME) as conn: conn.execute("INSERT INTO users (name, username, password, role) VALUES (?,?,?,?)", (d['name'], d['username'], d['password'], d['role']))
            flash('Usuário criado!', 'success')
        except: flash('Login já existe', 'error')
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        return render_template('index.html', view='users', users=conn.execute("SELECT * FROM users").fetchall(), user=session.get('user_name'), role=session.get('role'))

@app.route('/delete_user/<int:id>')
def delete_user(id):
    if 'user_id' not in session or session.get('role') != 'admin': return redirect(url_for('index'))
    if id != session.get('user_id'): 
        with sqlite3.connect(DB_NAME) as conn: conn.execute("DELETE FROM users WHERE id=?", (id,))
    return redirect(url_for('users'))

@app.route('/reports')
def reports():
    if 'user_id' not in session: return redirect(url_for('login'))
    with sqlite3.connect(DB_NAME) as conn:
        conn.row_factory = sqlite3.Row
        clients = conn.execute("SELECT * FROM clients ORDER BY id DESC").fetchall()
        stats = {'total': len(clients), 'val': sum(c['contract_val'] or 0 for c in clients)}
    return render_template('index.html', view='reports', clients=clients, stats=stats, user=session.get('user_name'), role=session.get('role'))

# --- FECHAMENTO (MÉTODO SALVAR EM DISCO) ---
@app.route('/closing', methods=['GET', 'POST'])
def closing():
    if 'user_id' not in session: return redirect(url_for('login'))
    processed = []
    
    if request.method == 'POST' and 'file' in request.files:
        file = request.files['file']
        if file.filename:
            # CAMINHO TEMPORÁRIO SEGURO
            temp_path = os.path.join(os.getcwd(), "temp_upload_file")
            try:
                # 1. Salva o arquivo no disco (Evita erros de memória/buffer)
                file.save(temp_path)
                
                df = None
                # 2. Tenta ler
                try:
                    # Tenta como Excel
                    df = pd.read_excel(temp_path, engine='openpyxl')
                except:
                    # Tenta como CSV
                    try: df = pd.read_csv(temp_path, encoding='latin1', sep=';', on_bad_lines='skip')
                    except: df = pd.read_csv(temp_path, encoding='latin1', sep=',', on_bad_lines='skip')

                # 3. Processa
                if df is not None:
                    df.columns = [str(c).strip() for c in df.columns]
                    if 'ID Cliente' not in df.columns:
                        flash(f'Erro: Coluna "ID Cliente" não achada. Colunas lidas: {list(df.columns)}', 'error')
                    else:
                        with sqlite3.connect(DB_NAME) as conn:
                            conn.row_factory = sqlite3.Row
                            db_clients = conn.execute("SELECT * FROM clients").fetchall()
                        map_cli = {str(c['external_id']): c for c in db_clients}

                        for cid_raw, group in df.groupby('ID Cliente'):
                            cid = str(cid_raw).replace('.0', '')
                            info = map_cli.get(cid)
                            
                            if 'Quantidade' in group.columns:
                                group['Quantidade'] = pd.to_numeric(group['Quantidade'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                                total = group['Quantidade'].sum()
                            else: total = 0
                            
                            limit = info['contract_limit'] if info else 0
                            cost = (info['contract_val'] if info else 0) + (max(0, total - limit) * (info['extra_val'] if info else 0))
                            
                            processed.append({
                                'id': cid,
                                'name': info['name'] if info else f"ID {cid}",
                                'status': 'OK' if info else 'Off',
                                'total_kg': round(total, 2),
                                'excess': round(max(0, total - limit), 2),
                                'total_final': round(cost, 2),
                                'coletas': group.fillna('').to_dict('records'),
                                'db_data': dict(info) if info else None
                            })
                else:
                    flash("Não foi possível ler o arquivo (formato inválido).", 'error')

            except Exception as e:
                flash(f'Erro crítico no arquivo: {e}', 'error')
            finally:
                # 4. Limpeza: Apaga o arquivo temporário
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                
    return render_template('index.html', view='closing', processed=processed, user=session.get('user_name'), role=session.get('role'))

@app.route('/generate_excel', methods=['POST'])
def generate_excel():
    try:
        import openpyxl
        from openpyxl.styles import Font
    except ImportError:
        return "Erro: openpyxl não instalado.", 500

    data = request.json
    wb = openpyxl.Workbook()
    ws = wb.active
    c, coletas = data['client'], data['coletas']
    
    ws.append([f"FECHAMENTO {datetime.now().strftime('%B/%Y').upper()}"])
    ws['A1'].font = Font(bold=True, size=14)
    ws.append(["TRANSER AMBIENTAL LTDA"]); ws.append(["CLIENTE:", c['name']])
    ws.append(["ENDEREÇO:", c['address']]); ws.append(["CONTRATO:", c['contract_num']]); ws.append([])
    ws.append(["Data", "Resíduo", "Kg", "Tipo"])
    
    total = 0
    for row in coletas:
        qtd = float(row.get('Quantidade', 0))
        total += qtd
        ws.append([row.get('Data'), row.get('Classe do Resíduo'), qtd, "Normal"])
    
    ws.append(["TOTAL", "", total]); ws.append([])
    ws.append(["Descrição", "Qtd", "Unit", "Total"])
    ws.append(["Pacote", 1, c['contract_val'], c['contract_val']])
    exc = max(0, total - c['contract_limit'])
    ws.append([f"Excedente ({c['contract_limit']}Kg)", exc, c['extra_val'], exc * c['extra_val']])
    ws.append(["TOTAL GERAL", "", "", c['contract_val'] + (exc * c['extra_val'])])
    
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name=f"Fechamento_{c['name']}.xlsx", as_attachment=True)

if __name__ == '__main__':
    init_db()
    Timer(1, lambda: webbrowser.open_new('http://127.0.0.1:5000/')).start()
    app.run(debug=True, use_reloader=False)