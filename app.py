from flask import Flask, render_template, jsonify, request, send_file
from flask_socketio import SocketIO, emit
import openpyxl
import json
import os
import time
import io
import random

# ReportLab imports for PDF Generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__)
app.config['SECRET_KEY'] = 'auction_secret_2026'
socketio = SocketIO(app, cors_allowed_origins="*")

EXCEL_FILE = 'Online Auction.xlsx'
STATE_FILE = 'auction_state.json'

def create_default_excel():
    if os.path.exists(EXCEL_FILE):
        return
    wb = openpyxl.Workbook()
    
    # 1. Settings Tab
    ws_settings = wb.active
    ws_settings.title = 'Settings'
    ws_settings.append(['key', 'value'])
    settings_rows = [
        ['tournament_name', 'Friends Premier League 2026'],
        ['initial_purse', 120.0],
        ['max_squad_size', 15],
        ['default_base_price', 0.5],
        ['base_timer_seconds', 15],
        ['final_call_seconds', 3]
    ]
    for r in settings_rows:
        ws_settings.append(r)

    # 2. Teams Tab
    ws_teams = wb.create_sheet('Teams')
    ws_teams.append(['team_id', 'team_name', 'short_name', 'team_code', 'is_admin'])
    teams_rows = [
        [1, 'Mumbai Indians', 'MI', 'mi2026', True],
        [2, 'Chennai Super Kings', 'CSK', 'csk2026', False],
        [3, 'Royal Challengers Bengaluru', 'RCB', 'rcb2026', False],
        [4, 'Kolkata Knight Riders', 'KKR', 'kkr2026', False],
        [5, 'Sunrisers Hyderabad', 'SRH', 'srh2026', False],
        [6, 'Rajasthan Royals', 'RR', 'rr2026', False],
        [7, 'Delhi Capitals', 'DC', 'dc2026', False],
        [8, 'Punjab Kings', 'PBKS', 'pbks2026', False],
        [9, 'Gujarat Titans', 'GT', 'gt2026', False],
        [10, 'Lucknow Super Giants', 'LSG', 'lsg2026', False]
    ]
    for r in teams_rows:
        ws_teams.append(r)

    # 3. Players Tab
    ws_players = wb.create_sheet('Players')
    ws_players.append(['player_id', 'name', 'role', 'base_price', 'photo_url'])
    players_rows = [
        [1, 'Virat Kohli', 'BAT', 2.0, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/2.png'],
        [2, 'Jasprit Bumrah', 'BOWL', 2.0, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/9.png'],
        [3, 'Hardik Pandya', 'AR', 2.0, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/54.png'],
        [4, 'MS Dhoni', 'WK', 2.0, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/57.png'],
        [5, 'Suryakumar Yadav', 'BAT', 1.5, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/174.png'],
        [6, 'Rashid Khan', 'BOWL', 1.5, 'https://documents.iplt20.com/ipl/IPLHeadshot2024/218.png']
    ]
    for r in players_rows:
        ws_players.append(r)

    wb.save(EXCEL_FILE)

def init_state_from_excel():
    create_default_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    settings = {}
    for row in wb['Settings'].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            settings[str(row[0]).strip()] = row[1]

    initial_purse = float(settings.get('initial_purse', 120.0))
    max_squad_size = int(settings.get('max_squad_size', 15))
    default_base = float(settings.get('default_base_price', 0.5))
    base_timer = int(settings.get('base_timer_seconds', 15))
    final_call = int(settings.get('final_call_seconds', 3))

    teams = []
    for row in wb['Teams'].iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[1] is not None:
            is_admin_val = str(row[4]).strip().upper() in ['TRUE', '1', 'YES']
            teams.append({
                'team_id': int(row[0]),
                'team_name': str(row[1]).strip(),
                'short_name': str(row[2]).strip(),
                'team_code': str(row[3]).strip(),
                'is_admin': is_admin_val,
                'purse': initial_purse,
                'players': []
            })

    players = []
    for row in wb['Players'].iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[1] is not None:
            players.append({
                'player_id': int(row[0]),
                'name': str(row[1]).strip(),
                'role': str(row[2]).strip().upper() if row[2] else 'BAT',
                'base_price': float(row[3]) if row[3] is not None else default_base,
                'photo_url': str(row[4]).strip() if row[4] else '',
                'status': 'pending',
                'sold_to': None,
                'sold_to_short': None,
                'sold_price': 0.0
            })

    first_player = random.choice(players) if players else None
    first_bid = first_player['base_price'] if first_player else default_base

    state = {
        'settings': {
            'tournament_name': str(settings.get('tournament_name', 'Cricket Auction')),
            'initial_purse': initial_purse,
            'max_squad_size': max_squad_size,
            'default_base_price': default_base,
            'base_timer_seconds': base_timer,
            'final_call_seconds': final_call,
            'total_timer_seconds': base_timer + final_call
        },
        'teams': teams,
        'players': players,
        'auction': {
            'round': 'MAIN',
            'current_player_id': first_player['player_id'] if first_player else None,
            'current_bid': first_bid,
            'highest_bidder': None,
            'highest_bidder_short': None,
            'prev_bid': None,
            'prev_bidder': None,
            'prev_bidder_short': None,
            'is_active': False,
            'is_paused': False,
            'deadline': None,
            'paused_remaining': base_timer + final_call,
            'transition': None
        }
    }
    save_state(state)
    return state

def load_state():
    if not os.path.exists(STATE_FILE):
        return init_state_from_excel()
    with open(STATE_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_state(state):
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=4)

def get_next_increment(cb):
    if cb < 1.0:
        return 0.1
    elif cb < 5.0:
        return 0.2
    else:
        return 0.5

def calculate_mab(team, settings):
    max_squad = settings['max_squad_size']
    base_price = settings['default_base_price']
    current_count = len(team['players'])
    remaining_slots = max_squad - current_count
    if remaining_slots <= 0:
        return 0.0
    reserved = (remaining_slots - 1) * base_price
    mab = round(team['purse'] - reserved, 2)
    return max(0.0, mab)

def advance_auction(state):
    auc = state['auction']
    players = state['players']
    settings = state['settings']
    total_sec = settings['total_timer_seconds']

    curr_p = next((p for p in players if p['player_id'] == auc['current_player_id']), None)
    if curr_p:
        bidder = auc['highest_bidder']
        if bidder:
            curr_p['status'] = 'sold'
            curr_p['sold_to'] = bidder
            curr_p['sold_to_short'] = auc['highest_bidder_short']
            curr_p['sold_price'] = auc['current_bid']
            for t in state['teams']:
                if t['team_name'] == bidder:
                    t['purse'] = round(t['purse'] - auc['current_bid'], 2)
                    t['players'].append({
                        'player_id': curr_p['player_id'],
                        'name': curr_p['name'],
                        'role': curr_p['role'],
                        'price': auc['current_bid'],
                        'photo_url': curr_p['photo_url']
                    })
                    break
            trans_status = 'SOLD'
        else:
            if auc['round'] == 'MAIN':
                curr_p['status'] = 'unsold_once'
            else:
                curr_p['status'] = 'unsold_final'
            trans_status = 'UNSOLD'

        auc['transition'] = {
            'status': trans_status,
            'player': dict(curr_p),
            'until': time.time() + 5.0
        }

    if auc['round'] == 'MAIN':
        pending_pool = [p for p in players if p['status'] == 'pending']
        if pending_pool:
            next_p = random.choice(pending_pool)
        else:
            unsold_pool = [p for p in players if p['status'] == 'unsold_once']
            if unsold_pool:
                auc['round'] = 'UNSOLD'
                next_p = random.choice(unsold_pool)
            else:
                auc['round'] = 'COMPLETED'
                auc['is_active'] = False
                auc['current_player_id'] = None
                return
    elif auc['round'] == 'UNSOLD':
        unsold_pool = [p for p in players if p['status'] == 'unsold_once']
        if unsold_pool:
            next_p = random.choice(unsold_pool)
        else:
            auc['round'] = 'COMPLETED'
            auc['is_active'] = False
            auc['current_player_id'] = None
            return
    else:
        return

    auc['current_player_id'] = next_p['player_id']
    auc['current_bid'] = next_p['base_price']
    auc['highest_bidder'] = None
    auc['highest_bidder_short'] = None
    auc['prev_bid'] = None
    auc['prev_bidder'] = None
    auc['prev_bidder_short'] = None
    auc['is_active'] = True
    auc['is_paused'] = False
    auc['deadline'] = time.time() + 5.0 + total_sec

def sync_timers(state):
    now = time.time()
    auc = state['auction']
    changed = False

    if auc['transition'] and now >= auc['transition']['until']:
        auc['transition'] = None
        changed = True

    if auc['is_active'] and not auc['is_paused'] and auc['round'] != 'COMPLETED':
        if auc['deadline'] and now >= auc['deadline']:
            advance_auction(state)
            changed = True

    if changed:
        save_state(state)

@app.route('/')
def index():
    state = load_state()
    team_param = request.args.get('team', '').strip()
    key_param = request.args.get('key', '').strip()

    matched_team = None
    for t in state['teams']:
        if (t['short_name'].lower() == team_param.lower() or t['team_name'].lower() == team_param.lower()) and t['team_code'] == key_param:
            matched_team = t
            break

    if not matched_team:
        return render_template('index.html', error="Invalid or Missing Team Link! Please open the exact personal URL shared by your Auctioneer.", teams=state['teams'], settings=state['settings'])

    return render_template('index.html', error=None, my_team=matched_team, settings=state['settings'])

@app.route('/api/state')
def get_state():
    state = load_state()
    sync_timers(state)
    now = time.time()
    auc = state['auction']

    remaining_sec = 0
    if auc['transition']:
        remaining_sec = state['settings']['total_timer_seconds']
    elif auc['is_paused']:
        remaining_sec = auc['paused_remaining']
    elif auc['is_active'] and auc['deadline']:
        remaining_sec = max(0, round(auc['deadline'] - now, 1))

    teams_with_mab = []
    for t in state['teams']:
        t_copy = dict(t)
        t_copy['mab'] = calculate_mab(t, state['settings'])
        teams_with_mab.append(t_copy)

    curr_p = next((p for p in state['players'] if p['player_id'] == auc['current_player_id']), None)

    return jsonify({
        'settings': state['settings'],
        'teams': teams_with_mab,
        'auction': auc,
        'current_player': curr_p,
        'remaining_seconds': remaining_sec,
        'next_bid_amount': round(auc['current_bid'] + get_next_increment(auc['current_bid']), 2) if auc['highest_bidder'] else auc['current_bid']
    })

@app.route('/api/bid', methods=['POST'])
def place_bid():
    req = request.json
    team_name = req.get('team_name')
    team_code = req.get('team_code')

    state = load_state()
    sync_timers(state)
    auc = state['auction']
    settings = state['settings']

    if auc['transition'] or not auc['is_active'] or auc['is_paused'] or auc['round'] == 'COMPLETED':
        return jsonify({'success': False, 'message': 'Bidding is currently not active.'})

    team = next((t for t in state['teams'] if t['team_name'] == team_name and t['team_code'] == team_code), None)
    if not team:
        return jsonify({'success': False, 'message': 'Unauthorized team.'})

    if auc['highest_bidder'] == team_name:
        return jsonify({'success': False, 'message': 'You are already the highest bidder!'})

    if len(team['players']) >= settings['max_squad_size']:
        return jsonify({'success': False, 'message': 'Your squad is already full!'})

    bid_amount = auc['current_bid'] if auc['highest_bidder'] is None else round(auc['current_bid'] + get_next_increment(auc['current_bid']), 2)

    mab = calculate_mab(team, settings)
    if bid_amount > mab:
        return jsonify({'success': False, 'message': f'Exceeds your Max Allowed Bid (₹{mab} Cr)!'})

    auc['prev_bid'] = auc['current_bid']
    auc['prev_bidder'] = auc['highest_bidder']
    auc['prev_bidder_short'] = auc['highest_bidder_short']

    auc['current_bid'] = bid_amount
    auc['highest_bidder'] = team['team_name']
    auc['highest_bidder_short'] = team['short_name']
    auc['deadline'] = time.time() + settings['total_timer_seconds']

    save_state(state)
    return jsonify({'success': True})

def verify_admin(state, team_code):
    return any(t['team_code'] == team_code and t['is_admin'] for t in state['teams'])

@app.route('/api/admin/<action>', methods=['POST'])
def admin_action(action):
    req = request.json or {}
    team_code = req.get('team_code')
    state = load_state()

    if not verify_admin(state, team_code):
        return jsonify({'success': False, 'message': 'Admin access denied!'})

    auc = state['auction']
    settings = state['settings']
    now = time.time()

    if action == 'start_or_resume':
        if not auc['is_active'] and auc['round'] != 'COMPLETED':
            auc['is_active'] = True
            auc['is_paused'] = False
            auc['deadline'] = now + settings['total_timer_seconds']
        elif auc['is_paused']:
            auc['is_paused'] = False
            auc['deadline'] = now + auc['paused_remaining']

    elif action == 'pause':
        if auc['is_active'] and not auc['is_paused']:
            auc['paused_remaining'] = settings['total_timer_seconds']
            auc['is_paused'] = True

    elif action == 'undo':
        if auc['highest_bidder'] is not None:
            auc['current_bid'] = auc['prev_bid']
            auc['highest_bidder'] = auc['prev_bidder']
            auc['highest_bidder_short'] = auc['prev_bidder_short']
            auc['prev_bid'] = None
            auc['prev_bidder'] = None
            auc['prev_bidder_short'] = None
            auc['deadline'] = now + settings['total_timer_seconds']

    elif action == 'next':
        advance_auction(state)

    elif action == 'complete':
        auc['round'] = 'COMPLETED'
        auc['is_active'] = False
        auc['is_paused'] = False
        auc['current_player_id'] = None
        auc['transition'] = None

    elif action == 'reload_excel':
        state = init_state_from_excel()
        return jsonify({'success': True, 'message': 'Reloaded from Online Auction.xlsx'})

    save_state(state)
    return jsonify({'success': True})

@app.route('/api/download_pdf', methods=['GET'])
def download_pdf():
    team_code = request.args.get('key', '').strip()
    state = load_state()
    team = next((t for t in state['teams'] if t['team_code'] == team_code), None)
    
    if not team:
        return "Unauthorized or Invalid Team", 403

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=20, textColor=colors.HexColor('#1e3a8a'), alignment=1, spaceAfter=6)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=11, textColor=colors.HexColor('#475569'), alignment=1, spaceAfter=15)
    
    elements.append(Paragraph(f"{state['settings']['tournament_name']} - Final Squad", title_style))
    elements.append(Paragraph(f"<b>Franchise:</b> {team['team_name']} ({team['short_name']}) &nbsp;&nbsp;|&nbsp;&nbsp; <b>Purse Spent:</b> ₹{round(state['settings']['initial_purse'] - team['purse'], 2)} Cr &nbsp;&nbsp;|&nbsp;&nbsp; <b>Purse Left:</b> ₹{team['purse']} Cr", sub_style))
    elements.append(Spacer(1, 10))

    table_data = [['#', 'Player Name', 'Role', 'Price (Cr)']]
    for idx, p in enumerate(team['players'], start=1):
        table_data.append([str(idx), p['name'], p['role'], f"₹{p['price']}"])

    if len(team['players']) == 0:
        table_data.append(['-', 'No players purchased yet', '-', '-'])

    t = Table(table_data, colWidths=[40, 280, 100, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2563eb')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8fafc')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f1f5f9')]),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
    ]))

    elements.append(t)
    doc.build(elements)
    buffer.seek(0)

    filename = f"{team['short_name']}_Final_Squad.pdf"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype='application/pdf')

@socketio.on('voice_stream')
def handle_voice_stream(data):
    emit('voice_broadcast', data, broadcast=True, include_self=False)

@socketio.on('voice_status')
def handle_voice_status(data):
    speaking = bool((data or {}).get('speaking'))
    if speaking:
        state = load_state()
        sync_timers(state)
        auc = state['auction']
        settings = state['settings']
        if auc['is_active'] and not auc['is_paused'] and auc['round'] != 'COMPLETED' and not auc.get('transition'):
            auc['deadline'] = time.time() + settings['total_timer_seconds']
            save_state(state)
    emit('speaker_status', data, broadcast=True, include_self=False)

# Server start hote hi Excel se state initialize karein (Local + Cloud dono ke liye)
if not os.path.exists(STATE_FILE):
    init_state_from_excel()

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5000, debug=True, allow_unsafe_werkzeug=True)
