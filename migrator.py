#!/usr/bin/env python3
"""
eM Client Universal Migrator - Core Engine
===========================================
Extracts Accounts, Emails, Calendars, and Contacts from eM Client and migrates them to:
  1. Mozilla Thunderbird (Accounts via prefs.js, MBOX stores, Unified Folders)
  2. Microsoft Outlook (via Outlook MAPI COM automation / PST generation)
  3. Standalone Files (MBOX, EML, ICS, VCF)

Zero external dependencies for core & Thunderbird (pure Python 3 standard library).
Optional: pywin32 (for direct Outlook COM migration on Windows).
License: MIT
"""

import os
import sys
import io
import re
import json
import shutil
import sqlite3
import datetime
import tempfile
from pathlib import Path
from urllib.parse import quote
from email import message_from_bytes
from email.header import decode_header as decode_rfc2047
import configparser

# Try to import win32com for Outlook support
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

TICKS_EPOCH = 621355968000000000
TICKS_PER_SEC = 10000000

FLAGS_INBOX = 134754324
FLAGS_SENT = 134750740
FLAGS_TRASH = 532756
FLAGS_JUNK = 1074274324
FLAGS_DRAFTS = 134751252
FLAGS_ARCHIVE = 134766612

def ticks_to_datetime(ticks):
    if not ticks or ticks == 0:
        return None
    try:
        seconds = (ticks - TICKS_EPOCH) / TICKS_PER_SEC
        return datetime.datetime.fromtimestamp(seconds, tz=datetime.timezone.utc)
    except Exception:
        return None

def clean_filename(name):
    if not name:
        return "Folder"
    cleaned = re.sub(r'[\/:*?"<>|\r\n\t\x00-\x1f]', '_', str(name))
    return cleaned.strip('. ')

def decode_subject(subj):
    if not subj:
        return ""
    try:
        fragments = decode_rfc2047(subj)
        parts = []
        for frag, enc in fragments:
            if isinstance(frag, bytes):
                parts.append(frag.decode(enc or 'utf-8', errors='replace'))
            else:
                parts.append(str(frag))
        return "".join(parts)
    except Exception:
        return str(subj)

def get_boundary(hdr_bytes):
    try:
        msg = message_from_bytes(hdr_bytes)
        return msg.get_param('boundary')
    except Exception:
        return None

def build_mime_node(parts, part_name='TEXT'):
    hdr, body = parts.get(part_name, (b'', b''))
    boundary = get_boundary(hdr)
    out = bytearray()
    if hdr:
        out.extend(hdr)
    out.extend(b'\r\n\r\n')
    
    if boundary:
        children = []
        if part_name == 'TEXT':
            i = 1
            while str(i) in parts or str(i)+'.1' in parts:
                if str(i) in parts:
                    children.append(str(i))
                i += 1
        else:
            i = 1
            while f"{part_name}.{i}" in parts or f"{part_name}.{i}.1" in parts:
                child = f"{part_name}.{i}"
                if child in parts:
                    children.append(child)
                i += 1
                
        if body:
            out.extend(body)
            
        for child in children:
            out.extend(b'\r\n--' + boundary.encode('utf-8', errors='replace') + b'\r\n')
            out.extend(build_mime_node(parts, child))
            
        out.extend(b'\r\n--' + boundary.encode('utf-8', errors='replace') + b'--\r\n')
    else:
        if body:
            out.extend(body)
            
    return bytes(out)

class EMClientReader:
    def __init__(self, data_path=None):
        self.data_path = Path(data_path) if data_path else self.find_default_path()
        if not self.data_path or not self.data_path.exists():
            raise FileNotFoundError(f"eM Client directory not found: {self.data_path}")
        self.guid_map = {}
        self.accounts_config = []

    @staticmethod
    def find_default_path():
        if sys.platform == 'win32':
            appdata = os.environ.get('APPDATA')
            if appdata:
                p = Path(appdata) / 'eM Client'
                if p.exists():
                    return p
        elif sys.platform == 'darwin':
            p = Path.home() / 'Library' / 'Application Support' / 'eM Client'
            if p.exists():
                return p
        return None

    def load_accounts(self):
        accounts_db = self.data_path / 'accounts.dat'
        if not accounts_db.exists():
            return []

        conn = sqlite3.connect(f'file:{accounts_db}?mode=ro', uri=True)
        cur = conn.cursor()
        cur.execute("SELECT key, value FROM AccountsJson")
        rows = cur.fetchall()
        conn.close()

        configs = {}
        for k, v in rows:
            try:
                configs[k] = json.loads(v)
            except Exception:
                pass

        by_binding = {}
        for k, obj in configs.items():
            t = obj.get('Type', '')
            val = obj.get('TypeValue', {})
            b_uid = val.get('BindingConfigurationUID') or k
            if b_uid not in by_binding:
                by_binding[b_uid] = {}
            short_t = t.split(',')[0].split('.')[-1]
            by_binding[b_uid][short_t] = val

            user = val.get('Credentials', {}).get('Username') if isinstance(val.get('Credentials'), dict) else val.get('ServerUserName', val.get('AccountName'))
            if user and '@' in str(user):
                self.guid_map[k] = user
                self.guid_map[b_uid] = user
            elif val.get('AccountName') and '@' in str(val.get('AccountName')):
                self.guid_map[k] = val.get('AccountName')
                self.guid_map[b_uid] = val.get('AccountName')

        self.guid_map['Local Folders'] = 'Local Folders'

        accounts_list = []
        for b_uid, types in by_binding.items():
            mail_acc = types.get('MailAccountConfiguration', {})
            imap = types.get('ImapAccountConfiguration', {})
            pop3 = types.get('Pop3AccountConfiguration', {})
            smtp = types.get('SmtpAccountConfiguration', {})
            cal = types.get('GdataCalendarAccountConfiguration') or types.get('CalDavAccountConfiguration') or types.get('WebCalAccountConfiguration')
            card = types.get('GdataContactsAccountConfiguration') or types.get('CardDavAccountConfiguration')

            email = None
            for sec in [mail_acc, imap, pop3, smtp]:
                if sec:
                    user = sec.get('Credentials', {}).get('Username') if isinstance(sec.get('Credentials'), dict) else sec.get('ServerUserName', sec.get('AccountName'))
                    if user and '@' in str(user):
                        email = user
                        break
            if not email:
                continue

            domain = email.split('@')[-1].lower()
            is_oauth_ms = domain in ['outlook.com', 'live.com', 'hotmail.com', 'hotmail.gr', 'windowslive.com']
            is_oauth_google = domain in ['gmail.com', 'googlemail.com']

            if is_oauth_google:
                inc_host, inc_port, inc_socket, inc_auth = 'imap.gmail.com', 993, 3, 10
                out_host, out_port, out_socket, out_auth = 'smtp.gmail.com', 465, 3, 10
                out_desc = 'Google Mail'
            elif is_oauth_ms:
                inc_host, inc_port, inc_socket, inc_auth = 'outlook.office365.com', 993, 3, 10
                out_host, out_port, out_socket, out_auth = 'smtp-mail.outlook.com', 587, 2, 10
                out_desc = 'Outlook.com (Microsoft)'
            elif 'yahoo' in domain:
                inc_host, inc_port, inc_socket, inc_auth = 'imap.mail.yahoo.com', 993, 3, 10
                out_host, out_port, out_socket, out_auth = 'smtp.mail.yahoo.com', 465, 3, 10
                out_desc = 'Yahoo! Mail'
            else:
                inc_sec = imap or pop3 or {}
                inc_host = inc_sec.get('Host') or (smtp.get('Host') if smtp else '')
                inc_port = inc_sec.get('Port') or (993 if inc_sec.get('SSL') == 'UseTlsOnSpecialPort' else 143)
                inc_socket = 3 if inc_port in (993, 995) else (2 if inc_port in (143, 110) else 0)
                inc_auth = 3

                smtp_sec = smtp or {}
                out_host = smtp_sec.get('Host') or inc_host
                out_port = smtp_sec.get('Port') or (465 if smtp_sec.get('SSL') == 'UseTlsOnSpecialPort' else 587)
                out_socket = 3 if out_port == 465 else (2 if out_port == 587 else 0)
                out_auth = 3
                out_desc = f"{email} SMTP"

            inc_type = 'imap' if imap else ('pop3' if pop3 else 'imap')
            inc_user = (imap or pop3 or {}).get('Credentials', {}).get('Username') if isinstance((imap or pop3 or {}).get('Credentials'), dict) else (imap or pop3 or {}).get('ServerUserName', email)
            out_user = (smtp or {}).get('Credentials', {}).get('Username') if isinstance((smtp or {}).get('Credentials'), dict) else (smtp or {}).get('ServerUserName', email)

            accounts_list.append({
                'binding_id': b_uid,
                'email': email,
                'domain': domain,
                'inc_type': inc_type,
                'inc_host': inc_host,
                'inc_port': inc_port,
                'inc_socket': inc_socket,
                'inc_auth': inc_auth,
                'inc_user': inc_user,
                'out_host': out_host,
                'out_port': out_port,
                'out_socket': out_socket,
                'out_auth': out_auth,
                'out_user': out_user,
                'out_desc': out_desc,
                'has_cal': cal is not None,
                'has_contacts': card is not None
            })

        self.accounts_config = accounts_list
        return accounts_list

    def scan_data_tree(self):
        """Scans eM Client databases to build a complete tree of accounts, folders, and item counts."""
        if not self.accounts_config:
            self.load_accounts()

        tree = {}
        for d in self.data_path.iterdir():
            if not d.is_dir(): continue
            acc_name = self.guid_map.get(d.name, d.name)
            
            # Check for folders.dat & mail_data.dat either directly or in subfolders
            acc_dirs = []
            if (d / 'folders.dat').exists() and (d / 'mail_data.dat').exists():
                acc_dirs.append(d)
            elif re.match(r'^[0-9a-fA-F\-]{36}$', d.name) or d.name in ('Local Folders', 'Archive'):
                for sub in d.iterdir():
                    if sub.is_dir() and (sub / 'folders.dat').exists() and (sub / 'mail_data.dat').exists():
                        acc_dirs.append(sub)

            if not acc_dirs:
                continue

            if acc_name not in tree:
                tree[acc_name] = {
                    'name': acc_name,
                    'folders': {},
                    'cal_count': 0,
                    'contact_count': 0,
                    'acc_dirs': []
                }

            for ad in acc_dirs:
                tree[acc_name]['acc_dirs'].append(ad)
                folders = self.read_account_folders(ad)
                for f_id, f_info in folders.items():
                    path_str = "/".join(f_info['path'])
                    if path_str not in tree[acc_name]['folders']:
                        tree[acc_name]['folders'][path_str] = {
                            'id': f_id,
                            'name': f_info['name'],
                            'path': f_info['path'],
                            'count': 0
                        }
                    tree[acc_name]['folders'][path_str]['count'] += f_info['count']

        # Scan Calendars
        for c_db in self.data_path.glob('**/event_data.dat'):
            try:
                parent = c_db.parent
                acc = 'Local Folders'
                while parent != self.data_path:
                    if parent.name in self.guid_map:
                        acc = self.guid_map[parent.name]
                        break
                    parent = parent.parent
                conn = sqlite3.connect(f'file:{c_db}?mode=ro', uri=True)
                cur = conn.cursor()
                cur.execute("SELECT count(*) FROM EventItemContents WHERE propertyName='calendar'")
                cnt = cur.fetchone()[0]
                conn.close()
                if acc not in tree:
                    tree[acc] = {'name': acc, 'folders': {}, 'cal_count': 0, 'contact_count': 0, 'acc_dirs': []}
                tree[acc]['cal_count'] += cnt
            except Exception:
                pass

        # Scan Contacts
        for c_db in self.data_path.glob('**/contact_index.dat'):
            try:
                parent = c_db.parent
                acc = 'Local Folders'
                while parent != self.data_path:
                    if parent.name in self.guid_map:
                        acc = self.guid_map[parent.name]
                        break
                    parent = parent.parent
                conn = sqlite3.connect(f'file:{c_db}?mode=ro', uri=True)
                cur = conn.cursor()
                cur.execute("SELECT count(*) FROM ContactItems")
                cnt = cur.fetchone()[0]
                conn.close()
                if acc not in tree:
                    tree[acc] = {'name': acc, 'folders': {}, 'cal_count': 0, 'contact_count': 0, 'acc_dirs': []}
                tree[acc]['contact_count'] += cnt
            except Exception:
                pass

        return tree

    def read_account_folders(self, acc_dir):
        f_db = acc_dir / 'folders.dat'
        m_idx = acc_dir / 'mail_index.dat'
        if not f_db.exists() or not m_idx.exists():
            return {}

        conn_f = sqlite3.connect(f'file:{f_db}?mode=ro', uri=True)
        cur_f = conn_f.cursor()
        cur_f.execute("SELECT id, name, parentFolderId FROM Folders")
        raw_folders = {r[0]: {'name': r[1], 'parent': r[2]} for r in cur_f.fetchall()}
        conn_f.close()

        conn_m = sqlite3.connect(f'file:{m_idx}?mode=ro', uri=True)
        cur_m = conn_m.cursor()
        cur_m.execute("SELECT folder, count(*) FROM MailItems GROUP BY folder")
        counts = dict(cur_m.fetchall())
        conn_m.close()

        folder_map = {}
        for fid, f in raw_folders.items():
            path_parts = []
            curr = fid
            while curr in raw_folders and raw_folders[curr]['parent'] != 0:
                name = raw_folders[curr]['name']
                if name not in ('Mails', 'Root', 'All Mail Client Folders'):
                    path_parts.append(clean_filename(name))
                curr = raw_folders[curr]['parent']
            if curr in raw_folders:
                name = raw_folders[curr]['name']
                if name not in ('Mails', 'Root', 'All Mail Client Folders'):
                    path_parts.append(clean_filename(name))
            path_parts.reverse()
            cnt = counts.get(fid, 0)
            folder_map[fid] = {
                'name': f['name'],
                'path': path_parts or ['INBOX'],
                'count': cnt
            }
        return folder_map

    def get_messages_generator(self, acc_dir, target_folder_path):
        """Yields reconstructed messages for a given folder path."""
        f_db = acc_dir / 'folders.dat'
        m_idx = acc_dir / 'mail_index.dat'
        m_dat = acc_dir / 'mail_data.dat'
        if not f_db.exists() or not m_idx.exists() or not m_dat.exists():
            return

        folder_map = self.read_account_folders(acc_dir)
        target_fids = [fid for fid, info in folder_map.items() if "/".join(info['path']) == target_folder_path]
        if not target_fids:
            return

        conn_idx = sqlite3.connect(f'file:{m_idx}?mode=ro', uri=True)
        cur_idx = conn_idx.cursor()
        placeholders = ",".join("?" * len(target_fids))
        cur_idx.execute(f"SELECT id, folder, subject, receivedDate, flags FROM MailItems WHERE folder IN ({placeholders})", target_fids)
        rows = cur_idx.fetchall()
        conn_idx.close()

        conn_data = sqlite3.connect(f'file:{m_dat}?mode=ro', uri=True)
        cur_data = conn_data.cursor()

        for mid, folder_id, subj, recv_ticks, flags in rows:
            cur_data.execute("SELECT partName, partHeader, partBody FROM LocalMailContents WHERE id=?", (mid,))
            parts = {r[0]: (r[1] or b'', r[2] or b'') for r in cur_data.fetchall()}
            if not parts:
                continue
            eml_bytes = build_mime_node(parts, 'TEXT')
            recv_dt = ticks_to_datetime(recv_ticks)
            is_read = False if (flags & 16) else True
            yield {
                'id': mid,
                'subject': decode_subject(subj),
                'received_date': recv_dt,
                'is_read': is_read,
                'flags': flags,
                'eml_bytes': eml_bytes,
                'parts': parts
            }
        conn_data.close()

class ThunderbirdTarget:
    def __init__(self, profile_path):
        self.profile = Path(profile_path)
        if not self.profile.exists():
            raise FileNotFoundError(f"Thunderbird profile path not found: {self.profile}")

    @staticmethod
    def find_default_profile():
        if sys.platform == 'win32':
            appdata = os.environ.get('APPDATA')
            if not appdata: return None
            tb_root = Path(appdata) / 'Thunderbird'
        elif sys.platform == 'darwin':
            tb_root = Path.home() / 'Library' / 'Thunderbird'
        else:
            tb_root = Path.home() / '.thunderbird'

        if not tb_root.exists(): return None
        ini_file = tb_root / 'profiles.ini'
        if ini_file.exists():
            config = configparser.ConfigParser()
            config.read(ini_file, encoding='utf-8')
            for sec in config.sections():
                if config.get(sec, 'Default', fallback=None) == '1':
                    p = config.get(sec, 'Path', fallback=None)
                    is_rel = config.get(sec, 'IsRelative', fallback='1') == '1'
                    if p:
                        cand = (tb_root / p) if is_rel else Path(p)
                        if cand.exists(): return cand
        profiles_dir = tb_root / 'Profiles'
        if profiles_dir.exists():
            for p in profiles_dir.iterdir():
                if p.is_dir() and (p / 'prefs.js').exists():
                    return p
        return None

    def backup_prefs(self):
        prefs_file = self.profile / 'prefs.js'
        if prefs_file.exists():
            bak = self.profile / f"prefs.js.bak_{int(datetime.datetime.now().timestamp())}"
            shutil.copy2(prefs_file, bak)
            return bak
        return None

    def deploy_accounts(self, accounts_to_add):
        prefs_file = self.profile / 'prefs.js'
        if not prefs_file.exists(): return

        with open(prefs_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        existing_text = "".join(lines)
        existing_emails = set(re.findall(r'mail\.server\.server\d+\.userName", "([^"]+)"', existing_text))
        filtered = [a for a in accounts_to_add if a['email'].lower() not in [e.lower() for e in existing_emails]]
        if not filtered: return

        acc_nums = [int(m) for m in re.findall(r'mail\.account\.account(\d+)\.', existing_text)]
        srv_nums = [int(m) for m in re.findall(r'mail\.server\.server(\d+)\.', existing_text)]
        id_nums = [int(m) for m in re.findall(r'mail\.identity\.id(\d+)\.', existing_text)]
        smtp_nums = [int(m) for m in re.findall(r'mail\.smtpserver\.smtp(\d+)\.', existing_text)]

        next_acc = max(acc_nums) + 1 if acc_nums else 5
        next_srv = max(srv_nums) + 1 if srv_nums else 7
        next_id = max(id_nums) + 1 if id_nums else 4
        next_smtp = max(smtp_nums) + 1 if smtp_nums else 3

        new_prefs, new_acc_keys, new_smtp_keys = [], [], []
        existing_dirs = set(os.listdir(self.profile / 'ImapMail') if (self.profile / 'ImapMail').exists() else [])

        for a in filtered:
            email = a['email']
            acc_key = f"account{next_acc}"
            srv_key = f"server{next_srv}"
            id_key = f"id{next_id}"
            smtp_key = f"smtp{next_smtp}"
            next_acc += 1; next_srv += 1; next_id += 1; next_smtp += 1
            new_acc_keys.append(acc_key); new_smtp_keys.append(smtp_key)

            clean_host = re.sub(r'[^a-zA-Z0-9\.\-]', '_', a['inc_host'] or 'custom.mail')
            dir_name = clean_host
            c = 1
            while dir_name in existing_dirs:
                dir_name = f"{clean_host}-{c}"; c += 1
            existing_dirs.add(dir_name)

            top_dir = 'Mail' if a['inc_type'] == 'pop3' else 'ImapMail'
            full_dir_path = self.profile / top_dir / dir_name
            full_dir_path.mkdir(parents=True, exist_ok=True)
            rel_dir_path = f"[ProfD]{top_dir}/{dir_name}"

            new_prefs.append(f'user_pref("mail.account.{acc_key}.identities", "{id_key}");\n')
            new_prefs.append(f'user_pref("mail.account.{acc_key}.server", "{srv_key}");\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.attachPgpKey", false);\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.fullName", "{email}");\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.reply_on_top", 1);\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.smtpServer", "{smtp_key}");\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.useremail", "{email}");\n')
            new_prefs.append(f'user_pref("mail.identity.{id_key}.valid", true);\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.authMethod", {a["inc_auth"]});\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.check_new_mail", true);\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.directory", "{str(full_dir_path).replace(chr(92), chr(92)*2)}");\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.directory-rel", "{rel_dir_path}");\n')
            inc_host_str = a['inc_host'] or ''
            out_host_str = a['out_host'] or ''
            new_prefs.append(f'user_pref("mail.server.{srv_key}.hostname", "{inc_host_str}");\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.login_at_startup", true);\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.name", "{email}");\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.port", {a["inc_port"]});\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.socketType", {a["inc_socket"]});\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.storeContractID", "@mozilla.org/msgstore/berkeleystore;1");\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.type", "{a["inc_type"]}");\n')
            new_prefs.append(f'user_pref("mail.server.{srv_key}.userName", "{a["inc_user"]}");\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.authMethod", {a["out_auth"]});\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.description", "{a["out_desc"]}");\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.hostname", "{out_host_str}");\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.port", {a["out_port"]});\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.try_ssl", {a["out_socket"]});\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.type", "smtp");\n')
            new_prefs.append(f'user_pref("mail.smtpserver.{smtp_key}.username", "{a["out_user"]}");\n')

        updated_lines = []
        for line in lines:
            if 'mail.accountmanager.accounts' in line:
                m = re.search(r'user_pref\("mail\.accountmanager\.accounts", "([^"]+)"\);', line)
                curr = m.group(1) if m else "account1"
                updated_lines.append(f'user_pref("mail.accountmanager.accounts", "{",".join(curr.split(",") + new_acc_keys)}");\n')
            elif 'mail.smtpservers' in line:
                m = re.search(r'user_pref\("mail\.smtpservers", "([^"]+)"\);', line)
                curr = m.group(1) if m else "smtp1"
                updated_lines.append(f'user_pref("mail.smtpservers", "{",".join(curr.split(",") + new_smtp_keys)}");\n')
            elif 'mail.account.lastKey' in line:
                updated_lines.append(f'user_pref("mail.account.lastKey", {next_acc - 1});\n')
            else:
                updated_lines.append(line)

        updated_lines.extend(new_prefs)
        with open(prefs_file, 'w', encoding='utf-8') as f:
            f.writelines(updated_lines)

    def get_server_map(self):
        prefs_file = self.profile / 'prefs.js'
        if not prefs_file.exists():
            return {}
        with open(prefs_file, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        srv_users = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.userName", "([^"]+)"\);', text))
        srv_dirs = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.directory", "([^"]+)"\);', text))
        srv_hosts = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.hostname", "([^"]+)"\);', text))
        srv_types = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.type", "([^"]+)"\);', text))

        server_map = {}
        for s_id, s_dir in srv_dirs.items():
            user = srv_users.get(s_id, 'N/A')
            host = srv_hosts.get(s_id, 'N/A')
            stype = srv_types.get(s_id, 'N/A')
            if user != 'nobody':
                server_map[user] = {
                    'server_id': s_id,
                    'type': stype,
                    'user': user,
                    'host': host,
                    'dir': s_dir.replace('\\\\', '\\')
                }
        return server_map

    def get_account_mbox_dir(self, acc_name):
        smap = self.get_server_map()
        if acc_name in smap:
            p = Path(smap[acc_name]['dir'])
            p.mkdir(parents=True, exist_ok=True)
            return p
        p = self.profile / 'Mail' / 'Local Folders' / 'eM Client Migration.sbd' / f"{clean_filename(acc_name)}.sbd"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def update_unified_folders(self):
        smap = self.get_server_map()
        if not smap:
            return
        inbox_scopes, sent_scopes, trash_scopes, junk_scopes, draft_scopes, all_tree = [], [], [], [], [], []
        fc_file = self.profile / 'folderCache.json'
        try:
            with open(fc_file, 'r', encoding='utf-8') as f:
                folder_cache = json.load(f)
        except Exception:
            folder_cache = {}

        for user, info in smap.items():
            stype = info['type']
            host = info['host']
            tdir = Path(info['dir'])
            tdir.mkdir(parents=True, exist_ok=True)
            escaped_user = quote(user)
            base_uri = f"imap://{escaped_user}@{host}" if stype == 'imap' else f"mailbox://{escaped_user}@{host}"
            all_tree.append(base_uri)

            # 1. INBOX
            inbox_name = 'INBOX' if stype == 'imap' else 'Inbox'
            inbox_f = tdir / inbox_name
            if not inbox_f.exists(): inbox_f.touch()
            inbox_msf = tdir / f"{inbox_name}.msf"
            if not inbox_msf.exists(): inbox_msf.touch()
            folder_cache[str(inbox_msf)] = {'flags': FLAGS_INBOX, 'onlineName': inbox_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': inbox_f.stat().st_size}
            inbox_uri = f"{base_uri}/{inbox_name}"
            inbox_scopes.append(inbox_uri)
            all_tree.append(inbox_uri)

            # 2. Sent
            sent_f = tdir / 'Sent'
            if not sent_f.exists(): sent_f.touch()
            sent_msf = tdir / 'Sent.msf'
            if not sent_msf.exists(): sent_msf.touch()
            folder_cache[str(sent_msf)] = {'flags': FLAGS_SENT, 'onlineName': 'Sent', 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': sent_f.stat().st_size}
            sent_uri = f"{base_uri}/Sent"
            sent_scopes.append(sent_uri)
            all_tree.append(sent_uri)

            # 3. Trash / Deleted
            trash_name = 'Deleted' if 'outlook.office365.com' in host else 'Trash'
            trash_f = tdir / trash_name
            if not trash_f.exists(): trash_f.touch()
            trash_msf = tdir / f"{trash_name}.msf"
            if not trash_msf.exists(): trash_msf.touch()
            folder_cache[str(trash_msf)] = {'flags': FLAGS_TRASH, 'onlineName': trash_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': trash_f.stat().st_size}
            trash_uri = f"{base_uri}/{trash_name}"
            trash_scopes.append(trash_uri)
            all_tree.append(trash_uri)

            # 4. Junk / Spam / Bulk
            junk_name = 'Bulk' if 'yahoo' in host else ('Spam' if 'spacemail' in host or 'wana' in host else 'Junk')
            junk_f = tdir / junk_name
            if not junk_f.exists(): junk_f.touch()
            junk_msf = tdir / f"{junk_name}.msf"
            if not junk_msf.exists(): junk_msf.touch()
            folder_cache[str(junk_msf)] = {'flags': FLAGS_JUNK, 'onlineName': junk_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': junk_f.stat().st_size}
            junk_uri = f"{base_uri}/{junk_name}"
            junk_scopes.append(junk_uri)
            all_tree.append(junk_uri)

            # 5. Drafts
            draft_name = 'Draft' if 'yahoo' in host else 'Drafts'
            draft_f = tdir / draft_name
            if not draft_f.exists(): draft_f.touch()
            draft_msf = tdir / f"{draft_name}.msf"
            if not draft_msf.exists(): draft_msf.touch()
            folder_cache[str(draft_msf)] = {'flags': FLAGS_DRAFTS, 'onlineName': draft_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': draft_f.stat().st_size}
            draft_uri = f"{base_uri}/{draft_name}"
            draft_scopes.append(draft_uri)

        with open(fc_file, 'w', encoding='utf-8') as f:
            json.dump(folder_cache, f, indent=2)

        vf_file = self.profile / 'virtualFolders.dat'
        vf_content = f"""version=1
uri=mailbox://nobody@smart%20mailboxes/tags/%24label1
searchFolderFlag=1000
scope=*
terms=AND (tag,contains,$label1)
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/tags/%24label2
searchFolderFlag=1000
scope=*
terms=AND (tag,contains,$label2)
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/tags/%24label3
searchFolderFlag=1000
scope=*
terms=AND (tag,contains,$label3)
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/tags/%24label4
searchFolderFlag=1000
scope=*
terms=AND (tag,contains,$label4)
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/tags/%24label5
searchFolderFlag=1000
scope=*
terms=AND (tag,contains,$label5)
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Inbox
searchFolderFlag=1000
scope={'|'.join(inbox_scopes)}
terms=ALL
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Drafts
searchFolderFlag=400
scope={'|'.join(draft_scopes)}
terms=ALL
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Sent
searchFolderFlag=200
scope={'|'.join(sent_scopes)}
terms=ALL
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Archives
searchFolderFlag=4000
scope=imap://dmtrsx%40yahoo.gr@imap.mail.yahoo.com/Archive
terms=ALL
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Junk
searchFolderFlag=40000000
scope={'|'.join(junk_scopes)}
terms=ALL
searchOnline=false
uri=mailbox://nobody@smart%20mailboxes/Trash
searchFolderFlag=100
scope={'|'.join(trash_scopes)}|mailbox://nobody@Local%20Folders/Trash
terms=ALL
searchOnline=false
"""
        with open(vf_file, 'w', encoding='utf-8') as f:
            f.write(vf_content)

        ft_file = self.profile / 'folderTree.json'
        smart_mailboxes = [
            "mailbox://nobody@smart%20mailboxes",
            "mailbox://nobody@smart%20mailboxes/Inbox",
            "mailbox://nobody@smart%20mailboxes/Drafts",
            "mailbox://nobody@smart%20mailboxes/Templates",
            "mailbox://nobody@smart%20mailboxes/Sent",
            "mailbox://nobody@smart%20mailboxes/Archives",
            "mailbox://nobody@smart%20mailboxes/Junk",
            "mailbox://nobody@smart%20mailboxes/Trash"
        ]
        folder_tree = {
            "open": {
                "all": list(set(all_tree + smart_mailboxes)),
                "smart": list(set(smart_mailboxes + inbox_scopes + sent_scopes + trash_scopes + junk_scopes))
            }
        }
        with open(ft_file, 'w', encoding='utf-8') as f:
            json.dump(folder_tree, f, indent=2)

    def write_mbox_message(self, mbox_file_handle, msg_dict):
        recv_dt = msg_dict['received_date']
        date_str = recv_dt.strftime("%a %b %d %H:%M:%S %Y") if recv_dt else "Mon Oct 16 10:00:00 2026"
        status_line = b"X-Mozilla-Status: 0001\r\n" if msg_dict['is_read'] else b"X-Mozilla-Status: 0000\r\n"
        
        mbox_file_handle.write(f"From - {date_str}\r\n".encode('ascii'))
        mbox_file_handle.write(status_line)
        mbox_file_handle.write(b"X-Mozilla-Status2: 00000000\r\n")
        
        for line in msg_dict['eml_bytes'].split(b'\n'):
            clean_l = line.rstrip(b'\r')
            if clean_l.startswith(b'From '):
                mbox_file_handle.write(b'>' + clean_l + b'\r\n')
            else:
                mbox_file_handle.write(clean_l + b'\r\n')
        mbox_file_handle.write(b'\r\n')

class OutlookTarget:
    def __init__(self, mode='pst', output_dir=None):
        self.mode = mode # 'pst' or 'profile'
        self.output_dir = Path(output_dir) if output_dir else Path(os.path.expanduser('~/Documents/Outlook Files'))
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.outlook = None
        self.ns = None

    def connect(self):
        if not HAS_WIN32COM:
            raise RuntimeError("pywin32 (win32com) is not installed. Outlook integration requires pywin32 on Windows.")
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")

    def get_or_create_store(self, account_name):
        if self.mode == 'profile':
            return self.ns.GetDefaultFolder(6).Parent
        
        pst_file = self.output_dir / f"eMClient_{clean_filename(account_name)}.pst"
        pst_path_str = str(pst_file.resolve())
        # Check if already added
        for s in self.ns.Stores:
            if s.FilePath and os.path.abspath(s.FilePath).lower() == os.path.abspath(pst_path_str).lower():
                return s.GetRootFolder()

        self.ns.AddStoreEx(pst_path_str, 2) # olStoreUnicode
        for s in self.ns.Stores:
            if s.FilePath and os.path.abspath(s.FilePath).lower() == os.path.abspath(pst_path_str).lower():
                root = s.GetRootFolder()
                try:
                    root.Name = account_name
                except Exception:
                    pass
                return root
        return None

    def get_target_folder(self, root_folder, folder_path_parts):
        curr = root_folder
        # Standard Outlook folder mapping (supports Greek & English)
        name_map = {
            'INBOX': 'Εισερχόμενα',
            'SENT': 'Απεσταλμένα',
            'TRASH': 'Διαγραμμένα',
            'DELETED': 'Διαγραμμένα',
            'JUNK': 'Ανεπιθύμητα email',
            'SPAM': 'Ανεπιθύμητα email',
            'DRAFTS': 'Πρόχειρα',
            'ARCHIVE': 'Αρχείο'
        }
        for part in folder_path_parts:
            mapped_name = name_map.get(part.upper(), part)
            try:
                curr = curr.Folders.Item(mapped_name)
            except Exception:
                try:
                    curr = curr.Folders.Add(mapped_name)
                except Exception:
                    curr = curr.Folders.Item(1) # Fallback to first folder
        return curr

    def import_message(self, target_folder, msg_dict):
        mail = target_folder.Items.Add("IPM.Note")
        mail.Subject = msg_dict['subject']
        if msg_dict['received_date']:
            try:
                mail.ReceivedTime = msg_dict['received_date']
            except Exception:
                pass
        mail.UnRead = not msg_dict['is_read']

        parts = msg_dict.get('parts', {})
        # Extract plain/HTML body
        body_plain = ""
        body_html = ""
        for p_name, (hdr, bdata) in parts.items():
            if bdata:
                text_content = bdata.decode('utf-8', errors='replace')
                if b'text/html' in hdr.lower() or '<html' in text_content.lower() or '<!doctype' in text_content.lower():
                    body_html = text_content
                elif b'text/plain' in hdr.lower() or not body_plain:
                    body_plain = text_content

        if body_html:
            mail.HTMLBody = body_html
        elif body_plain:
            mail.Body = body_plain

        mail.Save()
        return mail
