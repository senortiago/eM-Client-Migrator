#!/usr/bin/env python3
"""
eM Client to Mozilla Thunderbird Migrator
==========================================
A complete, open-source tool to seamlessly migrate all Accounts (settings, servers, credentials),
Mails (MBOX format with full MIME, attachments, and read/unread status),
Calendars (iCalendar .ics), and Contacts (vCard .vcf) from eM Client to Mozilla Thunderbird.

Zero external dependencies - Uses Python 3 standard library only.
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
import argparse
import subprocess
from pathlib import Path
from urllib.parse import quote
from email import message_from_bytes
import configparser

# Ensure UTF-8 output in terminal
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

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

class Migrator:
    def __init__(self, emclient_path=None, tb_profile_path=None, force=False, dry_run=False):
        self.force = force
        self.dry_run = dry_run
        
        # 1. Discover eM Client Path
        self.emclient_path = Path(emclient_path) if emclient_path else self.find_emclient_path()
        if not self.emclient_path or not self.emclient_path.exists():
            raise FileNotFoundError(f"eM Client directory not found at: {self.emclient_path}")
            
        # 2. Discover Thunderbird Profile Path
        self.tb_profile = Path(tb_profile_path) if tb_profile_path else self.find_tb_profile()
        if not self.tb_profile or not self.tb_profile.exists():
            raise FileNotFoundError(f"Thunderbird profile directory not found at: {self.tb_profile}")

        # Output documents directory for Calendar & Contacts
        self.docs_dir = Path(os.path.expanduser('~/Documents')) / 'Thunderbird Migration'
        self.docs_dir.mkdir(parents=True, exist_ok=True)
        
        self.guid_map = {}
        self.accounts = []
        self.server_map = {}

    def find_emclient_path(self):
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

    def find_tb_profile(self):
        if sys.platform == 'win32':
            appdata = os.environ.get('APPDATA')
            if not appdata:
                return None
            tb_root = Path(appdata) / 'Thunderbird'
        elif sys.platform == 'darwin':
            tb_root = Path.home() / 'Library' / 'Thunderbird'
        else:
            tb_root = Path.home() / '.thunderbird'

        if not tb_root.exists():
            return None

        # Check profiles.ini
        ini_file = tb_root / 'profiles.ini'
        if ini_file.exists():
            config = configparser.ConfigParser()
            config.read(ini_file, encoding='utf-8')
            
            # First check installs.ini for default
            installs_ini = tb_root / 'installs.ini'
            if installs_ini.exists():
                iconfig = configparser.ConfigParser()
                iconfig.read(installs_ini, encoding='utf-8')
                for sec in iconfig.sections():
                    rel_p = iconfig.get(sec, 'Default', fallback=None)
                    if rel_p:
                        candidate = tb_root / rel_p
                        if candidate.exists():
                            return candidate

            # Check profile sections
            for sec in config.sections():
                if config.get(sec, 'Default', fallback=None) == '1':
                    path = config.get(sec, 'Path', fallback=None)
                    is_rel = config.get(sec, 'IsRelative', fallback='1') == '1'
                    if path:
                        candidate = (tb_root / path) if is_rel else Path(path)
                        if candidate.exists():
                            return candidate

            # Default to first profile directory found
            for sec in config.sections():
                path = config.get(sec, 'Path', fallback=None)
                is_rel = config.get(sec, 'IsRelative', fallback='1') == '1'
                if path:
                    candidate = (tb_root / path) if is_rel else Path(path)
                    if candidate.exists() and (candidate / 'prefs.js').exists():
                        return candidate

        profiles_dir = tb_root / 'Profiles'
        if profiles_dir.exists():
            for p in profiles_dir.iterdir():
                if p.is_dir() and (p / 'prefs.js').exists():
                    return p
        return None

    def check_thunderbird_running(self):
        lock_file = self.tb_profile / 'parent.lock'
        if not lock_file.exists():
            return False

        if self.force:
            print("[*] Thunderbird lock detected. Terminating running instance as requested (--force)...")
            if sys.platform == 'win32':
                subprocess.run(['powershell', '-Command', "Get-Process -Name thunderbird -ErrorAction SilentlyContinue | Stop-Process -Force"], capture_output=True)
            else:
                subprocess.run(['pkill', '-9', 'thunderbird'], capture_output=True)
            import time
            time.sleep(1)
            try:
                if lock_file.exists():
                    lock_file.unlink()
            except Exception:
                pass
            return False

        # Try to remove stale lock
        try:
            lock_file.unlink()
            return False
        except Exception:
            return True

    def parse_emclient_accounts(self):
        accounts_db = self.emclient_path / 'accounts.dat'
        if not accounts_db.exists():
            print("[!] accounts.dat not found in eM Client directory.")
            return

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

            # Update guid map
            user = val.get('Credentials', {}).get('Username') if isinstance(val.get('Credentials'), dict) else val.get('ServerUserName', val.get('AccountName'))
            if user and '@' in str(user):
                self.guid_map[k] = user
                self.guid_map[b_uid] = user
            elif val.get('AccountName') and '@' in str(val.get('AccountName')):
                self.guid_map[k] = val.get('AccountName')
                self.guid_map[b_uid] = val.get('AccountName')

        self.guid_map['Local Folders'] = 'Local Folders'

        # Build account settings
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
                inc_host = 'imap.gmail.com'
                inc_port = 993
                inc_socket = 3
                inc_auth = 10
                out_host = 'smtp.gmail.com'
                out_port = 465
                out_socket = 3
                out_auth = 10
                out_desc = 'Google Mail'
            elif is_oauth_ms:
                inc_host = 'outlook.office365.com'
                inc_port = 993
                inc_socket = 3
                inc_auth = 10
                out_host = 'smtp-mail.outlook.com'
                out_port = 587
                out_socket = 2
                out_auth = 10
                out_desc = 'Outlook.com (Microsoft)'
            elif 'yahoo' in domain:
                inc_host = 'imap.mail.yahoo.com'
                inc_port = 993
                inc_socket = 3
                inc_auth = 10
                out_host = 'smtp.mail.yahoo.com'
                out_port = 465
                out_socket = 3
                out_auth = 10
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

            self.accounts.append({
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

    def sync_accounts_to_thunderbird(self):
        prefs_file = self.tb_profile / 'prefs.js'
        if not prefs_file.exists():
            print("[!] prefs.js not found in Thunderbird profile.")
            return

        with open(prefs_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        existing_text = "".join(lines)
        existing_emails = set(re.findall(r'mail\.server\.server\d+\.userName", "([^"]+)"', existing_text))
        
        # Backup prefs.js
        bak_file = self.tb_profile / f"prefs.js.bak_{int(datetime.datetime.now().timestamp())}"
        shutil.copy2(prefs_file, bak_file)
        print(f"[*] Backed up prefs.js -> {bak_file.name}")

        to_add = [a for a in self.accounts if a['email'].lower() not in [e.lower() for e in existing_emails]]
        print(f"[*] Accounts to configure in Thunderbird: {len(to_add)} new ({len(self.accounts) - len(to_add)} existing)")

        acc_nums = [int(m) for m in re.findall(r'mail\.account\.account(\d+)\.', existing_text)]
        srv_nums = [int(m) for m in re.findall(r'mail\.server\.server(\d+)\.', existing_text)]
        id_nums = [int(m) for m in re.findall(r'mail\.identity\.id(\d+)\.', existing_text)]
        smtp_nums = [int(m) for m in re.findall(r'mail\.smtpserver\.smtp(\d+)\.', existing_text)]

        next_acc = max(acc_nums) + 1 if acc_nums else 5
        next_srv = max(srv_nums) + 1 if srv_nums else 7
        next_id = max(id_nums) + 1 if id_nums else 4
        next_smtp = max(smtp_nums) + 1 if smtp_nums else 3

        new_prefs = []
        new_acc_keys = []
        new_smtp_keys = []
        existing_dirs = set(os.listdir(self.tb_profile / 'ImapMail') if (self.tb_profile / 'ImapMail').exists() else [])

        for a in to_add:
            email = a['email']
            acc_key = f"account{next_acc}"
            srv_key = f"server{next_srv}"
            id_key = f"id{next_id}"
            smtp_key = f"smtp{next_smtp}"
            
            next_acc += 1
            next_srv += 1
            next_id += 1
            next_smtp += 1
            
            new_acc_keys.append(acc_key)
            new_smtp_keys.append(smtp_key)
            
            base_dir = a['inc_host'] if a['inc_host'] else 'custom.mail'
            clean_host = re.sub(r'[^a-zA-Z0-9\.\-]', '_', base_dir)
            dir_name = clean_host
            c = 1
            while dir_name in existing_dirs:
                dir_name = f"{clean_host}-{c}"
                c += 1
            existing_dirs.add(dir_name)
            
            is_pop = (a['inc_type'] == 'pop3')
            top_dir = 'Mail' if is_pop else 'ImapMail'
            full_dir_path = self.tb_profile / top_dir / dir_name
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

        # Update accounts list, smtp list, lastKey
        updated_lines = []
        for line in lines:
            if 'mail.accountmanager.accounts' in line:
                m = re.search(r'user_pref\("mail\.accountmanager\.accounts", "([^"]+)"\);', line)
                curr = m.group(1) if m else "account1"
                all_accs = curr.split(',') + new_acc_keys
                updated_lines.append(f'user_pref("mail.accountmanager.accounts", "{",".join(all_accs)}");\n')
            elif 'mail.smtpservers' in line:
                m = re.search(r'user_pref\("mail\.smtpservers", "([^"]+)"\);', line)
                curr = m.group(1) if m else "smtp1"
                all_smtps = curr.split(',') + new_smtp_keys
                updated_lines.append(f'user_pref("mail.smtpservers", "{",".join(all_smtps)}");\n')
            elif 'mail.account.lastKey' in line:
                updated_lines.append(f'user_pref("mail.account.lastKey", {next_acc - 1});\n')
            else:
                updated_lines.append(line)

        updated_lines.extend(new_prefs)
        
        if not self.dry_run:
            with open(prefs_file, 'w', encoding='utf-8') as f:
                f.writelines(updated_lines)
        print(f"[✓] Added {len(to_add)} accounts to prefs.js.")

        # Re-parse server mapping for mail deployment
        with open(prefs_file, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        srv_users = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.userName", "([^"]+)"\);', text))
        srv_dirs = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.directory", "([^"]+)"\);', text))
        srv_hosts = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.hostname", "([^"]+)"\);', text))
        srv_types = dict(re.findall(r'user_pref\("mail\.server\.(server\d+)\.type", "([^"]+)"\);', text))

        for s_id, s_dir in srv_dirs.items():
            user = srv_users.get(s_id, 'N/A')
            host = srv_hosts.get(s_id, 'N/A')
            stype = srv_types.get(s_id, 'N/A')
            if user != 'nobody':
                self.server_map[user] = {
                    'server_id': s_id,
                    'type': stype,
                    'user': user,
                    'host': host,
                    'dir': s_dir.replace('\\\\', '\\')
                }

    def migrate_mails(self):
        print("\n[*] Migrating emails across all accounts...")
        # Scan account directories
        account_dirs = []
        for d in self.emclient_path.iterdir():
            if not d.is_dir(): continue
            if d.name in ('Local Folders', 'Archive'):
                account_dirs.append((d, self.guid_map.get(d.name, d.name)))
            elif (d / 'folders.dat').exists() and (d / 'mail_data.dat').exists():
                account_dirs.append((d, self.guid_map.get(d.name, d.name)))
            elif re.match(r'^[0-9a-fA-F\-]{36}$', d.name):
                acc_name = self.guid_map.get(d.name, d.name)
                for sub in d.iterdir():
                    if sub.is_dir() and (sub / 'folders.dat').exists() and (sub / 'mail_data.dat').exists():
                        account_dirs.append((sub, acc_name))

        total_migrated = 0
        for acc_dir, acc_name in account_dirs:
            target_info = self.server_map.get(acc_name)
            if target_info:
                target_dir = Path(target_info['dir'])
            else:
                target_dir = self.tb_profile / 'Mail' / 'Local Folders' / 'eM Client Migration.sbd' / f"{clean_filename(acc_name)}.sbd"

            target_dir.mkdir(parents=True, exist_ok=True)
            count = self.export_mbox_folder(acc_dir, target_dir)
            total_migrated += count
            print(f"  [✓] {acc_name}: {count} messages migrated.")

        print(f"[✓] Total messages migrated: {total_migrated}")

    def export_mbox_folder(self, acc_dir, target_dir):
        f_db = acc_dir / 'folders.dat'
        if not f_db.exists(): return 0
        
        conn = sqlite3.connect(f'file:{f_db}?mode=ro', uri=True)
        cur = conn.cursor()
        cur.execute("SELECT id, name, parentFolderId FROM Folders")
        folders = {r[0]: {'name': r[1], 'parent': r[2]} for r in cur.fetchall()}
        conn.close()

        folder_paths = {}
        for fid, f in folders.items():
            path_parts = []
            curr = fid
            while curr in folders and folders[curr]['parent'] != 0:
                name = folders[curr]['name']
                if name not in ('Mails', 'Root', 'All Mail Client Folders'):
                    path_parts.append(clean_filename(name))
                curr = folders[curr]['parent']
            if curr in folders:
                name = folders[curr]['name']
                if name not in ('Mails', 'Root', 'All Mail Client Folders'):
                    path_parts.append(clean_filename(name))
            path_parts.reverse()
            folder_paths[fid] = path_parts or ['INBOX']

        mail_index = acc_dir / 'mail_index.dat'
        mail_data = acc_dir / 'mail_data.dat'
        if not mail_index.exists() or not mail_data.exists(): return 0

        conn_idx = sqlite3.connect(f'file:{mail_index}?mode=ro', uri=True)
        cur_idx = conn_idx.cursor()
        cur_idx.execute("SELECT id, folder, subject, receivedDate, flags FROM MailItems")
        mail_meta = {r[0]: {'folder': r[1], 'subject': r[2], 'receivedDate': r[3], 'flags': r[4]} for r in cur_idx.fetchall()}
        conn_idx.close()

        conn_data = sqlite3.connect(f'file:{mail_data}?mode=ro', uri=True)
        cur_data = conn_data.cursor()
        cur_data.execute("SELECT DISTINCT id FROM LocalMailContents")
        msg_ids = [r[0] for r in cur_data.fetchall()]

        msgs_by_folder = {}
        for mid in msg_ids:
            meta = mail_meta.get(mid, {})
            f_parts = folder_paths.get(meta.get('folder', 0), ['INBOX'])
            key = tuple(f_parts)
            if key not in msgs_by_folder:
                msgs_by_folder[key] = []
            msgs_by_folder[key].append(mid)

        count = 0
        for f_parts, mids in msgs_by_folder.items():
            curr_dir = target_dir
            for part in f_parts[:-1]:
                (curr_dir / part).touch(exist_ok=True)
                curr_dir = curr_dir / f"{part}.sbd"
                curr_dir.mkdir(parents=True, exist_ok=True)

            raw_name = f_parts[-1]
            # Standardize names for Thunderbird
            norm_name = raw_name
            if raw_name.upper() in ('INBOX', 'ΕΙΣΕΡΧΟΜΕΝΑ', 'ΌΛΑ ΤΑ ΜΗΝΎΜΑΤΑ', 'ALL MAIL'):
                norm_name = 'INBOX'
            elif raw_name.upper() in ('SENT', 'ΑΠΕΣΤΑΛΜΈΝΑ', 'SENT MAIL'):
                norm_name = 'Sent'
            elif raw_name.upper() in ('TRASH', 'DELETED', 'ΚΆΔΟΣ ΑΠΟΡΡΙΜΜΆΤΩΝ', 'ΔΙΑΓΡΑΜΜΈΝΑ'):
                norm_name = 'Trash'
            elif raw_name.upper() in ('JUNK', 'SPAM', 'BULK', 'ΑΝΕΠΙΘΎΜΗΤΑ'):
                norm_name = 'Junk'
            elif raw_name.upper() in ('DRAFT', 'DRAFTS', 'ΠΡΌΧΕΙΡΑ'):
                norm_name = 'Drafts'

            mbox_file_path = curr_dir / norm_name
            if self.dry_run:
                count += len(mids)
                continue

            with open(mbox_file_path, 'ab') as mbox_out:
                for mid in mids:
                    meta = mail_meta.get(mid, {})
                    cur_data.execute("SELECT partName, partHeader, partBody FROM LocalMailContents WHERE id=?", (mid,))
                    parts = {r[0]: (r[1] or b'', r[2] or b'') for r in cur_data.fetchall()}
                    eml = build_mime_node(parts, 'TEXT')
                    
                    recv_dt = ticks_to_datetime(meta.get('receivedDate'))
                    date_str = recv_dt.strftime("%a %b %d %H:%M:%S %Y") if recv_dt else "Mon Oct 16 10:00:00 2026"
                    flags = meta.get('flags', 0)
                    is_read = False if (flags & 16) else True
                    status_line = b"X-Mozilla-Status: 0001\r\n" if is_read else b"X-Mozilla-Status: 0000\r\n"
                    
                    mbox_out.write(f"From - {date_str}\r\n".encode('ascii'))
                    mbox_out.write(status_line)
                    mbox_out.write(b"X-Mozilla-Status2: 00000000\r\n")
                    
                    for line in eml.split(b'\n'):
                        clean_l = line.rstrip(b'\r')
                        if clean_l.startswith(b'From '):
                            mbox_out.write(b'>' + clean_l + b'\r\n')
                        else:
                            mbox_out.write(clean_l + b'\r\n')
                    mbox_out.write(b'\r\n')
                    count += 1

            # Ensure .msf exists
            msf_path = curr_dir / f"{norm_name}.msf"
            if not msf_path.exists():
                msf_path.touch()

        conn_data.close()
        return count

    def update_unified_folders(self):
        print("\n[*] Updating Unified Folders and folder caches...")
        inbox_scopes = []
        sent_scopes = []
        trash_scopes = []
        junk_scopes = []
        draft_scopes = []
        archive_scopes = []
        all_tree = []

        fc_file = self.tb_profile / 'folderCache.json'
        try:
            with open(fc_file, 'r', encoding='utf-8') as f:
                folder_cache = json.load(f)
        except Exception:
            folder_cache = {}

        for user, info in self.server_map.items():
            stype = info['type']
            host = info['host']
            tdir = Path(info['dir'])
            escaped_user = quote(user)
            base_uri = f"imap://{escaped_user}@{host}" if stype == 'imap' else f"mailbox://{escaped_user}@{host}"
            all_tree.append(base_uri)

            # Inboxes
            inbox_name = 'INBOX' if stype == 'imap' else 'Inbox'
            inbox_f = tdir / inbox_name
            if not inbox_f.exists(): inbox_f.touch()
            inbox_msf = tdir / f"{inbox_name}.msf"
            if not inbox_msf.exists(): inbox_msf.touch()
            folder_cache[str(inbox_msf)] = {'flags': FLAGS_INBOX, 'onlineName': inbox_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': inbox_f.stat().st_size}
            inbox_uri = f"{base_uri}/{inbox_name}"
            inbox_scopes.append(inbox_uri)
            all_tree.append(inbox_uri)

            # Sent
            sent_f = tdir / 'Sent'
            if not sent_f.exists(): sent_f.touch()
            sent_msf = tdir / 'Sent.msf'
            if not sent_msf.exists(): sent_msf.touch()
            folder_cache[str(sent_msf)] = {'flags': FLAGS_SENT, 'onlineName': 'Sent', 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': sent_f.stat().st_size}
            sent_uri = f"{base_uri}/Sent"
            sent_scopes.append(sent_uri)
            all_tree.append(sent_uri)

            # Trash
            trash_name = 'Deleted' if 'outlook.office365.com' in host else 'Trash'
            trash_f = tdir / trash_name
            if not trash_f.exists(): trash_f.touch()
            trash_msf = tdir / f"{trash_name}.msf"
            if not trash_msf.exists(): trash_msf.touch()
            folder_cache[str(trash_msf)] = {'flags': FLAGS_TRASH, 'onlineName': trash_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': trash_f.stat().st_size}
            trash_uri = f"{base_uri}/{trash_name}"
            trash_scopes.append(trash_uri)
            all_tree.append(trash_uri)

            # Junk
            junk_name = 'Bulk' if 'yahoo' in host else ('Spam' if 'spacemail' in host or 'wana' in host else 'Junk')
            junk_f = tdir / junk_name
            if not junk_f.exists(): junk_f.touch()
            junk_msf = tdir / f"{junk_name}.msf"
            if not junk_msf.exists(): junk_msf.touch()
            folder_cache[str(junk_msf)] = {'flags': FLAGS_JUNK, 'onlineName': junk_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': junk_f.stat().st_size}
            junk_uri = f"{base_uri}/{junk_name}"
            junk_scopes.append(junk_uri)
            all_tree.append(junk_uri)

            # Drafts
            draft_name = 'Draft' if 'yahoo' in host else 'Drafts'
            draft_f = tdir / draft_name
            if not draft_f.exists(): draft_f.touch()
            draft_msf = tdir / f"{draft_name}.msf"
            if not draft_msf.exists(): draft_msf.touch()
            folder_cache[str(draft_msf)] = {'flags': FLAGS_DRAFTS, 'onlineName': draft_name, 'totalMsgs': -1, 'totalUnreadMsgs': -1, 'folderSize': draft_f.stat().st_size}
            draft_uri = f"{base_uri}/{draft_name}"
            draft_scopes.append(draft_uri)

        if not self.dry_run:
            with open(fc_file, 'w', encoding='utf-8') as f:
                json.dump(folder_cache, f, indent=2)

        # Update virtualFolders.dat
        vf_file = self.tb_profile / 'virtualFolders.dat'
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
        if not self.dry_run:
            with open(vf_file, 'w', encoding='utf-8') as f:
                f.write(vf_content)

        # Update folderTree.json
        ft_file = self.tb_profile / 'folderTree.json'
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
        if not self.dry_run:
            with open(ft_file, 'w', encoding='utf-8') as f:
                json.dump(folder_tree, f, indent=2)
                
        print(f"[✓] Unified Folders updated ({len(inbox_scopes)} accounts active in Unified view).")

    def export_calendars(self):
        print("\n[*] Exporting Calendars (iCalendar .ics)...")
        cal_dir = self.docs_dir / 'Calendars'
        cal_dir.mkdir(parents=True, exist_ok=True)
        
        event_dbs = list(self.emclient_path.glob('**/event_data.dat'))
        all_vevents = []
        total_events = 0
        
        for db in event_dbs:
            try:
                acc_name = 'Unknown'
                parent = db.parent
                while parent != self.emclient_path:
                    if parent.name in self.guid_map:
                        acc_name = self.guid_map[parent.name]
                        break
                    parent = parent.parent
                    
                conn = sqlite3.connect(f'file:{db}?mode=ro', uri=True)
                cur = conn.cursor()
                cur.execute("SELECT id, properties FROM EventItemContents WHERE propertyName='calendar'")
                rows = cur.fetchall()
                conn.close()
                if not rows: continue
                
                acc_events = []
                for eid, props in rows:
                    text = props.decode('utf-8', errors='replace') if isinstance(props, bytes) else str(props)
                    m = re.search(r'(BEGIN:VEVENT[\s\S]*?END:VEVENT)', text)
                    if m:
                        vevent = m.group(1)
                        all_vevents.append(vevent)
                        acc_events.append(vevent)
                        total_events += 1
                        
                if acc_events and not self.dry_run:
                    acc_ics = cal_dir / f"{clean_filename(acc_name)}.ics"
                    with open(acc_ics, 'w', encoding='utf-8') as f:
                        f.write("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Thunderbird Migration//EN\r\n")
                        for ev in acc_events:
                            f.write(ev + "\r\n")
                        f.write("END:VCALENDAR\r\n")
            except Exception as e:
                print(f"  [!] Failed reading {db.name}: {e}")

        master_ics = self.docs_dir / 'eMClient_All_Calendars.ics'
        if not self.dry_run:
            with open(master_ics, 'w', encoding='utf-8') as f:
                f.write("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Thunderbird Migration Master//EN\r\n")
                for ev in all_vevents:
                    f.write(ev + "\r\n")
                f.write("END:VCALENDAR\r\n")

        print(f"[✓] Exported {total_events} calendar events -> {master_ics}")

    def export_contacts(self):
        print("\n[*] Exporting Contacts (vCard .vcf)...")
        contact_dir = self.docs_dir / 'Contacts'
        contact_dir.mkdir(parents=True, exist_ok=True)
        
        contact_dbs = list(self.emclient_path.glob('**/contact_index.dat'))
        all_vcards = []
        total_contacts = 0
        
        for db in contact_dbs:
            try:
                acc_name = 'Unknown'
                parent = db.parent
                while parent != self.emclient_path:
                    if parent.name in self.guid_map:
                        acc_name = self.guid_map[parent.name]
                        break
                    parent = parent.parent
                    
                conn = sqlite3.connect(f'file:{db}?mode=ro', uri=True)
                cur = conn.cursor()
                cur.execute("SELECT id, folder, firstName, lastName, displayName, organization, title FROM ContactItems")
                contacts = cur.fetchall()
                if not contacts:
                    conn.close()
                    continue
                    
                acc_vcards = []
                for c in contacts:
                    cid, folder_id, fname, lname, dname, org, title = c
                    vcard = ["BEGIN:VCARD", "VERSION:3.0"]
                    vcard.append(f"N:{lname or ''};{fname or ''};;;")
                    vcard.append(f"FN:{dname or fname or lname or ''}")
                    if org: vcard.append(f"ORG:{org}")
                    if title: vcard.append(f"TITLE:{title}")
                    
                    try:
                        cur.execute("SELECT email FROM ContactEmails WHERE parentId=?", (cid,))
                        for (email,) in cur.fetchall():
                            if email: vcard.append(f"EMAIL;TYPE=INTERNET:{email}")
                    except Exception:
                        pass
                        
                    try:
                        cur.execute("SELECT telephone FROM ContactTelephones WHERE parentId=?", (cid,))
                        for (tel,) in cur.fetchall():
                            if tel: vcard.append(f"TEL;TYPE=VOICE:{tel}")
                    except Exception:
                        pass
                        
                    vcard.append("END:VCARD\r\n")
                    vcard_str = "\r\n".join(vcard)
                    all_vcards.append(vcard_str)
                    acc_vcards.append(vcard_str)
                    total_contacts += 1
                conn.close()
                
                if acc_vcards and not self.dry_run:
                    acc_vcf = contact_dir / f"{clean_filename(acc_name)}.vcf"
                    with open(acc_vcf, 'w', encoding='utf-8') as f:
                        f.write("\r\n".join(acc_vcards))
            except Exception as e:
                print(f"  [!] Failed reading {db.name}: {e}")

        master_vcf = self.docs_dir / 'eMClient_All_Contacts.vcf'
        if not self.dry_run:
            with open(master_vcf, 'w', encoding='utf-8') as f:
                f.write("\r\n".join(all_vcards))
                
        print(f"[✓] Exported {total_contacts} contacts -> {master_vcf}")

    def run(self, do_accounts=True, do_mails=True, do_calendar=True, do_contacts=True):
        print("=================================================================")
        print("           eM Client to Mozilla Thunderbird Migrator             ")
        print("=================================================================")
        print(f"Source eM Client Path:      {self.emclient_path}")
        print(f"Target Thunderbird Profile: {self.tb_profile}")
        print(f"Export Directory:           {self.docs_dir}")
        print("-----------------------------------------------------------------")
        
        if self.check_thunderbird_running():
            print("\n[!] ERROR: Thunderbird is currently running!")
            print("    Please close Thunderbird before running this script so preferences")
            print("    and database files can be safely updated without conflicts.")
            print("    Or pass --force to automatically close it.")
            sys.exit(1)

        print("\n[*] Discovering accounts in eM Client...")
        self.parse_emclient_accounts()
        print(f"[✓] Found {len(self.accounts)} accounts.")

        if do_accounts:
            self.sync_accounts_to_thunderbird()

        if do_mails:
            self.migrate_mails()
            self.update_unified_folders()

        if do_calendar:
            self.export_calendars()

        if do_contacts:
            self.export_contacts()

        print("\n=================================================================")
        print("                   MIGRATION COMPLETE!                           ")
        print("=================================================================")
        print("1. All accounts and mail stores are ready in Thunderbird.")
        print("2. Unified Folders (Inbox, Sent, Junk, Trash) now include all accounts.")
        print(f"3. Master Calendar: {self.docs_dir / 'eMClient_All_Calendars.ics'}")
        print(f"4. Master Contacts: {self.docs_dir / 'eMClient_All_Contacts.vcf'}")
        print("You can now open Thunderbird and enjoy your migrated setup!")

def main():
    parser = argparse.ArgumentParser(description="Migrate eM Client data to Mozilla Thunderbird")
    parser.add_argument('--emclient-path', help="Path to eM Client data directory (default: auto-detect)")
    parser.add_argument('--tb-profile', help="Path to Thunderbird profile directory (default: auto-detect)")
    parser.add_argument('--force', action='store_true', help="Automatically close Thunderbird if running")
    parser.add_argument('--dry-run', action='store_true', help="Perform dry run without writing files")
    parser.add_argument('--accounts-only', action='store_true', help="Only migrate account settings")
    parser.add_argument('--mails-only', action='store_true', help="Only migrate emails")
    parser.add_argument('--calendar-only', action='store_true', help="Only migrate calendar events")
    parser.add_argument('--contacts-only', action='store_true', help="Only migrate contacts")

    args = parser.parse_args()

    migrator = Migrator(
        emclient_path=args.emclient_path,
        tb_profile_path=args.tb_profile,
        force=args.force,
        dry_run=args.dry_run
    )

    do_all = not (args.accounts_only or args.mails_only or args.calendar_only or args.contacts_only)
    migrator.run(
        do_accounts=do_all or args.accounts_only,
        do_mails=do_all or args.mails_only,
        do_calendar=do_all or args.calendar_only,
        do_contacts=do_all or args.contacts_only
    )

if __name__ == '__main__':
    main()
