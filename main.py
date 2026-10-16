#!/usr/bin/env python3
"""
eM Client Universal Migrator
=============================
Main entry point. Automatically launches the modern Graphical User Interface (GUI),
or executes via command-line arguments when CLI flags are provided.

Usage:
  python main.py              # Opens the friendly UI
  python main.py --gui        # Opens the friendly UI
  python main.py --help       # Shows command-line options
"""

import sys
import io
import argparse

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def run_cli(args):
    from migrator import EMClientReader, ThunderbirdTarget, OutlookTarget
    print("=================================================================")
    print("                eM Client Universal Migrator                     ")
    print("=================================================================")
    print(f"Target: {args.target.upper()}")
    
    reader = EMClientReader(args.emclient_path)
    print(f"Source eM Client: {reader.data_path}")
    tree = reader.scan_data_tree()
    print(f"Discovered {len(tree)} accounts.")

    if args.dry_run:
        print("\n[DRY RUN] Preview of accounts and folders:")
        for acc_name, info in tree.items():
            total_m = sum(f['count'] for f in info['folders'].values())
            print(f"- {acc_name} ({total_m} mails, {info['cal_count']} cal, {info['contact_count']} contacts)")
            for fp, fi in info['folders'].items():
                print(f"    * {fp}: {fi['count']} mails")
        print("\n[DRY RUN] Complete. No changes were made.")
        return

    # Handle targets
    if args.target == 'thunderbird':
        tb = ThunderbirdTarget(args.tb_profile or ThunderbirdTarget.find_default_profile())
        print(f"Target Thunderbird Profile: {tb.profile}")
        tb.backup_prefs()
        if not args.mails_only:
            tb.deploy_accounts(reader.accounts_config)
            print("[✓] Accounts configured in Thunderbird.")
    elif args.target == 'outlook':
        ot = OutlookTarget(mode='pst')
        ot.connect()
        print("[✓] Connected to Outlook MAPI.")

    print("\nMigration complete! Launch with no arguments to use the graphical interface.")

def main():
    # If no arguments provided or --gui passed, launch GUI
    if len(sys.argv) == 1 or '--gui' in sys.argv:
        try:
            import gui
            gui.main()
        except ImportError as e:
            print(f"Error loading GUI: {e}")
            sys.exit(1)
        return

    parser = argparse.ArgumentParser(description="eM Client Universal Migrator (Thunderbird & Outlook)")
    parser.add_argument('--gui', action='store_true', help="Launch the graphical user interface")
    parser.add_argument('--target', choices=['thunderbird', 'outlook', 'export'], default='thunderbird', help="Migration target (default: thunderbird)")
    parser.add_argument('--emclient-path', help="Custom path to eM Client data folder")
    parser.add_argument('--tb-profile', help="Custom path to Thunderbird profile folder")
    parser.add_argument('--dry-run', action='store_true', help="Preview migration without modifying files")
    parser.add_argument('--mails-only', action='store_true', help="Only migrate emails")
    parser.add_argument('--calendar-only', action='store_true', help="Only export calendar events")
    parser.add_argument('--contacts-only', action='store_true', help="Only export contacts")

    args = parser.parse_args()
    run_cli(args)

if __name__ == '__main__':
    main()
