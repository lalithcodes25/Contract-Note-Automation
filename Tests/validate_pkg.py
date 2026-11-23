import importlib.metadata
import os
import sys
import imaplib
from dotenv import load_dotenv

load_dotenv()

email_user = os.getenv("EMAIL_USER")
email_pass = os.getenv("EMAIL_PASS")
imap_server = os.getenv("IMAP_SERVER") or "imap.gmail.com"

print(f"Using Python executable: {sys.executable}")

# Determine base directory (1 level up from current script location)
current_dir = os.path.dirname(os.path.abspath(__file__))
base_dir = os.path.dirname(current_dir)
requirements_file = os.path.join(base_dir, "requirements.txt")

if not os.path.exists(requirements_file):
  print(f"[ERROR] '{requirements_file}' not found!")
  sys.exit(1)

print(f"\n--- Validating Libraries from {requirements_file} ---")
failed = False

with open(requirements_file, "r") as f:
  for line in f:
    line = line.strip()
    # Skip empty lines and comments
    if not line or line.startswith("#"):
      continue

    # Clean package name: strip extras like [base] and version specifiers
    pkg_name = (
        line.split("[")[0].split("==")[0].split(">=")[0].split("<=")[0].strip()
    )

    try:
      # Check if package is installed and get its metadata distribution
      dist = importlib.metadata.distribution(pkg_name)

      # Automatically retrieve the package's actual importable modules from metadata
      modules = []
      top_level_content = dist.read_text("top_level.txt")
      if top_level_content:
        modules = [m.strip() for m in top_level_content.splitlines() if m.strip()]

      # Fallback if top_level.txt isn't present
      if not modules:
        modules = [pkg_name.replace("-", "_")]

      # Test importing at least one module associated with the package
      imported_ok = False
      for mod in modules:
        try:
          __import__(mod)
          imported_ok = True
          break
        except ImportError:
          continue

      if imported_ok:
        print(f" [OK] {line} (v{dist.version}, module: {modules[0]})")
      else:
        print(
            f" [FAIL] Package '{line}' is installed but its module(s)"
            f" {modules} failed to import."
        )
        failed = True

    except importlib.metadata.PackageNotFoundError:
      print(f" [FAIL] Package '{line}' is NOT installed.")
      failed = True
    except Exception as e:
      print(f" [FAIL] Error validating '{line}': {e}")
      failed = True

if failed:
  print(
      "\nSome packages are missing or failed to load. Please run:\n  pip"
      " install -r requirements.txt"
  )
  sys.exit(1)
else:
  print("\n All packages in requirements.txt successfully validated and ready!")


if not email_user or not email_pass:
  raise SystemExit(
      "Missing EMAIL_USER or EMAIL_PASS in environment / .env file."
  )

# Helpful sanity check (does not print the password)
print(f"Testing IMAP login for: {email_user} on {imap_server}")
print(f"Password length: {len(email_pass)} (no spaces: {email_pass.strip() == email_pass})")

mail = None
try:
  mail = imaplib.IMAP4_SSL(imap_server)
  status, response = mail.login(email_user.strip(), email_pass.strip())
  print(f"Login status: {status} | Response: {response}")

  box_status, box_data = mail.select("inbox", readonly=True)
  msg_count = box_data[0].decode() if box_data and box_data[0] else "unknown"
  print(f"Inbox selected ({box_status}). Total messages in INBOX: {msg_count}")

  # Optional: verify Gmail search works for contract notes
  search_status, search_data = mail.search(None, 'X-GM-RAW "contract-notes"')
  ids = search_data[0].split() if search_data and search_data[0] else []
  print(f"Matching contract-note emails found: {len(ids)}")

except imaplib.IMAP4.error as e:
  print(f"IMAP authentication/protocol error: {e}")
  raise
finally:
  if mail is not None:
    try:
      mail.logout()
    except Exception:
      pass
