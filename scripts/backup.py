"""
Exports all job applications from Supabase to JSON in backup/.
Run by GitHub Actions every 6 hours. Full rows so restore is lossless.
"""
import os
import json
from pathlib import Path
from datetime import datetime, timezone

url = os.environ.get("SUPABASE_URL", "").strip()
key = os.environ.get("SUPABASE_KEY", "").strip()

if not url or not key:
    raise SystemExit("ERROR: SUPABASE_URL and SUPABASE_KEY must be set in GitHub secrets.")

if "supabase.com/dashboard" in url:
    raise SystemExit(
        f"ERROR: SUPABASE_URL looks like a browser URL, not an API URL.\n"
        f"  Got:      {url}\n"
        f"  Expected: https://<project-id>.supabase.co"
    )

from supabase import create_client
sb = create_client(url, key)

Path("backup").mkdir(exist_ok=True)

# Applications — full rows
rows = sb.table("applications").select("*").order("created_at").execute()
applications = rows.data
with open("backup/applications_backup.json", "w", encoding="utf-8") as f:
    json.dump(applications, f, indent=2, ensure_ascii=False)

# Manifest
manifest = {
    "timestamp":          datetime.now(timezone.utc).isoformat(),
    "application_count":  len(applications),
    "supabase_url":       url,
}
with open("backup/manifest.json", "w", encoding="utf-8") as f:
    json.dump(manifest, f, indent=2)

print(f"OK Backed up {len(applications)} applications at {manifest['timestamp']}")
