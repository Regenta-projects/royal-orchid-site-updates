#!/usr/bin/env python3
"""
ROHL Site-Update — "RUNNER" (scheduled job, runs from company GitHub Actions)
------------------------------------------------------------------------------
Every few minutes it:
  1. Reads new records from the WorkDrive _inbox folder
  2. For each: finds that hotel's deck (by ROHL id) in the LIVE folder,
     downloads it + the submitted photo/render, runs the Builder engine,
  3. Saves the deck back OVER THE SAME FILE (new version) so the sheet link
     never changes, then moves the record to _processed.
It talks only to WorkDrive. It never touches the Zoho sheet.

ENV (set as GitHub Action secrets):
  ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, ZOHO_REFRESH_TOKEN
  LIVE_FOLDER      - WorkDrive id of the "Pan India Presentations - Live" folder
  INBOX_FOLDER     - WorkDrive id of the "_inbox" folder
  PROCESSED_FOLDER - WorkDrive id of a "_processed" folder (audit trail)
"""
import io, os, json, tempfile, traceback, requests
from pptx import Presentation
import builder as B

ACCOUNTS = "https://accounts.zoho.in"
WD = "https://workdrive.zoho.in/api/v1"

# ---- auth -------------------------------------------------------------------
def token():
    r = requests.post(f"{ACCOUNTS}/oauth/v2/token", data={
        "grant_type": "refresh_token",
        "client_id": os.environ["ZOHO_CLIENT_ID"],
        "client_secret": os.environ["ZOHO_CLIENT_SECRET"],
        "refresh_token": os.environ["ZOHO_REFRESH_TOKEN"],
    }, timeout=30)
    d = r.json()
    if "access_token" not in d:
        raise SystemExit(f"Zoho auth failed: {d}")
    return d["access_token"]

def B(t):  # Bearer — the style WorkDrive actually accepts here
    return {"Authorization": f"Bearer {t}"}
def Z(t):  # fallback
    return {"Authorization": f"Zoho-oauthtoken {t}", "Accept": "application/vnd.api+json"}

# ---- WorkDrive helpers ------------------------------------------------------
def list_files(t, folder_id):
    for h in (B(t), Z(t)):
        r = requests.get(f"{WD}/files/{folder_id}/files?page%5Blimit%5D=200", headers=h, timeout=30)
        try:
            d = r.json()
        except Exception:
            d = {}
        if isinstance(d.get("data"), list):
            return d["data"]
    return []

def download(t, file_id):
    for h in (B(t), Z(t)):
        r = requests.get(f"{WD}/files/{file_id}/content", headers=h, timeout=120)
        if r.ok and len(r.content) > 50:
            return r.content
    raise RuntimeError(f"download failed for file {file_id}")

def upload_version(t, parent_id, name, data, ctype):
    """Upload as a new version of the same-named file -> keeps its link stable."""
    files = {"content": (name, io.BytesIO(data), ctype)}
    dataf = {"parent_id": parent_id, "override-name-exist": "true"}
    r = requests.post(f"{WD}/upload", headers=B(t), files=files, data=dataf, timeout=180)
    r.raise_for_status()
    return r.json()

def move_file(t, file_id, new_parent):
    try:
        requests.patch(f"{WD}/files/{file_id}",
                       headers={**B(t), "Content-Type": "application/json"},
                       json={"data": {"type": "files", "id": file_id,
                                      "attributes": {"parent_id": new_parent}}}, timeout=30)
    except Exception:
        pass

# ---- deck lookup by ROHL id -------------------------------------------------
def find_deck(t, live_folder, rohl, hotel):
    files = list_files(t, live_folder)
    rl = (rohl or "").lower()
    for f in files:
        a = f.get("attributes", {})
        name = (a.get("name") or "")
        if a.get("is_folder"):
            continue
        if rl and name.lower().startswith(rl):
            return f["id"], name
    # fallback: match by hotel name fragment
    hl = (hotel or "").lower().strip()
    if hl:
        for f in files:
            a = f.get("attributes", {})
            if not a.get("is_folder") and hl and hl in (a.get("name") or "").lower():
                return f["id"], a.get("name")
    return None, None

def save_tmp(data, suffix):
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path

# ---- process one inbox record ----------------------------------------------
def process(t, rec_file, live, processed):
    rec = json.loads(download(t, rec_file["id"]).decode("utf-8"))
    rohl, hotel = rec.get("rohl", ""), rec.get("hotel", "")
    deck_id, deck_name = find_deck(t, live, rohl, hotel)
    if not deck_id:
        print(f"  ! no deck for {rohl or hotel} — leaving in inbox")
        return False

    deck_path = save_tmp(download(t, deck_id), ".pptx")
    sub = {
        "space": rec.get("space", ""),
        "variant": rec.get("variant") or None,
        "videos": rec.get("videos", {}) or {},
        "feedback": rec.get("feedback", ""),
        "date": rec.get("date", ""),
    }
    if rec.get("photoFileId"):
        sub["photo"] = save_tmp(download(t, rec["photoFileId"]), ".jpg")
    if rec.get("renderFileId"):
        sub["render"] = save_tmp(download(t, rec["renderFileId"]), ".jpg")

    prs = Presentation(deck_path)
    B.apply_submission(prs, sub)
    out = deck_path + ".out.pptx"
    prs.save(out)
    with open(out, "rb") as f:
        upload_version(t, live, deck_name, f.read(),
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    print(f"  ✓ updated {deck_name}  ({rec.get('space')}{' / ' + rec['variant'] if rec.get('variant') else ''})")
    if processed:
        move_file(t, rec_file["id"], processed)
    return True

def main():
    live = os.environ["LIVE_FOLDER"]
    inbox = os.environ["INBOX_FOLDER"]
    processed = os.environ.get("PROCESSED_FOLDER", "")
    t = token()
    records = [f for f in list_files(t, inbox)
               if not f.get("attributes", {}).get("is_folder")
               and (f.get("attributes", {}).get("name") or "").endswith(".json")]
    print(f"inbox: {len(records)} record(s)")
    ok = 0
    for rec in records:
        try:
            if process(t, rec, live, processed):
                ok += 1
        except Exception:
            print("  ! error on a record:\n" + traceback.format_exc())
    print(f"done: {ok}/{len(records)} applied")

if __name__ == "__main__":
    main()
