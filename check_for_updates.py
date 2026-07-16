import os
import logging
import re

import requests

CURRENT_VERSION = "2.3.2"

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

VERSION_URL = os.getenv("VERSION_URL")

logger = logging.getLogger(__name__)


def _version_key(version):
    parts = re.findall(r"\d+", str(version or ""))
    return tuple(int(part) for part in parts)

def get_latest_version():
    if not VERSION_URL:
        logger.warning("VERSION_URL is not set")
        return None

    headers = {
        "Accept": "application/vnd.github.v3.raw"
    }
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"

    try:
        response = requests.get(VERSION_URL, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logger.exception("Error checking for update: %s", e)
        print("Error checking for update:", e)
        return None
    except ValueError as e:
        logger.exception("Update payload was not valid JSON: %s", e)
        print("Error checking for update:", e)
        return None

def is_newer_version(latest_version, current_version):
    return _version_key(latest_version) > _version_key(current_version)

def check_for_update():
    latest_info = get_latest_version()
    if not latest_info:
        return

    latest_version = latest_info.get("version")
    if not latest_version:
        print("Update metadata did not include a version.")
        return

    if is_newer_version(latest_version, CURRENT_VERSION):
        print(f"🔔 Update available: {latest_version}")
        print("Release notes:", latest_info.get("release_notes", ""))
        download_url = latest_info.get("download_url")
        if download_url:
            print("Download here:", download_url)
    else:
        print("✅ You're on the latest version.")

if __name__ == "__main__":
    check_for_update()
