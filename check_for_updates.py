import requests

# Your current app version (can be imported or stored in a config)
CURRENT_VERSION = "2.3.2"

# Your GitHub access token (⚠️ Never hardcode in public production builds)
GITHUB_TOKEN = "ghp_your_token_here"

# Raw URL to version.json in your private GitHub repo
VERSION_URL = "https://raw.githubusercontent.com/johnwassfy/St.Mary-Maadi-Liturgies/master/version.json"

def get_latest_version():
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3.raw"
    }
    try:
        response = requests.get(VERSION_URL, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print("Error checking for update:", e)
        return None

def is_newer_version(latest_version, current_version):
    def parse(v): return [int(x) for x in v.split(".")]
    return parse(latest_version) > parse(current_version)

def check_for_update():
    latest_info = get_latest_version()
    if not latest_info:
        return

    latest_version = latest_info["version"]
    if is_newer_version(latest_version, CURRENT_VERSION):
        print(f"🔔 Update available: {latest_version}")
        print("Release notes:", latest_info.get("release_notes", ""))
        print("Download here:", latest_info["download_url"])
    else:
        print("✅ You're on the latest version.")

if __name__ == "__main__":
    check_for_update()
