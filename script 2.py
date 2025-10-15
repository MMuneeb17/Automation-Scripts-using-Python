import socket
import ipaddress
import os
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

NETWORK = "192.168.1.0/24"
PORT = 5900
VNC_PASSWORD = "Fccl@0987"
OUTPUT_FOLDER = "vnc_shortcuts"
TIMEOUT = 1

def ping(ip):
    command = ['ping', '-n', '1', '-w', '1000', str(ip)]
    result = subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return result.returncode == 0

def is_port_open(ip, port=PORT):
    try:
        with socket.create_connection((str(ip), port), timeout=TIMEOUT):
            return True
    except:
        return False

def get_hostname(ip):
    try:
        return socket.gethostbyaddr(str(ip))[0]
    except socket.herror:
        return None

def get_username(ip):
    try:
        output = subprocess.check_output(["nbtstat", "-A", str(ip)], stderr=subprocess.DEVNULL).decode(errors='ignore')
        for line in output.splitlines():
            if "<03>" in line or "<20>" in line:
                parts = line.split()
                if len(parts) >= 1:
                    name = parts[0].strip()
                    if name and not name.endswith("<03>") and not name.endswith("<20>"):
                        return name
        for line in output.splitlines():
            if "<20>" in line:
                parts = line.split()
                if len(parts) >= 1:
                    return parts[0].strip()
    except:
        return None
    return None

def sanitize_filename(name):
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in name)

def create_vnc_file(ip, hostname, username):
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    user_part = username if username else "Unknown"
    ip_part = str(ip).replace('.', '_')
    filename_base = f"{user_part}_{ip_part}"
    safe_filename = sanitize_filename(filename_base)
    filepath = os.path.join(OUTPUT_FOLDER, f"{safe_filename}.vnc")

    content = f"""[connection]
host={ip}
port={PORT}
name={hostname if hostname else 'Unknown'}
username={username if username else 'Unknown'}
password={VNC_PASSWORD}
"""

    with open(filepath, "w") as f:
        f.write(content)

    print(f"[+] Created: {filepath}")

def scan_ip(ip):
    if ping(ip) and is_port_open(ip):
        hostname = get_hostname(ip)
        username = get_username(ip)
        create_vnc_file(ip, hostname, username)
        return (str(ip), hostname if hostname else "Unknown", username if username else "Unknown")
    return None

def main():
    print(f"Starting scan on network {NETWORK}...\n")
    ips = list(ipaddress.ip_network(NETWORK).hosts())
    results = []

    with ThreadPoolExecutor(max_workers=50) as executor:
        futures = {executor.submit(scan_ip, ip): ip for ip in ips}
        for future in as_completed(futures):
            result = future.result()
            if result:
                ip, hostname, username = result
                print(f"{ip} -> {hostname} (User: {username})")
                results.append(result)

    print(f"\nScan complete. Found {len(results)} VNC servers.")
    print(f"Shortcut files saved in '{OUTPUT_FOLDER}' directory.")

if __name__ == "__main__":
    main()
