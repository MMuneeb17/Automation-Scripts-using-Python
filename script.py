import socket
import ipaddress
import os
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

NETWORK = "192.168.1.0/24"  # corrected subnet mask
PORT = 5900
VNC_PASSWORD = "Fccl@0987"
OUTPUT_FOLDER = "vnc_shortcuts"
TIMEOUT = 1  # seconds

def ping(ip):
    # Ping with 1 echo request, Windows style
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

def sanitize_filename(name):
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in name)

def create_vnc_file(ip, hostname):
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    filename_base = hostname if hostname else f"Unknown_{str(ip).replace('.', '_')}"
    safe_filename = sanitize_filename(filename_base)
    filepath = os.path.join(OUTPUT_FOLDER, f"{safe_filename}.vnc")

    content = f"""[connection]
host={ip}
port={PORT}
name={hostname if hostname else 'Unknown'}
password={VNC_PASSWORD}
"""

    with open(filepath, "w") as f:
        f.write(content)

    print(f"[+] Created: {filepath}")

def scan_ip(ip):
    if ping(ip) and is_port_open(ip):
        hostname = get_hostname(ip)
        create_vnc_file(ip, hostname)
        return (str(ip), hostname if hostname else "Unknown")
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
                ip, hostname = result
                print(f"{ip} -> {hostname}")
                results.append(result)

    print(f"\nScan complete. Found {len(results)} VNC servers.")
    print(f"Shortcut files saved in '{OUTPUT_FOLDER}' directory.")

if __name__ == "__main__":
    main()
