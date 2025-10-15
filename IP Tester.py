#!/usr/bin/env python3
"""
Ping a fixed list of IPs, show status & ping stats in console, and save results to Excel.
Saves file as ping_results_<YYYYMMDD_HHMMSS>.xlsx

Requirements:
    pip install pandas openpyxl
"""

import subprocess
import platform
import socket
import re
import concurrent.futures
from datetime import datetime
import pandas as pd

# -- Configuration: the named IPs you provided --
targets = [
    ("EMS", "192.168.1.25"),
    ("SAP", "10.10.254.202"),
    ("IP-3", "192.168.1.3"),
    ("IP-41", "192.168.1.41"),
    ("IP-96", "192.168.1.96"),
    ("IP-98", "192.168.1.98"),
    ("IP-36", "192.168.2.36"),
]

# number of pings to send per host
PING_COUNT = 4

def get_local_ipv4():
    """Return the primary IPv4 address of this machine (best-effort)."""
    try:
        # use UDP socket to avoid sending packets; connect to public IP to discover default interface
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            # doesn't actually send packets
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"

def run_ping(ip, count=PING_COUNT, timeout_seconds=5):
    """Run system ping and return (returncode, stdout_text)."""
    system = platform.system().lower()
    if system == "windows":
        cmd = ["ping", "-n", str(count), "-w", str(timeout_seconds * 1000), ip]
    else:
        # Linux / macOS
        # -c count, -W timeout (Linux seconds) / -t not portable for mac
        # Use -c only; ping's native timeout varies by platform. This is best-effort.
        cmd = ["ping", "-c", str(count), ip]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=(count * timeout_seconds) + 5)
        return proc.returncode, proc.stdout + proc.stderr
    except subprocess.TimeoutExpired as e:
        return 124, (str(e) or "")  # timeout code-like


def parse_ping_output(output):
    """
    Parse ping output to extract:
      - avg_rtt_ms (float or None)
      - packet_loss_percent (float or None)
    Works for common Windows and Unix formats (best-effort).
    """
    avg = None
    loss = None

    # Try packet loss patterns
    # Unix: "X% packet loss"
    m = re.search(r"(\d+(?:\.\d+)?)%\s*packet loss", output, re.IGNORECASE)
    if m:
        try:
            loss = float(m.group(1))
        except:
            loss = None
    else:
        # Windows: "Lost = X (Y% loss)" or "Lost = X (Y% loss)"
        m2 = re.search(r"Lost = \d+ \((\d+)% loss\)", output, re.IGNORECASE)
        if m2:
            try:
                loss = float(m2.group(1))
            except:
                loss = None
        else:
            # Another Windows variation: "lost = X (Y% loss)"
            m3 = re.search(r"lost = \d+ \((\d+)% loss\)", output, re.IGNORECASE)
            if m3:
                try:
                    loss = float(m3.group(1))
                except:
                    loss = None

    # Try average RTT patterns
    # Unix (Linux): "rtt min/avg/max/mdev = 0.123/0.234/0.345/0.045 ms"
    m = re.search(r"rtt [^=]*=\s*[\d\.]+/([\d\.]+)/[\d\.]+/[\d\.]+\s*ms", output, re.IGNORECASE)
    if m:
        try:
            avg = float(m.group(1))
        except:
            avg = None
    else:
        # macOS: "round-trip min/avg/max/stddev = 12.345/13.456/14.567/0.456 ms"
        m2 = re.search(r"round-trip [^=]*=\s*[\d\.]+/([\d\.]+)/[\d\.]+/[\d\.]+\s*ms", output, re.IGNORECASE)
        if m2:
            try:
                avg = float(m2.group(1))
            except:
                avg = None
        else:
            # Windows: "Minimum = 1ms, Maximum = 4ms, Average = 2ms"
            m3 = re.search(r"Average\s*=\s*(\d+)\s*ms", output, re.IGNORECASE)
            if m3:
                try:
                    avg = float(m3.group(1))
                except:
                    avg = None
            else:
                # Some locales or outputs show avg as "avg = 1.23 ms"
                m4 = re.search(r"avg(?:/| = )\s*([\d\.]+)\s*ms", output, re.IGNORECASE)
                if m4:
                    try:
                        avg = float(m4.group(1))
                    except:
                        avg = None

    return avg, loss


def ping_target(name, ip, src_ip):
    """Ping a single target and return a dict with results."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    retcode, output = run_ping(ip)
    avg_rtt, packet_loss = parse_ping_output(output)

    # Determine reachability: prefer packet_loss when available, otherwise retcode==0
    reachable = None
    if packet_loss is not None:
        reachable = (packet_loss < 100.0)
    else:
        # retcode 0 typically indicates success
        reachable = (retcode == 0)

    row = {
        "Name": name,
        "IP": ip,
        "Source IP": src_ip,
        "Timestamp": ts,
        "Reachable": bool(reachable),
        "Avg RTT (ms)": avg_rtt if avg_rtt is not None else "",
        "Packet Loss (%)": packet_loss if packet_loss is not None else "",
        "Raw Ping Return Code": retcode,
        "Raw Output (short)": ("\n".join(output.splitlines()[:6]) + ("..." if len(output.splitlines())>6 else "")),
    }
    return row


def main():
    print("Ping script starting...")
    print("Targets:")
    for n, ip in targets:
        print(f"  - {n}: {ip}")

    src_ip = get_local_ipv4()
    print(f"\nDetected local (source) IP: {src_ip}\n")

    results = []
    # Use a thread pool to ping concurrently (fast)
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(targets))) as ex:
        futures = [ex.submit(ping_target, name, ip, src_ip) for name, ip in targets]
        for fut in concurrent.futures.as_completed(futures):
            try:
                row = fut.result()
            except Exception as e:
                # In case a single ping job fails unexpectedly
                row = {
                    "Name": "ERROR",
                    "IP": "",
                    "Source IP": src_ip,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Reachable": False,
                    "Avg RTT (ms)": "",
                    "Packet Loss (%)": "",
                    "Raw Ping Return Code": -1,
                    "Raw Output (short)": f"Exception: {e}",
                }
            results.append(row)
            # Print row summary to console
            print(f"[{row['Timestamp']}] {row['Name']} ({row['IP']}) -> Reachable: {row['Reachable']}, Avg RTT: {row['Avg RTT (ms)']}, Loss: {row['Packet Loss (%)']}")

    # Sort results by name for nicer Excel output
    results_sorted = sorted(results, key=lambda r: r.get("Name", ""))

    # Save to Excel
    df = pd.DataFrame(results_sorted)
    filename = f"ping_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    try:
        df.to_excel(filename, index=False)
        print(f"\nResults written to Excel file: {filename}")
    except Exception as e:
        print(f"\nFailed to write Excel file: {e}")
        print("You may need to install pandas and openpyxl: pip install pandas openpyxl")

    print("\nDetailed results (first few columns):")
    # print a simple table-like output
    header = ["Name", "IP", "Reachable", "Avg RTT (ms)", "Packet Loss (%)"]
    print("{:<12} {:<16} {:<9} {:<13} {}".format(*header))
    for r in results_sorted:
        print("{:<12} {:<16} {:<9} {:<13} {}".format(
            r.get("Name", "")[:12],
            r.get("IP", "")[:16],
            str(r.get("Reachable", "")),
            str(r.get("Avg RTT (ms)", ""))[:13],
            str(r.get("Packet Loss (%)", ""))
        ))

    input("\nPing complete. Press Enter (or any key + Enter) to exit...")

if __name__ == "__main__":
    main()
