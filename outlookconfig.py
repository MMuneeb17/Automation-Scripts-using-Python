import os
import subprocess
import winreg
import getpass

# Define the path to the MDaemon Connector installer
installer_path = r"\\192.168.1.14\DB_Connection\MDaemonConnectorClient64.exe"


def install_mdaemon_connector():
    """Installs MDaemon Connector silently."""
    print("[+] Installing MDaemon Connector...")
    try:
        subprocess.run([installer_path, "/quiet", "/norestart"], check=True)
        print("[+] Installation successful.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Installation failed: {e}")


def configure_outlook(email, password):
    """Configures Outlook by modifying the registry."""
    print("[+] Configuring Outlook...")

    try:
        outlook_key_path = r"Software\Microsoft\Office\Outlook\OMI Account Manager"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, outlook_key_path) as key:
            winreg.SetValueEx(key, "DefaultProfile", 0, winreg.REG_SZ, email)

        profile_key_path = fr"Software\Microsoft\Office\Outlook\Profiles\{email}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, profile_key_path) as key:
            winreg.SetValueEx(key, "Email", 0, winreg.REG_SZ, email)
            winreg.SetValueEx(key, "Password", 0, winreg.REG_SZ, password)  # Avoid storing passwords in plain text
            winreg.SetValueEx(key, "Server", 0, winreg.REG_SZ, "mail.yourdomain.com")  # Modify this

        print("[+] Outlook configuration complete.")
    except Exception as e:
        print(f"[-] Outlook configuration failed: {e}")


if __name__ == "__main__":
    email = input("Enter your email: ")
    password = getpass.getpass("Enter your password: ")

    install_mdaemon_connector()
    configure_outlook(email, password)

    print("[+] Setup complete. Restart Outlook for changes to take effect.")
