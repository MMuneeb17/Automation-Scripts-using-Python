import xml.etree.ElementTree as ET
import os
import uuid

# Define SAP base directory in AppData
sap_appdata_path = os.path.expandvars(r"%APPDATA%\SAP")

# Define all necessary SAP directories
sap_folders = [
    sap_appdata_path,
    os.path.join(sap_appdata_path, "Common"),
    os.path.join(sap_appdata_path, "SAP GUI"),
    os.path.join(sap_appdata_path, "SAP GUI\\ABAP"),
    os.path.join(sap_appdata_path, "SAP GUI\\Themes"),
    os.path.join(sap_appdata_path, "SAP GUI\\Traces"),
    os.path.join(sap_appdata_path, "SAP GUI\\Scripts"),
    os.path.join(sap_appdata_path, "SAP GUI\\Security"),
    os.path.join(sap_appdata_path, "SAP GUI\\Logs"),
]

# Create all required SAP folders
for folder in sap_folders:
    os.makedirs(folder, exist_ok=True)

# Define SAP landscape file paths
sap_common_path = os.path.join(sap_appdata_path, "Common")
sap_landscape_path = os.path.join(sap_common_path, "SAPUILandscape.xml")
sap_global_landscape_path = os.path.join(sap_common_path, "SAPUILandscapeGlobal.xml")

# Function to create an empty SAPUILandscape XML if it does not exist
def create_sap_landscape_file(file_path):
    if not os.path.exists(file_path):
        root = ET.Element("Landscape", version="1", generator="SAP GUI for Windows")
        ET.SubElement(root, "Includes")
        ET.SubElement(root, "Workspaces")
        ET.SubElement(root, "Services")

        tree = ET.ElementTree(root)
        tree.write(file_path, encoding="utf-8", xml_declaration=True)
        print(f"Created new SAP landscape file: {file_path}")

# Ensure SAPUILandscape.xml exists
create_sap_landscape_file(sap_landscape_path)

# Try to load the SAPUILandscape.xml file
try:
    tree = ET.parse(sap_landscape_path)
    root = tree.getroot()
except ET.ParseError:
    print("Error: Unable to parse SAPUILandscape.xml. Creating a new one.")
    create_sap_landscape_file(sap_landscape_path)
    tree = ET.parse(sap_landscape_path)
    root = tree.getroot()

# Ensure the <Workspaces> section exists
workspaces = root.find("Workspaces")
if workspaces is None:
    workspaces = ET.SubElement(root, "Workspaces")

# Check if 'Local' workspace exists, if not, create it
local_workspace = next((ws for ws in workspaces.findall("Workspace") if ws.get("name") == "Local"), None)

if local_workspace is None:
    local_workspace = ET.SubElement(workspaces, "Workspace", name="Local", uuid=str(uuid.uuid4()))
    print("Created 'Local' workspace.")

# Ensure the <Services> section exists
services = root.find("Services")
if services is None:
    services = ET.SubElement(root, "Services")

# Define new SAP system details (without router string)
new_system = {
    "name": "FCCL SAP PRODUCTION",
    "sid": "FCP",
    "app_server": "10.10.254.202",
    "instance": "00",
    "port": f"32{int('00'):02d}",  # Ensure correct port format (e.g., 3200 for instance 00)
    "connection_type": "Custom Application Server",
}

# Check if the system already exists
if any(service.get("name") == new_system["name"] for service in services.findall("Service")):
    print(f"System '{new_system['name']}' already exists in SAP Logon.")
else:
    # Generate UUID for the new service
    new_service_uuid = str(uuid.uuid4())

    # Create new service entry **without router information**
    new_service = ET.Element("Service", type="SAPGUI", uuid=new_service_uuid,
                              name=new_system["name"], systemid=new_system["sid"],
                              mode="1", server=f"{new_system['app_server']}:{new_system['port']}",
                              sncop="-1", dcpg="2", description=new_system["name"],
                              server_type=new_system["connection_type"])

    # Add the new service to the <Services> section
    services.append(new_service)

    # Link the new service in the 'Local' workspace
    ET.SubElement(local_workspace, "Item", uuid=str(uuid.uuid4()), serviceid=new_service_uuid)

    # Save the updated SAPUILandscape.xml file
    tree.write(sap_landscape_path, encoding="utf-8", xml_declaration=True)

    print(f"New SAP connection '{new_system['name']}' added successfully!")

# Force SAP Logon to refresh
bak_file = f"{sap_landscape_path}.bak"
if os.path.exists(bak_file):
    os.remove(bak_file)
