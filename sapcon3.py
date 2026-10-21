import xml.etree.ElementTree as ET
import os
import sys
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
        ET.SubElement(root, "Routers")  # Ensure <Routers> section exists

        tree = ET.ElementTree(root)
        tree.write(file_path, encoding="utf-8", xml_declaration=True)
        print(f"Created new SAP landscape file: {file_path}")

# Ensure both SAPUILandscape.xml and SAPUILandscapeGlobal.xml exist
create_sap_landscape_file(sap_landscape_path)
create_sap_landscape_file(sap_global_landscape_path)

# Load the SAPUILandscape.xml file
tree = ET.parse(sap_landscape_path)
root = tree.getroot()

# Ensure the <Workspaces> section exists
workspaces = root.find("Workspaces")
if workspaces is None:
    workspaces = ET.SubElement(root, "Workspaces")

# Check if 'Local' workspace exists, if not, create it
local_workspace = None
for workspace in workspaces.findall("Workspace"):
    if workspace.get("name") == "Local":
        local_workspace = workspace
        break

if local_workspace is None:
    local_workspace = ET.SubElement(workspaces, "Workspace", name="Local", uuid=str(uuid.uuid4()))
    print("Created 'Local' workspace.")

# Ensure the <Services> section exists
services = root.find("Services")
if services is None:
    services = ET.SubElement(root, "Services")

# Ensure the <Routers> section exists
routers = root.find("Routers")
if routers is None:
    routers = ET.SubElement(root, "Routers")

# Define SAP systems to add (all share the same SAPRouter and settings)
saprouter = "/H/58.65.136.155"  # SAPRouter string

new_systems = [
    {"name": "FCCL SAP PRODUCTION", "sid": "FCP", "app_server": "10.10.254.202", "instance": "00"},
    {"name": "RISE_PROD", "sid": "PS4", "app_server": "10.210.10.243", "instance": "00"},
    {"name": "RISE_QAS", "sid": "QS4", "app_server": "10.210.10.192", "instance": "00"},
]

# Reuse an existing router with the same string, otherwise create one
router_uuid = None
for router in routers.findall("Router"):
    if router.get("router") == saprouter:
        router_uuid = router.get("uuid")
        break

if router_uuid is None:
    router_uuid = str(uuid.uuid4())
    new_router = ET.Element("Router", uuid=router_uuid, name=saprouter,
                             description=saprouter, router=saprouter)
    routers.append(new_router)

existing_names = {service.get("name") for service in services.findall("Service")}

added = 0
for system in new_systems:
    # Skip systems that already exist
    if system["name"] in existing_names:
        print(f"System '{system['name']}' already exists in SAP Logon.")
        continue

    port = f"32{system['instance']}"  # e.g., 3200 for instance 00
    new_service_uuid = str(uuid.uuid4())

    # Create new service entry with all parameters, including router reference
    new_service = ET.Element("Service", type="SAPGUI", uuid=new_service_uuid,
                              name=system["name"], systemid=system["sid"],
                              mode="1", server=f"{system['app_server']}:{port}",
                              sncop="-1", dcpg="2", router=saprouter,
                              description=system["name"], server_type="Custom Application Server",
                              routerid=router_uuid)  # Reference the router ID

    # Explicitly add the SAPRouter string
    router_elem = ET.SubElement(new_service, "Router")
    router_elem.text = saprouter

    # Add the new service to the <Services> section
    services.append(new_service)

    # Link the new service in the 'Local' workspace
    ET.SubElement(local_workspace, "Item", uuid=str(uuid.uuid4()), serviceid=new_service_uuid)
    print(f"Added '{system['name']}'.")
    added += 1

if added == 0:
    print("No new SAP connections to add.")
    sys.exit()

# Save the updated SAPUILandscape.xml file
tree.write(sap_landscape_path, encoding="utf-8", xml_declaration=True)

# Force SAP Logon to refresh
bak_file = f"{sap_landscape_path}.bak"
if os.path.exists(bak_file):
    os.remove(bak_file)

print(f"{added} new SAP connection(s) added successfully")
