import xml.etree.ElementTree as ET
import os

def get_sap_connections_from_xml():
    sap_config_path = os.path.expandvars(r"%APPDATA%\SAP\Common\SAPUILandscape.xml")
    
    try:
        tree = ET.parse(sap_config_path)
        root = tree.getroot()

        systems = [item.get("name") for item in root.findall(".//Service") if item.get("name")]

        if not systems:
            print("No SAP systems found in SAPUILandscape.xml.")
            return []

        print("Available SAP Connections from XML:")
        for i, system in enumerate(systems):
            print(f"{i+1}. {system}")

        return systems

    except Exception as e:
        print(f"Error reading SAP config file: {e}")
        return []

# Run function
sap_connections = get_sap_connections_from_xml()
