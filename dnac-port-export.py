# Release 1.0 - ports info queried one-by-one

import threading
import json
import argparse
import logging
from openpyxl import Workbook
import os, sys
from argparse import ArgumentParser
import time
import requests
from requests.auth import HTTPBasicAuth
import urllib3
urllib3.disable_warnings()

from ratelimit import limits, sleep_and_retry
api_rate_limit = 999           # Calls
api_rate_limit_period = 60*100 # Seconds

class TokenManager:
    def __init__(self, dnac_base_url, client_id, client_secret):
        self.token_url = dnac_base_url + '/dna/system/api/v1/auth/token'
        self.username  = client_id
        self.password  = client_secret
        self.token     = None
        self.token_expiry = 3600  # Token is valid for 1 hour
        self.token_refresh_time = 3000  # Refresh token 10 minutes before expiry

    def request_token(self):
        try:
            response = requests.post(self.token_url, auth=HTTPBasicAuth(username, password), verify=False)
            self.token = response.json()['Token']
            logging.info(f"Token received from DNAC")
            return self.token
        except requests.exceptions.RequestException as e:
            logging.error(f"Failed to request token: {e}")

    def refresh_token(self):
        while True:
            time.sleep(self.token_refresh_time)
            self.request_token()

    def get_token(self):
        return self.token

    def start(self):
        self.request_token()
        refresh_thread = threading.Thread(target=self.refresh_token)
        refresh_thread.daemon = True
        refresh_thread.start()

def create_user_api_session(dnac_base_url, username, password):
    
    login_url   = dnac_base_url + "/api/system/v1/auth/login"
    
    session = requests.session()

    session.auth   = username, password
    session.verify = False

    login_results = session.get(login_url)

    session.headers.update({'cookie':login_results.headers['Set-Cookie']})
    session.headers.update({'content-type':'application/json'})

    logging.info(f"Created User GUI API session with DNAC {dnac_base_url}")

    return session

@sleep_and_retry
@limits(calls=api_rate_limit, period=api_rate_limit_period)
def check_api_rate_limit():
    pass

def critical_api_error(device_hostname, device_id, response):
    logging.error(f"Critical API error: {response.status_code}")
    logging.error(f"Switch: {device_hostname}, ID: {device_id}")
    logging.error(f"{response._content}")
    sys.exit()

def check_system_variables(variables):
    missing_variables = []
    
    for variable in variables:
        if os.getenv(variable) is None:
            missing_variables.append(variable)
    
    return missing_variables

def set_variables_from_env_and_cli():
    run_options: ArgumentParser = argparse.ArgumentParser()
    run_options.add_argument('-device-filter', default='', dest='device_name_filter_string', required=False, help="Limit port query to device names that include the filter string, the default is not set")
    run_options.add_argument('-dump', default=False, dest='dump_dictionary', required=False, help="Set to True to dump device and interface dictionaries into txt files, the default is not set")

    device_name_filter_string = run_options.parse_args().device_name_filter_string
    logging.info(f"Using device name filter string {device_name_filter_string}")
    
    dump_dictionary = run_options.parse_args().dump_dictionary
    logging.info(f"Dumping device and interface dictionaries into txt files")
    
    device_credential_vars = ['DNAC_SCRIPT_USERNAME', 'DNAC_SCRIPT_PASSWORD']
    
    # Check if system variables exist
    missing_vars = check_system_variables(device_credential_vars)

    if len(missing_vars) == 0:
        username = os.getenv(device_credential_vars[0])
        password = os.getenv(device_credential_vars[1])
        logging.info("DNAC credentials successfully imported from OS environment variables")
    else:
        logging.error("The following system variables are missing:")
        for var in missing_vars:
            logging.error(var)
        logging.error("Execution stopped due to missing system variables\nSet them in the PowerShell session using commands:\n$Env:DNAC_SCRIPT_USERNAME = \"username\"\n$Env:DNAC_SCRIPT_PASSWORD = \"password\"")
        sys.exit()

    return username, password, device_name_filter_string, dump_dictionary
    
def get_headers():
    
    token = token_manager.get_token()

    headers = {
        'X-Auth-Token': token,
        'Content-Type': 'application/json'
    }

    return headers

def get_devices(base_url):
    devices_url = '/dna/intent/api/v1/network-device'
    params = {}

    check_api_rate_limit()
    headers = get_headers()
    response = requests.get(base_url + devices_url, headers=headers, params=params, verify=False)
 
    devices = []

    # Extract hostname as the first key in each dictionary
    for item in response.json()['response']:
        device_info = {'hostname': item['hostname']}  # Ensure 'hostname' is first
        device_info.update(item)  # Add remaining keys
        devices.append(device_info)

    logging.info(f"Devices data retrieved from DNAC {base_url}")

    return devices

def get_switch_device_info_data(user_api_session, base_url, device):
    """
    Retrieves access port information for a switch in the provided data.

    Args:
        device (dict) : Dictionary containing one switch data from DNA Center.
        base_url (str): Base URL of DNA Center.
        managementIpAddress (str): IP address of a switch

    Returns:
        list: List of dictionaries containing retrieved port information.
    """
    
    device_info_url = base_url + f"/api/v2/data/customer-facing-service/DeviceInfo?networkDeviceId={device['id']}"

    logging.debug(f"Retrieving device data for switch {device['hostname']}")
    
    check_api_rate_limit()
    response = user_api_session.get(device_info_url)
  
    try:
        data = response.json()['response'][0]['deviceInterfaceInfo']
    except Exception as e:
        logging.error(f"'response'][0]['deviceInterfaceInfo'] not found: {e}")
        dump_json_to_file(response.json(), f"{device['hostname']}_get_switch_device_info_data-received.txt")
        sys.exit()
    
    if response.status_code == 200:
        return data
    else:
        critical_api_error(device['hostname'], device['id'], response)

def get_switch_port_info_data(base_url, device):
    """
    Retrieves access port information for a switch in the provided data.

    Args:
        device (dict) : Dictionary containing one switch data from DNA Center.
        base_url (str): Base URL of DNA Center.
        managementIpAddress (str): IP address of a switch

    Returns:
        list: List of dictionaries containing retrieved port information.
    """
    device_port_url = f"/api/v1/interface/network-device/{device['id']}/"

    logging.debug(f"Retrieving port data for switch {device['hostname']}")
    
    check_api_rate_limit()
    headers = get_headers()
    response = requests.get(base_url + device_port_url, headers=headers, verify=False)
    data = response.json()['response']
	
    if response.status_code == 200:
        return data
    else:
        critical_api_error(device['hostname'], device['id'], response)

def merge_device_and_interface_data(switch_device_info_data, switch_port_info_data):
    switch_device_info_data_keys_to_copy = [
        "interfaceName",
        "description",
        "connectedDeviceType",
        "role"
    ]

    switch_device_info_data_nested_key = "authenticationProfile"

    switch_device_info_data_nested_keys_to_copy = [
        "name",
        "deploymentMode",
        "multi_auth",
        "order",
        "priority"
    ]

    switch_device_info_data_nested_keys_new_key_prefix = "dot1x_"

    switch_port_info_data_keys_to_copy = [
        "vlanId",
        "voiceVlan",
        "nativeVlanId",
        "description",
        "adminStatus",
        "status",        
        "speed",
        "duplex",
        "portMode",
        "ipv4Address"
    ]

    # Define speed mapping
    speed_mapping = {
        "10000": "10Mbps",
        "100000": "100Mbps",
        "1000000": "1Gbps",
        "10000000": "10Gbps",
        "40000000": "40Gbps",
        "100000000": "100Gbps",
    }

    data = []

    dict2 = {item['id']: item for item in switch_port_info_data}

    for item1 in switch_device_info_data:
        id1 = item1['interfaceId']

        if id1 in dict2:
            item2 = dict2[id1]

            combined_item = {key: item1[key] for key in switch_device_info_data_keys_to_copy if key in item1}
            combined_item.update({key: item2[key] for key in switch_port_info_data_keys_to_copy if key in item2})

            # Rewrite "speed" according to the mapping
            if 'speed' in combined_item and combined_item['speed'] in speed_mapping:
                combined_item['speed'] = speed_mapping[combined_item['speed']]

            if switch_device_info_data_nested_key in item1:
                nested_dict = item1[switch_device_info_data_nested_key]
                combined_item.update({switch_device_info_data_nested_keys_new_key_prefix + key: nested_dict[key] for key in switch_device_info_data_nested_keys_to_copy if key in nested_dict})

            data.append(combined_item)

    return data

def dump_json_to_file(data, filename):
    with open(filename, 'w') as file:
        json.dump(data, file, indent=4)
        logging.info(f"Saved data dump to {filename}")
        
def get_and_write_all_switches_ports_data_to_workbook(base_url, user_api_session, switches_data, device_name_filter_string, workbook):
    """
    Retrieves access port information for switches in the provided data and adds each switch's port info into a separate sheet in workbook.

    Args:
        switches_data (dict): Dictionary containing all switches data from DNA Center.
        base_url (str): Base URL of DNA Center.
        workbook (workbook): workbook object

    Returns:
        workbook: Updated workbook object with each switch's port info in a separate sheet. Only switches with assigned ports are included
    """

    switch_device_info_data = []
    switch_port_info_data   = []

    for device in switches_data:
        if device['role'] == "ACCESS" and device_name_filter_string in device['hostname']:  # Check if device role is ACCESS and hostname matches device name filter
           
            device_hostname_short = device['hostname'].split(".")[0]

            logging.info(f"Retrieving device info data from access switch {device_hostname_short}")
            logging.debug(f"Engaging function get_switch_device_info_data with parameters {base_url}, {device}\n")
            switch_device_info_data = get_switch_device_info_data(user_api_session, base_url, device)
            logging.debug(f"Retieved device info data")
            if dump_dictionary:
                dump_json_to_file(switch_device_info_data, f"{device['hostname']}_device_info_data.txt")
            
            logging.info(f"Retrieving port info data from access switch {device_hostname_short}")
            logging.debug(f"Engaging function get_switch_port_info_data with parameters {base_url}, {device}\n")
            switch_port_info_data = get_switch_port_info_data(base_url, device)
            logging.debug(f"Retieved device info data")
            if dump_dictionary:
                dump_json_to_file(switch_port_info_data, f"{device['hostname']}_switch_port_info_data.txt")

            if switch_device_info_data and switch_port_info_data:
                logging.debug(f"Merging device and port info for {device_hostname_short}")
                interface_data = merge_device_and_interface_data(switch_device_info_data, switch_port_info_data)
                if dump_dictionary:
                    dump_json_to_file(interface_data, f"{device['hostname']}_interface_data.txt")
                workbook = write_data_to_sheet_in_workbook(interface_data, device_hostname_short, workbook)

    return workbook

def write_data_to_sheet_in_workbook(data, sheet_name, workbook):
  if not data:
    logging.warning("Data is empty. No sheet created.")
    return workbook

  sheet = workbook.create_sheet(title=sheet_name)

  # Find the dictionary with maximum number of keys
  max_keys_dict = max(data, key=lambda item: len(item.keys()))

  # Use keys from max_keys_dict to create headings
  headings = list(max_keys_dict.keys())

  # Write table headings
  sheet.append(headings)

  # Sorting function for switch interface in value
  def sort_key(item):
    value = item.get(headings[0], "")
    # Split by "/" but keep the original value for full comparison
    parts = value.split("/")

    # If no "/" or single element, return the original value and two zeros
    if len(parts) <= 1:
      return value, 0, 0

    # Convert second and third parts to integers (handle potential errors)
    int_part2 = int(parts[1])
    int_part3 = int(parts[2])
 
    # Sort based on interface, second part (number), then third part (number)
    return (parts[0], int_part2, int_part3)

  # Sort data using the custom sort function
  sorted_data = sorted(data, key=sort_key)

  # Write sorted data
  for item in sorted_data:
    row = []
    for heading in headings:
      row.append(item.get(heading, "N/A"))
    sheet.append(row)

  logging.info(f"Data saved to workbook sheet '{sheet_name}'")

  return workbook

def save_results(workbook, filename):
  
    try:
        workbook.save(filename)
        logging.info(f"Data written successfully to '{filename}'")
    except Exception as e:
        logging.error(f"Error writing to Excel file: {e}")
        sys.exit()
    
if __name__ == '__main__':

    logging.basicConfig(level=logging.DEBUG, format='[%(asctime)s - %(levelname)s] %(message)s')
    logging.info(f"Setting API rate limit: {api_rate_limit} calls in {api_rate_limit_period/60} minutes")
    logging.info(f"The script will pause when the limit is reached")

    dnac_base_url = 'https://172.26.123.10'

    timestamp       = time.strftime('%Y-%m-%d_%H-%M')
    output_filename = f"dnac-report-{timestamp}.xlsx"
    
    username, password, device_name_filter_string, dump_dictionary = set_variables_from_env_and_cli()    

    wb = Workbook()
    wb.remove(wb.active)

    token_manager = TokenManager(dnac_base_url, username, password) # Creating a class instance for the token manager
    token_manager.start() # Starting a thread for automatic token refresh every 50 minutes
    
    user_api_session = create_user_api_session(dnac_base_url, username, password)

    devices       = get_devices(dnac_base_url)
    switches_data = [device for device in devices if device["family"] == "Switches and Hubs"] # Extract Switches data from devices
    wap_data      = [device for device in devices if device["family"] == "Unified AP"] # Extract WAPs data from devices
    
    wb = write_data_to_sheet_in_workbook(switches_data, "Switches", wb) # Add Switches data to workbook object
    wb = write_data_to_sheet_in_workbook(wap_data, "WAPs", wb) # Add WAPs data to workbook object
    
    wb = get_and_write_all_switches_ports_data_to_workbook(dnac_base_url, user_api_session, switches_data, device_name_filter_string, wb) # Add Ports data to workbook objects (stack per sheet)

    save_results(wb, output_filename)

    logging.info(f"Script execution completed successfully")