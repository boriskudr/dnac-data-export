import threading
import argparse
import json
import logging
import re
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
    def __init__(self, token_url, client_id, client_secret):
        self.token_url = token_url
        self.username = client_id
        self.password = client_secret
        self.token = None
        self.token_expiry = 3600  # Token is valid for 1 hour
        self.token_refresh_time = 3000  # Refresh token 10 minutes before expiry

    def request_token(self):
        try:
            check_api_rate_limit()
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

@sleep_and_retry
@limits(calls=api_rate_limit, period=api_rate_limit_period)
def check_api_rate_limit():
    pass

def critical_api_error(management_ip_address, interface_name, response):
    logging.error(f"Critical API error: {response.status_code}")
    logging.error(f"Switch {management_ip_address}, interface {interface_name}")
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
    run_options.add_argument('-file', default='commands.xlsx', dest='filename', required=False, help="Custom filename, the default is commands.xlsx")
    run_options.add_argument('-device-filter', default='', dest='device_name_filter_string', required=False, help="Limit port query to device names that include the filter string, the default is not set")

    device_name_filter_string = run_options.parse_args().device_name_filter_string
    logging.info(f"Using device name filter string {device_name_filter_string}")

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

    return username, password, device_name_filter_string
    
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

    headers = get_headers()
    check_api_rate_limit()
    response = requests.get(base_url + devices_url, headers=headers, params=params, verify=False)
    devices = []
    # Save the entire device data as a list of dictionaries
    for item in response.json()['response']:
        devices.append(item)

    logging.info(f"Devices data retrieved from DNAC {base_url}")

    return devices

def get_user_port_config(base_url, management_ip_address, interface_name):
    
    sda_port_assignment_url = '/dna/intent/api/v1/business/sda/hostonboarding/user-device'

    port_info = {
        "deviceManagementIpAddress": management_ip_address,
        "interfaceName": interface_name
    }

    logging.info(f"Retrieving data for switch {management_ip_address} {interface_name}")
    
    headers = get_headers()
    check_api_rate_limit()
    response = requests.get(base_url + sda_port_assignment_url, headers=headers, params=port_info, verify=False)
    data = response.json()
	
    if response.status_code == 200:
        # Removing extra keys from received dictionary
        del data['status']
        del data['description']
        del data['executionId']
        del data['scalableGroupName']
        return data
    else:
        logging.debug(f"{data}\n")
        try:
            if "This device is not provisioned to any site" in data['description']:
                logging.info(f"Skipping switch because it is not provisioned to any site - response from ({interface_name})")
                return None
            if "This interface is not assigned to any connected device" in data['description']:
                logging.info(f"Skipping {interface_name} because it is not assigned to any connected device")
                return "Unassigned"
            if "This interface is not assigned to user device" in data['description']:
                # Need to call another function because this port is connected to WAP
                logging.debug(f"Engaging get_wap_port_config function for {interface_name}\n")
                return get_wap_port_config(base_url, management_ip_address, interface_name)
            else:
                critical_api_error(management_ip_address, interface_name, response)
        except KeyError: # Handle API response without "description"
            critical_api_error(management_ip_address, interface_name, response)

def get_wap_port_config(base_url, management_ip_address, interface_name):
    
    sda_wap_port_assignment_url = '/dna/intent/api/v1/business/sda/hostonboarding/access-point'

    port_info = {
        "deviceManagementIpAddress": management_ip_address,
        "interfaceName": interface_name
    }

    logging.info(f"Retrieving data for switch {management_ip_address} {interface_name} (WAP)")
    
    headers = get_headers()
    check_api_rate_limit()
    response = requests.get(base_url + sda_wap_port_assignment_url, headers=headers, params=port_info, verify=False)
    data = response.json()
    
    if response.status_code == 200:
        # Removing extra keys from received dictionary
        del data['status']
        del data['description']
        del data['executionId']
        # Adding this key for compatibility with user ports data format
        data['voiceIpAddressPoolName'] = "" 
        return data
    else:
        logging.debug(f"{data}\n")
        critical_api_error(management_ip_address, interface_name, response)
        
def get_switch_ports_data(base_url, device):
    """
    Retrieves access port information for a switch in the provided data.

    Args:
        device (dict) : Dictionary containing one switch data from DNA Center.
        base_url (str): Base URL of DNA Center.
        managementIpAddress (str): IP address of a switch

    Returns:
        list: List of dictionaries containing retrieved port information.
    """
    port_info_list = []
    platform_id = device['platformId'].split(",")
    for switch_stack_member_id, platform in enumerate(platform_id):
        # Construct interface name pattern based on platform ID
        # Assuming that switch stack members are numbered consequently from 1
        platform_parts = platform.split("-")  # Split platform ID by dash
        platform_ports = platform_parts[1] # Take second part of platform
        match = re.search(r"(\d+)(.*)", platform_ports) # Match first digits in the ports part of the platform
        if match:
            end_port = int(match.group(1))  # Extract and convert captured group (numeric part)
        else:
            # Handle case where no numeric part is found
            logging.error(f"Error parsing platform ID '{platform}' for device {device['managementIpAddress']}")
            sys.exit()
        
        # Loop through port range and call get_port_config for each port
        for port_number in range(1, end_port + 1):
            # Assuming that all switch ports are Gigabit Ethernet
            interface_name = f"GigabitEthernet{switch_stack_member_id+1}/0/{port_number}"
            logging.debug(f"Engaging function get_port_config with parameters {base_url}, {device['managementIpAddress']}, {interface_name}\n")
            port_data = get_user_port_config(base_url, device['managementIpAddress'], interface_name)
            # Checking if switch is not provisioned to site - will get port_data = None from function
            if port_data:
                if port_data != "Unassigned":
                    port_info_list.append(port_data)  # Append retrieved data
                    logging.debug(f"Retrieved data for {interface_name}\n") 
                    logging.debug(f"{port_data}\n")
            else:
                break

    return port_info_list

def get_and_write_all_switches_ports_data_to_workbook(base_url, switches_data, device_name_filter_string, workbook):
    """
    Retrieves access port information for switches in the provided data and adds each switch's port info into a separate sheet in workbook.

    Args:
        switches_data (dict): Dictionary containing all switches data from DNA Center.
        base_url (str): Base URL of DNA Center.
        workbook (workbook): workbook object

    Returns:
        workbook: Updated workbook object with each switch's port info in a separate sheet. Only switches with assigned ports are included
    """

    ports_data = []

    for device in switches_data:
        if device['role'] == "ACCESS" and device_name_filter_string in device['hostname']:  # Check if device role is ACCESS and hostname matches device name filter
            
            device_hostname_short = device['hostname'].split(".")[0]
            
            logging.debug(f"Engaging function get_switch_ports_data with parameters {base_url}, {device}\n")
            
            logging.info(f"Retrieving ports data from access switch {device_hostname_short}")
            
            ports_data = get_switch_ports_data(base_url, device)
            
            logging.debug(f"Retieved ports data: {ports_data}\n")

            if ports_data:
                logging.info(f"Retrieved ports data for {device_hostname_short}")
                workbook = write_data_to_sheet_in_workbook(ports_data, device_hostname_short, workbook)

    return workbook

def write_data_to_sheet_in_workbook(data, sheet_name, workbook):
  
    sheet = workbook.create_sheet(title=sheet_name)

    # Write table headings
    headings = list(data[0].keys())    # Get headings from the first device data
    sheet.append(headings)

    # Write data
    for item in data:
        row = []
        for heading in headings:
            row.append(item[heading])
        sheet.append(row)
        
    logging.info(f"Data for {sheet_name} added to workbook")
    
    return workbook

def save_results(workbook, filename):
  
    try:
        workbook.save(filename)
        logging.info(f"Data written successfully to '{filename}'")
    except Exception as e:
        logging.error(f"Error writing to Excel file: {e}")
        sys.exit()
    
if __name__ == '__main__':

    logging.basicConfig(level=logging.INFO, format='[%(asctime)s - %(levelname)s] %(message)s')
    logging.info(f"Setting API rate limit: {api_rate_limit} calls in {api_rate_limit_period} seconds")
    
    timestamp = time.strftime('%Y-%m-%d_%H-%M')

    dnac_base_url      = 'https://172.26.123.10'
    output_filename    = f"dnac-report-{timestamp}.xlsx"
    
    username, password, device_name_filter_string = set_variables_from_env_and_cli()    

    wb = Workbook()
    wb.remove(wb.active)

    auth_url = '/dna/system/api/v1/auth/token'
    token_manager = TokenManager(dnac_base_url + auth_url, username, password) # Creating a class instance for the token manager
    token_manager.start() # Starting a thread for automatic token refresh every 50 minutes
    
    devices       = get_devices(dnac_base_url)
    switches_data = [device for device in devices if device["family"] == "Switches and Hubs"] # Extract Switches data from devices
    wap_data      = [device for device in devices if device["family"] == "Unified AP"] # Extract WAPs data from devices
    
    wb = write_data_to_sheet_in_workbook(switches_data, "Switches", wb) # Add Switches data to workbook object
    wb = write_data_to_sheet_in_workbook(wap_data, "WAPs", wb) # Add WAPs data to workbook object
    
    wb = get_and_write_all_switches_ports_data_to_workbook(dnac_base_url, switches_data, device_name_filter_string, wb) # Add Ports data to workbook objects (stack per sheet)

    save_results(wb, output_filename)

    logging.info(f"Script execution completed successfully")