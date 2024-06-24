import argparse
from argparse import ArgumentParser
import logging
import requests
import os, sys
from pprint import pprint
import urllib3
urllib3.disable_warnings()

def check_system_variables(variables):
    missing_variables = []
    
    for variable in variables:
        if os.getenv(variable) is None:
            missing_variables.append(variable)
    
    return missing_variables

def set_variables_from_env_and_cli():
    run_options: ArgumentParser = argparse.ArgumentParser()
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

if __name__ == '__main__':

    logging.basicConfig(level=logging.DEBUG, format='[%(asctime)s - %(levelname)s] %(message)s')

    username, password, device_name_filter_string = set_variables_from_env_and_cli()    

    hostname = "172.26.123.10"

    login_url   = f"https://{hostname}/api/system/v1/auth/login"
    devices_url = f"https://{hostname}/api/v2/data/customer-facing-service/DeviceInfo?id=d8947aff-64a3-43cf-8c30-9be2b6f3f035"

    session = requests.session()

    session.auth   = username, password
    session.verify = False

    login_results = session.get(login_url)

    session.headers.update({'cookie':login_results.headers['Set-Cookie']})
    session.headers.update({'content-type':'application/json'})

    device_results = session.get(devices_url)

    devices = device_results.json()

    pprint(devices)