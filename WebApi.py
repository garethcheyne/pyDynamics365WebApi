import os
import platform
import argparse
import requests
import json
import yaml
from datetime import datetime, timedelta

# DEFAULTS
config_file = 'config.yaml'

# Resource URI, ie you CRM Instance
RESOURCE_URI = ''
API_VERSION = ''

# Azure Token Created by Script
API_TOKEN = {'token': None, 'expire_on': None}


class WebApiException(Exception):
    pass


def get_token(config_location, instance='prod'):
    """
    Connect to the Azure Authorization URL and get a token to us the WebApi Calls.b
    """
    with open(config_location, 'r') as ymlfile:
        try:
            cfg = yaml.load(ymlfile)
            if instance == 'sandbox':
                resource_uri = str(cfg['INSTANCE']['SANDBOX'])
            else:
                resource_uri = str(cfg['INSTANCE']['PRODUCTION'])

            api_version = str(cfg['INSTANCE']['API_VERSION'])
            username = cfg['DYNAMICS_CREDS']['USERNAME']
            password = cfg['DYNAMICS_CREDS']['PASSWORD']
            authorization_url = cfg['AZURE']['AUTHORIZATION_URL']
            client_id = cfg['APP']['CLIENTID']
            client_secret = cfg['APP']['CLIENTSECRET']

        except yaml.YAMLError as err:
            print('')
            print(err)

    def check_token_expire(secs):
        """
        Sets the DateTime of when the Token Expires using the stand python datetime format.
        """
        now = datetime.now()
        expire = now + timedelta(0, int(secs))
        return expire

    if API_TOKEN['expire_on'] is None or API_TOKEN['expire_on'] < datetime.now():
        data = {
            'client_id': client_id,
            'client_secret': client_secret,
            'resource': resource_uri,
            'username': username,
            'password': password,
            'grant_type': 'password'
        }

        token_response = requests.post(authorization_url, data=data)

        if token_response.status_code is 200:
            API_TOKEN['token'] = token_response.json()['access_token']
            API_TOKEN['expire_on'] = check_token_expire(token_response.json()['expires_in'])
            print('pyDynamics365WebApi :: New Token Granted')
            return resource_uri, api_version, (API_TOKEN['token'])
        else:
            print(':( Sorry you have a connection error, please review your pyXRM config file.')
            print('=== Stack Trace - Start ===')
            print(token_response.json()['error_description'])
            print('=== Stack Trace - End ===')
            print('Exiting Script Now...')
            exit()
    else:
        print('pyDynamics365WebApi :: Old Token')
        return resource_uri, api_version, (API_TOKEN['token'])


class WebApi(object):
    """
    List of all the standard Web Api called based on the standardised calls listed on MS Dynamics Web Api Dev site
    https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi
    """

    def __init__(self, config_file_location=config_file):
        self._resource_uri, self._api_version, self._token = get_token(config_file_location)
        self._user = None
        self._headers = {
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
            'Authorization': 'Bearer ' + self._token,
            'Content-Type': 'application/json; charset=utf-8',
            'Prefer': 'return=representation',
            'MSCRMCallerID': self._user,
        }

    @staticmethod
    def __cli__(args):
        if args == 'options':
            options = {'create_record, Create a new Dynamics Record.',
                       'DeleteRecord, Delete a Dynamics Record.',
                       'UpdateRecord, Update a Dynamics Record.',
                       'UpsertRecord, Update or Create a Dynamics Record if does not exist.',
                       'retrieve_record, Retrieve Dynamics Record with GUID or Alternative key.',
                       'retrieve_multiple_records, Retrieve Multiple Dynamics Records with Query',
                       'IsAvailableOffline'
                       'Execute, Execute a Dynamics Workflow with GUID of Workflow',
                       'ExecuteMultiple'
                       }
            print('Error :: No valid option selected.\n')
            print('Options are as follows: (Not case sensitive)')
            for option in options:
                print('>> %s' % option)

    def get_user_guid(self, full_name=None):
        """
        Query's the Dynamics365 instance and return the guid for the selected user.
        :param full_name: The full name of the Dynamics 365 user.
        :return: The guid of that user if exists.
        """
        response = WebApi.retrieve_multiple_records(self, entity='systemusers', options="?$select=systemuserid&$filter=fullname eq '" + full_name + "'")

        if 'error' in response:
            return
        else:
            guid = response[0]['systemuserid']
            print(guid)
            return guid

    def connection_test(self):
        """
        Basis test that you have configured your yaml file, and your credentials works. Response should be
        OrganizationId, UserId, and BusinessUnitID
        :return: json response
        """
        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/WhoAmI', headers=self._headers)
        if response.status_code is not 200:
            print('pyDynamics365WebApi :: Connection Test Failed \n')
            print(' status code = %s' % response.status_code)
        else:
            for key, value in response.json().items():
                print(key, value)
            return

    def retrieve_record(self, entity, guid, options=None, user_guid=None, user_fullname=None):
        """
        Retrieve a single record from Dynamics CRM, you must supply that records GUID
        :param entity:
        :param guid:
        :param options:
        :param user_guid:
        :return:
        """
        headers = self._headers

        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)

        if user_guid is not None:
            headers.update({'MSCRMCallerID': user_guid})

        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '/' + guid + '?' + options, headers=headers).json()

        return response

    def retrieve_multiple_records(self, entity, options=None, maxPageSize=None, user_guid=None, user_fullname=None):

        headers = self._headers

        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)

        if user_guid is not None:
            headers.update({'MSCRMCallerID': user_guid})

        if maxPageSize is not None:
            headers.update({'Prefer': 'odata.maxpagesize=' + str(maxPageSize)})

        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + options, headers=headers).json()
        next_response = response

        while True:
            if '@odata.nextLink' in response:
                next_link = response['@odata.nextLink']
                response = requests.get(next_link, headers=self._headers).json()
                next_response['value'].extend(response['value'])
            if 'error' in response:
                print('pyDynamics365WebApi :: retrieve_multiple_records Failed\n')
                print(response)
                return None
            else:
                return next_response['value']

    def create_record(self, entity=None, data=None, user_guid=None, user_fullname=None):

        headers = self._headers

        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)

        if user_guid is not None:
            self._headers.update({'MSCRMCallerID': user_guid})

        data = json.dumps(data)
        
        response = requests.post(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity, data=data, headers=headers).json()

        if 'error' in response:
            print('pyDynamics365WebApi :: Create Record Failed\n')
            print(response)
            return None
        else:
            return response

    def upsert_record(self, entity, guid=None, alternate_key=None, data=None, user_guid=None, user_fullname=None):
        """
        Update or Create a Dynamics Entity Record
        :param entity: Required, A Dynamics 365 entity logical name.
        :param guid: Required, The Dynamics 364 record id.
        :param alternate_key:
        :param data:
        :param user_guid:
        :return: Dynamics365 Response
        """

        headers = self._headers
        headers.update({'If-Match': '*'})

        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)

        if user_guid is not None:
            headers.update({'MSCRMCallerID': user_guid})

        data = json.dumps(data)

        response = requests.patch(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '(' + guid + ')', data=data, headers=headers)

        if response.status_code is 204:
            response = response.json()
            return response

        elif 'error' in response:
            print('pyDynamics365WebApi :: Update Record Failed')
            print(response)
            return None

    def update_record(self, entity, guid, data, user_guid=None, user_fullname=None):
        """
        Update a Dynamics Entity Record
        :param entity: Required, A Dynamics 365 entity logical name.
        :param guid: Required, The Dynamics 365 record id.
        :param data: Required, A list of fields and the values you want updated.
        :param user_guid: Optional, A Dynamics 365 user id you may want to masquerade as.
        :param user_fullname, Optional, A Dyanmics 365 fullname as stored in the instance. 
        :return: Dynamics365 Response
        """

        headers = self._headers

        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)
            
        if user_guid is not None:
            headers.update({'MSCRMCallerID': user_guid})

        data = json.dumps(data)

        response = requests.patch(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '(' + guid + ')', data=data, headers=headers).json()

        if 'error' in response:
            print('pyDynamics365WebApi :: Update Record Failed')
            print('ServerResponse :: ' + response.json()['error']['message'])
            return None
        else:
            return response

    def delete_record(self, entity, guid, user_guid=None, user_fullname=None):
        """
        Deletes a record from Dynamics 365
        :param entity: Required, A Dynamics 365 entity logical name.
        :param guid: Required, The Dynamics 365 record id.
        :param user_guid: Optional, A Dynamics 365 user id you may want to masquerade as.
        :param user_fullname, Optional, A Dyanmics 365 fullname as stored in the instance. 
        :return: Null

        """
        headers = self._headers
        
        if user_fullname is not None:
            user_guid = get_user_guid(user_fullname)

        if user_guid is not None:
            headers.update({'MSCRMCallerID': user_guid})

        response = requests.delete(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '(' + guid + ')', headers=headers)

        if response.status_code is 204:
            print('ServerResponse :: Delete Record Successful')
            return
        else:
            if 'error' in response.json():
                print('pyDynamics365WebApi :: Delete Record Failed')
                print('ServerResponse :: ' + response.json()['error']['message'])
                return

    def tools():
        """
        A collection of helpful tools
        :function to_dict: 
        """
        def to_dict(response, index_key):
            """
            Converts the response from Dynamics 365 to a Nested Dictionary where you can control what field is used as an index key.
            :param response: The JSON formatted response from Dynamics 365.
            :param index_key: What field would you like as the Dictionary Key?
            :return: A Nested Dictionary Object with your desired key.
            """

            d = {}

            if 'value' in response: 
                response = response['value']

            for entry in response:
                d[entry[index_key]] = {}
                d[entry[index_key]] = entry

            return d


if __name__ == '__main__':
    # Check OS type and lear the screen accordingly.
    if platform.system() is 'Windows':
        os.system('cls')
    else:
        os.system('clear')

    # Initiate the parser
    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--version", help="Show program version", action="store_true")
    parser.add_argument("-t", "--test", help="Run a basis test on the Config", action="store_true")
    parser.add_argument("-r", "--readme", help="Show the help URL", action="store_true")
    parser.add_argument("-i", "--instance", help="Set the Dynamics CRM Instance")
    parser.add_argument("-c", "--config", help="Set location on the YAML Config file.")
    parser.add_argument("-x", "--execute", help="Execute a WebApi Function, requires to be used with -q, --query and -e, --entity", )
    parser.add_argument("-e", "--entity", help="Entity name for WebApi Function")
    parser.add_argument("-q", "--query", help="Query for WebApi Function")

    # Read arguments from the command line
    args = parser.parse_args()

    # Process the arguments and do something with them..
    if args.instance:
        instance = args.instance

    if args.config:
        config_file = args.config

    if args.version:
        '''
        Prints to screen the Version on the Script
        '''
        print('pyDynamics365WebApi :: Version 0.1.0.0 alpha')

    if args.test:
        '''
        Performs a connection test to the Dynamics 365 instance, and 
        prints to screen its result.
        '''
        print('pyDynamics365WebApi :: Running Connections Test, (WhoAmI) \n')
        WebApi(config_file_location=config_file).connection_test()

    if args.readme:
        print('pyDynamics365WebApi :: Click Link Below \n'
              'https://github.com/garethcheyne/pyDynamics365WebApi/blob/master/README.md')

    if args.execute:
        if args.entity and args.query:
            print('pyDynamics365WebApi :: Execute a WebApi Function - %s' % args.execute.lower())
            WebApi = WebApi()
            if args.execute.lower() == 'createrecord':
                print('create')

            elif args.execute.lower() == 'deleterecord':
                WebApi.delete_record(entity=args.entity, guid=args.query)

            elif args.execute.lower() == 'retrievemultiplerecords':
                response = WebApi.retrieve_multiple_records(entity=args.entity.lower(),
                                                            options=args.query.lower())
                print(response)

        elif args.execute.lower() == 'options':
            WebApi.__cli__(args='options')

        else:
            WebApi.__cli__(args='options')
