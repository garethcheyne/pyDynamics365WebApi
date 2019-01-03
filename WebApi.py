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


def getToken(config_location, instance='prod'):
    """
    Connect to the Azure Authorization URL and get a token to us the WebApi Calls.b
    """
    with open(config_location, 'r') as ymlfile:
        try:
            cfg = yaml.load(ymlfile)
            if instance == 'sandbox':
                RESOURCE_URI = str(cfg['INSTANCE']['SANDBOX'])
            else:
                RESOURCE_URI = str(cfg['INSTANCE']['PRODUCTION'])

            API_VERSION = str(cfg['INSTANCE']['API_VERSION'])
            XRM_USERNAME = cfg['DYNAMICS_CREDS']['USERNAME']
            XRM_PASSWORD = cfg['DYNAMICS_CREDS']['PASSWORD']
            TENANT_AUTHORIZATION_URL = cfg['AZURE']['AUTHORIZATION_URL']
            XRM_CLIENTID = cfg['APP']['CLIENTID']
            XRM_CLIENTSECRET = cfg['APP']['CLIENTSECRET']
        except yaml.YAMLError as err:
            print('')
            print(err)

    def CheckTokenExpire(secs):
        """
        Sets the DateTime of when the Token Expires using the stand python datetime format.
        """
        now = datetime.now()
        expire = now + timedelta(0, int(secs))
        return expire

    if API_TOKEN['expire_on'] is None or API_TOKEN['expire_on'] < datetime.now():
        data = {
            'client_id': XRM_CLIENTID,
            'client_secret': XRM_CLIENTSECRET,
            'resource': RESOURCE_URI,
            'username': XRM_USERNAME,
            'password': XRM_PASSWORD,
            'grant_type': 'password'
        }
        token_responce = requests.post(TENANT_AUTHORIZATION_URL, data=data)
        if token_responce.status_code is 200:
            API_TOKEN['token'] = token_responce.json()['access_token']
            API_TOKEN['expire_on'] = CheckTokenExpire(token_responce.json()['expires_in'])
            print('pyDynamics365WebApi :: New Token Granted')
            return RESOURCE_URI, API_VERSION, (API_TOKEN['token'])
        else:
            print(':( Sorry you have a connection error, please review your pyXRM config file.')
            print('=== Stack Trace - Start ===')
            print(token_responce.json()['error_description'])
            print('=== Stack Trace - End ===')
            print('Exiting Script Now...')
            exit()
    else:
        print('pyDynamics365WebApi :: Old Token')
        return RESOURCE_URI, API_VERSION, (API_TOKEN['token'])


class WebApi(object):
    """
    List of all the standard Web Api called based on the standardised calls listed on MS Dynamics Web Api Dev site
    https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi
    """

    def __init__(self, config_file_location=config_file):
        self._resource_uri, self._api_version, self._token = getToken(config_file_location)
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
            options = {'CreateRecord, Create a new Dynamics Record.',
                       'DeleteRecord, Delete a Dynamics Record.',
                       'UpdateRecord, Update a Dynamics Record.',
                       'UpsertRecord, Update or Create a Dynamics Record if does not exist.',
                       'RetrieveRecord, Retrieve Dynamics Record with GUID or Alternative key.',
                       'RetrieveMultipleRecords, Retrieve Multiple Dynamics Records with Query',
                       'IsAvailableOffline'
                       'Execute, Execute a Dynamics Workflow with GUID of Workflow',
                       'ExecuteMultiple'
                       }
            print('Error :: No valid option selected.\n')
            print('Options are as follows: (Not case sensitive)')
            for option in options:
                print('>> %s' % option)

    def __connection_test__(self):
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

    def RetrieveRecord(self, entityLogicalName=None, id=None, options=None, user=None):
        """
        Retrieve a single record from Dynamics CRM, you must supply that records GUID
        :param entityLogicalName:
        :param id:
        :param options:
        :param user:
        :return:
        """

        if user is not None:
            self._headers.update({'MSCRMCallerID': user})

        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/' + entityLogicalName + '/' + id + '?' + options, headers=self._headers).json()
        return response


    def RetrieveMultipleRecords(self, entityLogicalName=None, options=None, maxPageSize=None, user=None):

        if user is not None:
            self._headers.update({'MSCRMCallerID': user})

        if maxPageSize is not None:
            self._headers.update({'Prefer': 'odata.maxpagesize=' + str(maxPageSize)})

        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/' + entityLogicalName + options, headers=self._headers).json()
        next_response = response
        while True:
            if '@odata.nextLink' in response:
                next_link = response['@odata.nextLink']
                response = requests.get(next_link, headers=self._headers).json()
                next_response['value'].extend(response['value'])
            if 'error' in response:
                print('pyDynamics365WebApi :: RetrieveMultipleRecords Failed\n')
                print(response)
                return None
            else:
                return next_response['value']

    def CreateRecord(self, entityLogicalName=None, data=None, user=None):

        if user is not None:
            self._headers.update({'MSCRMCallerID': user})

        data = json.dumps(data)

        headers = {
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json;odata=verbose',
            'Authorization': 'Bearer ' + self._token,
            'Content-Type': 'application/json; charset=utf-8',
            'Prefer': 'return=representation',
            'MSCRMCallerID': user,
        }
        response = requests.post(self._resource_uri + '/api/data/v' + self._api_version + '/' + entityLogicalName, data=data, headers=headers).json()
        if 'error' in response:
            print('pyDynamics365WebApi :: Create Record Failed\n')
            print(response)
            return None
        else:
            return response

    def Upsert(self, entityLogicalName=None, AlternateKey=None, data=None, user=None):
        """
        Update or Create a Dynamics Entity Record
        :param entityLogicalName:
        :param altkey:
        :param data:
        :param user:
        :return:
        """
        data = json.dumps(data)

        return


    def updateRecord(self, entity=None, guid=None, data=None, user=None):
        """
        Update a Dynamics Entity Record
        :param entity: Required, A Dynamics entity logical name.
        :param guid: Required, The record id.
        :param data: Required, A list of fields and the values you want updated.
        :param user: Optional, A Dynamics user id you may want to masquerade as.
        :return: Dynamics365Response
        """

        data = json.dumps(data)

        headers = self._headers

        if user is not None:
            headers.update({'MSCRMCallerID': user})

        response = requests.patch(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '(' + guid + ')', data=data, headers=headers).json()

        if 'error' in response:
            print('pyDynamics365WebApi :: Update Record Failed')
            print(response)
            return None
        else:
            return response

    def deleteRecord(self, entity=None, guid=None):
        headers = self._headers
        response = requests.delete(self._resource_uri + '/api/data/v' + self._api_version + '/' + entity + '(' + guid + ')', headers=headers)

        if response.status_code is 204:
            print('ServerResponse :: Delete Record Successful')
            return
        else:
            if 'error' in response.json():
                print('pyDynamics365WebApi :: Delete Record Failed')
                print('ServerResponse :: ' + response.json()['error']['message'])
                return

    def isAvailableOffline(self):
        pass

    def execute(self):
        pass

    def executeMultiple(self):
        pass

    def StatusCode(code):
        if code is 200:
            return str(code) + ' :: Success: Result Found'

        elif code is 201:
            return str(code) + ' :: Success: Record Created/Updated.'

        elif code is '500':
            return str(code) + ' :: Failed: Internal Server Error.'

        else:
            print(code)
            return

    @staticmethod
    def ConvertToDictWithIndex(index_key, webapi_response):
        """
        Converts the response from Dynamics to a Dictionary where you can control what field is used as the index key.
        :param index_key: What field would you like as the Dictionary Key?
        :param webapi_response: The JSON formatted response from Dynamics.
        :return: A Dictionary Object with your desired key.
        """

        d = {}

        if 'value' in webapi_response:
            webapi_response = webapi_response['value']

        for entry in webapi_response:
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
        print('pyDynamics365WebApi :: Vesion 0.1.0.0 alpha')
        #print(setup.version)

    if args.test:
        print('pyDynamics365WebApi :: Running Connections Test, (WhoAmI) \n')
        WebApi(config_file_location=config_file).__connection_test__()

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
                WebApi.deleteRecord(entity=args.entity, guid=args.query)

            elif args.execute.lower() == 'retrievemultiplerecords':
                response = WebApi.RetrieveMultipleRecords(entityLogicalName=args.entity.lower(),
                                                          options=args.query.lower())
                print(response)

        elif args.execute.lower() == 'options':
            WebApi.__cli__(args='options')

        else:
            WebApi.__cli__(args='options')
