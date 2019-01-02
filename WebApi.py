import requests
import json
import yaml
from datetime import datetime, timedelta


# Resource URI, ie you CRM Instance
RESOURCE_URI = ''
API_VERSION = ''

# Azure Token Created by Script
API_TOKEN = {'token': None, 'expire_on': None}


class WebApiException(Exception):
    pass

def GetToken(config_file_location):
    """
    Connect to the Azure Authorization URL and get a token to us the WebApi Calls.
    """
    with open(config_file_location) as ymlfile:
        cfg = yaml.load(ymlfile)
        RESOURCE_URI = str(cfg['RESOURCE_URI'])
        API_VERSION = str(cfg['API_VERSION'])
        XRM_USERNAME = cfg['XRM_USERNAME']
        XRM_PASSWORD = cfg['XRM_PASSWORD']
        TENANT_AUTHORIZATION_URL = cfg['TENANT_AUTHORIZATION_URL']
        XRM_CLIENTID = cfg['XRM_CLIENTID']
        XRM_CLIENTSECRET = cfg['XRM_CLIENTSECRET']

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
            print('pyXRM :: New Token')
            return RESOURCE_URI, API_VERSION, (API_TOKEN['token'])
        else:
            print(':( Sorry you have a connection error, please review your pyXRM config file.')
            print('=== Stack Trace - Start ===')
            print(token_responce.json()['error_description'])
            print('=== Stack Trace - End ===')
            print('Exiting Script Now...')
            exit()
    else:
        print('pyXRM :: Old Token')
        return RESOURCE_URI, API_VERSION, (API_TOKEN['token'])


class WebApi(object):
    """
    List of all the standard Web Api called based on the standardised calls listed on MS Dynamics Web Api Dev site
    https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi
    """

    def __init__(self, config_file_location='xrm_config.yaml'):
        self._resource_uri, self._api_version, self._token = GetToken(config_file_location)
        self._user = None
        self._headers = {
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Accept': 'application/json',
            'Authorization': 'Bearer ' + self._token,
            'Content-Type': 'application/json; charset=utf-8',
            'MSCRMCallerID': self._user,
        }

    def __connection_test__(self):
        """
        Basis test that you have configured your yaml file, and your credentials works. Response should be OrganizationId, UserId, and BusinessUnitID
        :return: json responce
        """
        response = requests.get(self._resource_uri + '/api/data/v' + self._api_version + '/WhoAmI', headers=self._headers)
        if response.status_code is not 200:
            print('pyXRM :: Connection Test Failed')
            print(response.status_code)
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
                print(response)
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
            print(response)

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


    def UpdateRecord(self, entityLogicalName=None, id=None, data=None, user=None):
        """
        Update a Dynamics Entity Record
        :param entityLogicalName: Required, A Dynamics entity logical name.
        :param id: Required, The record id.
        :param data: Required, A list of fields and the values you want updated.
        :param user: Optional, A Dynamics user id you may want to masquerade as.
        :return: Dynamics Responce
        """

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

        response = requests.patch(self._resource_uri + '/api/data/v' + self._api_version + '/' + entityLogicalName + '(' + id + ')', data=data, headers=headers).json()

        if 'error' in response:
            print(response)

        return response


    def DeleteRecord(self, entityLogicalName=str, id=str, user=str):
        pass

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
    def ConvertToDictWithIndex(index_key, xrm_response):
        """
        Converts the responce from Dynamics to a Dictionary where you can control what field is used as the index key.
        :param index_key: What field would you like as the Dictionary Key?
        :param xrm_response: The JSON formatted responce from Dynamics.
        :return: A Dictionary Object with your desired key.
        """

        d = {}

        if 'value' in xrm_response:
            xrm_response = xrm_response['value']

        for entry in xrm_response:
            d[entry[index_key]] = {}
            d[entry[index_key]] = entry

        return d


if __name__ == '__main__':
    WebApi(config_file_location="../xrm_config.yaml").__connection_test__()