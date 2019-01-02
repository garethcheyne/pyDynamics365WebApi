### pyDynamicsWebApi

A Python Module that utilises `Dynamics365 WebApi` and connects via `Azure App Token`.

Basic design around `Microsoft Project` [MS Developer Documentation](https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi)

1). Simple syntax of using the package.
> ```python WebApi.py -H``` Show the package argument help.

> ```python WebApi.py -V``` Show the current installed version of this package.

> ```python WebApi.py -I <xxx.dynamcis.com>``` Change the Dynamics365 instance of what is listed in the Config.yaml file.

> ```python WebApi.py -T``` Test Connection to your Dynamics365 instance.

> ```python WebApi.py -C <file location>``` Change the default location of the Config.yaml file.



2). You will need to create a yaml config file that includes all your creds to connect to your instance of Dynamics CRM. The default file name should be called xrm_config.yaml

**`Example xrm_config.yaml file.`**

 >Download Sample File https://github.com/garethcheyne/pyDynamics365WebApi/blob/master/sample_config.yaml
```
  Resource URI, ie you CRM Instance  
  RESOURCE_URI: https://???.crm6.dynamics.com  
  API_VERSION: 9.0  

  Username and Password to CRM Instance  
  XRM_USERNAME: ??
  XRM_PASSWORD: ??

  Azure Tenant Auth Url  
  TENANT_AUTHORIZATION_URL: ??  

  Azure ClientID and Secret for Dynamics Api  
  XRM_CLIENTID: ??  
  XRM_CLIENTSECRET: ??
```

**`Basic use as follows`**
```
from pyDynamicsWebApi import WebApi

WebApi = WebApi(' _location for your config file_ '})
WebApi.__connection_test__()

response = WebApi.RetrieveMultipleRecords('accounts', options="?$filter ???")

```
#### Things To Do:
- [x] DeleteRecord
- [ ] isAvaiableOffline
- [ ] execute
- [ ] executeMultiple
- [ ] Generic other stuff
