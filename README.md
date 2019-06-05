### pyDynamicsWebApi - αlpha 

A Python Module that utilises `Dynamics365 WebApi` and connects via `Azure App Token`. The basic design around `Microsoft Project` [MS Developer Documentation](https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/clientapi/reference/xrm-webapi). 

Also what you will need to do is to create a `Azure App Token`, see this blog here from my friends at `Magnetism Solutions` [Blog](https://www.magnetismsolutions.com/blog/johntowgood/2018/03/08/dynamics-365-online-authenticate-with-client-credentials).



#### 1). Simple Use Command Line (CLI)
Simple syntax of using the package.
> ```python WebApi.py -h```, Show the package argument help.

> ```python WebApi.py -v```, Show the current installed version of this package.

> ```python WebApi.py -i <xxx.dynamcis.com>```, Change the Dynamics365 instance of what is listed in the Config.yaml file.

> ```python WebApi.py -t```, Test Connection to your Dynamics365 instance.

> ```python WebApi.py -c <file location>```, Change the default location of the Config.yaml file.

> ```python WebApi.py -x <other arguments>```, Query the Dynamics365 Instance from the CLI.

You are able also to the the CLI interface to perform basic queries.
_Here is an example of Querying Dynamics365 and return you a result of all the Account Name in the Dynamics365 Instance._
> `python WebApi.py -x retrievemultiplerecords -e accounts -q "?$select=name"`

_Or here to should the CLI Options_
> `python WebApi.py -x options`



#### 2). Config File
You will need to create a yaml config file that includes all your creds to connect to your instance of Dynamics CRM. The default file name should be called xrm_config.yaml

**`Example config.yaml file.`**

 >Download [config.yaml](https://github.com/garethcheyne/pyDynamics365WebApi/blob/master/sample_config.yaml)

#### 3). Import and Basic Use
```
from pyDynamicsWebApi import WebApi

WebApi = WebApi(' _location for your config file_ '})
WebApi.__connection_test__()

response = WebApi.retrieve_multiple_records('accounts', options="?$filter ???")

```
#### Things To Do:
- [x] delete_record
- [ ] is_avaiable_offline
- [ ] execute
- [ ] execute_multiple
- [ ] Generic other stuff (Tidy Code.)
