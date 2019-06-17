from WebApi import WebApi


def example_00():
    webapi.connection_test()

def example_01():
    guid = webapi.get_user_guid('Gareth Cheyne')
    print(guid)

def example_02():
    data = {'name': 'TEST ACCOUNT'}
    print("Create Account")
    webapi.create_record(entity='accounts', data=data, user_fullname='Gareth Cheyne')

def example_03():
    print("Get all active Accounts")
    response = webapi.retrieve_multiple_records()
    response_dict = WebApi.mro()


if __name__ == '__main__':
    webapi = WebApi()
    example_00()
    example_01()
    #example_02()


    
