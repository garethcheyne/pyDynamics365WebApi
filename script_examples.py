from WebApi import WebApi


def example_00():
    WebApi.connection_test()

def example_01():
    guid = WebApi.get_user_guid('Gareth Cheyne')
    print(guid)

def example_02():
    data = {'name': 'TEST ACCOUNT'}
    print("Create Account")
    WebApi.create_record(entity='accounts', data=data, user_fullname='Gareth Cheyne')

if __name__ == '__main__':
    WebApi = WebApi()
    example_00()
    example_01()
    example_02()
