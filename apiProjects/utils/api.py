import requests
from utils.ValidateUser import ValidateUser



"""methods for testing ecommerce api"""

ACTIVATEURL = 'http://localhost:5000/api/users/activate/'
CREATEUSERURL = 'http://localhost:5000/api/users'
USERLOGINURL = 'http://localhost:5000/api/users/login'
USERDATAURL = 'http://localhost:5000/api/users/infor/'
ADDCART = 'http://localhost:5000/api/users/addcart/'


class Api_ski_store:
    @staticmethod
    def start_USERS_API_testing():
        print('start USERS API tests')
    @staticmethod
    def negative_create_user_empty_email(dataOne, dataTwo, dataThree):
        """check empty registration fields"""

        """empty name"""
        res = requests.post(CREATEUSERURL, dataOne)
        assert res.status_code == 403
        ValidateUser.checkNegativeUserCreationEmptyFields(res.json())
        """empty email"""
        res = requests.post(CREATEUSERURL, dataTwo)
        assert res.status_code == 403
        ValidateUser.checkNegativeUserCreationEmptyFields(res.json())
        """empty pass"""
        res = requests.post(CREATEUSERURL, dataThree)
        assert res.status_code == 403
        ValidateUser.checkNegativeUserCreationEmptyFields(res.json())

    @staticmethod
    def negative_create_user_bad_email(dataOne, dataTwo):
        """check bas email conditions"""
        """not for a - @ an emailAddress"""
        res = requests.post(CREATEUSERURL, dataOne)
        assert res.status_code == 403
        ValidateUser.checkNegativeUserCreationBadEmail(res.json())
        """not for a domain in emailAddress"""
        res = requests.post(CREATEUSERURL, dataTwo)
        assert res.status_code == 403
        ValidateUser.checkNegativeUserCreationBadEmail(res.json())

    @staticmethod
    def negative_user_pass_length(dataOne, dataTwo):
        """check that pass is not required length"""

        """short password"""
        res = requests.post(CREATEUSERURL, dataOne)
        ValidateUser.checkNegativePassUserLength(res.json())
        """special character"""
        res = requests.post(CREATEUSERURL, dataTwo)
        ValidateUser.checkNegativeWeakPassUser(res.json())

    @staticmethod
    def create_new_user_and_validate(data):
        """create and validate user"""
        res = requests.post(CREATEUSERURL, data)
        # print( f"Satus code {res.status_code}")
        """call Validation User JSON"""
        try:
            assert res.status_code == 200
            ValidateUser.checkResponseCreatedUser(res.json(), data)
        except:
            assert res.status_code == 403
            ValidateUser.checkIfUserExist(res.json())

    @staticmethod
    def login(data):
        res = requests.post(USERLOGINURL, data)
        return res
    @staticmethod
    def check_user_login(data):
        """check user is login"""
        res = Api_ski_store.login(data)
        assert res.status_code == 200
        """call validate User Login"""
        ValidateUser.checkResponseLoginUser(res.json(), data)

    @staticmethod
    def check_user_activate(data):
        """check email activation positive"""
        userActivationLink = Api_ski_store.login(data)
        res = requests.get(f"{ACTIVATEURL}{ userActivationLink.json()['activationLink']}")
        assert res.status_code == 200
        ValidateUser.checkActivationLink(res.json())

        """check email activation negative"""
        res = requests.get(f"{ACTIVATEURL}{ userActivationLink.json()['activationLink']}+1")
        assert res.status_code == 500
        ValidateUser.checkNegativeActivationLink(res.json())

    @staticmethod
    def check_user_data_is_return(user_id , headers):
        """check User Information for localStore Cart items"""
        USERURL = f'{USERDATAURL}{user_id}'
        response = requests.get(USERURL, headers=headers)
        assert response.status_code == 200
        # print(response.json()['cart'][0]['_id'])
        ValidateUser.checkUSERdataisExistforLocalStore(response.json(), user_id)
    @staticmethod
    def check_user_data_is_return_Negative(user_id , headers):
        """check Negative User Information for localStore Cart items"""
        USERURL = f'{USERDATAURL}{user_id}'
        response = requests.get(USERURL, headers=headers)
        ValidateUser.checkNegativeUSERdataisExistforLocalStore(response.json())

    @staticmethod
    def check_addcart_product(user, headers, product):
        """check add to cart for UserID"""
        CARTURL = f"{ADDCART}{user}"
        res = requests.patch(CARTURL, headers=headers, json=product)
        ValidateUser.checkUSERdataisExistforLocalStore(res.json(), user)

