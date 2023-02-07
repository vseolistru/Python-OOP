import requests
from utils.ValidateUser import ValidateUser



"""methods for testing ecommerce api"""

ACTIVATEURL = 'http://localhost:5000/api/users/activate/'
CREATEUSERURL = 'http://localhost:5000/api/users'
USERLOGINURL = 'http://localhost:5000/api/users/login'
data = {
    "email": "some@email.ru",
    "password": "!qw2Er4Ty6"
}

class Api_ski_store:


    @staticmethod
    def create_new_user_and_validate():
        """create and validate user"""
        data = {
            "name": "some_seo",
            "email": "some@email.ru",
            "password": "!qw2Er4Ty6"
        }
        res = requests.post(CREATEUSERURL, data)
        print( f"Satus code {res.status_code}")
        """call Validation User JSON"""
        try:
            assert res.status_code == 200
            ValidateUser.checkResponseCreatedUser(res.json(), data)
        except:
            assert res.status_code == 403
            ValidateUser.checkIfUserExist(res.json())

    @staticmethod
    def login():
        res = requests.post(USERLOGINURL, data)
        return res
    @staticmethod
    def check_user_login():
        """check user is login"""
        res = Api_ski_store.login()
        assert res.status_code == 200
        """call validate User Login"""
        ValidateUser.checkResponseLoginUser(res.json(), data)

    @staticmethod
    def check_user_activate():
        """check email activation positive"""
        userActivationLink = Api_ski_store.login()
        res = requests.get(f"{ACTIVATEURL}{ userActivationLink.json()['activationLink']}")
        assert res.status_code == 200
        ValidateUser.checkActivationLink(res.json())

        """check email activation negative"""
        res = requests.get(f"{ACTIVATEURL}{ userActivationLink.json()['activationLink']}+1")
        assert res.status_code == 500
        ValidateUser.checkNegativeActivationLink(res.json())


