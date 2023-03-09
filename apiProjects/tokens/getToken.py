import requests


def userLogin():
    USERLOGINURL = 'http://localhost:5000/api/users/login'
    data = {
        "email": "some@email.ru",
        "password": "!qw2Er4Ty6"
    }
    res = requests.post(USERLOGINURL, data)
    return res.json()['token']

a = userLogin()
print(a)

