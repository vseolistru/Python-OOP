import requests
# from tokens.getToken import userLogin

def userLogin():
    USERLOGINURL = 'http://localhost:5000/api/users/login'
    data = {
        "email": "some@email.ru",
        "password": "!qw2Er4Ty6"
    }
    res = requests.post(USERLOGINURL, data)
    return res.json()['token']


def newUserData():
    data = {
        "name": "some_seo",
        "email": "some@email.ru",
        "password": "!qw2Er4Ty6"
    }
    return data

def userLoginData():
    data = {
        "email": "some@email.ru",
        "password": "!qw2Er4Ty6"
    }
    return data

def userIDLocalStoreData():
    user_id = '63b4329f77327792981df485'
    return user_id
def userIDLocalStoreDataNegative():
    user_id = '63b4329f77327792981df484'
    return user_id

def headers():
    token = userLogin()
    print(token)
    headers = {
        'authorization': f'Bearer {token}'
    }
    return headers

def negativeEmptyFields():
    listData = [
        {
            "name": "",
            "email": "some@email.ru",
            "password": "!qw2Er4Ty6"
        },
        {
            "name": "some_seo",
            "email": "",
            "password": "!qw2Er4Ty6"
        },
        {
            "name": "some_seo",
            "email": "some@email.ru",
            "password": ""
        }
    ]
    return listData

def badEmailData():
    listData = [
        {
            "name": "some_seo",
            "email": "someemail.ru",
            "password": "!qw2Er4Ty6"
        },
        {
            "name": "some_seo",
            "email": "somee@mail",
            "password": "!qw2Er4Ty6"
        }
    ]
    return listData
def incorrectPasswordLenghtData():
    listData = [
        {
            "name": "some_seo",
            "email": "some@email.ru",
            "password": "!qw2Er"
        },
        {
            "name": "some_seo",
            "email": "some@email.ru",
            "password": "qw2er4ty6ui8"
        }
    ]
    return listData

def productData():
    product = {
        "cart":
            {
                "_id": "63b4132177327792981defcf",
                "productId": "6",
                "slug": "rossignol-x-ium-skating-wcs--s2--ifp",
                "title": "rossignol x-ium skating wcs -s2 -ifp",
                "description": " Гоночные беговые лыжи высшего класса для конькового хода от компании Rossignol. Предназначены для профессиональных спортсменов для использования на Кубке Мира и соревнованиях высочайшего уровня. Благодаря конструкции Active cap, комбинирующей два материала с различными свойствами, строению 3D CARBON PROFILE, легчайшему сотовому сердечнику Nomex, усиленному углеродным волокном, укреплённым боковым граням SUPRA EDGE и геометрии COBRA RACING лыжи обладают неповторимой гоночной точностью, качественной передачей усилия и высочайшей производительностью. Скользящая поверхность лыж K7000 / UNIVERSAL с универсальной структурой и двойными бороздками SKATING DOUBLE GUIDE GROOVE, повышающими устойчивость и способность контролировать направление, создаёт непревзойдённое скольжение, а заниженный и короткий мыс LOWER TIP придаёт лыжам повышенную манёвренность. На лыжах установлена платформа IFP для использования креплений системы TURNAMIC®, обеспечивающей лыже естественную гибкость и хорошее ощущение рельефа трассы.",
                "img1": "rossignol-xium.jpg",
                "img2": "rossignol-xium_3.jpg",
                "img3": "rossignol-xium_2.jpg",
                "category": "ski",
                "brand": "rossignol",
                "size": [
                    "186",
                    "192"
                ],
                "price": 59000,
                "checked": False,
                "sold": 1,
                "createdAt": "2023-01-03T11:36:01.852Z",
                "updatedAt": "2023-01-03T11:36:01.860Z",
                "__v": 0,
                "quantity": 1,
                "sizesToSell": "192"
            }
    }
    return product
