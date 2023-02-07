import requests
class TestGoogleLocation:
    """"""

    def createNewGoogleLocation(self):
        BASEURL = 'https://rahulshettyacademy.com'
        KEY = "?key=qalick123"
        POSTURL = '/maps/api/place/add/json'

        URL = BASEURL+POSTURL+KEY

        jsonForCreateNewLocation = {
            "location": {
                "lat": -38.383494,
                "lang": 33.427362
            },
            "accuracy": 50,
            "name": "Frontline house",
            "phone_number": "(+91) 983 893 3937",
            "address": "29, side layout, cohen 09",
            "types": [
                "shoe park",
                "shop"
            ],
            "website": "http://google.com",
            "language": "French-IN"
        }

        resultPost = requests.post(URL, json=jsonForCreateNewLocation)
        print(resultPost)

page = TestGoogleLocation()
page.createNewGoogleLocation()
