from utils.api import Api_ski_store
from userDate import headers
from userDate import newUserData, userLoginData, userIDLocalStoreData
from userDate import negativeEmptyFields, badEmailData, incorrectPasswordLenghtData
from userDate import userIDLocalStoreDataNegative, productData



""" python3 -m pytest -s -v """
"""test data usersApi init"""
data = newUserData()
loginData = userLoginData()


emptyFiledsDataOne = negativeEmptyFields()[0]
emptyFiledsDataTwo = negativeEmptyFields()[1]
emptyFiledsDataThree = negativeEmptyFields()[2]
badEmailDataOne = badEmailData()[0]
badEmailDataTwo = badEmailData()[1]
incorrectPassDataOne = incorrectPasswordLenghtData()[0]
incorrectPassDataTwo = incorrectPasswordLenghtData()[1]


def test_create_user_negative():
    Api_ski_store.start_USERS_API_testing()
    """Negative registration new User"""

    Api_ski_store.negative_create_user_empty_email(emptyFiledsDataOne, emptyFiledsDataTwo, emptyFiledsDataThree)
    Api_ski_store.negative_create_user_bad_email(badEmailDataOne, badEmailDataTwo)
    Api_ski_store.negative_user_pass_length(incorrectPassDataOne, incorrectPassDataTwo)
def test_create_user():
    """create new User"""
    Api_ski_store.create_new_user_and_validate(data)

def test_login_user_and_activate():
    """login created user"""
    Api_ski_store.check_user_login(loginData)
    Api_ski_store.check_user_activate(loginData)




