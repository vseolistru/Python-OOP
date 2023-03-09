from utils.api import Api_ski_store
from userDate import headers
from userDate import newUserData, userLoginData, userIDLocalStoreData
from userDate import userIDLocalStoreDataNegative, productData



""" python3 -m pytest -s -v """
"""test data usersApi init"""
data = newUserData()
loginData = userLoginData()
user = userIDLocalStoreData()
negativeUser = userIDLocalStoreDataNegative()
headers = headers()
product = productData()


def test_local_store_positive_negative():
    """user infor for localStore"""
    Api_ski_store.check_user_data_is_return(user, headers)
    Api_ski_store.check_user_data_is_return_Negative(negativeUser, headers)

def test_addproduct_to_cart():
    """add cart to User"""
    Api_ski_store.check_addcart_product(user, headers, product)