from utils.api import Api_ski_store

def test_create_user():
    Api_ski_store.create_new_user_and_validate()
    Api_ski_store.check_user_login()
    Api_ski_store.check_user_activate()

