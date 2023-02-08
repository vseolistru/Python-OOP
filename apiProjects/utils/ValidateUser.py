import requests

# data = {
#
# }

#
json = {
    'message': 'Пользователь с таким email уже существует'
}


"""Check User Json"""
class ValidateUser:

    @staticmethod
    def checkNegativeUserCreationEmptyFields(json):
        print('function validate Empty Email || Name || Pass')
        """assertion empty Email"""
        assert json['message'] == 'Значения полей name, email и password не должны быть пустыми'
    @staticmethod
    def checkNegativeUserCreationBadEmail(json):
        print ('function validate Bad Email params')
        """asssertion bad Email"""
        assert json['message'] == 'Email должен содержать - @ и домен'

    @staticmethod
    def checkNegativePassUserLength(json):
        print('function - validate short password')
        """assertion short Password"""
        assert json['message'] == 'Пароль должен содержать 8 символов'

    @staticmethod
    def checkNegativeWeakPassUser(json):
        print('function - validate weak password')
        """assertion short Password"""
        assert json['message'] == 'Пароль должен содержать 1 спецсимвол, 1 символ верхнего регистра, 1 число'

    @staticmethod
    def checkResponseCreatedUser (json, data):
        """assertion Creation new User"""
        print('function - validateUserJson')
        assert data['name'] == json['name']
        assert data['email'] == json['email']
        assert isinstance(json['cart'], list)
        assert isinstance(json['activationLink'], str)
        assert isinstance(json['_id'], str)
        assert isinstance(json['createdAt'], str)
        assert isinstance(json['updatedAt'], str)
        assert isinstance(json['token'], str)

    @staticmethod
    def checkResponseLoginUser(json, data):
        """assertion Login User"""
        print ('function - validate User login')
        assert isinstance(json['_id'], str)
        assert data['email'] == json['email']
        assert isinstance(json['isActivated'], bool)
        assert isinstance(json['role'], int)
        assert isinstance(json['cart'], list)
        assert isinstance(json["name"], str)
        assert isinstance(json['activationLink'], str)
        assert isinstance(json['createdAt'], str)
        assert isinstance(json['updatedAt'], str)
        assert isinstance(json['token'], str)

    @staticmethod
    def checkIfUserExist(json):
        """assertion If User have Existed"""
        print('function - validate Error User have is Exist')
        assert json['message'] == 'Пользователь с таким email уже существует'

    @staticmethod
    def checkActivationLink(json):
        """assertion login activationLink"""
        print('function - validate User activationLink')
        assert json['message'] == 'all fine'

    @staticmethod
    def checkNegativeActivationLink(json):
        """assertion link in activation"""
        print('function - validate invalid link')
        assert json['message'] == 'Incorrect link'

    @staticmethod
    def checkUSERdataisExistforLocalStore(json, user_id):
        """assertion User Date for localStore"""
        print('function - validate data infor browser localStore')
        assert json['_id'] == user_id
        assert isinstance(json['name'], str)
        assert isinstance(json['email'], str)
        assert isinstance(json['isActivated'], bool)
        assert isinstance(json['role'], int)
        assert isinstance(json['cart'], list)

        assert isinstance(json['activationLink'], str)
        assert isinstance(json['createdAt'], str)
        assert isinstance(json['updatedAt'], str)
        if len(json['cart']):
            for item in (json['cart']):
                assert isinstance(item['_id'], str)
                assert isinstance(item['productId'], str)
                assert isinstance(item['slug'], str)
                assert isinstance(item['title'], str)
                assert isinstance(item['description'], str)
                assert isinstance(item['img1'], str)
                assert isinstance(item['img2'], str)
                assert isinstance(item['img3'], str)
                assert isinstance(item['category'], str)
                assert isinstance(item['brand'], str)
                assert isinstance(item['size'], list)
                assert isinstance(item['price'], int)
                # assert isinstance(item['checked'], bool)
                assert isinstance(item['sold'], int)
                assert isinstance(item['createdAt'], str)
                assert isinstance(item['updatedAt'], str)
                assert isinstance(item['quantity'], int)
                assert isinstance(item['sizesToSell'], str)

    @staticmethod
    def checkNegativeUSERdataisExistforLocalStore(json):
        """assertion Nedative ID User Date for localStore"""
        print('function - validate user localStore Data if Negative userID')
        assert json['message'] == 'User is not found'

