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
        """check login activationLink"""
        print('function - validate User activationLink')
        assert json['message'] == 'all fine'

    @staticmethod
    def checkNegativeActivationLink(json):
        """incorrect link in activation"""
        print('function - invalid link')
        assert json['message'] == 'Incorrect link'


