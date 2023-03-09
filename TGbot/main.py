import os
from pyrogram import Client
from dotenv import load_dotenv
from pyrogram.handlers import MessageHandler
load_dotenv()

api_id = os.getenv('APP_API_ID')
api_hash = os.getenv('APP_API_HASH')

with Client ('account', api_id, api_hash) as app:
    app.send_message('me', 'Some from bot')
# ________________________

newApp = Client ('account', api_id, api_hash)
#
# async def main():
#     async with newApp:
#         await newApp.send_message('me', 'newMSG')
# newApp.run(main())

# @newApp.on_message
# def myHandler(client, message):
#     message.forward('me')
# newApp.run()

def listenerMessage(client, message):
    message.forward('me')

my_handler = MessageHandler(listenerMessage)
newApp.add_handler(my_handler)
newApp.run()
