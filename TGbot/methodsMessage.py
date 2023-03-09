import os
from pyrogram import Client, filters
from dotenv import load_dotenv
from pyrogram.handlers import MessageHandler
load_dotenv()

api_id = os.getenv('APP_API_ID')
api_hash = os.getenv('APP_API_HASH')
group = 'testmyhandler'

app = Client ('account', api_id, api_hash)

# app.send_message(group, 'Some from bot')

@app.on_message()
def echo (client, message):
    # print(message.chat.id)
    app.send_message(group, "Hello from bot")
    # app.send_document(group, 'email_ClaimFile90053004.docx')
    # app.send_photo(group, "mobileHeader.png")
    app.get_chat_history(group)

app.run()


# with Client ('account', api_id, api_hash) as app:
#     app.send_message(group, 'Some from bot')

