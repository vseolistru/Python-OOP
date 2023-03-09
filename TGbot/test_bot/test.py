import os
from dotenv import load_dotenv
load_dotenv()


value = os.getenv('SOME_DATA')


print(value)