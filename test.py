import os
import dotenv
from dotenv import load_dotenv


load_dotenv()

print(os.environ.get("EMAIL"))
print(os.environ.get("USER"))