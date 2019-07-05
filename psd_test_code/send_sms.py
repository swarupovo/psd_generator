import sys;
from googlevoice import Voice

voice = Voice()
voice.login("swarupovo@gmail.com","Python123@")

voice.send_sms("+917908504352", "hi how are u")