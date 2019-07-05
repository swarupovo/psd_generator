import time 
from time import sleep 
from sinchsms import SinchSMS 
  
# function for sending SMS 
def sendSMS(): 
  
    # enter all the details 
    # get app_key and app_secret by registering 
    # a app on sinchSMS 
    number = '+917908504352'
    app_key = 'c646f3de015f4c83a41edea5ed57e830'
    app_secret = 'ae36d1aa396b4f10950c19f2407276b6'
    
    # enter the message to be sent 
    message = 'Hello Message!!!'
    print("sending sms")
  
    client = SinchSMS(app_key, app_secret) 
    print("Sending '%s' to %s" % (message, number)) 
  
    response = client.send_message(number, message) 
    message_id = response['messageId'] 
    response = client.check_status(message_id) 
  
    # keep trying unless the status retured is Successful 
    while response['status'] != 'Successful': 
        print(response['status']) 
        time.sleep(1) 
        response = client.check_status(message_id) 
  
    print(response['status']) 
  
 
sendSMS() 