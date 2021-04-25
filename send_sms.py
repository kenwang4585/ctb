
from twilio.rest import Client
import os
import sys




def send_me_sms(to_num,message):
	'''
	Send sms to cellphone
	:param message:  message content
	:param to_num: receiving phone#
	:return: None
	'''
	try:
		sid=os.getenv('SMS_SID')
		auth=os.getenv('SMS_AUTH')
		myTwilio_num = '+15203293250'

		twilio_client=Client(sid,auth)

		twilio_client.messages.create(body=message,to=to_num,from_=myTwilio_num)
	except Exception as e:
		func_name=sys._getframe().f_code.co_name

		print(func_name,':',e)


if __name__=='__main__':
	to_num = '+8618665932236'
	message='this is a test message from twilio- 12：46'
	send_me_sms(to_num, message)

