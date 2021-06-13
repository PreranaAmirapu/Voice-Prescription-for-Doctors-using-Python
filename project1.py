import speech_recognition as sr
import pyttsx3
import xlsxwriter


r = sr.Recognizer() 

 
def SpeakText(command): 
	
	
	engine = pyttsx3.init() 
	engine.say(command) 
	engine.runAndWait() 
try: 
		
	with sr.Microphone() as source: 	
		print("Your name is:")
		r.adjust_for_ambient_noise(source,duration = 0.5)
		audio = r.listen(source) 
			
			
		MyText = r.recognize_google(audio) 
			
		print(MyText) 
			
			
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear") 


try: 
	from word2number import w2n

	with sr.Microphone() as source1: 	
		print("Your age is:")
		r.adjust_for_ambient_noise(source1,duration = 0.5)
		audio1 = r.listen(source1) 
			
			
		MyText1 = r.recognize_google(audio1)
		number = w2n.word_to_num(MyText1)
		print(number) 
			
			
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear") 	


try: 
	glist= ["male","female","other"]
		
	with sr.Microphone() as source2: 	
		print("Your gender is: select one : male ,female,other")
		r.adjust_for_ambient_noise(source2,duration = 0.5)
		audio2 = r.listen(source2) 
			
			
		MyText2 = r.recognize_google(audio2) 
		if MyText2 in glist:
		   print(MyText2)	
		else:
		   print("choose from the mentioned")   

except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("Select from the mentioned") 

try: 
	
		
	with sr.Microphone() as source3:
		print("Symptoms:") 	
		r.adjust_for_ambient_noise(source3,duration = 0.5)
		audio3 = r.listen(source3)
		MyText3 = r.recognize_google(audio3) 
		if("symptoms" in MyText3):
			new = MyText3.split();
			word=["symptoms"]
			SYMPTOMS= " ".join([i for i in new if i not in word ])
			print(SYMPTOMS)
		else:
			print("symptoms are not audible")	
		
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear") 


try: 
		
	with sr.Microphone() as source4:
		print("Diagnosis:") 	
		r.adjust_for_ambient_noise(source4,duration = 0.5)
		audio4 = r.listen(source4)
		MyText4 = r.recognize_google(audio4) 
		if("diagnosis" in MyText4):
			new1 = MyText4.split();
			word1=["diagnosis"]
			DIAGNOSIS= " ".join([i for i in new1 if i not in word1])
			print(DIAGNOSIS)
		else:
			print("voice not audible")	
		
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear") 


try: 
		
	with sr.Microphone() as source5:
		print("prescription:") 	
		r.adjust_for_ambient_noise(source5,duration = 0.5)
		audio5 = r.listen(source5)
		MyText5 = r.recognize_google(audio5) 
		if("prescription" in MyText5):
			new2 = MyText5.split();
			word2=["prescription"]
			PRESCRIPTION= " ".join([i for i in new2 if i not in word2])
			print(PRESCRIPTION)
		else:
			print("voice not audible")	
		
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear")


try: 
		
	with sr.Microphone() as source6:
		print("advice") 	
		r.adjust_for_ambient_noise(source6,duration = 0.5)
		audio6 = r.listen(source6)
		MyText6= r.recognize_google(audio6) 
		if("advice" in MyText6):
			new3 = MyText6.split();
			word3=["advice"]
			ADVICE = " ".join([i for i in new3 if i not in word3])
			print(ADVICE)
		else:
			print("voice not audible")	
		
except sr.RequestError as e: 
	print("Could not request results; {0}".format(e)) 
		
except sr.UnknownValueError: 
	print("voice not clear")	

Workbook = xlsxwriter.Workbook("project1.xlsx")
worksheet = Workbook.add_worksheet()
data = (
			{'name of the patient': MyText,
			'age' : MyText1,
			'Gender':MyText2,
			'symptoms' : SYMPTOMS,
		    'diagnosis':DIAGNOSIS,
			'prescription':PRESCRIPTION,
			'advice' : ADVICE})
row = 0
col = 0
for i, d in data.items():
		worksheet.write(row,col,   i)
		worksheet.write(row,col+1,d)
		row +=1
Workbook.close()
