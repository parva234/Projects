import speech_recognition as sr
import pyttsx3
import subprocess
import os
import webbrowser
import datetime
from os import path
import ctypes
import pyautogui
import cv2
import time as t
import win32com.client
import ecapture as ec 
import numpy as np
import requests

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
         speak("Good Morning sir, I am Velter, What can I do for you?")
    elif hour>=12 and hour<18:
         speak("Good Afternoon sir, I am Velter, What can I do for you?")
    else:
         speak("Good Evening sir, I am Velter, What can I do for you?")       
         
def time():
     strTime = (datetime.datetime.now().strftime('%H:%M:%S'))
     speak(f"Sir, the time is {strTime}")

def takecommand():
    r =sr.Recognizer()
    with sr.Microphone() as source:
        print("Listen...")
        audio = r.listen(source)        
    try:
        print("Recognition.....")
        query = r.recognize_google(audio)
        print(f"User said: {query}\n")         
    except Exception as e:
        print(e)        
        return None
    return query.lower()

def capture_image(filename='captured_image.jpg', gallery_dir='gallery'):
    """Capture an image from the webcam and save it to the specified directory."""
    if not os.path.exists(gallery_dir):
        os.makedirs(gallery_dir)

    filepath = os.path.join(gallery_dir, filename)
    
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open video capture device.")
        return None

    ret, frame = cap.read()
    if not ret:
        print("Error: Could not read frame from video capture device.")
        return None

    cv2.imwrite(filepath, frame)
    cap.release()
    return frame, filepath

def compare_images(image1, image2):
    """Compare two images using ORB feature matching."""

    orb = cv2.ORB_create()
    gray1 = cv2.cvtColor(image1, cv2.COLOR_BGR2GRAY)
    gray2 = cv2.cvtColor(image2, cv2.COLOR_BGR2GRAY)

    kp1, des1 = orb.detectAndCompute(gray1, None)
    kp2, des2 = orb.detectAndCompute(gray2, None)

    bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = bf.match(des1, des2)

    if len(matches) == 0:
        return float('inf')  

    matches = sorted(matches, key=lambda x: x.distance)

    mean_distance = sum(match.distance for match in matches) / len(matches)
    
    return mean_distance

def password():
    speak("which password would you like to use")
    speak("option 1,speak the password")
    speak("option 2,face password")
    query = takecommand().lower()
    if "speak" in query:
        speak("Please tell the password")
        password = takecommand().lower()
        array = ["welter", "user"]
        if password in array:
            speak("Access granted.")
            print("Access granted.")
            return True
        else:
            speak("Access denied.")
            print("Access denied.")
            return False
    elif "face password" in query:
        password_filename = 'password_image.jpg'
        gallery_dir = 'gallery'

        password_filepath = os.path.join(gallery_dir, password_filename)
        if not os.path.exists(password_filepath):
            print("No password image found. Capturing a new password image.")
            password_image, _ = capture_image(password_filename, gallery_dir)
            if password_image is None:
                print("Error: Failed to capture password image.")
                return
            print(f"Password image captured and saved at {password_filepath}.")
        else:
            print("Password image found.")

        password_image = cv2.imread(password_filepath)

    
        print("Please look at the camera to verify your identity.")
        new_image, _ = capture_image()

        if new_image is None:
            print("Error: Failed to capture image for comparison.")
            return
        mean_distance = compare_images(password_image, new_image)

        threshold = 50  
        if mean_distance < threshold:
            print("Access granted.")
            speak("Access granted.")
            return True
        else:
            print("Access denied.")
            speak("Access denied.")
            return False
 
def open_word_and_draw_circle():
    speak("opening Word and performing multiple actions")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    action = takecommand().lower()
    if 'draw circle' in action:
        shape = doc.Shapes.AddShape(9, 100, 100, 100, 100)  
        shape.Fill.ForeColor.RGB = 0x0000FF
        shape.Line.ForeColor.RGB = 0x0000FF  
    elif 'add paragraph' in action:
            speak("adding Paragraph")
            paragraph = doc.Paragraphs.Add()
            paragraph.Range.Text = f"Hello "+takecommand()
            paragraph.Range.Font.Name = "Blackadder ITC"
            paragraph.Range.Font.Size = 30
    else:
        speak("Invalid action. Please try again.")

def open_excel_and_create_sheet():
    speak("opening Excel and creating a new sheet")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbook = excel.Workbooks.Add()
    worksheet = workbook.Sheets.Add()
    action = takecommand().lower()
    if 'create table' in action:
        speak("creating table")
        worksheet.Range("A1:B2").Value = "Hello World"
    elif 'insert data' in action:
            speak("inserting data")
            worksheet.Range("A3").Value = f"Hello "+takecommand()
    else:
        speak("Invalid action. Please try again.")

def open_powerpoint_and_create_slide():
    speak("opening PowerPoint and creating a new slide")
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True
    presentation = powerpoint.Presentations.Add()
    slide = presentation.Slides.Add(1, 12)
    action = takecommand().lower()
    if 'add a text' in action:
        speak("adding text")
        shape = slide.Shapes.Add(1, 100, 100, 100, 100)
        shape.TextFrame.TextRange.Text = f"Hello "+takecommand()
    elif 'add a image' in action:
            speak("adding image")
            shape = slide.Shapes.AddPicture("C:\velter\b.jpg", 1, 100, 100, 100, 100)
            
    else:
        speak("Invalid action. Please try again.")

def open_notepad_and_write_text():
    speak("opening Notepad and writing text")
    subprocess.Popen([r'C:\Windows\System32\notepad.exe'])
    action = takecommand().lower()
    if 'start' in action:
        speak("writing text")
        text_to_write = takecommand()  
        pyautogui.write(f"{text_to_write}") 
    else:
        speak("Invalid action. Please try again.")

def record_screen(output_filename, duration=20, fps=10):
    speak("Recording start")
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    screen_size = pyautogui.size()
    out = cv2.VideoWriter(output_filename, fourcc, fps, screen_size)
    
    start_time = t.time()
    while (t.time() - start_time) < duration:
        img = pyautogui.screenshot()
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        out.write(frame)
        t.sleep(1 / fps)
    out.release()

def take_screenshot():
    screenshot = pyautogui.screenshot()
    filename = f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    screenshot.save(filename)
    print(f"Screenshot saved as {filename}")

def searching():
    s =sr.Recognizer()
    with sr.Microphone() as source:
        print("Listen...")
        audio = s.listen(source)        
    try:
        print("Searching.....")
        search = s.recognize_google(audio)
        speak(f"Searching Successfully: {search}\n") 
        print(f"Searching Successfully: {search}\n")         
    except Exception as e:
        print(e)        
    return search  

def web():
     speak("which web open ples tell")
     query = takecommand()
     if 'open youtube' in query:
          speak ("What do you search in youtube")
          u = searching()
          speak("searching in youtube")
          webbrowser.open(f"https://www.youtube.com/results?search_query={u}")
     elif 'open google' in query:
          speak("What do you search in google")
          s = searching()
          speak("searching in Google")
          webbrowser.open(f"https://www.google.com/search?q={s}")
     elif 'open wikipedia'in query:
           speak("What do you search")
           search = searching()
           webbrowser.open(f"https://en.wikipedia.org/wiki/{search}")
     elif 'open instagram' in query:
          speak("opening Instagram")
          webbrowser.open("https://www.instagram.com/")

def open_whatsapp():
    whatsapp_path = "start whatsapp:"
    os.system(whatsapp_path)
    t.sleep(10)  

def send_whatsapp_message(phone_number, message):
                open_whatsapp()
                pyautogui.hotkey('ctrl', 'f')
                t.sleep(2)
                pyautogui.write(phone_number)
                t.sleep(2)
                pyautogui.press('enter')
                t.sleep(2)
                pyautogui.click(x=200, y=250) 
                t.sleep(2)
                pyautogui.write(message)
                t.sleep(1)
                pyautogui.press('enter')

def get_contact_number(contact_name):
                contacts = {
                    "mummy": "Parva Suthar",
                    "Kathan" : "Kathan"
                }
                
                if contact_name in contacts:
                    return contacts[contact_name]
                else:
                    return None

def mcq():
  speak("Please select the question 1 or 2")
  query=takecommand()
  if 'question  1' in query:
    speak("what is your age")
    speak("option A:18")
    speak("option b:17")
    speak("option c:16")
    speak("option d:20")
    query=takecommand()
    if 'b' or 'answer is b' in query:
        speak("it is right")
        print("it is right")
    else:
        speak("it is wrong")
        print("it is wrong")
  elif 'question 2' in query:
      speak("what is your name")
      speak("option A:user")
      speak("option b:parv")
      speak("option c:kathan")
      speak("option d:wetler")
      if 'B' or 'answer is b' in query:
          speak("answer is right")
          print("answer is right")
      else:
          speak("answer is wrong")
          print("answer is wrong")
      
def application():
     speak("which application do you want to open")
     query = takecommand()
     if 'open notepad' in query:
        speak("Opening Notepad")
        open_notepad_and_write_text()
     elif 'open word' in query:   
         speak("Opening worldpad")
         open_word_and_draw_circle()
     elif 'open powerpoint' in query:   
         speak("Opening powerpoint")
         open_powerpoint_and_create_slide()
     elif 'open excel' in query:   
         speak("Opening excel")
         open_excel_and_create_sheet()
     elif 'open calculator' in query:
             speak("opening calculator")
             subprocess.Popen([r'C:\Windows\System32\calc.exe'])  
     elif 'open vs' in query:
            speak("opening Vs code")
            subprocess.Popen([r'C:\Users\DELL\AppData\Local\Programs\Microsoft VS Code\Code.exe'])
     elif 'open game' in query:    
         speak("opning Asphalt 9") 
         subprocess.Popen([r"C:\XboxGames\Asphalt Legends Unite\Content\Asphalt9_gdk_x64_rtl.exe"])
         while True:
            query = takecommand().lower()
            if "start" in query:
             t.sleep(120)
             pyautogui.keyDown('w')  
            elif " left" in query:
             pyautogui.keyDown('a')  
             t.sleep(2)
            elif "right" in query:
             pyautogui.keyDown('d')  
             t.sleep(2)
     elif 'open whatsapp' in query:  
         speak("Say the name of the contact...")
         contact_name = takecommand()
         if contact_name:
          print(f"Opening chat for {contact_name}...")
          contact_number = get_contact_number(contact_name)
         if contact_number:
            speak("Say the message you want to send...")
            message = takecommand()
            if message:
                send_whatsapp_message(contact_number, message)
                speak(f"Message sent to {contact_name}")
                print(f"Message sent to {contact_name}")
            else:
                print(f"Contact {contact_name} not found.")
         else:
          print("Failed to recognize the contact name.")
     elif 'open spotify' in query:
                speak("Opening Spotify")
                os.system("start spotify:")
                t.sleep(10)
                speak("What song would you like to play?")
                song_name = searching().lower()
                search_and_play_song_on_spotify(song_name)     
     elif 'close' in query:
         speak("closing ai")
         exit()
     
def search_and_play_song_on_spotify(song_name):
    pyautogui.hotkey('ctrl', 'l')
    t.sleep(1)
    pyautogui.write(song_name)
    t.sleep(1)
    pyautogui.press('enter')
    t.sleep(2)
    pyautogui.press('tab', presses=14)
    pyautogui.press('enter')
     
def play_song(song_name):
     music_dir = "C:" 
     song_path = os.path.join(music_dir, f"{song_name}.mp3")
     os.startfile(song_path) 

def stop_song():
    os.system("taskkill /im wmplayer.exe /f") 

def song():
    speak("Please select a song to play.")
    song_name = takecommand()
    if song_name == 'believer':
        play_song('believer')
    elif song_name == 'music':
        play_song('hey')

    while True:
        speak("Say 'stop' to stop the song.")
        command = takecommand()
        if "stop" in command:
            stop_song()
            speak("Song stopped.")
            break

def system():
    speak("which system would you like to use")
    query = takecommand().lower()
    if 'shutdown' in query:
        speak("Shutting down the computer...")
        os.system("shutdown /s /t 1")
    elif  'restart' in query:   
        speak("restart pc...") 
        os.system("restart /r /t 1")
    elif 'close' in query:
        speak("closing ai")
        exit()
    elif 'take a screenshot' in query:
        speak("taking screenshot")
        take_screenshot()    
    elif "camera" in query or "take a photo" in query:
         ec.capture(0, " Camera ", "photo.jpg")
         t.sleep(2)
         pyautogui.press('q')
    elif "recording" in query:
        output_filename = "screen_recording.avi"
        duration = 20 
        fps = 10  
        record_screen(output_filename, duration, fps)
        print(f"Screen recording saved as {output_filename}")     
    elif 'change background' in query:
         w = os.path.abspath("b.jpg")
         ctypes.windll.user32.SystemParametersInfoW(20, 0,w,0)
         speak("Background changed successful")
    elif "check weather" in query or "weather" in query:
        api_key = "47822dd95146c9cbebe6be6a0124807c"
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        speak("ples tell the  City name ")
        print("City name : ")
        city_name = takecommand()
        complete_url = base_url + "appid=" + api_key + "&q=" + city_name
        response = requests.get(complete_url)
        x = response.json()
        
        if x["cod"] != "404":
            y = x["main"]
            current_temperature = y["temp"]
            current_pressure = y["pressure"]
            current_humidity = y["humidity"]
            z = x["weather"]
            weather_description = z[0]["description"]
            print(" Temperature  = " +
            str(current_temperature - 273.15) +
            "\n atmospheric pressure (in hPa unit) =" +
            str(current_pressure) +
            "\n humidity (in percentage) = " +
            str(current_humidity) +
            "\n description = " +
            str(weather_description))
        else:
            speak(" City Not Found ")
     
def calculator():
    speak("Enter the first number")
    num1 = int(takecommand())
    speak("Enter the second number")
    num2 = int(takecommand())
    speak("select the action")
    query = takecommand().lower()
    if 'addition' in query:
        add = num1 + num2
        print(f"Addition of {num1} and {num2} is {add}")
        speak(f"the sum is {add}")
    elif 'subtraction' in query:
        sub = num1 - num2
        print(f"Subtraction of {num1} and {num2} is {sub}")
        speak(f"the sub is {sub}")
    elif 'multiplication' in query:
        mul = num1 * num2
        print(f"Multiplication of {num1} and {num2} is {mul}")
        speak(f"the mul is {mul}")
    elif 'division' in query:
        div = num1 / num2
        print(f"Division of {num1} and {num2} is {div}")
        speak(f"the div is {div}")
    else:
        exit()

def audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        speak("Say something!")
        print("Say something!")
        audio = r.listen(source)
    try:
        with open("microphone-results.wav", "wb") as f:
            f.write(audio.get_wav_data())
        AUDIO_FILE = path.join(path.dirname(__file__), "microphone-results.wav")
        with sr.AudioFile(AUDIO_FILE) as source:
            audio = r.record(source)
            try:
                guj = r.recognize_google(audio, language="gu-IN")
                print("Gujarati:", guj)
                print("English :", r.recognize_google(audio))
            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Google Speech Recognition service; {0}".format(e))
    except Exception as e:
        print("Error while saving audio:", e)

def get_directory_path():
    speak("Where would you like to create the folder? Say 'Desktop', 'Documents', or specify a custom path.")
    directory_choice = takecommand()

    if directory_choice == "desktop":
        return os.path.join(os.path.expanduser("~"), "Desktop")
    elif directory_choice == "documents":
        return os.path.join(os.path.expanduser("~"), "Documents")
    else:
        if os.path.exists(directory_choice):
            return directory_choice
        else:
            speak("The specified path does not exist. Please provide a valid directory path.")
            return None

def create_folder_and_file():
    directory_path = get_directory_path()

    if directory_path is None:
        speak("Invalid directory. Exiting...")
        return

    speak("What would you like to name the folder?")
    folder_name = takecommand()

    if not folder_name:
        speak("Folder name not provided. Exiting...")
        return

    folder_path = os.path.join(directory_path, folder_name)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        speak(f"Folder '{folder_name}' created successfully in the chosen directory.")
    else:
        speak(f"Folder '{folder_name}' already exists in the chosen directory.")

def  final():      
        speak("What would you like to use?")  
        while True:
            query = takecommand()
            if query: 
                if  "web browser" in query:
                     web()
                elif "hello welter" in query:
                     speak("Yes boss")
                elif "good morning" in query:
                     speak("Good Moring")             
                elif "application" in query:
                    application()
                elif "song" in query:
                    song()
                elif "calculator" in query:
                    calculator()
                elif "system" in query:
                    system()
                elif "mcq" in query:
                    mcq()
                elif 'record' in query:
                    speak("audio start")
                    audio()
                elif 'thanks' in query:
                     speak("Thank you Sir")     
                elif 'create folder' in query:
                     create_folder_and_file()
                elif 'create file' in query:
                    speak("Please provide the name of the file with extension.")
                    file_name = takecommand()
                    folder_path = os.getcwd()
                    file_path = os.path.join(folder_path, file_name)
                    if not os.path.exists(file_path):
                        open(file_path, 'w').close()
                        speak(f"File '{file_name}' created successfully.")
                    else:
                        speak(f"File '{file_name}' already exists.")    
                elif 'close' in query:
                            speak("closing velter ")
                            exit()
                else:
                            print("Waiting for next command...")
                            continue

if __name__ == '__main__':
   #if password():
        wishme()
        final()
    #else:
        exit()            