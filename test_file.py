from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from _datetime import datetime
from win32com.client import Dispatch


def speak(str1):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)


video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('C:/Users/mdran/PycharmProjects/Smart Attendance '
                                   'Project/data/haarcascade_frontalface_default.xml')

with open('C:/Users/mdran/PycharmProjects/Smart Attendance Project/data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)
with open('C:/Users/mdran/PycharmProjects/Smart Attendance Project/data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

kmn = KNeighborsClassifier(n_neighbors=5)
kmn.fit(FACES, LABELS)

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    col = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(
        col,
        scaleFactor=1.1,
        minNeighbors=5,
        flags=cv2.CASCADE_SCALE_IMAGE,
        minSize=(30, 30))
    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = kmn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile(
            "C:/Users/mdran/PycharmProjects/Smart Attendance Project/Attendence/Attendence_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 1)
        attendance = [str(output[0]), str(timestamp)]
    cv2.imshow("Frame", frame)
    k = cv2.waitKey(1)
    if k == ord('p'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("C:/Users/mdran/PycharmProjects/Smart Attendance Project/Attendence/Attendence_" + date + ".csv",
                      "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("C:/Users/mdran/PycharmProjects/Smart Attendance Project/Attendence/Attendence_" + date + ".csv",
                      "+a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()
    if k == ord('q'):
        break
video.release()
cv2.destroyAllWindows()
