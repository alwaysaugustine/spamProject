import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
from win32com.client import Dispatch
import pythoncom

pythoncom.CoInitialize()


def speak(text):
    speakk = Dispatch(("SAPI.SpVoice"))
    speakk.Speak(text)


model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))


def main():
    st.title("SMS Spam Detector")
    st.write("Built using Python & Streamlit")
    activites = ["Classification", "About Us"]
    choices = st.sidebar.selectbox("Select Activities", activites)
    if choices == "Classification":
        st.subheader("Check Your Message Here!")
        msg: str = st.text_input("Enter Message:")
        if st.button("Process"):
            print(msg)
            print(type(msg))
            data = [msg]
            print(data)
            vec = cv.transform(data).toarray()
            result = model.predict(vec)
            if result[0] == 0:
                st.success("This is Not A Spam SMS")
                speak("This is Not A Spam SMS")
            else:
                st.error("This is A Spam SMS")
                speak("This is A Spam SMS")

    if choices == "About Us":
        st.title("About Us")
        msg1: str = st.subheader(
            "Classification problems can be broadly split into two categories: Binary Classification means there are only two possible label classes. Example: A patient's condition is cancerous or it isn't; or a financial transaction is fraudulent or not. Multi-class classification refers to the cases where there are more than two label classes. Spam Detection is a similar case of Binary Classification.")
        print("\n")
        msg1: str = st.subheader(
            "This website is specifically built to detect and predict whether a message is SPAM or HAM(Legitimate)")
        print("\n")
        msg1: str = st.subheader(
            "We have used a dataset containing a large number of SMS which include spam and those which aren't. Using Multinomial Naive Bayes Machine Learning Algorithm which can predict whether a message is spam or not. The backend of the project has been written in Jupyter Notebook and PyCharm Community Edition using Python Language. Packages such as 'sklearn' , 'pickle' and 'win32 client' have been used to make the execution possible.")
        print("\n")
        msg1: str = st.subheader(
            "The website is currently being displayed on Streamlit using the imported 'streamlit' package which displays the content in the web browser.")
        print(msg1)


main()
