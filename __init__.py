
# Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International Public License  
# (CC BY-NC 4.0)
#
# You may share or modify this code as long as you provide proper credit.  
# You may NOT use it for commercial purposes.  
#
"""
{Non-exhaustive exerpt from the license}
    License Grant  
        1. Subject to the terms and conditions of this Public License,  
        the Licensor hereby grants You a worldwide, royalty-free,  
        non-sublicensable, non-exclusive, irrevocable license to  
        exercise the Licensed Rights in the Licensed Material to:  

        a. Reproduce and Share the Licensed Material, in whole or  
            in part, for NonCommercial purposes only.  

        b. Produce, reproduce, and Share Adapted Material for  
            NonCommercial purposes only.  

        2. Exceptions and Limitations. Where Exceptions and Limitations apply to  
        Your use, this Public License does not apply, and You do not need to  
        comply with its terms and conditions.  

        3. Term. The term of this Public License is specified in Section 6(a).  

        4. Media and Formats; Technical Modifications Allowed.  
        The Licensor authorizes You to exercise the Licensed Rights in  
        all media and formats, whether now known or hereafter created,  
        and to make technical modifications necessary to do so. The  
        Licensor waives and/or agrees not to assert any right or  
        authority to forbid You from making technical modifications  
        necessary to exercise the Licensed Rights, including  
        technical modifications necessary to circumvent Effective  
        Technological Measures. Simply making modifications authorized  
        by this section does not create Adapted Material.  
"""
#
# PangolinAnkiTTS was made by Pangolin-with-a-bow on GitHub.  
#
# Any modifications to this file must keep this entire header intact.  


import re
import threading
import time
import queue
import win32com.client
import pythoncom
from anki.hooks import wrap
from aqt.reviewer import Reviewer

speech_queue = queue.Queue()
cancel_event = threading.Event()
current_speaker = None  

def html_to_text(html):
    html = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</p>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'<p.*?>', '', html, flags=re.IGNORECASE)
    text = re.sub(r'<.*?>', '', html)
    return text.replace("&nbsp;", " ")
def speech_worker():
    """
    Worker thread that initializes COM once, processes speech tasks sequentially
    """
    global current_speaker
    pythoncom.CoInitialize()
    try:
        while True:
            text = speech_queue.get() 
            if text is None:
                break
            cancel_event.clear()
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            current_speaker = speaker
            speaker.Speak(text, 1)
            while speaker.Status.RunningState == 2: 
                if cancel_event.is_set():
                    speaker.Speak("", 3) 
                    break
                time.sleep(0.1)
            speech_queue.task_done()
    finally:
        pythoncom.CoUninitialize()

def enqueue_speech(text):
    """
    Cancel any current speech, clear pending tasks, enqueue new speech text
    """
    cancel_event.set()
    while not speech_queue.empty():
        try:
            _ = speech_queue.get_nowait()
            speech_queue.task_done()
        except queue.Empty:
            break
    speech_queue.put(text)

def stop_speech():
    """
    Signal cancellation, clear any pending speech tasks
    """
    cancel_event.set()
    while not speech_queue.empty():
        try:
            _ = speech_queue.get_nowait()
            speech_queue.task_done()
        except queue.Empty:
            break

def log_question(self):
    card = self.card
    if card:
        question_html = card.note().fields[0]
        question_text = html_to_text(question_html)
        enqueue_speech(question_text)
def log_answer(self):
    card = self.card
    if card:
        answer_html = card.note().fields[1]
        answer_text = html_to_text(answer_html)
        enqueue_speech(answer_text)
        print(answer_text)

def stop_speech_on_answer(self):
    stop_speech()

speech_worker_thread = threading.Thread(target=speech_worker, daemon=True)
speech_worker_thread.start()

Reviewer._showQuestion = wrap(Reviewer._showQuestion, log_question, "after")
Reviewer._showAnswer = wrap(Reviewer._showAnswer, log_answer, "before")
Reviewer._showAnswer = wrap(Reviewer._showAnswer, stop_speech_on_answer, "after")
