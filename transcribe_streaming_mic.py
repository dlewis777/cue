#!/usr/bin/env python

# Modified the Google Tutorial Code for streaming audio using their speech API
# Copyright 2017 Google Inc. All Rights Reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Google Cloud Speech API sample application using the streaming API.
NOTE: This module requires the additional dependency `pyaudio`. To install
using pip:
    pip install pyaudio
Example usage:
    python transcribe_streaming_mic.py
"""

# [START import_libraries]
from __future__ import division

import re
import sys
import pyautogui
import easygui
import win32gui
import pywinauto

from time import sleep
from google.cloud import speech
from google.cloud.speech import enums
from google.cloud.speech import types
import pyaudio
from six.moves import queue
# [END import_libraries]

# Audio recording parameters
RATE = 16000
CHUNK = int(RATE / 10)  # 100ms

backdict = {
    'zero': 0,
    'one': 1,
    'two': 2,
    'three': 3,
    'four': 4,
    'five': 5,
    'six': 6,
    'seven': 7,
    'eight': 8,
    'nine': 9,
    'ten': 10
}

ilist = [1,2,3,4,5,6,7,8,9,10]

class MicrophoneStream(object):
    """Opens a recording stream as a generator yielding the audio chunks."""
    def __init__(self, rate, chunk):
        self._rate = rate
        self._chunk = chunk

        # Create a thread-safe buffer of audio data
        self._buff = queue.Queue()
        self.closed = True

    def __enter__(self):
        self._audio_interface = pyaudio.PyAudio()
        self._audio_stream = self._audio_interface.open(
            format=pyaudio.paInt16,
            # The API currently only supports 1-channel (mono) audio
            # https://goo.gl/z757pE
            channels=1, rate=self._rate,
            input=True, frames_per_buffer=self._chunk,
            # Run the audio stream asynchronously to fill the buffer object.
            # This is necessary so that the input device's buffer doesn't
            # overflow while the calling thread makes network requests, etc.
            stream_callback=self._fill_buffer,
        )

        self.closed = False

        return self

    def __exit__(self, type, value, traceback):
        self._audio_stream.stop_stream()
        self._audio_stream.close()
        self.closed = True
        # Signal the generator to terminate so that the client's
        # streaming_recognize method will not block the process termination.
        self._buff.put(None)
        self._audio_interface.terminate()

    def _fill_buffer(self, in_data, frame_count, time_info, status_flags):
        """Continuously collect data from the audio stream, into the buffer."""
        self._buff.put(in_data)
        return None, pyaudio.paContinue

    def generator(self):
        while not self.closed:
            # Use a blocking get() to ensure there's at least one chunk of
            # data, and stop iteration if the chunk is None, indicating the
            # end of the audio stream.
            chunk = self._buff.get()
            if chunk is None:
                return
            data = [chunk]

            # Now consume whatever other data's still buffered.
            while True:
                try:
                    chunk = self._buff.get(block=False)
                    if chunk is None:
                        return
                    data.append(chunk)
                except queue.Empty:
                    break

            yield b''.join(data)
# [END audio_stream]



def listen_for_cue(responses, cues, mode):
    curindex = 0
    print(cues[curindex])
    num_chars_printed = 0
    for response in responses:
        if not response.results:
            continue

        # The `results` list is consecutive. For streaming, we only care about
        # the first result being considered, since once it's `is_final`, it
        # moves on to considering the next utterance.
        result = response.results[0]
        if not result.alternatives:
            continue

        # Display the transcription of the top alternative.
        transcript = result.alternatives[0].transcript

        # Display interim results, but with a carriage return at the end of the
        # line, so subsequent lines will overwrite them.
        #
        # If the previous result was longer than this one, we need to print
        # some extra spaces to overwrite the previous result
        overwrite_chars = ' ' * (num_chars_printed - len(transcript))

        if not result.is_final:
            sys.stdout.write(transcript + overwrite_chars + '\r')
            sys.stdout.flush()

            num_chars_printed = len(transcript)

        else:
            print(transcript + overwrite_chars)

            if 'back' in transcript:
                print('LA')
                for x in range(len(transcript.split(' '))):
                    print(x)
                    print(transcript.split(' ')[x])
                    if 'back' == str(transcript.split(' ')[x]):
                        print('HERE')
                        if x < (len(transcript.split(' ')) - 1):
                            print('THERE')
                            
                            if transcript.split(' ')[x+1] in backdict:
                                print('DONE2')
                                for y in range(int(backdict[transcript.split(' ')[x+1]])):
                                    pyautogui.hotkey('left')
                                    print('left')
                                break              
                            elif int(transcript.split(' ')[x+1]) in ilist:
                                print('DONE') 
                                for y in range(int(transcript.split(' ')[x+1])):
                                    pyautogui.hotkey('left')
                                    print('left')
                                    sleep(.1)
                                break    
            elif mode is 'Advanced':
                if cues[curindex].lower() in transcript:
                    pyautogui.hotkey('right')
                    if curindex == (len(cues) - 1):
                        break
                    curindex += 1
                num_chars_printed = 0
            elif mode is 'Simple':
                for element in cues:
                    if element.lower() in transcript:
                        pyautogui.hotkey('right')
                        break


def choose_mode():

    msg = "Choose which mode to run in. Simple lets you use any of the words to change. Advanced gives you finer control."
    title = "Cue"
    choices = ['Simple', 'Advanced']
    mode = easygui.buttonbox(msg, title, choices)
    return mode
def get_word_list(mode):
    if mode is 'Advanced':
        msg = "Enter the number of slides"
        title = "Cue"
        slidenum = easygui.integerbox(msg, title)

        msg2 = "Enter transition phrase for each slide"
        fieldNames = []
        for x in range(int(slidenum)):
            fieldNames.append("Slide " + str(x))
        fieldValues = []
        fieldValues = easygui.multenterbox(msg2, title, fieldNames)

        return fieldValues
    elif mode is 'Simple':
        msg = "Enter words to change slides separated by spaces"
        title = "Cue"
        fieldValue = easygui.enterbox(msg,title)
        return fieldValue.split(' ')

def launchppt():

    msg = "Enter whole or part of the file name you wish to use"
    fieldname = ['File']
    fieldValue = []
    title = "Cue"
    fieldValue = easygui.multenterbox(msg, title, fieldname)
    app = pywinauto.application.Application()
    for window in pywinauto.findwindows.enum_windows():
        temp = win32gui.GetWindowText(window)
        if fieldValue[0] in str(temp):
            handle = pywinauto.findwindows.find_windows(title=temp, class_name=win32gui.GetClassName(window))[0]
            win32gui.SetForegroundWindow(handle)
            pyautogui.hotkey('f5')


def main():
    
    mode = choose_mode()
    #WORD LIST
    cues = get_word_list(mode)
    launchppt() 
    # See http://g.co/cloud/speech/docs/languages
    # for a list of supported languages.
    language_code = 'en-US'  # a BCP-47 language tag

    client = speech.SpeechClient()
    config = types.RecognitionConfig(
        encoding=enums.RecognitionConfig.AudioEncoding.LINEAR16,
        sample_rate_hertz=RATE,
        language_code=language_code)
    streaming_config = types.StreamingRecognitionConfig(
        config=config,
        interim_results=True)

    with MicrophoneStream(RATE, CHUNK) as stream:
        audio_generator = stream.generator()
        requests = (types.StreamingRecognizeRequest(audio_content=content)
                    for content in audio_generator)

        responses = client.streaming_recognize(streaming_config, requests)

       
        listen_for_cue(responses, cues, mode)

if __name__ == '__main__':
    main()