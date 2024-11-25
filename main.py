import os
import vosk
import pyaudio
import json
from langchain_ollama import OllamaLLM
from langchain_core.prompts import ChatPromptTemplate
import win32com.client as incomer


model_path = "Paste the path to your vosk model"

speak = incomer.Dispatch("SAPI.SpVoice")

voices = speak.GetVoices()
voice_index = 1  
speak.Voice = voices.Item(voice_index)

model = vosk.Model(model_path)


recognizer = vosk.KaldiRecognizer(model, 16000)


p = pyaudio.PyAudio()
stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=1024)
stream.start_stream()


ollama_model = OllamaLLM(model="model-name")
template = '''
Answer the question below.

Here is the conversation history: {context}

Question: {question}

Answer:
'''
prompt = ChatPromptTemplate.from_template(template)


def ollama_response(context, user_input):
    """Generate response from Ollama AI model based on conversation history."""
    result = ollama_model.invoke(context=context, input=user_input)
    return result


def start_conversation():
    """Handles conversation with Ollama after detecting the wake word."""
    context = ""
    speak.Speak("I'm listening. Say 'bye' to end the conversation.")
    print("Leo: I'm listening. Say 'bye' to end the conversation.")

    while True:
        # Read audio data from the microphone
        data = stream.read(1024, exception_on_overflow=False)

        if len(data) == 0:
            break

        if recognizer.AcceptWaveform(data):
            result = recognizer.Result()
            result_dict = json.loads(result)
            recognized_text = result_dict.get('text', '').lower()

            if recognized_text:
                print(f"Recognized: {recognized_text}")

                if "bye" in recognized_text or "exit" in recognized_text:
                    speak.Speak("Goodbye!")
                    print("Leo: Goodbye!")
                    break

            
                ai_response = ollama_response(context, recognized_text)
                speak.Speak(ai_response)
                print(f"Leo: {ai_response}")

       
                context += f"\nUser: {recognized_text}\nLeo: {ai_response}"


print("Listening for the wake word 'Leo'...")


while True:
    try:

        data = stream.read(1024, exception_on_overflow=False)

        if len(data) == 0:
            break

        if recognizer.AcceptWaveform(data):
            result = recognizer.Result()
            result_dict = json.loads(result)
            recognized_text = result_dict.get('text', '').lower()

            if recognized_text:
                print(f"Sir: {recognized_text}")

               
                if "leo" in recognized_text:
                    start_conversation()

    except OSError as e:
        if e.args[0] == -9981:  
            print("Input overflow, skipping...")
            continue
        else:
            raise


stream.stop_stream()
stream.close()
p.terminate()
