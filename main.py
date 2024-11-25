import os
import vosk
import pyaudio
import json
from langchain_ollama import OllamaLLM
from langchain_core.prompts import ChatPromptTemplate
import win32com.client as incomer

# Vosk model path
model_path = "C:/Users/COLE/Downloads/vosk-model-small-en-us-0.15"

speak = incomer.Dispatch("SAPI.SpVoice")
# List available voices
voices = speak.GetVoices()
voice_index = 1  # Choose the index of the voice you want (0 for the first voice, 1 for the second, etc.)
speak.Voice = voices.Item(voice_index)
# Load the Vosk model
model = vosk.Model(model_path)

# Set up the audio stream using PyAudio
recognizer = vosk.KaldiRecognizer(model, 16000)

# Initialize PyAudio to capture audio from the microphone
p = pyaudio.PyAudio()
stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=1024)
stream.start_stream()

# Ollama AI model setup
ollama_model = OllamaLLM(model="smollm:360m")
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

                # Generate and print AI response
                ai_response = ollama_response(context, recognized_text)
                speak.Speak(ai_response)
                print(f"Leo: {ai_response}")

                # Update the context with the new conversation history
                context += f"\nUser: {recognized_text}\nLeo: {ai_response}"


print("Listening for the wake word 'Leo'...")

# Main loop to detect the wake word and start conversation
while True:
    try:
        # Read audio data from the microphone
        data = stream.read(1024, exception_on_overflow=False)

        if len(data) == 0:
            break

        if recognizer.AcceptWaveform(data):
            result = recognizer.Result()
            result_dict = json.loads(result)
            recognized_text = result_dict.get('text', '').lower()

            if recognized_text:
                print(f"Sir: {recognized_text}")

                # If the wake word 'Leo' is detected, start the conversation
                if "leo" in recognized_text:
                    start_conversation()

    except OSError as e:
        if e.args[0] == -9981:  # Handle buffer overflow error
            print("Input overflow, skipping...")
            continue
        else:
            raise

# Stop the stream when the script exits
stream.stop_stream()
stream.close()
p.terminate()
