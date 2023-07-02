# Basic_Assistant:
This has been an attempt to create a basic version of an Assistant which has the ability to take in an input from a user in the form of voice, process the audio, get a text transcription and process this text throught the GPT 3.5 transformer to provide us with an answer for our questions. There is a provision to listen to the output of the response as well.
## Installing libraries:
The following libraries are required to be installed:
1. The OpenAI Whisper API:
   - It is an API tool which has been developed by OpenAI in order to make speech to text translations easy. It consists of an Automatic           Speech Recognition model which has been trained on multilingual data to make translations much effective.
   - Installed as `!pip install -U openai-whisper`
2. Gradio:
   - It is an open source python library which is used to create User Interfaces for Machine Learning projects and further create APIs           for the project so that they could be hosted online effectively.
   - Installed as `!pip install gradio`
3. OpenAI:
   - It is an open source library which has been used to gain access to the GPT transformer which will be able to process our queries and        statements and hence provide us with a suitable response.
   - Installed as `!pip install openai`
4. TTS:
   - It is an advanced library which uses DeepLearning Models for tet to speech conversions.
   - Installed as `!pip install TTS`
## Importing Necessary Libraries:
Using the above mentioned methods, we can install the necessary libraries. Now we can import them using the `import` function.

`import whisper`

`import gradio as gr`

`import openai`

`from TTS.api import TTS`

`import warnings`

`warnings.filterwarnings('ignore')`

# The major components of the project: 
## The Text to Speech Part:
The following the method to use the TTS library. Here we have used the `tts_models/en/ljspeech/tacotron2-DDC_ph` model from the TTS library in order to perform the text to speech conversion. It is an english language model. The models loaded here are passed later into the Transcription function in order to make the conversion of text in audio possible.

## The Whisper part
The whisper part is created in order to pass the model which will be used for the Speech to Text conversion. Tiny.en is an english language module which has been trained on multiple accents and languages. The module loaded here will be passed into the transcribe function in order to convert speech into text.
Furthermore the API key is passed using a .json file which has the API key stored in the following format:
{"OPENAI_API_KEY":"Your API Key here"}
## The ChatGPT Function:
This is the major function where the queries are proceesed and the output is generated for the statement or query you have presented.

## The Transcription Function
This is the function which converts audio to text and text to audio as required.

### The following are the outputs generated: 
### 1-Your Speech converted to text
`output_1 = gr.Textbox(label="YOUR QUESTION: ")`
### 2- Your Query answered in the text format
`output_2 = gr.Textbox(label="ANSWER(TEXT): ")`
### 3- The answer presented in the audio format
`output_3 = gr.Audio(label="ANSWER(AUDIO): ", upload="output.wav")`

## THE GRADIO INTERFACE:
The part which creates the user interface which is visible to the user and allows the hosting of your project online.
