import whisper
import ollama

# Load the Whisper model
model = whisper.load_model("base")

def transcribe(audio_path):
    print(f"Transcribing: {audio_path}")
    result = model.transcribe(audio_path)
    transcript = result["text"]
    print("Done!")
    return transcript

def generate_flashcards(transcript):
    print("Generating flashcards and summary...")
    response = ollama.chat(
        model="llama3.2",
        messages=[
            {
                "role": "user",
                "content": f"Here is a lecture transcript. First give a 2-3 sentence summary. Then generate 3 flashcards in this exact format:\nFront: [question]\nBack: [answer]\n\nTranscript: {transcript}"
            }
        ]
    )
    result = response["message"]["content"]
    print("Done!")
    return result

# Test it
transcript = transcribe("test.mp3")
print(generate_flashcards(transcript))