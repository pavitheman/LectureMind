import whisper

# Load the Whisper model
model = whisper.load_model("base")

def transcribe(audio_path):
    print(f"Transcribing: {audio_path}")
    result = model.transcribe(audio_path)
    transcript = result["text"]
    print("Done!")
    print(transcript)
    return transcript

# Test it
transcribe("test.mp3")