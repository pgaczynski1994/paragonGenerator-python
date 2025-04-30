from flask import Flask, request, jsonify
from google.cloud import vision
import os

app = Flask(__name__)

def ocr_google_vision(image_bytes: bytes) -> str:
    client = vision.ImageAnnotatorClient()
    image = vision.Image(content=image_bytes)
    response = client.text_detection(image=image)
    texts = response.text_annotations
    if not texts:
        return ""
    return texts[0].description

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "Brak pliku"}), 400

    file = request.files['file']
    image_bytes = file.read()
    ocr_text = ocr_google_vision(image_bytes)

    return jsonify({"ocr_text": ocr_text})

if __name__ == "__main__":
    app.run(debug=True)
