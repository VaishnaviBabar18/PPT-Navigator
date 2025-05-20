# 🖐️🗣️ Hand & Voice Controlled PowerPoint Navigator

A Python-based project that lets you control your PowerPoint presentations **using hand gestures and voice commands**. Ideal for presenters, educators, and individuals seeking hands-free accessibility while delivering presentations.

---

## 🚀 Features

### 🖐️ Hand Gesture Control (via webcam)
- ✋ **All five fingers raised** → Move to **Next Slide**
- 👍 **Only thumb raised** → Move to **Previous Slide**

### 🗣️ Voice Command Control (via microphone)
- Say `"next slide"` → Moves forward
- Say `"previous slide"` → Moves backward
- Say `"zoom in"` or `"zoom out"` → Adjusts the zoom factor *(future-ready)*

### 👀 Real-Time Interface
- Hand tracking with live webcam feed
- Voice commands processed in parallel using threads
- Runs PowerPoint in slideshow mode automatically

---

## 🧠 Technologies Used

| Tool/Library         | Purpose                                      |
|----------------------|----------------------------------------------|
| `OpenCV`             | Video capture and image processing           |
| `cvzone`             | Simplified hand tracking (based on MediaPipe)|
| `speech_recognition` | Captures and recognizes voice input          |
| `pywin32`            | Interfaces with Microsoft PowerPoint         |
| `threading`          | Runs voice and hand control simultaneously   |

---

## 💻 Setup Instructions

### ✅ Prerequisites
- Windows OS with **Microsoft PowerPoint installed**
- Python 3.8+  
- Webcam and Microphone connected

📂 File Setup
Place your PowerPoint file (.pptx) inside the project folder.

Modify this line in the Python script to point to your file:

python
Copy
Edit


### 📦 Install Dependencies

```bash
pip install opencv-python cvzone SpeechRecognition pywin32
📂 File Setup


