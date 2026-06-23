# Face Detector Examination System

This project is a Python-based examination monitoring system. It uses the webcam to watch for face-related exam violations, such as no face detected, suspicious face movement, gaze changes, and hand activity. When a violation is detected, the system can save screenshot evidence locally.

## Features

- Detects whether a face is visible during an exam
- Tracks face position and head/gaze direction
- Detects hand activity using MediaPipe
- Shows live status messages on the camera window
- Plays alert sounds or voice warnings on supported Windows systems
- Saves evidence screenshots in a local ignored folder

## Project Structure

```text
exam detection/
  TheFaceDetector.py      Main Python script
  exam_evidence/          Generated evidence screenshots, ignored by Git
```

## Libraries Used

### OpenCV

`opencv-python` is used for webcam access, live video display, face detection, eye detection, drawing boxes/text on the frame, and saving screenshot evidence.

### MediaPipe

`mediapipe` is used for more advanced face and hand tracking. It helps detect facial landmarks, estimate gaze/head movement, and identify hand gestures in the camera frame.

### NumPy

`numpy` is used for math calculations, landmark positions, distances, and angle estimation.

### winsound

`winsound` is a built-in Windows library used to play beep alerts. It does not need to be installed separately.

### pywin32

`pywin32` is optional. It allows the system to use Windows text-to-speech voice alerts. If it is not installed, the program can still run with beep alerts.

## Installation

Install the required Python libraries:

```bash
pip install opencv-python numpy mediapipe pywin32
```

If you do not need voice alerts, `pywin32` can be skipped:

```bash
pip install opencv-python numpy mediapipe
```

## How To Run

From the project folder, run:

```bash
python "exam detection/TheFaceDetector.py"
```

The webcam window will open and start monitoring. Press `q` to close the program.

## Evidence Photos

Violation screenshots are saved inside:

```text
exam detection/exam_evidence/
```

This folder is ignored by Git because evidence photos may contain private exam or face data. The code is shared on GitHub, but generated evidence images are kept only on your local computer.

## Notes

- A working webcam is required.
- The system works best in good lighting.
- Windows is recommended because sound and voice alerts use Windows-specific libraries.
- This project is intended as an exam monitoring helper, not a complete replacement for human supervision.
