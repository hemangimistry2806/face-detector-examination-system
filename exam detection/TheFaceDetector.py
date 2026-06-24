import cv2
import os
import time
import threading
import wave

import numpy as np

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "2")

try:
    import winsound
except ImportError:
    winsound = None

try:
    import win32com.client
except ImportError:
    win32com = None

try:
    import mediapipe as mp
except Exception:
    mp = None

try:
    import sounddevice as sd
except Exception:
    sd = None

mp_face_mesh = mp.solutions.face_mesh if mp else None
try:
    from mediapipe.tasks import vision
    from mediapipe.framework.formats import landmark_pb2
    iris_detector_available = True
except:
    iris_detector_available = False

# Configuration
CAMERA_WIDTH = 560
CAMERA_HEIGHT = 420
DISPLAY_WIDTH = 960
DISPLAY_HEIGHT = 720
ALARM_INTERVAL = 2.5
MIRROR_CAMERA = True
MIN_FACE_AREA_RATIO = 0.035
MAX_FACE_AREA_RATIO = 0.55
MIN_FACE_CONFIDENCE = 1.5
STRONG_FACE_CONFIDENCE = 4.0
GAZE_THRESHOLD = 25.0
EYE_GAZE_THRESHOLD = 0.18
GAZE_VIOLATION_DURATION = 5.0
GAZE_SMOOTHING_ALPHA = 0.25
FACE_MISSING_SCREENSHOT_DELAY = 2.0
SCREENSHOT_COOLDOWN = 15.0
TALKING_VIOLATION_DURATION = 3.0
MOUTH_OPEN_THRESHOLD = 0.08
MOUTH_SMOOTHING_ALPHA = 0.35
AUDIO_RECORD_SECONDS = 10
AUDIO_SAMPLE_RATE = 16000

voice_alarm_running = False
last_gaze_screenshot_time = 0
gaze_violation_start_time = None
gaze_violation_screenshot_taken = False
smoothed_gaze_score = 0.0
sustained_gaze_direction = "looking center"
face_missing_start_time = None
face_missing_screenshot_taken = False
multiple_faces_screenshot_taken = False
talking_start_time = None
talking_screenshot_taken = False
talking_audio_taken = False
smoothed_mouth_open_ratio = 0.0
last_screenshot_by_type = {}
audio_recording_running = False
SCREENSHOT_FOLDER = os.path.join(os.path.dirname(__file__), "exam_evidence")

if not os.path.exists(SCREENSHOT_FOLDER):
    os.makedirs(SCREENSHOT_FOLDER)

hands_available = mp is not None
mp_hands = mp.solutions.hands if mp else None
mp_drawing = mp.solutions.drawing_utils if mp else None
mp_face_mesh_available = mp_face_mesh is not None
hand_detector = None
face_mesh_detector = None


def save_violation_screenshot(frame, violation_type, details):
    """Save screenshot evidence when violations occur."""
    global last_gaze_screenshot_time
    
    current_time = time.time()
    screenshot_key = f"{violation_type}:{details}"
    last_screenshot_time = last_screenshot_by_type.get(screenshot_key, 0)
    if current_time - last_screenshot_time < SCREENSHOT_COOLDOWN:
        return
    
    last_gaze_screenshot_time = current_time
    last_screenshot_by_type[screenshot_key] = current_time
    
    try:
        timestamp = time.strftime("%Y%m%d_%H%M%S_") + f"{int((current_time % 1) * 1000):03d}"
        filename = f"{violation_type}_{details}_{timestamp}.jpg"
        filepath = os.path.join(SCREENSHOT_FOLDER, filename)
        
        # Add timestamp text on the screenshot
        screenshot = frame.copy()
        cv2.putText(
            screenshot,
            f"Violation: {violation_type} - {time.strftime('%Y-%m-%d %H:%M:%S')}",
            (10, 30),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.8,
            (0, 0, 255),
            2,
        )
        cv2.putText(
            screenshot,
            f"Details: {details}",
            (10, 65),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.7,
            (0, 0, 255),
            2,
        )
        
        cv2.imwrite(filepath, screenshot)
        print(f"[EVIDENCE] Screenshot saved: {filename}")
    except Exception as e:
        print(f"Error saving screenshot: {e}")


def record_audio_evidence(details, duration=AUDIO_RECORD_SECONDS):
    """Save a short microphone recording for a talking violation."""
    global audio_recording_running

    if sd is None or audio_recording_running:
        if sd is None:
            print("[AUDIO] sounddevice not installed. Run: pip install sounddevice")
        return

    audio_recording_running = True

    def _record():
        global audio_recording_running
        try:
            timestamp = time.strftime("%Y%m%d_%H%M%S_") + f"{int((time.time() % 1) * 1000):03d}"
            filename = f"talking_audio_{details}_{timestamp}.wav"
            filepath = os.path.join(SCREENSHOT_FOLDER, filename)

            print(f"[AUDIO] Recording {duration} seconds: {filename}")
            audio = sd.rec(
                int(duration * AUDIO_SAMPLE_RATE),
                samplerate=AUDIO_SAMPLE_RATE,
                channels=1,
                dtype="int16",
            )
            sd.wait()

            with wave.open(filepath, "wb") as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(AUDIO_SAMPLE_RATE)
                wav_file.writeframes(audio.tobytes())

            print(f"[AUDIO] Evidence saved: {filename}")
        except Exception as e:
            print(f"Error recording audio evidence: {e}")
        finally:
            audio_recording_running = False

    threading.Thread(target=_record, daemon=True).start()


def simple_alarm():
    if winsound is None:
        return
    winsound.Beep(1200, 250)


def voice_alarm(message):
    """Play a strong no-face alert without freezing the camera window."""
    global voice_alarm_running

    if voice_alarm_running:
        return

    voice_alarm_running = True

    def _play():
        global voice_alarm_running
        try:
            if winsound is not None:
                for frequency in (1300, 1700, 1300):
                    winsound.Beep(frequency, 220)

            if win32com is not None:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                speaker.Volume = 100
                speaker.Rate = 1
                speaker.Speak(message)
            elif winsound is not None:
                for _ in range(4):
                    winsound.Beep(1500, 300)
        finally:
            voice_alarm_running = False

    threading.Thread(target=_play, daemon=True).start()


def draw_error_banner(frame, message):
    height, width = frame.shape[:2]
    cv2.rectangle(frame, (0, height - 78), (width, height), (0, 0, 255), -1)
    cv2.putText(
        frame,
        message,
        (12, height - 30),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.8,
        (255, 255, 255),
        2,
    )


def box_iou(first, second):
    x1, y1, w1, h1 = first
    x2, y2, w2, h2 = second

    left = max(x1, x2)
    top = max(y1, y2)
    right = min(x1 + w1, x2 + w2)
    bottom = min(y1 + h1, y2 + h2)

    if right <= left or bottom <= top:
        return 0.0

    intersection = (right - left) * (bottom - top)
    area1 = w1 * h1
    area2 = w2 * h2
    return intersection / float(area1 + area2 - intersection)


def draw_eye_detections(frame, eye_cascade, gray, face_box):
    x, y, w, h = face_box
    upper_face = gray[y : y + int(h * 0.65), x : x + w]
    eyes = eye_cascade.detectMultiScale(
        upper_face,
        scaleFactor=1.08,
        minNeighbors=5,
        minSize=(max(18, w // 8), max(18, h // 8)),
    )
    for ex, ey, ew, eh in eyes:
        cv2.rectangle(frame, (x + ex, y + ey), (x + ex + ew, y + ey + eh), (255, 255, 0), 1)
    return list(eyes)


def estimate_head_rotation(face_landmarks, image_width, image_height):
    """Estimate where user is looking (left/right/center) and the horizontal gaze angle."""
    if face_landmarks is None:
        return None

    def pt(index):
        lm = face_landmarks.landmark[index]
        return np.array([lm.x * image_width, lm.y * image_height], dtype=np.float32)

    # Get facial landmarks
    left_eye = pt(33)
    right_eye = pt(263)
    nose_tip = pt(1)
    left_mouth = pt(61)
    right_mouth = pt(291)

    # Calculate face center (horizontal)
    face_center_x = (left_eye[0] + right_eye[0]) / 2.0
    nose_x = nose_tip[0]

    # Horizontal gaze angle: positive = looking right, negative = looking left
    # Based on where nose is relative to eyes center
    eye_distance = right_eye[0] - left_eye[0]
    if eye_distance > 0:
        gaze_angle = (nose_x - face_center_x) / (eye_distance / 2.0) * 30.0
    else:
        gaze_angle = 0.0

    # Clamp angle
    gaze_angle = np.clip(gaze_angle, -90, 90)

    # Determine direction
    if gaze_angle > GAZE_THRESHOLD:
        direction = "looking right"
    elif gaze_angle < -GAZE_THRESHOLD:
        direction = "looking left"
    else:
        direction = "looking center"

    return {"angle": gaze_angle, "direction": direction}


def face_box_from_landmarks(face_landmarks, image_width, image_height):
    if face_landmarks is None:
        return None

    xs = [lm.x * image_width for lm in face_landmarks.landmark]
    ys = [lm.y * image_height for lm in face_landmarks.landmark]
    x1 = max(0, int(min(xs)))
    y1 = max(0, int(min(ys)))
    x2 = min(image_width, int(max(xs)))
    y2 = min(image_height, int(max(ys)))

    if x2 <= x1 or y2 <= y1:
        return None

    padding_x = int((x2 - x1) * 0.12)
    padding_y = int((y2 - y1) * 0.18)
    x1 = max(0, x1 - padding_x)
    y1 = max(0, y1 - padding_y)
    x2 = min(image_width, x2 + padding_x)
    y2 = min(image_height, y2 + padding_y)
    return (x1, y1, x2 - x1, y2 - y1, True)


def detect_talking(face_landmarks, image_width, image_height):
    """Estimate talking from sustained mouth opening."""
    global smoothed_mouth_open_ratio

    if face_landmarks is None:
        smoothed_mouth_open_ratio = 0.0
        return False, 0.0

    def pt(index):
        lm = face_landmarks.landmark[index]
        return np.array([lm.x * image_width, lm.y * image_height], dtype=np.float32)

    upper_lip = pt(13)
    lower_lip = pt(14)
    left_mouth = pt(61)
    right_mouth = pt(291)

    mouth_width = np.linalg.norm(right_mouth - left_mouth)
    if mouth_width <= 1:
        return False, smoothed_mouth_open_ratio

    mouth_open_ratio = np.linalg.norm(lower_lip - upper_lip) / mouth_width
    smoothed_mouth_open_ratio = (
        MOUTH_SMOOTHING_ALPHA * mouth_open_ratio
        + (1.0 - MOUTH_SMOOTHING_ALPHA) * smoothed_mouth_open_ratio
    )
    return smoothed_mouth_open_ratio >= MOUTH_OPEN_THRESHOLD, smoothed_mouth_open_ratio


def detect_eye_gaze(frame, face_landmarks, face_box):
    """Detect eye direction from MediaPipe iris landmarks with a pupil fallback."""
    if face_landmarks is None:
        return None

    def pt(index):
        lm = face_landmarks.landmark[index]
        return np.array([lm.x * frame.shape[1], lm.y * frame.shape[0]], dtype=np.float32)

    left_eye_left = pt(33)
    left_eye_right = pt(133)
    left_eye_top = pt(160)
    left_eye_bottom = pt(158)

    right_eye_left = pt(263)
    right_eye_right = pt(362)
    right_eye_top = pt(387)
    right_eye_bottom = pt(385)

    def direction_from_ratio(ratio):
        if ratio < 0.5 - EYE_GAZE_THRESHOLD:
            return "left"
        if ratio > 0.5 + EYE_GAZE_THRESHOLD:
            return "right"
        return "center"

    def iris_ratio(iris_indexes, eye_left, eye_right):
        if len(face_landmarks.landmark) <= max(iris_indexes):
            return None

        iris_center = np.mean([pt(index) for index in iris_indexes], axis=0)
        eye_vector = eye_right - eye_left
        eye_width = np.linalg.norm(eye_vector)
        if eye_width <= 1:
            return None

        ratio = np.dot(iris_center - eye_left, eye_vector) / (eye_width * eye_width)
        return float(np.clip(ratio, 0.0, 1.0))

    def get_pupil_direction(frame, eye_left, eye_right, eye_top, eye_bottom):
        """Find pupil position within eye region to determine gaze direction."""
        padding = 4
        x1 = max(0, int(min(eye_left[0], eye_right[0])) - padding)
        x2 = min(frame.shape[1], int(max(eye_left[0], eye_right[0])) + padding)
        y1 = max(0, int(min(eye_top[1], eye_bottom[1])) - padding)
        y2 = min(frame.shape[0], int(max(eye_top[1], eye_bottom[1])) + padding)

        if x2 <= x1 or y2 <= y1:
            return None

        eye_region = frame[y1:y2, x1:x2]
        gray_eye = cv2.cvtColor(eye_region, cv2.COLOR_BGR2GRAY) if len(eye_region.shape) == 3 else eye_region

        gray_eye = cv2.GaussianBlur(gray_eye, (5, 5), 0)
        _, _, min_loc, _ = cv2.minMaxLoc(gray_eye)
        pupil_x = min_loc[0]
        eye_width = x2 - x1
        pupil_ratio = pupil_x / max(eye_width, 1)

        return {
            "direction": direction_from_ratio(pupil_ratio),
            "ratio": float(pupil_ratio),
            "center": (x1 + pupil_x, y1 + min_loc[1]),
            "source": "pupil",
        }

    try:
        left_ratio = iris_ratio([468, 469, 470, 471, 472], left_eye_left, left_eye_right)
        right_ratio = iris_ratio([473, 474, 475, 476, 477], right_eye_left, right_eye_right)

        left_eye = None
        right_eye = None
        if left_ratio is not None:
            left_eye = {"direction": direction_from_ratio(left_ratio), "ratio": left_ratio, "source": "iris"}
        if right_ratio is not None:
            right_eye = {"direction": direction_from_ratio(right_ratio), "ratio": right_ratio, "source": "iris"}

        if left_eye is None:
            left_eye = get_pupil_direction(frame, left_eye_left, left_eye_right, left_eye_top, left_eye_bottom)
        if right_eye is None:
            right_eye = get_pupil_direction(frame, right_eye_left, right_eye_right, right_eye_top, right_eye_bottom)

        if left_eye is None or right_eye is None:
            return None

        average_ratio = (left_eye["ratio"] + right_eye["ratio"]) / 2.0
        combined_direction = direction_from_ratio(average_ratio)
        return {
            "left": left_eye["direction"],
            "right": right_eye["direction"],
            "direction": combined_direction,
            "ratio": average_ratio,
            "source": "iris" if left_eye["source"] == "iris" and right_eye["source"] == "iris" else "pupil",
        }
    except Exception as e:
        print(f"Error in eye gaze detection: {e}")
        return None


def smooth_gaze_state(pose, eye_gaze):
    """Blend head pose and eye movement so brief noisy frames do not trigger alerts."""
    global smoothed_gaze_score, sustained_gaze_direction

    raw_score = 0.0
    raw_direction = "looking center"

    if eye_gaze is not None and eye_gaze["direction"] != "center":
        eye_offset = abs(eye_gaze["ratio"] - 0.5)
        raw_score = max(raw_score, min(1.0, eye_offset / 0.35))
        raw_direction = f"looking {eye_gaze['direction']}"

    if pose is not None and pose["direction"] != "looking center":
        head_score = min(1.0, abs(pose["angle"]) / 45.0)
        if head_score > raw_score:
            raw_direction = pose["direction"]
        raw_score = max(raw_score, head_score)

    smoothed_gaze_score = (
        GAZE_SMOOTHING_ALPHA * raw_score
        + (1.0 - GAZE_SMOOTHING_ALPHA) * smoothed_gaze_score
    )

    if smoothed_gaze_score >= 0.55:
        sustained_gaze_direction = raw_direction
        return False, sustained_gaze_direction, smoothed_gaze_score

    if smoothed_gaze_score <= 0.35:
        sustained_gaze_direction = "looking center"
        return True, sustained_gaze_direction, smoothed_gaze_score

    return sustained_gaze_direction == "looking center", sustained_gaze_direction, smoothed_gaze_score


def initialize_face_mesh_detector():
    global face_mesh_detector
    if not mp_face_mesh_available:
        return
    face_mesh_detector = mp_face_mesh.FaceMesh(
        static_image_mode=False,
        max_num_faces=2,
        refine_landmarks=True,
        min_detection_confidence=0.5,
        min_tracking_confidence=0.5,
    )


def close_face_mesh_detector():
    global face_mesh_detector
    if face_mesh_detector is not None:
        face_mesh_detector.close()
        face_mesh_detector = None


def detect_face_candidates(face_cascade, gray):
    try:
        faces, _, weights = face_cascade.detectMultiScale3(
            gray,
            scaleFactor=1.08,
            minNeighbors=8,
            minSize=(100, 100),
            outputRejectLevels=True,
        )
        return list(faces), [float(weight) for weight in weights]
    except Exception:
        faces = face_cascade.detectMultiScale(
            gray,
            scaleFactor=1.08,
            minNeighbors=8,
            minSize=(100, 100),
        )
        return list(faces), [STRONG_FACE_CONFIDENCE for _ in faces]


def has_face_like_eyes(eye_cascade, gray, face_box):
    x, y, w, h = face_box
    upper_face = gray[y : y + int(h * 0.65), x : x + w]
    eyes = eye_cascade.detectMultiScale(
        upper_face,
        scaleFactor=1.08,
        minNeighbors=5,
        minSize=(max(18, w // 8), max(18, h // 8)),
    )
    return len(eyes) > 0


def filter_face_detections(face_cascade, eye_cascade, gray, frame_shape):
    raw_faces, weights = detect_face_candidates(face_cascade, gray)
    frame_h, frame_w = frame_shape[:2]
    frame_area = frame_w * frame_h
    candidates = []

    for face_box, weight in zip(raw_faces, weights):
        x, y, w, h = [int(value) for value in face_box]
        area_ratio = (w * h) / float(frame_area)
        aspect_ratio = w / float(h)

        if weight < MIN_FACE_CONFIDENCE:
            continue
        if area_ratio < MIN_FACE_AREA_RATIO or area_ratio > MAX_FACE_AREA_RATIO:
            continue
        if aspect_ratio < 0.72 or aspect_ratio > 1.35:
            continue

        eye_confirmed = has_face_like_eyes(eye_cascade, gray, (x, y, w, h))
        if not eye_confirmed and weight < STRONG_FACE_CONFIDENCE:
            continue

        candidates.append(((x, y, w, h), weight, eye_confirmed))

    candidates.sort(key=lambda item: (item[2], item[1], item[0][2] * item[0][3]), reverse=True)
    filtered_faces = []

    for face_box, _, eye_confirmed in candidates:
        if all(box_iou(face_box, kept_box[:4]) < 0.35 for kept_box in filtered_faces):
            filtered_faces.append((*face_box, eye_confirmed))

    return filtered_faces


def initialize_hand_detector():
    global hand_detector
    if not hands_available:
        return
    hand_detector = mp_hands.Hands(
        static_image_mode=False,
        max_num_hands=2,
        model_complexity=1,
        min_detection_confidence=0.35,
        min_tracking_confidence=0.35,
    )


def close_hand_detector():
    global hand_detector
    if hand_detector is not None:
        hand_detector.close()
        hand_detector = None


def classify_hand_gesture(hand_landmarks):
    lm = hand_landmarks.landmark

    def distance(first, second):
        dx = lm[first].x - lm[second].x
        dy = lm[first].y - lm[second].y
        return (dx * dx + dy * dy) ** 0.5

    # Thumb is considered open when its tip moves away from the index finger base.
    thumb_extended = distance(4, 5) > distance(3, 5) * 1.15
    thumb_vertical = abs(lm[4].y - lm[2].y) > abs(lm[4].x - lm[2].x) * 1.25

    finger_extended = [
        lm[8].y < lm[6].y and lm[7].y < lm[6].y,
        lm[12].y < lm[10].y and lm[11].y < lm[10].y,
        lm[16].y < lm[14].y and lm[15].y < lm[14].y,
        lm[20].y < lm[18].y and lm[19].y < lm[18].y,
    ]

    extended_count = sum(finger_extended) + int(thumb_extended)
    if extended_count == 5:
        return "open hand"
    if extended_count <= 1 and not any(finger_extended):
        return "fist"
    if finger_extended[0] and not any(finger_extended[1:]) and not thumb_extended:
        return "point"
    if finger_extended[0] and finger_extended[1] and not any(finger_extended[2:]) and not thumb_extended:
        return "peace"
    if finger_extended[1] and not finger_extended[0] and not finger_extended[2] and not finger_extended[3]:
        return "middle finger"
    if thumb_extended and thumb_vertical and not any(finger_extended):
        if lm[4].y < lm[2].y:
            return "thumbs up"
        return "thumbs down"
    if thumb_extended and not any(finger_extended):
        return "thumbs up"
    return "hand"


def main():
    global hand_detector, gaze_violation_start_time, gaze_violation_screenshot_taken
    global smoothed_gaze_score, sustained_gaze_direction
    global face_missing_start_time, face_missing_screenshot_taken, multiple_faces_screenshot_taken
    global talking_start_time, talking_screenshot_taken, talking_audio_taken, smoothed_mouth_open_ratio
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    eye_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_eye.xml")
    if face_cascade.empty():
        print("ERROR: Face cascade file could not be loaded.")
        return
    if eye_cascade.empty():
        print("ERROR: Eye cascade file could not be loaded.")
        return

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("ERROR: Could not open camera. If running in a headless environment, camera isn't available.")
        return

    cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)

    cv2.namedWindow("Exam Face Detector", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Exam Face Detector", DISPLAY_WIDTH, DISPLAY_HEIGHT)

    initialize_hand_detector()
    initialize_face_mesh_detector()
    last_alarm_time = 0

    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                print("Camera frame not available")
                break
            if MIRROR_CAMERA:
                frame = cv2.flip(frame, 1)

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            gray = cv2.equalizeHist(gray)
            faces = filter_face_detections(face_cascade, eye_cascade, gray, frame.shape)
            current_time = time.time()

            rgb_frame = None
            mesh_face_landmarks = []
            if mp_face_mesh_available and face_mesh_detector is not None:
                rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                rgb_frame.flags.writeable = False
                mesh_results = face_mesh_detector.process(rgb_frame)
                rgb_frame.flags.writeable = True
                mesh_face_landmarks = list(getattr(mesh_results, 'multi_face_landmarks', None) or [])

                if mesh_face_landmarks and len(faces) < len(mesh_face_landmarks):
                    existing_boxes = [face[:4] for face in faces]
                    for face_landmarks in mesh_face_landmarks:
                        mesh_box = face_box_from_landmarks(face_landmarks, frame.shape[1], frame.shape[0])
                        if mesh_box is not None and all(box_iou(mesh_box[:4], box) < 0.25 for box in existing_boxes):
                            faces.append(mesh_box)
                            existing_boxes.append(mesh_box[:4])

            face_count = len(faces)
            alarm_needed = False
            no_face_alarm = False
            alarm_message = None

            if face_count == 0:
                status_text = "ERROR: No face detected"
                status_color = (0, 0, 255)
                alarm_needed = True
                no_face_alarm = True
                alarm_message = "Warning. Face not detected. Please stay in front of the camera."
                if face_missing_start_time is None:
                    face_missing_start_time = current_time
                if (
                    current_time - face_missing_start_time >= FACE_MISSING_SCREENSHOT_DELAY
                    and not face_missing_screenshot_taken
                ):
                    save_violation_screenshot(frame, "face_violation", "no_face_detected")
                    face_missing_screenshot_taken = True
            elif face_count > 1:
                status_text = f"ERROR: Multiple faces detected ({face_count})"
                status_color = (0, 0, 255)
                alarm_needed = True
                alarm_message = "Warning. Multiple faces detected. Only one person is allowed."
                face_missing_start_time = None
                face_missing_screenshot_taken = False
                if not multiple_faces_screenshot_taken:
                    save_violation_screenshot(frame, "face_violation", f"multiple_faces_{face_count}")
                    multiple_faces_screenshot_taken = True
            else:
                status_text = "One face detected"
                status_color = (0, 255, 0)
                face_missing_start_time = None
                face_missing_screenshot_taken = False
                multiple_faces_screenshot_taken = False

            face_poses = []
            eye_gazes = []
            gaze_violation = False
            gaze_message = None
            current_gaze_direction = "looking center"
            current_gaze_score = 0.0
            talking_detected = False
            talking_violation = False
            talking_ratio = smoothed_mouth_open_ratio
            
            if mesh_face_landmarks:
                    for idx, face_landmarks in enumerate(mesh_face_landmarks):
                        pose = estimate_head_rotation(face_landmarks, frame.shape[1], frame.shape[0])
                        
                        if pose is not None:
                            face_poses.append(pose)
                        
                        # Use corresponding face box if available
                        if idx < len(faces):
                            face_box = faces[idx][:4]
                            eye_gaze = detect_eye_gaze(frame, face_landmarks, face_box)
                            if eye_gaze is not None:
                                eye_gazes.append(eye_gaze)
                        
                        if pose is not None and face_count == 1:
                            eye_gaze = eye_gazes[-1] if eye_gazes else None
                            looking_center, current_gaze_direction, current_gaze_score = smooth_gaze_state(pose, eye_gaze)

                            if not looking_center:
                                # Start timer if not already started
                                if gaze_violation_start_time is None:
                                    gaze_violation_start_time = current_time
                                
                                # Check if user has been looking away for GAZE_VIOLATION_DURATION seconds
                                time_looking_away = current_time - gaze_violation_start_time
                                if time_looking_away >= GAZE_VIOLATION_DURATION:
                                    gaze_violation = True
                                    gaze_message = "Warning. Please focus on your paper or exam."
                                    if not gaze_violation_screenshot_taken:
                                        screenshot_detail = current_gaze_direction.replace(" ", "_")
                                        save_violation_screenshot(frame, "gaze_violation", screenshot_detail)
                                        gaze_violation_screenshot_taken = True
                            else:
                                # User is looking center, reset timer
                                gaze_violation_start_time = None
                                gaze_violation_screenshot_taken = False

                            talking_detected, talking_ratio = detect_talking(face_landmarks, frame.shape[1], frame.shape[0])
                            if talking_detected:
                                if talking_start_time is None:
                                    talking_start_time = current_time
                                if current_time - talking_start_time >= TALKING_VIOLATION_DURATION:
                                    talking_violation = True
                                    if not talking_screenshot_taken:
                                        save_violation_screenshot(frame, "talking_violation", "mouth_open_or_talking")
                                        talking_screenshot_taken = True
                                    if not talking_audio_taken:
                                        record_audio_evidence("mouth_open_or_talking")
                                        talking_audio_taken = True
                            else:
                                talking_start_time = None
                                talking_screenshot_taken = False
                                talking_audio_taken = False
                        elif face_count != 1:
                            gaze_violation_start_time = None
                            gaze_violation_screenshot_taken = False
                            smoothed_gaze_score = 0.0
                            sustained_gaze_direction = "looking center"
            else:
                gaze_violation_start_time = None
                gaze_violation_screenshot_taken = False
                talking_start_time = None
                talking_screenshot_taken = False
                talking_audio_taken = False
                smoothed_gaze_score = 0.0
                sustained_gaze_direction = "looking center"
                smoothed_mouth_open_ratio = 0.0

            for idx, face_info in enumerate(faces):
                x, y, w, h, eye_confirmed = face_info
                box_color = (0, 255, 0) if face_count == 1 else (0, 0, 255)
                cv2.rectangle(frame, (x, y), (x + w, y + h), box_color, 2)
                draw_eye_detections(frame, eye_cascade, gray, (x, y, w, h))
                
                # Display head gaze direction
                if idx < len(face_poses):
                    pose = face_poses[idx]
                    gaze_text_color = (0, 0, 255) if gaze_violation else (0, 255, 255)
                    cv2.putText(
                        frame,
                        f"Head Gaze: {pose['direction']}",
                        (x, y - 10),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.7,
                        gaze_text_color,
                        2,
                    )
                
                # Display eye gaze direction
                if idx < len(eye_gazes):
                    eye_gaze = eye_gazes[idx]
                    eye_text = f"L:{eye_gaze['left']} R:{eye_gaze['right']} ({eye_gaze['source']})"
                    cv2.putText(
                        frame,
                        f"Eyes: {eye_text}",
                        (x, y + h + 25),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.6,
                        (255, 165, 0),
                        2,
                    )
                    if face_count == 1:
                        cv2.putText(
                            frame,
                            f"Focus score: {current_gaze_score:.2f}",
                            (x, y + h + 50),
                            cv2.FONT_HERSHEY_SIMPLEX,
                            0.6,
                            (0, 255, 255) if current_gaze_score < 0.55 else (0, 0, 255),
                            2,
                        )

            hand_gesture_summary = "hands unavailable"
            gesture_violation = False
            if hands_available and hand_detector is not None:
                rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                rgb_frame.flags.writeable = False
                results = hand_detector.process(rgb_frame)
                rgb_frame.flags.writeable = True
                gestures = []
                if getattr(results, 'multi_hand_landmarks', None):
                    for hand_landmarks in results.multi_hand_landmarks:
                        gesture = classify_hand_gesture(hand_landmarks)
                        gestures.append(gesture)
                        try:
                            mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                        except Exception:
                            pass
                hand_gesture_summary = ", ".join(gestures) if gestures else "no hands"
                gesture_violation = "middle finger" in gestures

            if gaze_violation:
                status_text = "ERROR: Focus away"
                status_color = (0, 0, 255)
                alarm_needed = True
                no_face_alarm = False
                alarm_message = gaze_message
            elif talking_violation:
                status_text = "ERROR: Talking detected"
                status_color = (0, 0, 255)
                alarm_needed = True
                no_face_alarm = False
                alarm_message = "Warning. Talking detected. Please stay silent during the exam."
            elif gesture_violation:
                status_text = "ERROR: Conduct violation"
                status_color = (0, 0, 255)
                alarm_needed = True
                no_face_alarm = False
                alarm_message = "Warning. Conduct violation detected. Inappropriate gesture."
                save_violation_screenshot(frame, "conduct_violation", "inappropriate_gesture")

            cv2.putText(frame, status_text, (10, 32), cv2.FONT_HERSHEY_SIMPLEX, 0.8, status_color, 2)
            cv2.putText(frame, f"Hand gestures: {hand_gesture_summary}", (10, 64), cv2.FONT_HERSHEY_SIMPLEX, 0.7,
                        (0, 255, 255) if hand_gesture_summary != "no hands" else (255, 255, 0), 2)
            talking_text = "talking" if talking_detected else "silent"
            cv2.putText(
                frame,
                f"Talking: {talking_text} ({talking_ratio:.2f})",
                (10, 96),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.7,
                (0, 0, 255) if talking_violation else (255, 255, 0),
                2,
            )

            if gaze_violation:
                draw_error_banner(frame, "ERROR: Please focus on your paper or exam")
            elif talking_violation:
                draw_error_banner(frame, "ERROR: Talking detected")
            elif gesture_violation:
                draw_error_banner(frame, "ERROR: Conduct violation")
            elif face_count == 0:
                draw_error_banner(frame, "ERROR: Face not detected")
            elif face_count > 1:
                draw_error_banner(frame, "ERROR: Only one face allowed")

            if alarm_needed:
                now = time.time()
                if now - last_alarm_time >= ALARM_INTERVAL:
                    if alarm_message is not None:
                        voice_alarm(alarm_message)
                    else:
                        simple_alarm()
                    last_alarm_time = now

            cv2.imshow("Exam Face Detector", frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

    except KeyboardInterrupt:
        print("Interrupted by user")
    finally:
        close_hand_detector()
        close_face_mesh_detector()
        cap.release()
        cv2.destroyAllWindows()


if __name__ == '__main__':
    main()
