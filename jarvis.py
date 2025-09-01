# --- Imports ---
import speech_recognition as sr
import pyttsx3
import subprocess
import threading
import time
import os
import re
import psutil
import pyautogui
import pyperclip
import win32gui
import win32process
import win32con
from datetime import datetime
import tkinter as tk
from tkinter import ttk
import sys
import traceback
import logging
import webbrowser
import ctypes
import pytesseract
from PIL import Image

# For advanced web automation
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    print("Selenium or its dependencies not available. Install with: pip install selenium webdriver-manager")

# --- Basic Setup ---
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Configure logging
logging.basicConfig(filename='jarvis.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Groq LLM Configuration ---
try:
    from groq import Groq
    GROQ_AVAILABLE = True
    GROQ_API_KEY = "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"
except ImportError:
    GROQ_AVAILABLE = False
    GROQ_API_KEY = None
    print("Groq library not available. Install with: pip install groq")

# --- Global State ---
is_sleeping = False
VALID_ACTIONS = {"open", "search", "create", "delete", "focus", "type", "press", "hotkey", "click", "wait", "screenshot", "speak", "scroll", "click_text", "web_search_and_play", "spotify_search_and_play"}
speech_lock = threading.Lock()
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
except Exception as e:
    logging.error(f"Pytesseract setup error: {e}. Please ensure Tesseract is installed and the path is correct.")

# --- Initialize Speech Engine (TTS) ---
try:
    engine = pyttsx3.init()
    engine.setProperty('rate', 175)
    engine.setProperty('volume', 1.0)
except Exception as e:
    logging.error(f"Failed to initialize pyttsx3: {e}")
    class DummyEngine:
        def say(self, text): pass
        def runAndWait(self): pass
    engine = DummyEngine()

# --- Initialize Recognizer (STT) ---
try:
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    recognizer.pause_threshold = 1.0
    recognizer.energy_threshold = 4000
    SPEECH_AVAILABLE = True
except Exception as e:
    recognizer = None
    microphone = None
    SPEECH_AVAILABLE = False
    print("Speech recognition library not available. Install with: pip install SpeechRecognition")

# --- GUI Class ---
class JarvisGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("J.A.R.V.I.S. Assistant")
        self.root.geometry("500x400")
        self.root.attributes('-topmost', True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.is_running = True
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TProgressbar", thickness=12)
        main_frame = tk.Frame(self.root)
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        self.status_label = tk.Label(main_frame, text="J.A.R.V.I.S. Initializing...", font=("Segoe UI", 16, "bold"))
        self.status_label.pack(pady=(0, 10))
        self.listening_label = tk.Label(main_frame, text="Status: Booting...", fg="orange", font=("Segoe UI", 12))
        self.listening_label.pack()
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', style="TProgressbar")
        self.progress.pack(pady=10, fill=tk.X)
        self.context_label = tk.Label(main_frame, text="Context: None", font=("Segoe UI", 10), wraplength=450, justify=tk.LEFT)
        self.context_label.pack(pady=10)
        self.history_text = tk.Text(main_frame, height=8, width=55, font=("Segoe UI", 9), relief=tk.SOLID, borderwidth=1)
        self.history_text.pack(fill=tk.BOTH, expand=True)

    def on_closing(self): self.is_running = False; self.root.destroy()
    def update_listening_status(self, status_text, color):
        if self.is_running:
            self.listening_label.config(text=f"Status: {status_text}", fg=color)
            if "Listening" in status_text or "Processing" in status_text: self.progress.start(15)
            else: self.progress.stop()
            self.root.update()
    def add_to_history(self, command):
        if self.is_running:
            self.history_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {command}\n"); self.history_text.see(tk.END); self.root.update()

gui = JarvisGUI()

# --- Core Assistant Functions ---
def speak_threaded(text):
    def run_speech():
        with speech_lock:
            engine.say(text)
            engine.runAndWait()
    print(f"J.A.R.V.I.S.: {text}")
    threading.Thread(target=run_speech).start()

def listen_for_input():
    if not SPEECH_AVAILABLE: return None
    with microphone as source:
        try:
            print("Listening...")
            audio = recognizer.listen(source, timeout=None, phrase_time_limit=12)
            gui.update_listening_status("Recognizing...", "cyan")
            text = recognizer.recognize_google(audio)
            print(f"You said: {text}")
            gui.add_to_history(text)
            return text.lower()
        except sr.UnknownValueError:
            return None
        except Exception as e:
            logging.error(f"Listening error: {e}")
            return None

def find_text_coordinates(text_to_find):
    try:
        screenshot = pyautogui.screenshot()
        text_data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT)
        for i, text in enumerate(text_data['text']):
            if text_to_find.lower() in text.lower():
                x = text_data['left'][i]
                y = text_data['top'][i]
                w = text_data['width'][i]
                h = text_data['height'][i]
                center_x = x + w // 2
                center_y = y + h // 2
                return (center_x, center_y)
    except Exception as e:
        logging.error(f"Error finding text coordinates: {e}")
    return None

def web_search_and_play(query):
    if not SELENIUM_AVAILABLE:
        speak_threaded("The web automation library is not installed.")
        return
    try:
        speak_threaded(f"Searching YouTube for {query}")
        chrome_options = Options()
        chrome_options.add_argument("--disable-infobars")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get("https://www.youtube.com")
        wait = WebDriverWait(driver, 10)
        
        # More robust locator for search box
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@id="search" and @name="search_query"]')))
        search_box.send_keys(query)
        
        search_button = wait.until(EC.element_to_be_clickable((By.ID, "search-icon-legacy")))
        search_button.click()
        
        # Use a more general locator for video links
        video_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'ytd-video-renderer a#video-title, ytd-compact-video-renderer a#video-title')))
        video_link.click()

        speak_threaded("Playing the first result.")
    except Exception as e:
        logging.error(f"Web automation error: {e}")
        speak_threaded("I'm sorry, sir. I encountered an error while trying to play the video.")
        if 'driver' in locals():
            driver.quit()

def spotify_web_search_and_play(query):
    if not SELENIUM_AVAILABLE:
        speak_threaded("The web automation library is not installed.")
        return
    try:
        speak_threaded(f"Opening Spotify Web Player and searching for {query}")
        chrome_options = Options()
        chrome_options.add_argument("--disable-infobars")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get("https://open.spotify.com/")

        wait = WebDriverWait(driver, 20)
        
        # Locate the search button (magnifying glass)
        try:
            search_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="search-link"]')))
            search_button.click()
        except:
            # If search is already on screen, try to find the input directly
            pass

        search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-testid="search-input"]')))
        search_input.send_keys(query)
        
        time.sleep(2) # Wait for search results to load
        
        # This locator selects the first track play button. This is a common and reliable pattern on Spotify.
        first_track_play_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="play-button"]')))
        first_track_play_button.click()
        
        speak_threaded("Now playing on Spotify Web Player.")
    except Exception as e:
        logging.error(f"Spotify web automation error: {e}")
        speak_threaded("I'm sorry, sir. I encountered an error while trying to play the song on the Spotify Web Player.")
        if 'driver' in locals():
            driver.quit()

def execute_command(command):
    logging.info(f"Executing command: {command}")
    action = command.split()[0].lower()
    try:
        if action == "open":
            target = re.search(r'target="(.+?)"', command, re.IGNORECASE)
            if target:
                url_or_app = target.group(1).lower()
                if url_or_app.startswith("http://") or url_or_app.startswith("https://"):
                    webbrowser.open(url_or_app)
                elif "chrome" in url_or_app:
                    os.startfile("chrome.exe")
                elif "notepad" in url_or_app:
                    os.startfile("notepad.exe")
                else:
                    os.startfile(url_or_app)
        elif action == "search":
            query = re.search(r'query="(.+?)"', command, re.IGNORECASE)
            if query: 
                webbrowser.open(f"https://www.google.com/search?q={query.group(1)}")
        elif action == "create":
            path_match = re.search(r'path="(.+?)"', command, re.IGNORECASE)
            content_match = re.search(r'content="(.+?)"', command, re.IGNORECASE)
            if path_match:
                file_path = os.path.expanduser(path_match.group(1))
                content = content_match.group(1) if content_match else ""
                content = content.replace('\\n', '\n').replace('\\t', '\t')
                with open(file_path, 'w') as f: f.write(content)
        elif action == "delete":
            path_match = re.search(r'path="(.+?)"', command, re.IGNORECASE)
            if path_match:
                file_path = os.path.expanduser(path_match.group(1))
                if os.path.exists(file_path):
                    os.remove(file_path)
                else:
                    logging.warning(f"File not found: {file_path}")
        elif action == "type":
            text_match = re.search(r'text="(.+?)"', command, re.IGNORECASE)
            if text_match: 
                text_to_type = text_match.group(1).replace('\\n', '\n').replace('\\t', '\t')
                pyautogui.typewrite(text_to_type, interval=0.01)
        elif action == "press":
            key_match = re.search(r'key="(.+?)"', command, re.IGNORECASE)
            if key_match: pyautogui.press(*key_match.group(1).split(','))
        elif action == "hotkey":
            keys_match = re.search(r'keys="(.+?)"', command, re.IGNORECASE)
            if keys_match: pyautogui.hotkey(*keys_match.group(1).split('+'))
        elif action == "click": pyautogui.click()
        elif action == "wait":
            duration_match = re.search(r'duration="(.+?)"', command, re.IGNORECASE)
            if duration_match: time.sleep(float(duration_match.group(1)))
        elif action == "focus":
            title_match = re.search(r'title="(.+?)"', command, re.IGNORECASE)
            if title_match:
                hwnd = win32gui.FindWindow(None, title_match.group(1))
                if hwnd: 
                    win32gui.SetForegroundWindow(hwnd)
                else:
                    logging.warning(f"Could not find window: {title_match.group(1)}")
        elif action == "screenshot":
            pyautogui.screenshot(f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        elif action == "speak":
            text_match = re.search(r'text="(.+?)"', command, re.IGNORECASE)
            if text_match: speak_threaded(text_match.group(1))
        elif action == "scroll":
            amount_match = re.search(r'amount="(.+?)"', command, re.IGNORECASE)
            if amount_match:
                try:
                    scroll_amount = int(amount_match.group(1))
                    pyautogui.scroll(scroll_amount)
                except (ValueError, TypeError) as e:
                    logging.error(f"Invalid scroll amount: {e}")
        elif action == "click_text":
            text_match = re.search(r'text="(.+?)"', command, re.IGNORECASE)
            if text_match:
                target_text = text_match.group(1)
                coordinates = find_text_coordinates(target_text)
                if coordinates:
                    pyautogui.click(coordinates[0], coordinates[1])
                    speak_threaded(f"Clicked on {target_text}.")
                else:
                    speak_threaded(f"Sorry, I could not find '{target_text}' on the screen.")
        elif action == "web_search_and_play":
            query_match = re.search(r'query="(.+?)"', command, re.IGNORECASE)
            if query_match: web_search_and_play(query_match.group(1))
        elif action == "spotify_search_and_play":
            query_match = re.search(r'query="(.+?)"', command, re.IGNORECASE)
            if query_match: spotify_web_search_and_play(query_match.group(1))
    except Exception as e:
        error_msg = f"Error executing '{command}': {str(e)}"; print(error_msg); logging.error(error_msg)
        speak_threaded("I'm sorry, sir. I failed to execute that command.")

def get_system_context():
    active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()) or "None"
    return f"""<CONTEXT>
    <TIMESTAMP>{datetime.now().strftime("%A, %B %d, %Y, %I:%M %p")}</TIMESTAMP>
    <USER>{os.getlogin()}</USER>
    <ACTIVE_WINDOW>{active_window}</ACTIVE_WINDOW>
    <CWD>{os.getcwd()}</CWD>
</CONTEXT>"""

def ask_llm(user_input):
    if not GROQ_AVAILABLE: return 'speak text="My connection to the Groq reasoning engine is unavailable."'
    prompt = f"""<TASK>
You are J.A.R.V.I.S., a hyper-competent AI controlling a Windows PC. Your ONLY output must be a sequence of structured commands. Adhere strictly to the command syntax. Do not add any conversational text or explanations.
</TASK>
{get_system_context()}
<RULES>
1. Your entire response must be a sequence of commands, each on a new line.
2. Use absolute paths for file operations (e.g., "C:\\Users\\{os.getlogin()}\\Desktop\\file.txt").
3. Always `focus` on a window before you `type` in it.
4. Parameters must be enclosed in double quotes.
5. Use `\\n` for newlines and `\\t` for tabs in `type` or `create` commands.
6. Use `scroll` to move the screen view. A positive value scrolls down, a negative one scrolls up.
7. Use `click_text` to click on text visible on the screen.
8. If the user asks to play a song or video from YouTube, use the `web_search_and_play` command.
9. If the user asks to play a song on Spotify, use the `spotify_search_and_play` command.
</RULES>
<COMMANDS>
- `open target="<app_name_or_URL>"` (e.g., "chrome", "notepad")
- `search query="<search_terms>"`
- `create path="<full_path>" content="<text_to_write>"`
- `delete path="<full_path>"`
- `focus title="<exact_window_title>"`
- `type text="<text_to_type>"`
- `press key="<key_name>"` (e.g., "enter", "f5", "win")
- `hotkey keys="<key1>+<key2>"` (e.g., "ctrl+c", "alt+f4")
- `click`
- `wait duration="<seconds>"`
- `screenshot`
- `speak text="<text_to_say>"`
- `scroll amount="<pixels>"`
- `click_text text="<exact_text>"`
- `web_search_and_play query="<search_terms>"`
- `spotify_search_and_play query="<search_terms>"`
</COMMANDS>
<EXAMPLE_TASK>
User said: "play 'Bohemian Rhapsody' by Queen on the Spotify web player"
Your response would be:
spotify_search_and_play query="Bohemian Rhapsody Queen"
</EXAMPLE_TASK>
<EXAMPLE_TASK>
User said: "play 'Bohemian Rhapsody' by Queen on YouTube"
Your response would be:
web_search_and_play query="Bohemian Rhapsody Queen"
</EXAMPLE_TASK>
<USER_REQUEST>
{user_input}
</USER_REQUEST>"""
    try:
        client = Groq(api_key=GROQ_API_KEY)
        chat_completion = client.chat.completions.create(messages=[{"role": "user", "content": prompt}], model="llama-3.3-70b-versatile", temperature=0.0)
        return chat_completion.choices[0].message.content
    except Exception as e:
        logging.error(f"LLM API error: {str(e)}"); return 'speak text="I am having trouble connecting to my reasoning engine."'

def process_llm_response(response):
    commands = [line.strip() for line in response.strip().split('\n') if line.strip() and line.split()[0].lower() in VALID_ACTIONS]
    if not commands:
        speak_threaded("I'm not sure how to do that, sir. My apologies.")
        return
    speak_threaded("Right away, sir.")
    for command in commands:
        if not gui.is_running: break
        execute_command(command)
        time.sleep(0.3)
    speak_threaded("Task complete.")

def main_loop():
    global is_sleeping
    speak_threaded("J.A.R.V.I.S. systems are online.")
    if SPEECH_AVAILABLE:
        with microphone as source: recognizer.adjust_for_ambient_noise(source, duration=1.5)
    while gui.is_running:
        try:
            status, color = ("Sleeping... (Say 'Jarvis wake up')", "orange") if is_sleeping else ("Listening...", "green")
            gui.update_listening_status(status, color)
            command = listen_for_input()
            if not command: continue
            if is_sleeping:
                if "jarvis wake up" in command: is_sleeping = False; speak_threaded("I am back online.")
                continue
            if "jarvis sleep" in command: is_sleeping = True; speak_threaded("Entering sleep mode."); continue
            if any(w in command for w in ["exit", "quit", "goodbye"]): speak_threaded("Shutting down."); gui.is_running = False; break
            gui.update_listening_status("Processing...", "purple")
            llm_response = ask_llm(command)
            process_llm_response(llm_response)
        except Exception as e:
            logging.error(f"Main loop error: {str(e)}"); traceback.print_exc()
            speak_threaded("A critical error occurred."); time.sleep(2)
    gui.root.destroy(); sys.exit()

if __name__ == "__main__":
    if is_admin():
        main_thread = threading.Thread(target=main_loop, daemon=True)
        main_thread.start()
        gui.root.mainloop()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    print("J.A.R.V.I.S. has shut down.")