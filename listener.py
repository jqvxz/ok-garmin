import speech_recognition as sr
import tkinter as tk
from PIL import Image, ImageTk
import time
import winsound
import threading
import collections.abc
import json
import os
import webbrowser
from fuzzywuzzy import fuzz
from colorama import Fore, Style, init
init(autoreset=True)

# --- Terminal Interface Prefixes ---
PREFIX_INFO = f"{Fore.CYAN}[*]{Style.RESET_ALL}"
PREFIX_OK = f"{Fore.GREEN}[+]{Style.RESET_ALL}"
PREFIX_ERROR = f"{Fore.RED}[-]{Style.RESET_ALL}"
PREFIX_DEBUG = f"{Fore.YELLOW}[DEBUG]{Style.RESET_ALL}"

# --- Compatibility Fix ---
collections.Callable = collections.abc.Callable
import pyautogui

# --- Configuration ---
HOTWORDS = ["ok garmin", "okay garmin", "okay gamin", "oke garmin", "ok gamin"]
IMAGE_PATH = "blend.png"
ANIMATION_DURATION_S = 4.0
COMMAND_TIMEOUT_SECONDS = 5.0
FUZZY_MATCH_THRESHOLD = 80
MIN_COMMAND_LENGTH_RATIO = 0.7 # Prevents short words from matching long commands.
READY_BEEP = (1000, 200)
SUCCESS_BEEP = (1500, 150)

# ==============================================================================
#  Action Functions
# ==============================================================================

def play_success_beep():
    """Plays two short, high-pitched beeps to confirm a successful action."""
    winsound.Beep(SUCCESS_BEEP[0], SUCCESS_BEEP[1])
    time.sleep(0.05)
    winsound.Beep(SUCCESS_BEEP[0], SUCCESS_BEEP[1])

def take_screenshot():
    """Captures the screen and saves it as a PNG file."""
    print(f"{PREFIX_DEBUG} Action: take_screenshot()")
    try:
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"screenshot_{timestamp}.png"
        pyautogui.screenshot().save(filename)
        print(f"{PREFIX_OK} Screenshot saved as {filename}")
    except Exception as e:
        print(f"{PREFIX_ERROR} Could not take screenshot: {e}")

def open_explorer():
    """Opens a new Windows File Explorer window."""
    print(f"{PREFIX_DEBUG} Action: open_explorer()")
    os.system("explorer")
    print(f"{PREFIX_OK} Opened File Explorer.")

def open_settings():
    """Opens the Windows Settings application."""
    print(f"{PREFIX_DEBUG} Action: open_settings()")
    os.system("start ms-settings:")
    print(f"{PREFIX_OK} Opened Windows Settings.")

def open_discord():
    """Attempts to launch the Discord application."""
    print(f"{PREFIX_DEBUG} Action: open_discord()")
    os.system("start discord://")
    print(f"{PREFIX_OK} Attempted to open Discord.")

def close_explorer():
    """Closes all open File Explorer windows safely."""
    print(f"{PREFIX_DEBUG} Action: close_explorer()")
    command = 'powershell -command "(New-Object -comObject Shell.Application).Windows() | foreach-object { $_.Quit() }"'
    os.system(command)
    print(f"{PREFIX_OK} Closed all File Explorer windows.")

def close_settings():
    """Closes the Windows Settings application."""
    print(f"{PREFIX_DEBUG} Action: close_settings()")
    os.system("taskkill /F /IM SystemSettings.exe > nul 2>&1")
    print(f"{PREFIX_OK} Closed Windows Settings.")

def close_discord():
    """Closes the Discord application."""
    print(f"{PREFIX_DEBUG} Action: close_discord()")
    os.system("taskkill /F /IM Discord.exe > nul 2>&1")
    print(f"{PREFIX_OK} Closed Discord.")

def close_browsers():
    """Closes all major web browsers."""
    print(f"{PREFIX_DEBUG} Action: close_browsers()")
    browsers = ["chrome.exe", "msedge.exe", "firefox.exe", "opera.exe", "brave.exe"]
    for browser in browsers:
        os.system(f"taskkill /F /IM {browser} > nul 2>&1")
    print(f"{PREFIX_OK} Closed major web browsers.")

def open_browser():
    """Opens the default web browser to Google."""
    print(f"{PREFIX_DEBUG} Action: open_browser()")
    webbrowser.open("https://google.com")
    print(f"{PREFIX_OK} Opened default browser.")

def open_github():
    """Opens a specific GitHub profile."""
    print(f"{PREFIX_DEBUG} Action: open_github()")
    webbrowser.open("https://github.com/jqvxz")
    print(f"{PREFIX_OK} Opened GitHub profile.")
    
def open_youtube():
    """Opens YouTube in the default browser."""
    print(f"{PREFIX_DEBUG} Action: open_youtube()")
    webbrowser.open("https://youtube.com")
    print(f"{PREFIX_OK} Opened YouTube.")

def open_twitter():
    """Opens Twitter/X in the default browser."""
    print(f"{PREFIX_DEBUG} Action: open_twitter()")
    webbrowser.open("https://twitter.com")
    print(f"{PREFIX_OK} Opened Twitter/X.")
    
def open_twitch():
    """Opens Twitch in the default browser."""
    print(f"{PREFIX_DEBUG} Action: open_twitch()")
    webbrowser.open("https://twitch.tv")
    print(f"{PREFIX_OK} Opened Twitch.")
    
def open_wikipedia():
    """Opens Wikipedia in the default browser."""
    print(f"{PREFIX_DEBUG} Action: open_wikipedia()")
    webbrowser.open("https://wikipedia.org")
    print(f"{PREFIX_OK} Opened Wikipedia.")

def open_ph():
    """Opens special site on request."""
    print(f"{PREFIX_DEBUG} Action: open_ph()")
    webbrowser.open("https://pornhub.com")
    print(f"{PREFIX_OK} Opened PH.")   

# ==============================================================================
#  Command Dictionary
# ==============================================================================

COMMANDS = {
    "öffne explorer": open_explorer, "öffne einstellungen": open_settings,
    "öffne discord": open_discord, "öffne browser": open_browser,
    "schließe explorer": close_explorer, "schließe einstellungen": close_settings,
    "schließe discord": close_discord, "schließe browser": close_browsers,
    "öffne youtube": open_youtube, "öffne twitter": open_twitter,
    "öffne twitch": open_twitch, "öffne wikipedia": open_wikipedia,
    "öffne github": open_github, "video speichern": take_screenshot, "save video": take_screenshot,
    "rubbel ihn langsam" :open_ph, "rubbel ihn schnell": open_ph
}

def recognize_speech_flexible(recognizer, audio):
    """Recognizes speech, prioritizing German, and returns all alternatives."""
    print(f"{PREFIX_DEBUG} Attempting recognition (German first)...")
    try:
        return recognizer.recognize_google(audio, language="de-DE", show_all=True)
    except sr.UnknownValueError:
        print(f"{PREFIX_DEBUG} German model failed. Trying English model...")
        try:
            return recognizer.recognize_google(audio, language="en-US", show_all=True)
        except sr.UnknownValueError:
            print(f"{PREFIX_DEBUG} English model failed. No recognition possible.")
            return None

def show_blended_image():
    """Displays a borderless image with a fade-in and fade-out animation."""
    print(f"{PREFIX_DEBUG} Animation thread started.")
    try:
        root = tk.Tk(); root.withdraw()
        window = tk.Toplevel(root); window.overrideredirect(True)
        window.wm_attributes('-topmost', 1); window.attributes('-alpha', 0)
        pil_image = Image.open(IMAGE_PATH); tk_image = ImageTk.PhotoImage(pil_image)
        screen_width, screen_height = window.winfo_screenwidth(), window.winfo_screenheight()
        pos_x, pos_y = (screen_width // 2) - (pil_image.width // 2), (screen_height // 2) - (pil_image.height // 2)
        window.geometry(f"+{pos_x}+{pos_y}")
        label = tk.Label(window, image=tk_image, bd=0); label.pack()
        fade_time, hold_time, steps = ((ANIMATION_DURATION_S / 2) * 0.75), (ANIMATION_DURATION_S - (((ANIMATION_DURATION_S / 2) * 0.75) * 2)), 30
        for i in range(steps + 1): window.attributes('-alpha', i / steps); window.update(); time.sleep(fade_time / steps)
        time.sleep(hold_time)
        for i in range(steps + 1): window.attributes('-alpha', 1.0 - (i / steps)); window.update(); time.sleep(fade_time / steps)
        window.destroy(); root.destroy()
        print(f"{PREFIX_DEBUG} Animation thread finished.")
    except FileNotFoundError: print(f"{PREFIX_ERROR} Image file not found at '{IMAGE_PATH}'")
    except Exception as e: print(f"{PREFIX_ERROR} A display error occurred: {e}")

def listen_for_command():
    """Listens for a command and executes the best match from the COMMANDS dictionary."""
    print(f"{PREFIX_DEBUG} listen_for_command() function called.")
    winsound.Beep(READY_BEEP[0], READY_BEEP[1])
    print(f"{PREFIX_INFO} Listening for a command...")
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            audio = recognizer.listen(source, timeout=COMMAND_TIMEOUT_SECONDS)
            print(f"{PREFIX_DEBUG} Command audio captured. Sending to API...")
            response = recognize_speech_flexible(recognizer, audio)
            
            if response and "alternative" in response:
                best_match_score = 0
                best_match_action = None
                best_match_transcript = ""

                # Find the best possible command match across all recognized alternatives.
                for alternative in response["alternative"]:
                    transcript = alternative["transcript"]
                    for command_phrase, action_function in COMMANDS.items():
                        
                        # Prevent short phrases like "schließe" from matching longer commands.
                        if len(transcript) < len(command_phrase) * MIN_COMMAND_LENGTH_RATIO:
                            continue
                            
                        ratio = fuzz.partial_ratio(command_phrase, transcript.lower())
                        if ratio > best_match_score:
                            best_match_score = ratio
                            best_match_action = action_function
                            best_match_transcript = transcript
                
                print(f"{PREFIX_DEBUG} Best overall command match: '{best_match_transcript}' (Score: {best_match_score}%)")
                
                # If the best match found is good enough, execute it and play the success beep.
                if best_match_score >= FUZZY_MATCH_THRESHOLD:
                    print(f"{PREFIX_OK} Executing command for '{best_match_transcript}'")
                    best_match_action()
                    play_success_beep()
                else:
                    print(f"{PREFIX_ERROR} Unknown command. Best guess was '{best_match_transcript}'.")
            else:
                print(f"{PREFIX_ERROR} Could not understand command.")
        except sr.WaitTimeoutError:
            print(f"{PREFIX_ERROR} No command heard within the time limit.")
        except sr.RequestError as e:
            print(f"{PREFIX_ERROR} API Error during command recognition: {e}")

def start_listening():
    """The main loop of the application."""
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    print(f"{PREFIX_INFO} Calibrating for ambient noise...")
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source, duration=1.5)
    
    # --- ADDED ---
    # Play a success beep to indicate that calibration is complete and the assistant is ready.
    play_success_beep()
    # --- END ADDED ---

    print(f"{PREFIX_INFO} Ready. Listening for hotwords (Threshold: {FUZZY_MATCH_THRESHOLD}%)")
    
    while True:
        with microphone as source:
            try:
                audio = recognizer.listen(source)
                print(f"{PREFIX_DEBUG} Hotword audio captured. Sending to API...")
                response = recognize_speech_flexible(recognizer, audio)
                
                if response and "alternative" in response:
                    hotword_match_found = False
                    for alternative in response["alternative"]:
                        transcript = alternative["transcript"]
                        best_score = 0
                        for hotword in HOTWORDS:
                            ratio = fuzz.partial_ratio(hotword.lower(), transcript.lower())
                            if ratio > best_score:
                                best_score = ratio
                        
                        print(f"{PREFIX_DEBUG}   Checking Alt: '{transcript}' (Best Hotword Score: {best_score}%)")

                        if best_score >= FUZZY_MATCH_THRESHOLD:
                            print(f"{PREFIX_OK} Hotword matched with '{transcript}' (Score: {best_score}%)")
                            animation_thread = threading.Thread(target=show_blended_image)
                            animation_thread.start()
                            listen_for_command()
                            print(f"\n{PREFIX_INFO} Ready. Listening for hotwords...")
                            hotword_match_found = True
                            break
                    
                    if not hotword_match_found:
                        print(f"{PREFIX_DEBUG} No alternative met the threshold of {FUZZY_MATCH_THRESHOLD}%.")
            except sr.RequestError as e:
                print(f"{PREFIX_ERROR} API Error: {e}")
                time.sleep(3)
            except Exception as e:
                print(f"{PREFIX_ERROR} An unexpected error occurred: {e}")

if __name__ == "__main__":
    print(f"{PREFIX_INFO} --- Initializing Voice Assistant ---")
    try:
        mic_names = sr.Microphone.list_microphone_names()
        print(f"{PREFIX_INFO} Found {len(mic_names)} microphone(s):")
        for index, name in enumerate(mic_names):
            print(f"  {PREFIX_DEBUG} Mic {index}: {name}")
        print(f"{PREFIX_INFO} Using default microphone (Index 0). Change in code if needed.")
    except Exception as e:
        print(f"{PREFIX_ERROR} Could not list microphones. Ensure PyAudio is installed correctly. Error: {e}")

    start_listening()