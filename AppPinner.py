import win32gui
import win32con
import time
import threading
from tkinter import messagebox
import tkinter as tk

# Opslag voor gepinde vensters
pinned_windows = {}
running = True

def pin_window(window_title):
    """Zet een venster altijd bovenaan."""
    found = False
    handles = []

    def callback(handle, data):
        title = win32gui.GetWindowText(handle)
        if window_title.lower() in title.lower():
            handles.append(handle)
        return True

    win32gui.EnumWindows(callback, None)

    if handles:
        for handle in handles:
            try:
                win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                pinned_windows[window_title] = handle
                found = True
            except:
                continue
        
        if found:
            print(f"'{window_title}' is vastgezet.")
            return True

    print(f"Geen venster gevonden met titel: '{window_title}'")
    return False

def unpin_window(window_title):
    """Maak een venster los van de bovenste laag."""
    to_remove = []
    for title, handle in pinned_windows.items():
        if window_title.lower() in title.lower():
            try:
                win32gui.SetWindowPos(handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                to_remove.append(title)
            except:
                continue
    
    for title in to_remove:
        del pinned_windows[title]
        print(f"'{title}' is losgemaakt.")
    
    if not to_remove:
        print(f"Geen venster gevonden met titel: '{window_title}'")

def list_pinned_windows():
    """Toon alle momenteel vastgezette vensters."""
    if pinned_windows:
        print("Gepinde vensters:")
        for title in pinned_windows.keys():
            print(f"- {title}")
    else:
        print("Er zijn momenteel geen vensters vastgezet.")

def keep_windows_pinned():
    """Houd vastgezette vensters continu bovenaan."""
    while running:
        for title, handle in list(pinned_windows.items()):
            try:
                if win32gui.IsWindow(handle):
                    win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                else:
                    print(f"Venster '{title}' bestaat niet meer.")
                    del pinned_windows[title]
            except:
                del pinned_windows[title]
        time.sleep(1)

def main_menu():
    """Toon het hoofdmenu en verwerk de gebruikersinvoer."""
    while True:
        print("\nOpties:")
        print("1. Pin een venster")
        print("2. Unpin een venster")
        print("3. Toon gepinde vensters")
        print("4. Stop het programma")

        keuze = input("Maak een keuze: ")

        if keuze == "1":
            title = input("Voer de titel van het venster in dat je wilt pinnen: ")
            pin_window(title)
        elif keuze == "2":
            title = input("Voer de titel van het venster in dat je wilt unpinnen: ")
            unpin_window(title)
        elif keuze == "3":
            list_pinned_windows()
        elif keuze == "4":
            return False
        else:
            print("Ongeldige keuze. Probeer opnieuw.")
        
        input("Druk op Enter om door te gaan...")

def main():
    global running
    
    # Start de thread om vensters vastgezet te houden
    thread = threading.Thread(target=keep_windows_pinned, daemon=True)
    thread.start()

    # Start het hoofdmenu
    running = main_menu()

    print("Programma wordt afgesloten.")

if __name__ == "__main__":
    main()
