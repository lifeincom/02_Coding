# -*- coding: utf-8 -*-
import sys
import os

# Add the current directory to sys.path to ensure modules can be imported correctly
# when running from the directory itself or via -m
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    USE_CTK = True
    try:
        import customtkinter
    except ImportError:
        USE_CTK = False
        print("customtkinter module was not found. Running with standard tkinter.")

    if USE_CTK:
        try:
            from ui.ctk_app import App
            app = App()
            app.mainloop()
        except Exception as e:
            print(f"Failed to launch CustomTkinter App: {e}")
            print("Falling back to Tkinter...")
            USE_CTK = False

    if not USE_CTK:
        try:
            from ui.tk_app import App
            app = App()
            app.mainloop()
        except Exception as e:
            print(f"Critical Error: Failed to launch Tkinter App: {e}")
            input("Press Enter to exit...")

if __name__ == "__main__":
    main()
