"""Entry point for PPLX GUI application."""

from tkinter import messagebox

from src.gui.app import PPLXGUIApp


def main():
    """Run the GUI application."""
    try:
        app = PPLXGUIApp()
        app.run()
    except Exception as e:
        print(f"Error starting application: {e}")
        messagebox.showerror(
            "Startup Error", f"Failed to start application:\n{str(e)}"
        )


if __name__ == "__main__":
    main()
