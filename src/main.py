import ctypes

try:
    ctypes.WinDLL("C:\\Windows\\System32\\msvcp140.dll")
    ctypes.WinDLL("C:\\Windows\\System32\\vcruntime140.dll")
except Exception:
    pass

import os
import sys
import time
from datetime import datetime
from audioplayer import AudioPlayer
from pynput.keyboard import Controller
from PyQt5.QtCore import QObject, QProcess
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QAction, QMessageBox

from key_listener import KeyListener
from result_thread import ResultThread
from ui.main_window import MainWindow
from ui.settings_window import SettingsWindow
from ui.status_window import StatusWindow
from transcription import create_local_model
from input_simulation import InputSimulator
from utils import ConfigManager


class WhisperWriterApp(QObject):
    def __init__(self):
        """
        Initialize the application, opening settings window if no configuration file is found.
        """
        super().__init__()
        self.app = QApplication(sys.argv)
        self.app.setWindowIcon(QIcon(os.path.join("assets", "ww-logo.png")))

        ConfigManager.initialize()

        self.settings_window = SettingsWindow()
        self.settings_window.settings_closed.connect(self.on_settings_closed)
        self.settings_window.settings_saved.connect(self.restart_app)

        if ConfigManager.config_file_exists():
            self.initialize_components()
        else:
            print("No valid configuration file found. Opening settings window...")
            self.settings_window.show()

    def initialize_components(self):
        """
        Initialize the components of the application.
        """
        self.input_simulator = InputSimulator()
        self.file_output_mode = False  # None = keyboard, "overwrite" = overwrite file, "append" = append to file
        self.file_output_append_mode = False
        self.active_context_tag = None  # Persistent context tag across transcriptions

        self.key_listener = KeyListener()
        self.key_listener.add_callback("on_activate", self.on_activation)
        self.key_listener.add_callback("on_deactivate", self.on_deactivation)
        self.key_listener.add_callback(
            "on_file_output_activate", self.on_file_output_activation
        )
        self.key_listener.add_callback(
            "on_file_output_deactivate", self.on_file_output_deactivation
        )
        self.key_listener.add_callback(
            "on_file_output_append_activate", self.on_file_output_append_activation
        )
        self.key_listener.add_callback(
            "on_file_output_append_deactivate", self.on_file_output_append_deactivation
        )

        model_options = ConfigManager.get_config_section("model_options")
        model_path = model_options.get("local", {}).get("model_path")
        self.local_model = (
            create_local_model() if not model_options.get("use_api") else None
        )

        self.result_thread = None

        self.main_window = MainWindow()
        self.main_window.openSettings.connect(self.settings_window.show)
        self.main_window.startListening.connect(self.key_listener.start)
        self.main_window.closeApp.connect(self.exit_app)

        if not ConfigManager.get_config_value("misc", "hide_status_window"):
            self.status_window = StatusWindow()

        self.create_tray_icon()
        self.main_window.show()

    def create_tray_icon(self):
        """
        Create the system tray icon and its context menu.
        """
        self.tray_icon = QSystemTrayIcon(
            QIcon(os.path.join("assets", "ww-logo.png")), self.app
        )

        tray_menu = QMenu()

        show_action = QAction("WhisperWriter Main Menu", self.app)
        show_action.triggered.connect(self.main_window.show)
        tray_menu.addAction(show_action)

        settings_action = QAction("Open Settings", self.app)
        settings_action.triggered.connect(self.settings_window.show)
        tray_menu.addAction(settings_action)

        exit_action = QAction("Exit", self.app)
        exit_action.triggered.connect(self.exit_app)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def cleanup(self):
        if self.key_listener:
            self.key_listener.stop()
        if self.input_simulator:
            self.input_simulator.cleanup()

    def exit_app(self):
        """
        Exit the application.
        """
        self.cleanup()
        QApplication.quit()

    def restart_app(self):
        """Restart the application to apply the new settings."""
        self.cleanup()
        QApplication.quit()
        QProcess.startDetached(sys.executable, sys.argv)

    def on_settings_closed(self):
        """
        If settings is closed without saving on first run, initialize the components with default values.
        """
        if not os.path.exists(os.path.join("src", "config.yaml")):
            QMessageBox.information(
                self.settings_window,
                "Using Default Values",
                "Settings closed without saving. Default values are being used.",
            )
            self.initialize_components()

    def on_activation(self):
        """
        Called when the activation key combination is pressed.
        """
        if self.result_thread and self.result_thread.isRunning():
            recording_mode = ConfigManager.get_config_value(
                "recording_options", "recording_mode"
            )
            if recording_mode == "press_to_toggle":
                self.result_thread.stop_recording()
            elif recording_mode == "continuous":
                self.stop_result_thread()
            return

        self.start_result_thread()

    def on_deactivation(self):
        """
        Called when the activation key combination is released.
        """
        if (
            ConfigManager.get_config_value("recording_options", "recording_mode")
            == "hold_to_record"
        ):
            if self.result_thread and self.result_thread.isRunning():
                self.result_thread.stop_recording()

    def on_file_output_activation(self):
        """
        Called when the file output activation key combination is pressed (overwrite mode).
        """
        if self.result_thread and self.result_thread.isRunning():
            recording_mode = ConfigManager.get_config_value(
                "recording_options", "recording_mode"
            )
            if recording_mode == "press_to_toggle":
                self.result_thread.stop_recording()
            elif recording_mode == "continuous":
                self.stop_result_thread()
            return

        self.file_output_mode = "overwrite"
        self.start_result_thread()

    def on_file_output_deactivation(self):
        """
        Called when the file output activation key combination is released.
        """
        if (
            ConfigManager.get_config_value("recording_options", "recording_mode")
            == "hold_to_record"
        ):
            if self.result_thread and self.result_thread.isRunning():
                self.result_thread.stop_recording()

    def on_file_output_append_activation(self):
        """
        Called when the file output append activation key combination is pressed.
        """
        if self.result_thread and self.result_thread.isRunning():
            recording_mode = ConfigManager.get_config_value(
                "recording_options", "recording_mode"
            )
            if recording_mode == "press_to_toggle":
                self.result_thread.stop_recording()
            elif recording_mode == "continuous":
                self.stop_result_thread()
            return

        self.file_output_mode = "append"
        self.start_result_thread()

    def on_file_output_append_deactivation(self):
        """
        Called when the file output append activation key combination is released.
        """
        if (
            ConfigManager.get_config_value("recording_options", "recording_mode")
            == "hold_to_record"
        ):
            if self.result_thread and self.result_thread.isRunning():
                self.result_thread.stop_recording()

    def start_result_thread(self):
        """
        Start the result thread to record audio and transcribe it.
        """
        if self.result_thread and self.result_thread.isRunning():
            return

        self.result_thread = ResultThread(self.local_model, self.active_context_tag)
        if not ConfigManager.get_config_value("misc", "hide_status_window"):
            self.result_thread.statusSignal.connect(self.status_window.updateStatus)
            self.status_window.closeSignal.connect(self.stop_result_thread)
        self.result_thread.resultSignal.connect(self.on_transcription_complete)
        self.result_thread.start()

    def stop_result_thread(self):
        """
        Stop the result thread.
        """
        if self.result_thread and self.result_thread.isRunning():
            self.result_thread.stop()

    def on_transcription_complete(self, result, tags=None, new_context_tag=None):
        """
        When the transcription is complete, either type the result (keyboard) or write to file.

        :param result: The transcribed text
        :param tags: List of matched tags (from tag detection)
        :param new_context_tag: New context tag signal ('clear', tag name, or None for no change)
        """
        if tags is None:
            tags = []

        # Update context tag state
        if new_context_tag == "clear":
            self.active_context_tag = None
        elif new_context_tag is not None:
            self.active_context_tag = new_context_tag

        if self.file_output_mode:
            # Write to file instead of typing
            output_file = os.getenv("WHISPER_OUTPUT_FILE")
            if output_file and output_file != "null":
                try:
                    mode = "a" if self.file_output_mode == "append" else "w"
                    with open(output_file, mode) as f:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

                        # Format tags as [tag1] [tag2]
                        tags_str = ""
                        if tags:
                            tags_str = " " + " ".join([f"[{tag}]" for tag in tags])

                        if self.file_output_mode == "append":
                            f.write("\n\n")
                            f.write(f"[{timestamp}]{tags_str} ")
                        else:
                            f.write(f"[{timestamp}]{tags_str} ")
                        f.write(result)
                    action = (
                        "appended to"
                        if self.file_output_mode == "append"
                        else "written to"
                    )
                    if ConfigManager.get_config_value("misc", "print_to_terminal"):
                        print(f"Transcription {action} {output_file}: {result}")
                except Exception as e:
                    print(f"Error writing to file {output_file}: {e}")
            else:
                print(
                    "WHISPER_OUTPUT_FILE not set or is null. Transcription: " + result
                )
            self.file_output_mode = False
        else:
            # Type the result using keyboard simulation
            self.input_simulator.typewrite(result)

        if ConfigManager.get_config_value("misc", "noise_on_completion"):
            AudioPlayer(os.path.join("assets", "beep.wav")).play(block=True)

        if (
            ConfigManager.get_config_value("recording_options", "recording_mode")
            == "continuous"
        ):
            self.start_result_thread()
        else:
            self.key_listener.start()

    def run(self):
        """
        Start the application.
        """
        sys.exit(self.app.exec_())


if __name__ == "__main__":
    app = WhisperWriterApp()
    app.run()
