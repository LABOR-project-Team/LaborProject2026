import subprocess
import os
import sys


def create_and_run():
    script_name = "app.py"

    # --add-data "source_path;dest_path"
    # On Windows, use ; as the separator. On Linux/Mac, use :
    # This example assumes you have an 'images' folder and an 'icon.ico' in your root
    add_data = [
        "--add-data", "images;images",
        "--add-data", "icon.ico;.",
        # If you have other folders (e.g., 'phases', 'ui', 'utils'),
        # add them here if they are not picked up automatically:
        # "--add-data", "phases;phases",
        # "--add-data", "ui;ui",
        # "--add-data", "utils;utils"
    ]

    print("Building executable...")
    # Clean up previous builds to avoid conflicts
    if os.path.exists("dist"):
        import shutil
        shutil.rmtree("dist")

    cmd = [
              sys.executable, "-m", "PyInstaller",
              "--onefile",
              "--noconsole",  # Hides the black terminal window
              "--icon=icon.ico"
          ] + add_data + [script_name]

    try:
        subprocess.check_call(cmd)
        print("Build successful.")
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        return

    # Run it
    exe_path = os.path.join("dist", "app.exe")
    if os.path.exists(exe_path):
        subprocess.Popen([exe_path])


if __name__ == "__main__":
    create_and_run()