import tkinter as tk
from tkinter import filedialog
import os
import shutil
from ppt2png import ppt2png
import sys
from PIL import Image, ImageTk
from GestureControl import main

# Variable to keep track of whether PNG files have been created
png_files_created = False


# Function to upload and copy a PowerPoint (PPTX) file
def upload_and_copy_ppt():
    global png_files_created  # Use the global variable
    if os.path.exists('Presentation'):
        shutil.rmtree('Presentation')
    os.makedirs('Presentation')
    file_path = filedialog.askopenfilename(
        filetypes=[("PowerPoint Files", "*.pptx")],
        title="Select a PowerPoint (PPTX) file"
    )

    if file_path:
        try:
            # Define the destination folder and filename
            destination_folder = os.path.join(os.getcwd(), 'Presentation')
            destination_file = os.path.join(destination_folder, 'Presentation.pptx')

            # Create the destination folder if it doesn't exist
            os.makedirs(destination_folder, exist_ok=True)

            # Copy the uploaded file to the destination with the desired filename
            shutil.copy(file_path, destination_file)

            # Convert the PowerPoint to PNG images (if needed)
            ppt2png.ppt_to_png()

            png_files_created = True  # Set the flag to True
            result_label.config(text=f"PPT file Is Ready For Presentation.", fg="green")

            # Enable the "PRESENT" button
            present_button.config(state="normal")
            upload_button.config(state="disabled")

        except Exception as e:
            result_label.config(text=f"Error: {e}", fg="red")


# Function to restart the application
def restart_application():
    os.execv(sys.executable, ['python'] + sys.argv)


# Function to start the presentation (to be called when "PRESENT" button is clicked)
def start_presentation():
    if png_files_created:
        main.Capture_gestures()

    else:
        result_label.config(text="Error: PNG files not created.", fg="red")


# Delete and create the 'Presentation' folder using folderUtils
if os.path.exists('Presentation'):
    shutil.rmtree('Presentation')
os.makedirs('Presentation')

root = tk.Tk()
root.title("Gesture Controlled Presentation")
root.geometry("600x300")  # Increase the window size

# Set background color to #a9dce3
root.configure(bg="#a9dce3")

info_label = tk.Label(root, text="Upload a PowerPoint (PPTX) file", font=("Arial", 16), bg="#a9dce3")
info_label.pack(pady=20)

upload_button = tk.Button(root, text="Select File", command=upload_and_copy_ppt, width=20, height=2, font=("Arial", 14))
upload_button.pack()

# Load a refresh icon image and convert it to a Tkinter PhotoImage
refresh_image = Image.open("Static\\refresh_icon.png")  # Replace "refresh_icon.png" with your image file
refresh_image = refresh_image.resize((32, 32))  # Resize the image as needed
refresh_image = ImageTk.PhotoImage(refresh_image)

# Create a button with the refresh icon
refresh_button = tk.Button(root, image=refresh_image, command=restart_application, bd=0, bg="#a9dce3")
refresh_button.image = refresh_image  # Keep a reference to the image to prevent it from being garbage collected
refresh_button.place(x=10, y=10)

# Create a "PRESENT" button (initially disabled) after successfully creating PNG files
present_button = tk.Button(root, text="Present", command=start_presentation, width=20, height=2, font=("Arial", 14))
present_button.pack()
present_button.config(state="disabled")  # Initially disable the button
present_button.place(x=187, y=187)

result_label = tk.Label(root, text="", font=("Arial", 14), bg="#a9dce3")
result_label.pack(pady=10)

root.mainloop()
