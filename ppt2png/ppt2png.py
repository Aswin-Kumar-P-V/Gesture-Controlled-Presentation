import win32com.client
import os

# Specify the full path to the PowerPoint presentation
presentation_path = r"C:\s7MultiMediaProject\Presentation\Presentation.pptx"

# Create a PowerPoint application object
Application = win32com.client.Dispatch("PowerPoint.Application")
def ppt_to_png():
    try:
        # Open the presentation without making it visible
        Presentation = Application.Presentations.Open(presentation_path, WithWindow=False)

        # Create a folder to save the slides as images
        slides_folder = os.path.join(os.path.dirname(presentation_path), "Slides")
        if not os.path.exists(slides_folder):
            os.makedirs(slides_folder)

        # Export each slide as an image
        for i, slide in enumerate(Presentation.Slides):
            image_path = os.path.join(slides_folder, f"{i + 1}.png")
            slide.Export(image_path, "JPG")

        # Close the presentation
        Presentation.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the PowerPoint application
        Application.Quit()
