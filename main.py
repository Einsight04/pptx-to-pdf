import os
import comtypes.client
from reportlab.pdfgen import canvas


def save_slide_as_image(presentation, slide_number, output_path):
    export_path = os.path.abspath(output_path)
    presentation.Slides[slide_number].Export(export_path, 'PNG')


def convert_ppt_to_pdf(input_folder_path, output_folder_path):
    input_folder_path = os.path.abspath(input_folder_path)
    output_folder_path = os.path.abspath(output_folder_path)

    input_file_paths = os.listdir(input_folder_path)

    powerpoint_app = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint_app.Visible = 1

    for input_file_name in input_file_paths:
        if not input_file_name.lower().endswith((".ppt", ".pptx")):
            continue

        input_file_path = os.path.join(input_folder_path, input_file_name)
        prs = powerpoint_app.Presentations.Open(input_file_path)
        file_name = os.path.splitext(input_file_name)[0]
        output_file_path = os.path.join(output_folder_path, file_name + ".pdf")

        c = canvas.Canvas(output_file_path)

        for slide_number in range(1, len(prs.Slides) + 1):
            slide_image_path = os.path.join(output_folder_path, f"{file_name}_slide_{slide_number}.png")
            save_slide_as_image(prs, slide_number, slide_image_path)
            c.drawImage(slide_image_path, 0, 0, width=c._pagesize[0], height=c._pagesize[1])
            c.showPage()
            os.remove(slide_image_path)

        c.save()
        prs.Close()

    powerpoint_app.Quit()


if __name__ == "__main__":
    INPUT_FOLDER_PATH = ""
    OUTPUT_FOLDER_PATH = ""

    convert_ppt_to_pdf(INPUT_FOLDER_PATH, OUTPUT_FOLDER_PATH)
