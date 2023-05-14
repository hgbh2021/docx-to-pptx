from __future__ import print_function

import os
import json
import aspose.slides as slides
import aspose.words as aw
from utils.detector import is_pdf, is_ppt, is_docx
from utils.convertor import convert_ppt_to_pdf, convert_docx_to_pdf


class CheckMail: 
    def __init__(self, root_dir=os.getcwd()):
        self.root_dir = root_dir

    def convert_to_pdf(self, file_path, folder_path):
        # This will convert docx to pdf
        file = file_path.split('/')[-1]
        if is_docx(file_path):
            pdf_file = f'{folder_path}/{file.split(".")[0]}.pdf'
            convert_docx_to_pdf(file_path, pdf_file)


    def remove_empty_dir(self, dir):
        #This will check if the directory is empty or not and if empty, it will remove it.
        try:
            if(os.listdir(dir) == []):
                os.rmdir(dir)
        except Exception as e:
            print(f"Error: {e}.")



    def check_dir(self, dir):
        """This will check if the directory is present or not. If not then it will create it."""
        try:
            os.listdir(dir)
        except FileNotFoundError as e:
            os.mkdir(dir)

    def run(self, dir='Attachments'):
        self.check_dir(dir)
        self.navigate_folder(dir)
        
    def pdf_to_pptx(self, file_path, folder_path):
        with slides.Presentation() as presentation:
            file = file_path.split('/')[-1]
            presentation.slides.remove_at(0)

            presentation.slides.add_from_pdf( f'{folder_path}/{file.split(".")[0]}.pdf')
            presentation.save("presentation.ppt", slides.export.SaveFormat.PPT)
if __name__ == '__main__':
    check_mail = CheckMail()
    check_mail.run()