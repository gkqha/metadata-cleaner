import argparse
import os
import win32com.client as win32com

__author__ = "Istvan David"
__license__ = "MIT"
__version__ = "1.0.0"


class MdCleaner:
    def run(self, rootFolder):
        print("Running cleaner on directory {}".format(rootFolder))
        for root, dirs, files in os.walk(rootFolder):
            path = root.split(os.sep)
            for file in files:
                if file.endswith(".xlsx"):
                    fileFullPath = root + "\\" + file
                    print(fileFullPath)

                    excel = win32com.gencache.EnsureDispatch("Excel.Application")
                    wb = excel.Workbooks.Open(fileFullPath, Local=True)
                    wb.RemovePersonalInformation = True
                    wb.Close(SaveChanges=1)
                    excel.Quit()
                elif file.endswith(".docx"):
                    fileFullPath = root + "\\" + file
                    print(fileFullPath)

                    word = win32com.gencache.EnsureDispatch("Word.Application")
                    print(word)
                    wb = word.Documents.Open(fileFullPath)
                    wb.RemovePersonalInformation = True
                    wb.Close()
                    word.Quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-root",
        default=str(os.getcwd),
        help=(
            "Provide root directory without trailing separator."
            "Example '-root d:\my_folder'"
        ),
    )

    options = parser.parse_args()
    if options.root == "":
        rootFolder = str(pathlib.Path.cwd())
    else:
        rootFolder = options.root

    print(rootFolder)

    MdCleaner().run(rootFolder)
