import os
import glob
import shutil
import fnmatch
import easygui


source = easygui.enterbox("Source Location", "Attachement Sorting")
bom_destination = source + "\BOMs"
production_destination = source + "\RBM Production"
specs_destination = source + "\Specs"
req_destination = source + "\Part Requests"
eco_destination = source + "\ECOs"
misc_sw_destination = source + "\Misc Solidworks"
bom_pattern = "??-??-????*"
production_pattern = "*86?????*"
spec_pattern = "*M1????*"

#move files
def move_files(file, destination_path):
    file_path = source + "\\" + file
    shutil.move(file_path, destination_path)

def is_BOM(file_name):
    if fnmatch.fnmatch(file_name[:10], bom_pattern) is True:
        for i in [".xlsx", ".xlsb", ".xlsm"]:
            if file_name.endswith(i) is True:
                return True

def is_Production(filename):
    if fnmatch.fnmatch(filename, production_pattern) is True:
        for i in [".pdf", ".dxf", ".dwg", ".sld*"]:
            if filename.endswith(i) is True:
                return True

def is_PartRequest(filename):
    for i in [".xlsx", ".xlsb", ".xlsm"]:
        if filename.endswith(i) is True:
            try:
                if int(filename[:-5]) > 50000:
                    return True
            except ValueError:
                return False


def is_ECO(filename):
    for i in [".xlsx", ".xlsb", ".xlsm"]:
        if filename.endswith(i) is True:
            try:
                if int(filename[:-5]) < 50000:
                    return True
            except ValueError:
                return False

def is_Spec(filename):
    if fnmatch.fnmatch(filename[:7], spec_pattern) is True:
        if filename.endswith("*.pdf") is True:
            return True

def is_Misc_SW(filename):
    if fnmatch.fnmatch(filename, "*.sld*") is True:
        return True

def build_structure(path):
    dirs = ["\BOMs","\ECOs","\RBM Production","\Misc","\Misc Solidworks","\Part Requests","\Specs","\Misc\Images",
            "\Misc\Excel","\Misc\Zip Files", "\Misc\Docs", "\Misc\PDF Files"]
    for i in dirs:
        os.mkdir(path + i)

def parse_misc(file):
    try:
        file_type = file.split(".", 1)[1]
        if file_type == "xlsx" or file_type == "xlsb" or file_type == "xlsm":
            move_files(file, source + "\Misc\Excel")
        elif file_type == "pdf":
            move_files(file, source + "\Misc\PDF Files")
        elif file_type == "jpg" or file_type == "bmp" or file_type == "png":
            move_files(file, source + "\Misc\Images")
        elif file.endswith(".zip"):
            move_files(file, source + "\Misc\Zip Files")
        elif file_type == "docx" or file_type == "doc":
            move_files(file, source + "\Misc\Docs")
    except IndexError:
        pass


# iterate on all files to move them to destination folder
def parse_files():
    #create folder structure if it doesn't exist
    try:
        build_structure(source)
    except FileExistsError:
        pass
    # gather all files
    attachements_folder = os.listdir(source)
    allfiles = [f for f in attachements_folder if os.path.isfile(source+'/'+f)]
    for file_name in allfiles:#allfiles:
        if is_BOM(file_name) is True:
            move_files(file_name,bom_destination)
        elif is_Production(file_name) is True:
            move_files(file_name, production_destination)
        elif is_Spec(file_name) is True:
            move_files(file_name, specs_destination)
        elif is_Misc_SW(file_name) is True:
            move_files(file_name, misc_sw_destination)
        elif is_ECO(file_name) is True:
            move_files(file_name, eco_destination)
        elif is_PartRequest(file_name) is True:
            move_files(file_name, req_destination)
        else:
            parse_misc(file_name)
