clr.AddReference("mscorlib")
clr.AddReference("SmartCore.Model")
clr.AddReference("Extensions.GDnt")
import os
from datetime import datetime
import time
from System import Console
from System.IO import File, Path
from System.Threading import Thread
from SmartCore.Model.Public import Metadata


class ImportMetadata:
    def __init__(self, part_number, employee_id, job_number, machine_number):
        self.PartNumber = part_number
        self.EmployeeID = employee_id
        self.JobNumber = job_number
        self.MachineNumber = machine_number


def get_import_metadata(import_filepath):
    with open(import_filepath, "r") as file:
        data = file.readlines()
        part_number_line = data[0].split("\t")[1].strip('"').strip('\n').rstrip('"')
        employee_id_line = data[1].split("\t")[1].strip('"').strip('\n').rstrip('"')
        job_number_line = data[2].split("\t")[1].strip('"').strip('\n').rstrip('"')
        machine_number_line = data[3].split("\t")[1].strip('"').strip('\n').rstrip('"')
    return ImportMetadata(part_number_line, employee_id_line, job_number_line, machine_number_line)


def get_report_filepath(import_metadata):
    report_path = "S:\\Smart Profile\\"
    report_path += import_metadata.PartNumber
    report_path += "_" + import_metadata.MachineNumber
    report_path += "_" + datetime.now().strftime('%Y%m%d%H%M%S') + ".pdf"
    return report_path.upper()


def get_project_filepath(import_metadata):
    smart_profile_file_path = "V:\\Inspect Programs\\Micro-Vu\\Approved Programs\\Smart Profile\\"
    smart_profile_file_path += import_metadata.PartNumber + ".spp"
    return smart_profile_file_path.upper()


def get_importer_options():
    options = CreatePointImporterOptions()
    options.CommaIsDecimalPoint = False
    options.CommaIsSeparator = False
    options.CommentMarkerStr = "/\r\n"
    options.CutCommentEnable = False
    options.DoubleQuoteIsTextMarker = True
    options.IgnoreLeadingWhiteSpace = False
    options.IgnoreWhiteSpacesAroundSeparators = False
    options.JoinConclusiveWhiteSeparators = False
    options.JoinConclusiveWhiteSpaces = True
    options.EnableOtherSeparators = False
    options.OtherSeparators = ":"
    options.SemicolonIsSeparator = False
    options.SingleQuoteIsTextMarker = False
    options.SkipLineMarker = "/"
    options.EnableSkipLineMarker = False
    options.SpaceIsWhiteSpace = False
    options.TabIsWhiteSpace = True
    options.TipRadiusSourceColumn = -1
    options.ZSourceColumn = -1
    return options


def execute():
    import_filepath = "C:\\Text\\OUTPUT.txt"

    if not os.path.exists(import_filepath):
        for count in range(5):
            time.sleep(1)
            if os.path.exists(import_filepath):
                break

    if not os.path.exists(import_filepath):
        MessageBox("'C:\\Text\\OUTPUT.txt' does not exist.")
        return

    try:
        import_metadata = get_import_metadata(import_filepath)
    except ValueError:
        MessageBox("Invalid data in c:\\text\\output.txt")
        return
    except IOError:
        MessageBox("Output file is locked.")
        return

    report_filepath = get_report_filepath(import_metadata)
    project_file_path = get_project_filepath(import_metadata)
    importer_options = get_importer_options()

    StartNewIteration()
    ChangeProject(project_file_path)
    SelectAllMeasuredPoints()
    DeleteSelected()
    ImportMeasuredPoints(import_filepath, importer_options)

    ResetAlignment()
    QuickAlign()
    Evaluate()
    File.Delete(import_filepath)
    context.Project.Document.ProjectData.Metadata = Metadata("JOB_NUMBER", import_metadata.JobNumber)
    context.Project.Document.ProjectData.Metadata = Metadata("EMPLOYEE_NUMBER", import_metadata.EmployeeID)
    context.Project.Document.ProjectData.Metadata = Metadata("MACHINE_NUMBER", import_metadata.MachineNumber)

    reporting = GetExtensionContext("reporting")
    reporting.ExportReport("Default", report_filepath, False)


execute()
Quit()
