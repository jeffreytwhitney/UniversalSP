clr.AddReference("mscorlib")
clr.AddReference("SmartCore.Model")
clr.AddReference("Extensions.GDnt")
import os
import sys
from datetime import date
from datetime import datetime
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


def GetImportMetadata(importFilepath):	
	with open(importFilepath, "r") as file:
		data = file.readlines()
		partNumberLine = data[0].split("\t")[1].strip('"').strip('\n').rstrip('"')
		employeeIdLine = data[1].split("\t")[1].strip('"').strip('\n').rstrip('"')
        jobNumberLine = data[2].split("\t")[1].strip('"').strip('\n').rstrip('"')
        machineNumberLine = data[3].split("\t")[1].strip('"').strip('\n').rstrip('"')
        return ImportMetadata(partNumberLine, employeeIdLine, jobNumberLine, machineNumberLine)


def GetReportFilepath(importMetadata):
    reportPath = "S:\\Smart Profile\\"
    reportPath += importMetadata.PartNumber
    reportPath += "_" + importMetadata.MachineNumber
    reportPath += "_" + datetime.now().strftime('%Y%m%d%H%M%S') + ".pdf"
    return reportPath.upper()	


def GetProjectFilepath(importMetadata):
    smartProfileFilePath = "V:\\Inspect Programs\\Micro-Vu\\Approved Programs\\Smart Profile\\"
    smartProfileFilePath += importMetadata.PartNumber + ".spp"
    return smartProfileFilePath.upper()


def GetImporterOptions():
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

	
def Execute():

	importFilepath = "C:\\Text\\OUTPUT.txt"
	importMetadata = GetImportMetadata(importFilepath)
	reportFilepath = GetReportFilepath(importMetadata)
	projectFilePath = GetProjectFilepath(importMetadata)
	importerOptions = GetImporterOptions()

	StartNewIteration()
	ChangeProject(projectFilePath)
	SelectAllMeasuredPoints()
	DeleteSelected()
	ImportMeasuredPoints(importFilepath, importerOptions)
	
	ResetAlignment()
	QuickAlign()
	Evaluate()
	File.Delete(importFilepath)
	context.Project.Document.ProjectData.Metadata = Metadata("JOB_NUMBER", importMetadata.JobNumber)
	context.Project.Document.ProjectData.Metadata = Metadata("EMPLOYEE_NUMBER", importMetadata.EmployeeID)
	context.Project.Document.ProjectData.Metadata = Metadata("MACHINE_NUMBER", importMetadata.MachineNumber)
	
	reporting = GetExtensionContext("reporting")
	reporting.ExportReport("Default", reportFilepath, False)



Execute()
Quit()

