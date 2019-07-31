import os, sys, re
from tkinter import messagebox
from datetime import datetime 
import xlsxwriter

def getFunctionName(fileNameWithDir):
    base=os.path.basename(fileNameWithDir)
    funName = os.path.splitext(base)[0]
    funName = funName.replace("_dyn_results","")
    funName = funName.replace("_Dyn_results","")
    return funName

def writeUTResultsIntoReport(UTRes, statusBarText, TkObject_ref, userInfoFileActivated):

    global worksheet_html, workbookObj
    

    ###############33 Creating Xlsheet to write the data into that if it is not exist 	#########
    
    dt = datetime.now()
    day_hour_min_sec = str(dt.day)+"_"+str(dt.hour)+"_"+str(dt.minute)+"_"+str(dt.second)
    testSheetName = 'Unit testing coverage status report_Build XXX'+str(day_hour_min_sec)+'.xlsx'	

    try:
        workbookObj = xlsxwriter.Workbook(testSheetName)
    except:
        messagebox.showerror("Error","Please close "+str(testSheetName)+" xlsx file if it is opened")
        sys.exit()

    worksheet_html = workbookObj.add_worksheet("XXXXXX Coverage Status")


    #################### Write header ##########################
    merge_format = workbookObj.add_format({'bg_color': '#DDD9C4', 'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 14, "font_name": "Alstom"})
    merge_format2 = workbookObj.add_format({'bg_color': '#DDD9C4', 'bold': 1, 'align': 'center', 'valign': 'vcenter', 'size': 12, "font_name": "Alstom"})
    bold_header = workbookObj.add_format({'bg_color': '#A6A6A6', 'bold': 1, 'align': 'center', 'size': 12, "font_name": "Alstom"})		

    worksheet_html.merge_range('B2:J2', '                Unit Test Report - Build XXX - XXXXX', merge_format)		
    worksheet_html.merge_range('B3:J3', '                  Build Tested : XXXXXXXXXX', merge_format2)		
    worksheet_html.merge_range('B4:J4', 'Note:\nThis report contains the report of Files that has changed\ functionally from build XXXXXXXXXX to Build XXX.\nNA = Not Applicable because there are no branches and no conditions', merge_format2)		

    worksheet_html.write('B6',".C /.H File Changed", bold_header)		
    worksheet_html.write('C6',"Current revision Tested", bold_header)		
    worksheet_html.write('D6',"Functions", bold_header)		
    worksheet_html.write('E6',"Status", bold_header)		
    worksheet_html.write('F6',"Statement Coverage", bold_header)		
    worksheet_html.write('G6',"Branch Coverage", bold_header)		
    worksheet_html.write('H6',"MCDC", bold_header)		
    worksheet_html.write('I6',"Tester's comments", bold_header)		
    worksheet_html.write('J6',"Verifier's comments", bold_header)	


    #Print UT results into report	
    reportWriteIndx = 7
    for report in UTRes:
        worksheet_html.write('B'+str(reportWriteIndx), UTRes[report]["SourcefileName"])
        worksheet_html.write('D'+str(reportWriteIndx), report)

        if 'Statement' in UTRes[report]:
            worksheet_html.write('F'+str(reportWriteIndx), UTRes[report]["Statement"])

        if 'Branch' in UTRes[report]:			
            worksheet_html.write('G'+str(reportWriteIndx), UTRes[report]["Branch"])

        if 'MC/DC' in UTRes[report]:
            worksheet_html.write('H'+str(reportWriteIndx), UTRes[report]["MC/DC"])

        reportWriteIndx = reportWriteIndx+1


    ################### Final closure  ##########
    workbookObj.close()
    statusBarText.set("DONE!!")	

    if userInfoFileActivated:
        messagebox.showinfo('DONE!!',"UT Report genarated!. Please look at userInfo.txt file")
    else:		
        messagebox.showinfo('DONE!!',"UT Report genarated!")


    TkObject_ref.destroy()


def script_exe(DirLoc, TkObject_ref, statusBarText):

    statusBarText.set("Selected Directory analysing for report")
    
    listOfReportNames = []
    for file in os.listdir(DirLoc):
        if file.endswith(".htm"):
            listOfReportNames.append(os.path.join(DirLoc, file))
    
    finalCoverageReport = {}
    progressCounter = 1
    fPtr = open("userInfo.txt", "w")	

    if (len(listOfReportNames) == 0):
        messagebox.showerror('No Htm files','There are no Htm files in selected path!')
        TkObject_ref.destroy()
        sys.exit()
    
    try:
        for fileName in listOfReportNames:
            functionName = getFunctionName(fileName)
            statusBarText.set("Number of Reports processed: ("+str(progressCounter)+"/"+str(len(listOfReportNames))+")"+". Funtion Name:"+functionName+".")
            progressCounter = progressCounter+1
        
            functionName_Report = functionName
            #remove macros if it has anything
        
            if "!PTCP_APPLICATION" in functionName:
                functionName = functionName.replace("_!PTCP_APPLICATION","")
            elif "PTCP_APPLICATION" in functionName:
                functionName = functionName.replace("_PTCP_APPLICATION","")		
        
            #Pattern which fetches coverage report line number from given report	
            patternfileName = re.compile("<H1> File :")
            patternHeader = re.compile("<TT> Procedure </TT>")
            patternRes = re.compile(".*<a href=.*> "+functionName+" </a>")
           
            
            ############# Fetch File Name   ###############
        
            fileNameMatchLineNumer = 0
            SourcefileName = ''
            
            for i, line in enumerate(open(fileName)):
                for match in re.finditer(patternfileName, line):
                    fileNameMatchLineNumer = i
            
        
            line = open(fileName, "r").readlines()[fileNameMatchLineNumer]		
            substr = re.findall(r'<H1> File :(.*)',line)
            if(len(substr) > 0):
                SourcefileName = substr[0].strip()
                fileName_afterslipt = SourcefileName.split("\\")
                SourcefileName = fileName_afterslipt[len(fileName_afterslipt)-1]

            ############# Fetch report header   ###############
            
            headerMatchLineNumer = 0
            fileHeader = []
            exitLoop = False
        
            for i, line in enumerate(open(fileName)):
                for match in re.finditer(patternHeader, line):
                    headerMatchLineNumer = i
            
            while ((not exitLoop) and (headerMatchLineNumer > 0)): 		
                line = open(fileName, "r").readlines()[headerMatchLineNumer+1]		
                if "</TR>" not in line:
                    substr = re.findall(r'<TT>(.*)</TT>',line)
                    if(len(substr) > 0):
                        fileHeader.append(substr[0].strip())
                else:
                    exitLoop = True
            
                headerMatchLineNumer = headerMatchLineNumer+1

            ############### Fetch the line numbers where coverage report will be given ##########3
            
            exitLoop = False
            stringMatchLineNumer = 0
            
            for i, line in enumerate(open(fileName)):
                for match in re.finditer(patternRes, line):
                    stringMatchLineNumer = i
            
            functionReport = {}
            indx = 0 
            userInfoFileActivated = False
            functionReport["SourcefileName"] = SourcefileName
            
            #Fetch the coverage report
            while ((not exitLoop) and (stringMatchLineNumer > 0)):
            
                line = open(fileName, "r").readlines()[stringMatchLineNumer]
            
                #Break both the loops if </TR> found which indicates end of the row
                if "</TR>" in line:
                    exitLoop = True
                else:
                    substr = re.findall(r'> .*\+([\d]+).* </font>',line)
                    if(len(substr) > 0): 
                        functionReport[fileHeader[indx]] = substr[0]
                    indx = indx+1
                    stringMatchLineNumer = stringMatchLineNumer+1

            if ((stringMatchLineNumer == 0) and (headerMatchLineNumer == 0)):
                userInfoFileActivated = True
                fPtr.writelines("\nLooks like function: "+functionName_Report+" is not a dynamic coverage report or report format is incorrect.")
            else:
                finalCoverageReport[functionName_Report] = functionReport

        #Write Results into Excel sheet
        fPtr.close()
        writeUTResultsIntoReport(finalCoverageReport, statusBarText, TkObject_ref, userInfoFileActivated)

    except Exception as e:
        messagebox.showerror('Exception raised','Following exception raised while UT Report generation! \n\n'+str(e))
        TkObject_ref.destroy()
        sys.exit()        
