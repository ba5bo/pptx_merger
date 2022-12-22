#from asyncore import readwrite
#from pptx import Presentation
import sys as sys
import glob
import os
import win32com.client


outputPresentationName = "merged_output.pptx"

# displays a [message] and exists the program


def exitWithMessage(message=''):
    print(message)
    sys.exit()

# displays the help information on running the program and calls exitWithMessage with input [message]


def displayHelpAndExit(message=''):
    print('please provide the names of the powerpoint files to merge (i.e. *.pptx for all pptx files in the directory)')
    exitWithMessage(message)


def getWorkingDirectory():
    # get the current working directory
    try:
        return os.getcwd()
    except:
        print('\nERROR> Could not get current directory: ',
              sys.exc_info()[0], 'occurred.')
        displayHelpAndExit()


def getFileNamesFromArguments():
    # looks into the program input arguments to create a list of powerpoint files to merge and returns it
    powerPointFileNames = []

    print('cmd entry:', sys.argv)
    numArgs = len(sys.argv)

    if numArgs == 1:
        # this means no arguments were given
        displayHelpAndExit()
    for i in range(1, numArgs):
        argument = sys.argv[i]
        if argument.find('*') != -1 or argument.find('?') != -1:
            # if wild cards are used, a list of names will be returned by the glob function. The + operator will merge the lists together
            powerPointFileNames = powerPointFileNames + glob.glob(argument)
        else:
            if argument.find('.ppt') != -1:
                powerPointFileNames.append(sys.argv[i])
            else:
                print(
                    f'Filename "{argument}" has no valid ".ppt" or ".pptx" extension. Skipped!')
    # remove possible duplicated file names and return the resulting list
    return list(dict.fromkeys(powerPointFileNames))


inputPresentationFiles = glob.glob("*.pptx")


def mergeSlides(inputPresentationFilesList, outputPresentationFileName):
    # function to copy all slides in all given input files [inputPresentationFilesList] into an output file named [outputPresentationFileName]
    initialRun = True
    outputPresentation = None
    # get the current working directory
    workingDirectory = getWorkingDirectory()
    outputFilePath = workingDirectory + '\\' + outputPresentationFileName
    # create the powerpoint instance object
    ppt_instance = win32com.client.Dispatch('PowerPoint.Application')

    for PresentationFileName in inputPresentationFilesList:
        filePath = workingDirectory + '\\' + PresentationFileName
        # open the powerpoint presentation in background
        try:
            inputPresentation = ppt_instance.Presentations.Open(
                filePath, ReadOnly=False, WithWindow=False)
        except:
            print(
                f'ERROR> Could not open file : "{filePath}". File Skipped!')
            continue
        if (initialRun == True):
            # uses the first file in the list as the initial slides to build upon
            print(
                f'Starting presentation {outputFilePath} with presentation {PresentationFileName}')
            try:
                inputPresentation.saveAs(outputFilePath)
            except:
                print(f'\nERROR> Could not save file : "{outputFilePath}"',
                      sys.exc_info()[0], 'occurred.')
                inputPresentation.Close()
                exitWithMessage('Quitting the program now!')
            inputPresentation.Close()
            try:
                outputPresentation = ppt_instance.Presentations.Open(
                    outputFilePath, ReadOnly=False, WithWindow=False)
            except:
                print(f'\nERROR> Could not open file : "{outputFilePath}"',
                      sys.exc_info()[0], 'occurred.')
                exitWithMessage('Quitting the program now!')
            initialRun = False
            continue  # back to the for loop

        print(
            f'Copying {inputPresentation.Slides.Count} slides from Presentation {PresentationFileName}')
        inputPresentation.Slides.Range(
            [1, inputPresentation.Slides.Count]).Copy()
        # Index=insert_index is omitted it will paste after the last slide
        outputPresentation.Slides.Paste()
        inputPresentation.Close()
    try:
        # save the powerpoint file
        print(f'Saving resulting presentation into {outputFilePath}')
        outputPresentation.save()
    except:
        print(f'\nERROR> Could not save pptx file : "{outputFilePath}"',
              sys.exc_info()[0], 'occurred.')
        outputPresentation.close()
        exitWithMessage('Quitting the program now!')
    try:
        # save as a pdf file for portability
        pdfFilePath = outputFilePath.replace(".pptx", ".pdf")
        print(f'Saving resulting presentation into {pdfFilePath}')
        # 32 is the value for ppSaveAsPDF
        outputPresentation.saveAs(pdfFilePath, 32)
    except:
        print(f'\nERROR> Could not save pdf file : "{pdfFilePath}"',
              sys.exc_info()[0], 'occurred.')
        outputPresentation.close()
        exitWithMessage('Quitting the program now!')
    outputPresentation.close()
    # kills ppt_instance
    ppt_instance.Quit()
    del ppt_instance


# main
# get all the file names given by command line argument, including wildcard characters based lists
inputPresentationFiles = getFileNamesFromArguments()
# exclude the output file from the input list, in case it exists
if outputPresentationName in inputPresentationFiles:
    inputPresentationFiles.remove(outputPresentationName)
# call the function to copy all slides in all given input files [inputPresentationFiles] into an output file named [outputPresentationName]
mergeSlides(inputPresentationFiles, outputPresentationName)
print('All done!')
