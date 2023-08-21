from openpyxl import load_workbook
from warnings import filterwarnings
from threading import Thread
from time import time
from os.path import join, dirname, splitext, basename, exists
from tkinter import Tk, filedialog, Label, Button, Entry, Toplevel, Canvas, Checkbutton, IntVar
from datetime import datetime
from tkinter.ttk import Progressbar, Style, Combobox
from tkinter.scrolledtext import ScrolledText
from random import randint
from xlwings import Book, App

fileHandler = open(f"logs_{datetime.now().strftime('%Y%m%d%H%M%S')}.txt", 'a')

iconFile = join(dirname(__file__), 'Excel_Mac_23559.ico')
aboutIcon = join(dirname(__file__), 'info.ico')

sourceExcelFile = ''

targetExcelFile = ''


def threadingSourceExcelFileDialogFunc(progressBar, progressStyle, sourceExcelEntry, sourceExcelFileDialogBtn,
                                       sourceExcelSheetCombo, resetBtn, messageText, sourceExcelSheetComboValues):
    threadSourceExcelFileDialog = Thread(target=sourceExcelFileDialogFunc, args=(progressBar, progressStyle,
                                                                                 sourceExcelEntry,
                                                                                 sourceExcelFileDialogBtn,
                                                                                 sourceExcelSheetCombo,
                                                                                 resetBtn, messageText,
                                                                                 sourceExcelSheetComboValues))
    threadSourceExcelFileDialog.start()


def sourceExcelFileDialogFunc(progressBar, progressStyle, sourceExcelEntry, sourceExcelFileDialogBtn,
                              sourceExcelSheetCombo, resetBtn, messageText, sourceExcelSheetComboValues):
    global sourceExcelFile
    messageText.config(text='')
    progressBar.config(value=0)
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='0 %')
    sourceExcelSheetComboValues.clear()
    sourceExcelSheetComboValues.append('Select WorkSheet')
    sourceExcelFile = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),))
    sourceExcelEntry.config(state='normal')
    sourceExcelEntry.delete(0, 'end')
    sourceExcelEntry.insert(0, sourceExcelFile)
    sourceExcelEntry.config(state='readonly')
    filterwarnings("ignore", category=UserWarning)
    if len(sourceExcelEntry.get()) > 0:
        fileHandler.write(
            f'{datetime.now().replace(microsecond=0)} [{sourceExcelFile}] is selected as source excel file\n')
        sourceExcelFileDialogBtn.config(state='disabled', bg='light grey')
        try:
            for sourceExcelSheet in load_workbook(sourceExcelEntry.get()).sheetnames:
                sourceExcelSheetComboValues.append(sourceExcelSheet)
            sourceExcelSheetCombo.config(state='normal')
            sourceExcelSheetCombo.config(values=sourceExcelSheetComboValues)
            resetBtn.config(state='normal', bg='red')
            load_workbook(sourceExcelEntry.get()).close()
        except PermissionError as permError:
            fileHandler.write(
                f'{datetime.now().replace(microsecond=0)} [Permission Error] [Permission Denied] [{sourceExcelFile}] '
                f'is opened by another application, close it first\n')
            fileHandler.write(f'{datetime.now().replace(microsecond=0)} [ERROR][{permError}]\n')
            messageText.config(text='Close excel file, in use by another app')
            sourceExcelFileDialogBtn.config(state='normal', bg='green')
        except IndexError as indError:
            fileHandler.write(
                f'{datetime.now().replace(microsecond=0)} [Index Error] Unable to access the sheets of excel file '
                f'[{sourceExcelFile}]. Possibly because excel file is not correctly saved before\n')
            fileHandler.write(f'{datetime.now().replace(microsecond=0)} [ERROR][{indError}]\n')
            messageText.config(text='Error, Try again with different excel file')
            sourceExcelFileDialogBtn.config(state='normal', bg='green')

    else:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} No source excel file selected\n')
        messageText.config(text='Select source Excel')
    fileHandler.flush()


def threadingTargetExcelFileDialogFunc(progressBar, progressStyle, targetExcelSheetComboValues, targetExcelEntry,
                                       targetExcelFileDialogBtn, targetExcelSheetCombo, messageText, resetBtn):
    threadTargetExcelFileDialog = Thread(target=TargetExcelFileDialogFunc, args=(progressBar, progressStyle,
                                                                                 targetExcelSheetComboValues,
                                                                                 targetExcelEntry,
                                                                                 targetExcelFileDialogBtn,
                                                                                 targetExcelSheetCombo,
                                                                                 messageText, resetBtn))
    threadTargetExcelFileDialog.start()


def TargetExcelFileDialogFunc(progressBar, progressStyle, targetExcelSheetComboValues, targetExcelEntry,
                              targetExcelFileDialogBtn, targetExcelSheetCombo, messageText, resetBtn):
    global targetExcelFile
    messageText.config(text='')
    progressBar.config(value=0)
    resetBtn.config(state='disabled', bg='light grey')
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='0 %')
    targetExcelSheetComboValues.clear()
    targetExcelSheetComboValues.append('Select WorkSheet')
    targetExcelFile = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"),))
    targetExcelEntry.config(state='normal')
    targetExcelEntry.delete(0, 'end')
    targetExcelEntry.insert(0, targetExcelFile)
    targetExcelEntry.config(state='readonly')
    if len(targetExcelEntry.get()) > 0:
        if sourceExcelFile == targetExcelFile:
            fileHandler.write(
                f'{datetime.now().replace(microsecond=0)} Same source and target excel file selected. Choose different '
                f'file\n')
            messageText.config(text='Choose different Excel file; target and source cannot be same')
        else:
            fileHandler.write(
                f'{datetime.now().replace(microsecond=0)} [{targetExcelFile}] is selected as target excel file\n')
            targetExcelFileDialogBtn.config(state='disabled', bg='light grey')
            try:
                for targetExcelSheet in load_workbook(targetExcelEntry.get()).sheetnames:
                    targetExcelSheetComboValues.append(targetExcelSheet)
                load_workbook(targetExcelEntry.get()).close()
                targetExcelSheetCombo.config(state='normal')
                targetExcelSheetCombo.config(values=targetExcelSheetComboValues)
            except PermissionError as permError:
                fileHandler.write(
                    f'{datetime.now().replace(microsecond=0)} [Permission Error] [Permission Denied] '
                    f'[{targetExcelFile}] is opened by another application, close it first\n')
                fileHandler.write(f'{datetime.now().replace(microsecond=0)} [ERROR][{permError}]\n')
                messageText.config(text='Close excel file, in use by another app')
                targetExcelFileDialogBtn.config(state='normal', bg='green')
            except IndexError as indError:
                fileHandler.write(
                    f'{datetime.now().replace(microsecond=0)} [Index Error] Unable to access the sheets of excel file '
                    f'[{targetExcelFile}]. Possibly because excel file is not correctly saved before\n')
                fileHandler.write(f'{datetime.now().replace(microsecond=0)} [ERROR][{indError}]\n')
                messageText.config(text='Error, Try again with different excel file')
                targetExcelFileDialogBtn.config(state='normal', bg='green')
    else:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} No target excel file selected\n')
        messageText.config(text='Select target Excel')
    resetBtn.config(state='normal', bg='red')
    fileHandler.flush()


def threadingMatchCopyPaste(sourceExcelEntry, sourceExcelSheetCombo, sourceReferenceColumnCombo,
                            sourceExcelFileDialogBtn, sourceCopyColumnsScrolledText, targetExcelEntry,
                            targetExcelSheetCombo, targetReferenceColumnCombo, targetReferenceColumnComboValues,
                            targetExcelFileDialogBtn, targetPasteColumnsScrolledText, resetBtn, messageText, window,
                            progressBar, progressStyle):
    threadMatchCopyPaste = Thread(target=matchCopyPaste, args=(sourceExcelEntry, sourceExcelSheetCombo,
                                                               sourceReferenceColumnCombo, sourceExcelFileDialogBtn,
                                                               sourceCopyColumnsScrolledText, targetExcelEntry,
                                                               targetExcelSheetCombo, targetReferenceColumnCombo,
                                                               targetReferenceColumnComboValues,
                                                               targetExcelFileDialogBtn, targetPasteColumnsScrolledText,
                                                               resetBtn, messageText, window, progressBar,
                                                               progressStyle))
    threadMatchCopyPaste.start()


def matchCopyPaste(sourceExcelEntry, sourceExcelSheetCombo, sourceReferenceColumnCombo, sourceExcelFileDialogBtn,
                   sourceCopyColumnsScrolledText, targetExcelEntry, targetExcelSheetCombo, targetReferenceColumnCombo,
                   targetReferenceColumnComboValues, targetExcelFileDialogBtn, targetPasteColumnsScrolledText, resetBtn,
                   messageText, window, progressBar, progressStyle):
    messageText.config(text='')
    for checkbox in copyCheckboxes.keys():
        checkbox.config(state='disabled')
    submitBtn.config(state='disabled', bg='light grey')
    resetBtn.config(state='disabled', bg='light grey')
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Column {copySelectedValues} index {copySelectedIndex} '
                      f'chosen to be copied\n')
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Column {pasteSelectedValues} index '
                      f'{pasteSelectedIndex} chosen to be pasted into\n')
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Searching... Copy & Paste\n')
    messageText.config(text="Processing...")
    startTime = time()
    filterwarnings("ignore", category=DeprecationWarning)
    sourceWorkbook = load_workbook(sourceExcelFile)
    sourceExcelSheet = sourceExcelSheetCombo.get()
    sourceWorkSheet = sourceWorkbook[sourceExcelSheet]
    sourceReferenceColumn = excelColumnIndex(sourceReferenceColumnCombo.get())
    sourceTotalRows = sourceWorkSheet.max_row
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} {sourceTotalRows} are the total unfiltered source excel sheet rows\n')
    targetWorkbook = load_workbook(targetExcelFile)
    targetExcelSheet = targetExcelSheetCombo.get()
    targetWorkSheet = targetWorkbook[targetExcelSheet]
    targetReferenceColumn = excelColumnIndex(targetReferenceColumnCombo.get())
    targetTotalRows = targetWorkSheet.max_row
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} {targetTotalRows} are the total unfiltered target excel sheet rows\n')
    copyPasteDone = set()
    totalSources = set()
    pasteSuccess = 0
    sourceVisibleRowCount = len([row for row in range(2, sourceTotalRows + 1) if
                                 not sourceWorkSheet.row_dimensions[row].hidden])
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} Total filtered source rows are {sourceVisibleRowCount}\n')
    source_rows = sourceWorkSheet.iter_rows(min_row=2, max_row=sourceTotalRows, values_only=True)
    for sourceRowIndex, sourceRow in enumerate(source_rows, start=2):
        if sourceWorkSheet.row_dimensions[sourceRowIndex].hidden:
            fileHandler.write(
                f'{datetime.now().replace(microsecond=0)} Source row [{sourceRowIndex}] is hidden(Filtered)\n')
            continue
        sourceCellValue = sourceRow[sourceReferenceColumn - 1]
        fileHandler.write(
            f'{datetime.now().replace(microsecond=0)} [{sourceCellValue}] is source cell value from '
            f'[row_({sourceRowIndex}),column_({sourceReferenceColumn})] \n')
        totalSources.add(sourceCellValue)
        target_rows = targetWorkSheet.iter_rows(min_row=2, max_row=targetTotalRows,
                                                values_only=True)
        for targetRowIndex, targetRow in enumerate(target_rows, start=2):
            if not targetWorkSheet.row_dimensions[targetRowIndex].hidden:
                targetCellValue = targetRow[targetReferenceColumn - 1]
                fileHandler.write(
                    f'{datetime.now().replace(microsecond=0)} {targetCellValue} is target cell value from '
                    f'[row_({targetRowIndex}),column_({targetReferenceColumn})]\n')
                fileHandler.write(
                    f'{datetime.now().replace(microsecond=0)} comparing source cell value [{sourceCellValue}] '
                    f'with target cell value [{targetCellValue}]\n')
                if sourceCellValue == targetCellValue:
                    targetRow = list(targetRow)

                    for sourceColumn, targetColumn in zip(copySelectedIndex, pasteSelectedIndex):
                        targetRow[targetColumn - 1] = sourceRow[sourceColumn - 1]

                    for sourceColumn, targetColumn in zip(copySelectedIndex, pasteSelectedIndex):
                        targetCell = targetWorkSheet.cell(row=targetRowIndex, column=targetColumn)
                        sourceCell = sourceWorkSheet.cell(row=sourceRowIndex, column=sourceColumn)
                        targetCell.value = sourceCell.value
                    fileHandler.write(
                        f'{datetime.now().replace(microsecond=0)} Matched, cell Value Copied Successfully for '
                        f'[{targetCellValue}] \n')
                    if targetCellValue is not None:
                        copyPasteDone.add(targetCellValue)
                    pasteSuccess += 1
                    updateProgress(progressBar, pasteSuccess, sourceVisibleRowCount, window, progressStyle)
                    fileHandler.flush()
                    break
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} Data for {len(copyPasteDone)} cells whose values are '
        f'[{copyPasteDone}] copied\n')
    missedItem = [item.strip() for item in totalSources if item.strip() not in copyPasteDone]
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} Data for {len(missedItem)} cells whose values are '
        f'[{missedItem}] failed to copy because it is not exact match.\n')

    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Saving file...\n')
    messageText.config(text='Saving File...')
    excelFileName = splitext(basename(targetExcelFile))[0]
    excelFileSave = f'{splitext(basename(targetExcelFile))[0]}_modified.xlsx'
    excelFIleCounter = 1
    while exists(excelFileSave):
        excelFileSave = f"{excelFileName}_modified_({excelFIleCounter}).xlsx"
        excelFIleCounter = excelFIleCounter + 1
    targetWorkbook.save(excelFileSave)
    fileHandler.write(
        f'{datetime.now().replace(microsecond=0)} modified file [{excelFileSave}] saved successfully\n')
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Applying filter\n')
    progressBar['value'] = 99
    progressStyle.configure("Custom.Horizontal.TProgressbar", background='#90EE90', text='99%')
    messageText.config(text='Applying Filter...')
    modifiedFilePath = join(dirname(__file__), excelFileSave)
    filterApplied = list(copyPasteDone)
    try:
        filterProcess = Book(modifiedFilePath, App(visible=False))
        filterProcess.sheets[f'{targetExcelSheet}'].api.Range(
            f"A1:{targetReferenceColumnComboValues[-1]}{filterProcess.sheets[f'{targetExcelSheet}'].range('A' + str(filterProcess.sheets[f'{targetExcelSheet}'].cells.last_cell.row)).end('up').row}").AutoFilter(
            Field := targetReferenceColumn, Criterial := filterApplied, Operator := 7)
        filterProcess.save()
        filterProcess.close()
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} Filter applied successfully and file saved.\n')
        progressBar['value'] = 100
        progressStyle.configure("Custom.Horizontal.TProgressbar", background='#7CFC00', text='100%')
    except Exception as error:
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} Error during filter, filter process failed\n')
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [ERROR] {error}\n')
    endTime = time()
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} TASK COMPLETED\n')
    elapsedTime = endTime - startTime
    hours, remainder = divmod(elapsedTime, 3600)
    minutes, remainder = divmod(remainder, 60)
    secondsTotal, millisecondsTotal = divmod(remainder, 1)
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} ELAPSED TIME TO COMPLETE THIS TASK IS '
                      f'{int(hours)} hours {int(minutes)} minutes {int(secondsTotal)} seconds '
                      f'{int(millisecondsTotal * 1000)} milliseconds\n')
    resetBtnFunc(sourceExcelEntry, sourceExcelSheetCombo, sourceReferenceColumnCombo, sourceExcelFileDialogBtn,
                 sourceCopyColumnsScrolledText, targetExcelEntry, targetExcelSheetCombo, targetReferenceColumnCombo,
                 targetExcelFileDialogBtn, targetPasteColumnsScrolledText, resetBtn, messageText)
    messageText.config(text=f'Changes saved in {excelFileSave}, check logs file for the status')
    fileHandler.flush()


def updateProgress(progressBar, newVal, totalVal, window, progressStyle):
    resultVal = round((newVal / totalVal) * 98, 2)
    progressBar['value'] = resultVal
    progressStyle.configure("Custom.Horizontal.TProgressbar", text=f"{resultVal}%")
    window.update()


copySelectedIndex = []
copyCheckBoxVars = []
copyCheckboxesText = []
copyCheckboxes = {}
copySelectedIndices = []
copyColors = {}


def enableSourceReferenceColumnCombo(sourceExcelSheetCombo, sourceReferenceColumnCombo,
                                     sourceReferenceColumnComboValues, sourceCopyColumnsScrolledText, resetBtn):
    sourceCopyColumnsScrolledText.delete(1.0, 'end')
    if not sourceExcelSheetCombo.get() == "Select WorkSheet":
        sourceReferenceColumnCombo.config(state='disabled')
        sourceExcelSheetCombo.config(state='disabled')
        resetBtn.config(state='disabled', bg='light grey')
        sourceReferenceColumnComboValues.clear()
        sourceReferenceColumnComboValues.append('Select Reference Column')
        for sourceColumnIndex in range(1, load_workbook(sourceExcelFile)[sourceExcelSheetCombo.get()].max_column + 1):
            sourceColumnLetter = getColumnLetter(sourceColumnIndex)
            sourceReferenceColumnComboValues.append(sourceColumnLetter)
        sourceExcelSheetCombo.config(state='normal')
        sourceReferenceColumnCombo.config(state='normal')
        sourceReferenceColumnCombo.config(values=sourceReferenceColumnComboValues)
        del sourceReferenceColumnComboValues[0]
        load_workbook(sourceExcelFile).close()
    else:
        sourceReferenceColumnCombo.config(state='disabled')
    resetBtn.config(state='normal', bg='red')


def copyCheckboxClicked(index):
    global copySelectedValues
    if copyCheckBoxVars[index].get() == 1:
        assign_color(copyCheckBoxVars[index], copyCheckboxes, copyColors)
        copySelectedIndices.append(index)
        disable_checkboxes(copyCheckboxes)
        enableCheckboxes(pasteCheckboxes)
    elif copyCheckBoxVars[index].get() == 0:
        enableCheckboxes(copyCheckboxes)
        copySelectedIndices.remove(index)
    copySelectedValues = [copyCheckboxesText[i] for i in copySelectedIndices]
    copySelectedIndex.clear()
    for item in copySelectedValues:
        copySelectedIndex.append(excelColumnIndex(item))


def threadingEnableSourceReferenceColumnCombo(sourceExcelSheetCombo, sourceReferenceColumnCombo,
                                              sourceReferenceColumnComboValues, sourceCopyColumnsScrolledText,
                                              resetBtn):
    threadEnableSourceReferenceColumnCombo = Thread(target=enableSourceReferenceColumnCombo,
                                                    args=(sourceExcelSheetCombo, sourceReferenceColumnCombo,
                                                          sourceReferenceColumnComboValues,
                                                          sourceCopyColumnsScrolledText, resetBtn))
    threadEnableSourceReferenceColumnCombo.start()


def enableTargetExcelDialogBtn(sourceExcelSheetCombo, sourceReferenceColumnCombo, targetExcelFileDialogBtn,
                               sourceReferenceColumnComboValues, sourceCopyColumnsScrolledText):
    if sourceExcelSheetCombo.get() and sourceReferenceColumnCombo.get():
        sourceExcelSheetCombo.config(state='disabled')
        sourceReferenceColumnCombo.config(state='disabled')
        sourceReferenceColumnComboValues.remove(sourceReferenceColumnCombo.get())
        for index, columnsText in enumerate(sourceReferenceColumnComboValues):
            var = IntVar()
            copyCheckBoxVars.append(var)
            copyCheckbox = Checkbutton(sourceCopyColumnsScrolledText, variable=var, text=columnsText, bg='light grey',
                                       state='disabled', cursor="arrow", command=lambda e=index: copyCheckboxClicked(e))
            copyCheckbox.pack()
            copyCheckboxesText.append(columnsText)
            copyCheckboxes[copyCheckbox] = var
            sourceCopyColumnsScrolledText.window_create('end', window=copyCheckbox)
            sourceCopyColumnsScrolledText.config(bd=4, width=12, bg='light grey', height=5)
            var.trace_add('write', enablePasteCheckBtnSubmitBtn)
        targetExcelFileDialogBtn.config(state='normal', bg='green')
        sourceExcelFileName = basename(sourceExcelFile)
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourceExcelSheetCombo.get()}] is selected as '
                          f'source work sheet of the excel file [{sourceExcelFileName}]\n')
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{sourceReferenceColumnCombo.get()}] is selected '
                          f'as source reference column from the excel work sheet [{sourceExcelSheetCombo.get()}]\n')
    else:
        targetExcelFileDialogBtn.config(state='disabled', bg='light grey')
    fileHandler.flush()


pasteSelectedIndex = []
pasteCheckBoxVars = []
pasteCheckboxesText = []
pasteCheckboxes = {}
pasteSelectedIndices = []
pasteColors = {}


def enableTargetReferenceColumnCombo(targetExcelSheetCombo, targetReferenceColumnCombo,
                                     targetReferenceColumnComboValues, targetPasteColumnsScrolledText, resetBtn):
    targetPasteColumnsScrolledText.delete(1.0, 'end')
    if not targetExcelSheetCombo.get() == "Select WorkSheet":
        targetReferenceColumnCombo.config(state='disabled')
        targetExcelSheetCombo.config(state='disabled')
        resetBtn.config(state='disabled', bg='light grey')
        targetReferenceColumnComboValues.clear()
        targetReferenceColumnComboValues.append('Select Reference Column')
        for targetColumnIndex in range(1, load_workbook(targetExcelFile)[targetExcelSheetCombo.get()].max_column + 1):
            targetColumnLetter = getColumnLetter(targetColumnIndex)
            targetReferenceColumnComboValues.append(targetColumnLetter)
        targetReferenceColumnCombo.config(state='normal')
        targetExcelSheetCombo.config(state='normal')
        targetReferenceColumnCombo.config(values=targetReferenceColumnComboValues)
        del targetReferenceColumnComboValues[0]
        load_workbook(targetExcelFile).close()
    else:
        targetReferenceColumnCombo.config(state='disabled')
    resetBtn.config(state='normal', bg='red')


def pasteCheckboxClicked(index):
    global pasteSelectedValues
    if pasteCheckBoxVars[index].get() == 1:
        pasteSelectedIndices.append(index)
        disable_checkboxes(pasteCheckboxes, except_checkbox=list(pasteCheckboxes.keys())[index])
        if len(pasteCheckboxes.keys()) != len(pasteSelectedIndices):
            enableCheckboxes(copyCheckboxes)
        if len(copyColors) > 0:
            color = list(copyColors.values())[-1]
            list(pasteCheckboxes.keys())[index].configure(bg=color)
            list(pasteCheckboxes.keys())[index].config(state='disabled')
            pasteColors[list(pasteCheckboxes.keys())[index]] = color
        else:
            assign_color(pasteCheckBoxVars[index], pasteCheckboxes, pasteColors)
    elif pasteCheckBoxVars[index].get() == 0:
        pasteSelectedIndices.remove(index)
        enableCheckboxes(pasteCheckboxes)
    pasteSelectedValues = [pasteCheckboxesText[i] for i in pasteSelectedIndices]
    pasteSelectedIndex.clear()
    for item in pasteSelectedValues:
        pasteSelectedIndex.append(excelColumnIndex(item))


def disable_checkboxes(checkboxes, except_checkbox=None):
    for checkbox, var in checkboxes.items():
        if checkbox != except_checkbox:
            checkbox.config(state='disabled')


def enableCheckboxes(checkboxes):
    for checkbox, var in checkboxes.items():
        if var.get() == 0:
            checkbox.config(state='normal')
        else:
            checkbox.config(state='disabled')


def threadingEnableTargetReferenceColumnCombo(targetExcelSheetCombo, targetReferenceColumnCombo,
                                              targetReferenceColumnComboValues, targetPasteColumnsScrolledText,
                                              resetBtn):
    threadEnableTargetReferenceColumnCombo = Thread(target=enableTargetReferenceColumnCombo,
                                                    args=(targetExcelSheetCombo, targetReferenceColumnCombo,
                                                          targetReferenceColumnComboValues,
                                                          targetPasteColumnsScrolledText, resetBtn))
    threadEnableTargetReferenceColumnCombo.start()


def getColumnLetter(index):
    if index <= 0:
        fileHandler.write(
            f'{datetime.now().replace(microsecond=0)} [ValueError][Column index must be greater than 0]\n')
    result = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    fileHandler.flush()
    return result


def enableCopyColumnsCheckBtn(targetExcelSheetCombo, targetReferenceColumnCombo, targetReferenceColumnComboValues,
                              targetPasteColumnsScrolledText):
    if not targetReferenceColumnCombo.get() == 'Select Reference Column':
        targetExcelFileName = basename(targetExcelFile)
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{targetExcelSheetCombo.get()}] is selected as '
                          f'target work sheet of the excel file [{targetExcelFileName}]\n')
        fileHandler.write(f'{datetime.now().replace(microsecond=0)} [{targetReferenceColumnCombo.get()}] is selected '
                          f'as target reference column from the excel work sheet [{targetExcelSheetCombo.get()}]\n')
        targetExcelSheetCombo.config(state='disabled')
        targetReferenceColumnCombo.config(state='disabled')
        targetReferenceColumnComboValues.remove(targetReferenceColumnCombo.get())
        for index, columnsText in enumerate(targetReferenceColumnComboValues):
            var = IntVar()
            pasteCheckBoxVars.append(var)
            pasteCheckbox = Checkbutton(targetPasteColumnsScrolledText, variable=var, text=columnsText, bg='light grey',
                                        state='disabled', cursor="arrow",
                                        command=lambda e=index: pasteCheckboxClicked(e))
            pasteCheckbox.pack()
            pasteCheckboxesText.append(columnsText)
            pasteCheckboxes[pasteCheckbox] = var
            targetPasteColumnsScrolledText.window_create('end', window=pasteCheckbox)
            targetPasteColumnsScrolledText.config(bd=4, width=12, bg='light grey', height=5)
            var.trace_add('write', enablePasteCheckBtnSubmitBtn)
        for copyCheckbox in copyCheckboxes:
            copyCheckbox.config(state='normal')
            copyCheckbox.configure(cursor="hand2")
    fileHandler.flush()


def enablePasteCheckBtnSubmitBtn(*args):
    copyCheckboxesState = [var.get() for var in copyCheckBoxVars]
    if any(copyCheckboxesState):
        for pasteCheckbox in pasteCheckboxes:
            pasteCheckbox.config(state='normal')
            pasteCheckbox.configure(cursor="hand2")
    else:
        for index, pasteCheckbox in enumerate(pasteCheckboxes):
            if pasteCheckBoxVars[index].get() == 1:
                pasteCheckbox.deselect()
                pasteSelectedIndices.clear()
            pasteCheckbox.config(state='disabled')
            pasteCheckbox.configure(cursor="arrow")
    copyCheckboxSelected = sum(copyVar.get() for copyVar in copyCheckBoxVars)
    pasteCheckboxSelected = sum(pasteVar.get() for pasteVar in pasteCheckBoxVars)
    if copyCheckboxSelected > 0 and copyCheckboxSelected == pasteCheckboxSelected:
        submitBtn.config(state='normal', bg='green')
    else:
        submitBtn.config(state='disabled', bg='light grey')


def excelColumnIndex(columnRef):
    base = ord('A') - 1
    index = 0
    for char in columnRef:
        index = index * 26 + (ord(char) - base)
    return index


def random_color():
    return f'#{randint(0, 0xFFFFFF):06x}'


def assign_color(check_var, checkboxes, colors):
    color = random_color()
    for checkbox, var in checkboxes.items():
        if var == check_var:
            checkbox.configure(bg=color)
            checkbox.config(state='disabled')
            colors[checkbox] = color


def resetBtnFunc(sourceExcelEntry, sourceExcelSheetCombo, sourceReferenceColumnCombo, sourceExcelFileDialogBtn,
                 sourceCopyColumnsScrolledText, targetExcelEntry, targetExcelSheetCombo, targetReferenceColumnCombo,
                 targetExcelFileDialogBtn, targetPasteColumnsScrolledText, resetBtn, messageText):
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Resetting...\n')
    sourceExcelEntry.config(state='normal')
    sourceExcelEntry.delete(0, 'end')
    sourceExcelEntry.config(state='disabled')
    sourceExcelSheetCombo.config(state='normal')
    sourceExcelSheetCombo.delete(0, 'end')
    sourceExcelSheetCombo.set('Select WorkSheet')
    sourceExcelSheetCombo.config(state='disabled')
    sourceReferenceColumnCombo.config(state='normal')
    sourceReferenceColumnCombo.delete(0, 'end')
    sourceReferenceColumnCombo.set('Select Reference Column')
    sourceReferenceColumnCombo.config(state='disabled')
    sourceExcelFileDialogBtn.config(state='normal', bg='green')
    sourceExcelFile = ''
    copySelectedIndices.clear()
    copyCheckBoxVars.clear()
    copyCheckboxes.clear()
    copySelectedIndex.clear()
    copyCheckboxesText.clear()
    copyColors.clear()
    sourceCopyColumnsScrolledText.delete(1.0, 'end')
    targetExcelEntry.config(state='normal')
    targetExcelEntry.delete(0, 'end')
    targetExcelEntry.config(state='disabled')
    targetExcelSheetCombo.config(state='normal')
    targetExcelSheetCombo.delete(0, 'end')
    targetExcelSheetCombo.set('Select WorkSheet')
    targetExcelSheetCombo.config(state='disabled')
    targetReferenceColumnCombo.config(state='normal')
    targetReferenceColumnCombo.delete(0, 'end')
    targetReferenceColumnCombo.set('Select Reference Column')
    targetReferenceColumnCombo.config(state='disabled')
    targetExcelFileDialogBtn.config(state='disabled', bg='light grey')
    pasteSelectedIndices.clear()
    pasteCheckBoxVars.clear()
    pasteCheckboxes.clear()
    pasteSelectedIndex.clear()
    pasteCheckboxesText.clear()
    pasteColors.clear()
    targetPasteColumnsScrolledText.delete(1.0, 'end')
    submitBtn.config(state='disabled', bg='light grey')
    resetBtn.config(state='disabled', bg='light grey')
    targetExcelFile = ''
    messageText.config(text='')
    fileHandler.write(f'{datetime.now().replace(microsecond=0)} Reset done..\n')


def mainGUI():
    global submitBtn
    window = Tk()
    window.config(bg='lemon chiffon')
    window.title('ExcelBridge v1.1')
    window.geometry('450x525')
    window.resizable(False, False)
    window.iconbitmap(iconFile)
    mainLabel = Label(window, text='Excel Data Transfer', font=('Arial', 15, 'bold'), fg='blue', bg='lemon chiffon')
    mainLabel.place(x=230, y=15, anchor='center')

    sourceEntryCanvas = Canvas(window, highlightthickness=3, highlightbackground="black", relief='solid', height=100,
                               width=435, bg='gray56')
    sourceEntryCanvas.place(x=5, y=30)
    sourceExcelLabel = Label(sourceEntryCanvas, text='Source Excel:', font=('Arial', 9, 'bold italic'), bg='gray56')
    sourceExcelLabel.place(x=5, y=10)
    sourceExcelEntry = Entry(sourceEntryCanvas, bd=4, width=47, bg='white', state='disabled')
    sourceExcelEntry.place(x=120, y=8)
    sourceExcelFileDialogBtn = Button(sourceEntryCanvas, text='...', bg='green', fg='white', font=('Arial', 8),
                                      command=lambda: threadingSourceExcelFileDialogFunc(progress, progressStyle,
                                                                                         sourceExcelEntry,
                                                                                         sourceExcelFileDialogBtn,
                                                                                         sourceExcelSheetCombo,
                                                                                         resetBtn, messageLabel,
                                                                                         sourceExcelSheetComboValues))
    sourceExcelFileDialogBtn.place(x=418, y=8)
    sourceExcelSheetLabel = Label(sourceEntryCanvas, text='Source Excel Sheet:', font=('Arial', 8, 'bold italic'),
                                  bg='gray56')
    sourceExcelSheetLabel.place(x=5, y=40)
    sourceExcelSheetComboValues = ['Select WorkSheet']
    sourceExcelSheetCombo = Combobox(sourceEntryCanvas, values=sourceExcelSheetComboValues, width=47, state='disabled')
    sourceExcelSheetCombo.place(x=120, y=38)
    sourceExcelSheetCombo.current(0)
    sourceReferenceColumnLabel = Label(sourceEntryCanvas, text='Source Ref. column:', font=('Arial', 8, 'bold italic'),
                                       bg='gray56')
    sourceReferenceColumnLabel.place(x=5, y=70)
    sourceReferenceColumnComboValues = ['Select Reference Column']
    sourceReferenceColumnCombo = Combobox(sourceEntryCanvas, values=sourceReferenceColumnComboValues, width=47,
                                          state='disabled')
    sourceReferenceColumnCombo.place(x=120, y=68)
    sourceReferenceColumnCombo.current(0)

    targetEntryCanvas = Canvas(window, highlightthickness=3, highlightbackground="yellow", relief='solid', height=100,
                               width=435, bg='gray56')
    targetEntryCanvas.place(x=5, y=140)
    targetExcelLabel = Label(targetEntryCanvas, text='Target Excel:', font=('Arial', 8, 'bold italic'), bg='grey56')
    targetExcelLabel.place(x=5, y=10)
    targetExcelEntry = Entry(targetEntryCanvas, bd=4, width=47, bg='white', state='disabled')
    targetExcelEntry.place(x=120, y=8)
    targetExcelFileDialogBtn = Button(targetEntryCanvas, text='...', bg='light grey', fg='white', font=('Arial', 8),
                                      state='disabled', command=lambda: threadingTargetExcelFileDialogFunc(
            progress, progressStyle, targetExcelSheetComboValues, targetExcelEntry, targetExcelFileDialogBtn,
            targetExcelSheetCombo, messageLabel, resetBtn))
    targetExcelFileDialogBtn.place(x=418, y=8)
    targetExcelSheetLabel = Label(targetEntryCanvas, text='Target Excel Sheet:', font=('Arial', 8, 'bold italic'),
                                  bg='grey56')
    targetExcelSheetLabel.place(x=5, y=40)
    targetExcelSheetComboValues = ['Select WorkSheet']
    targetExcelSheetCombo = Combobox(targetEntryCanvas, values=targetExcelSheetComboValues, width=47, state='disabled')
    targetExcelSheetCombo.place(x=120, y=38)
    targetExcelSheetCombo.current(0)
    targetReferenceColumnLabel = Label(targetEntryCanvas, text='Target Ref. column:', font=('Arial', 8, 'bold italic'),
                                       bg='grey56')
    targetReferenceColumnLabel.place(x=5, y=70)
    targetReferenceColumnComboValues = ['Select Reference Column']
    targetReferenceColumnCombo = Combobox(targetEntryCanvas, values=targetReferenceColumnComboValues, width=47,
                                          state='disabled')
    targetReferenceColumnCombo.place(x=120, y=68)
    targetReferenceColumnCombo.current(0)

    copyPasteCanvas = Canvas(window, highlightthickness=3, highlightbackground="light green", relief='solid',
                             height=265,
                             width=435, bg='gray56')
    copyPasteCanvas.place(x=5, y=250)

    sourceCopyColumnsLabel = Label(copyPasteCanvas, text='Copy Columns', font=('Arial', 8, 'bold italic'),
                                   bg='grey56')
    sourceCopyColumnsLabel.place(x=100, y=5)
    sourceCopyColumnsScrolledText = ScrolledText(copyPasteCanvas, bd=4, width=12, bg='white', height=5)
    sourceCopyColumnsScrolledText.place(x=100, y=25)
    targetPasteColumnsLabel = Label(copyPasteCanvas, text='Paste Columns:', font=('Arial', 8, 'bold italic'),
                                    bg='grey56')
    targetPasteColumnsLabel.place(x=245, y=5)
    targetPasteColumnsScrolledText = ScrolledText(copyPasteCanvas, bd=4, width=12, bg='white', height=5)
    targetPasteColumnsScrolledText.place(x=245, y=25)

    submitBtn = Button(copyPasteCanvas, text='Submit', bg='light grey', fg='white', font=('Arial', 12, 'bold'),
                       state='disabled', command=lambda: threadingMatchCopyPaste(sourceExcelEntry,
                                                                                 sourceExcelSheetCombo,
                                                                                 sourceReferenceColumnCombo,
                                                                                 sourceExcelFileDialogBtn,
                                                                                 sourceCopyColumnsScrolledText,
                                                                                 targetExcelEntry,
                                                                                 targetExcelSheetCombo,
                                                                                 targetReferenceColumnCombo,
                                                                                 targetReferenceColumnComboValues,
                                                                                 targetExcelFileDialogBtn,
                                                                                 targetPasteColumnsScrolledText,
                                                                                 resetBtn, messageLabel, window,
                                                                                 progress, progressStyle))
    submitBtn.place(x=155, y=120)
    resetBtn = Button(copyPasteCanvas, text='Reset', bg='light grey', fg='white',
                      font=('Arial', 12, 'bold'), state='disabled',
                      command=lambda: resetBtnFunc(sourceExcelEntry, sourceExcelSheetCombo,
                                                   sourceReferenceColumnCombo, sourceExcelFileDialogBtn,
                                                   sourceCopyColumnsScrolledText, targetExcelEntry,
                                                   targetExcelSheetCombo,
                                                   targetReferenceColumnCombo, targetExcelFileDialogBtn,
                                                   targetPasteColumnsScrolledText, resetBtn, messageLabel))

    resetBtn.place(x=245, y=120)
    progress = Progressbar(copyPasteCanvas, length=430, mode="determinate", style="Custom.Horizontal.TProgressbar")
    progress.place(x=5, y=160)
    progressStyle = Style()
    progressStyle.theme_use('default')
    progressStyle.configure("Custom.Horizontal.TProgressbar", thickness=20, troughcolor='#E0E0E0', background='#FFFF00',
                            troughrelief='flat', relief='flat', text='0 %')
    progressStyle.layout('Custom.Horizontal.TProgressbar', [('Horizontal.Progressbar.trough',
                                                             {'children': [('Horizontal.Progressbar.pbar',
                                                                            {'side': 'left', 'sticky': 'ns'})],
                                                              'sticky': 'nswe'}),
                                                            ('Horizontal.Progressbar.label', {'sticky': ''})])
    warnLabel = Label(copyPasteCanvas, text='Note:It does not supports Data_Validation_Rules',
                      font=('Arial', 8, 'bold italic'), bg='grey56')
    warnLabel.place(x=5, y=245)
    messageLabel = Label(copyPasteCanvas, justify='left', wraplength=420, font=('Arial', 8, 'bold'), bg='grey56')
    messageLabel.place(x=5, y=195)
    aboutBtn = Button(copyPasteCanvas, text='About', bg='brown', command=lambda: aboutWindow(window))
    aboutBtn.place(x=390, y=240)

    sourceExcelSheetCombo.bind("<<ComboboxSelected>>", lambda event: threadingEnableSourceReferenceColumnCombo(
        sourceExcelSheetCombo, sourceReferenceColumnCombo, sourceReferenceColumnComboValues,
        sourceCopyColumnsScrolledText, resetBtn))
    sourceReferenceColumnCombo.bind('<<ComboboxSelected>>', lambda event: enableTargetExcelDialogBtn(
        sourceExcelSheetCombo, sourceReferenceColumnCombo, targetExcelFileDialogBtn,
        sourceReferenceColumnComboValues, sourceCopyColumnsScrolledText))
    targetExcelSheetCombo.bind("<<ComboboxSelected>>", lambda event: threadingEnableTargetReferenceColumnCombo(
        targetExcelSheetCombo, targetReferenceColumnCombo, targetReferenceColumnComboValues,
        targetPasteColumnsScrolledText, resetBtn))
    targetReferenceColumnCombo.bind('<<ComboboxSelected>>', lambda event: enableCopyColumnsCheckBtn(
        targetExcelSheetCombo, targetReferenceColumnCombo, targetReferenceColumnComboValues,
        targetPasteColumnsScrolledText))
    window.mainloop()


def aboutWindow(mainWin):
    aboutWin = Toplevel(mainWin)
    aboutWin.grab_set()
    aboutWin.geometry('285x90')
    aboutWin.resizable(False, False)
    aboutWin.title('About')
    aboutWin.iconbitmap(aboutIcon)
    aboutWinLabel = Label(aboutWin, text=f'Version - 1.1\nDeveloped by Priyanshu\nFor any improvement please reach on '
                                         f'below email\nEmail : chandelpriyanshu8@outlook.com\nMobile : '
                                         f'+91-8285775109 '
                                         f'', font=('Helvetica', 9)).place(x=1, y=6)


def updateProgressSaveExcel(listingSuccess, progressBar, totalFiles, window, progressStyle):
    resultVal = (listingSuccess / totalFiles) * 100
    progressBar['value'] = resultVal
    progressStyle.configure("Custom.Horizontal.TProgressbar", text='{:g} %'.format(resultVal))
    window.update()


mainGUI()
