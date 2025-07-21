Attribute VB_Name = "模块1"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' 适用于64位版本的Word
Sub superPrint()
    Dim fso As Object
    Dim fd As FileDialog
    Dim folderPath As String '文件打印位置
    Dim logFilePath As String '日志文件位置
    Dim logFolderPath As String '日志文件所在目录
    Dim fileName As String
    Dim doc As Document
    Dim userInput As String '用户输入
    Dim stepPrintNum As Byte
    Dim stepTime As Byte
    
    ' 创建FileDialog对象实例
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    ' 配置FileDialog
    fd.Title = "要打印文件所在文件夹"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        ' 获取用户选择的文件路径
        folderPath = fd.SelectedItems(1)
    Else
        MsgBox "未选择文件夹"
        Exit Sub
    End If
    
    fileName = Dir(folderPath & "\*.*")
    Set fd = Nothing
    
     ' 获取用户输入的每次打印页数
    userInput = InputBox("请输入每次打印的页数：", "每次打印页数", "10")
    
    ' 验证用户输入
    If IsNumeric(userInput) Then
        stepPrintNum = CByte(userInput)
    Else
        MsgBox "无法打印，输入的不是数字", vbExclamation
        Exit Sub
    End If
    
    userInput = "" '置空
    
     ' 获取用户输入的打印间隔时间
    userInput = InputBox("请输入打印间隔时间（秒）：", "间隔时间", "90")
    
    ' 验证用户输入
    If IsNumeric(userInput) Then
        stepTime = CByte(userInput)
    Else
        MsgBox "无法打印，输入的不是数字", vbExclamation
        Exit Sub
    End If
    
    '初始化日志文件
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFolderPath = fso.BuildPath(fso.GetParentFolderName(folderPath), "log")
    logFilePath = fso.BuildPath(logFolderPath, "printDoc_log.txt")
    
    Call InitializeLog(fso, logFolderPath, logFilePath)
    Set fso = Nothing
    
    While fileName <> ""
        If LCase(Right(fileName, 4)) = ".doc" Or LCase(Right(fileName, 5)) = ".docx" Or _
           LCase(Right(fileName, 5)) = ".docm" Or LCase(Right(fileName, 4)) = ".rtf" Then
            On Error Resume Next ' 开始错误处理
            Set doc = Documents.Open(folderPath & "\" & fileName)
            If Err.Number <> 0 Then
                Debug.Print "无法打开文档：" & fileName
                Call WriteLog(logFilePath, Format(Now, "yyyy-mm-dd hh:mm:ss") & "无法打开文档：" & fileName)
                Call WriteLog(logFilePath, "    Source:" & Err.Source & ",    Num:" & Err.Number)
                Err.Clear
            Else
                On Error GoTo 0 ' 正常进行，结束错误处理
                
                ' 打印文档
                Call singleWordDocPrint(doc, stepTime, stepPrintNum, logFilePath)
    
                ' 关闭文档
                doc.Close SaveChanges:=wdDoNotSaveChanges
                Set doc = Nothing ' 释放对象
            End If
        Else
            Call WriteLog(logFilePath, Format(Now, "yyyy-mm-dd hh:mm:ss    ") & fileName & "  不是word文档，无法打印")
        End If
        ' 获取下一个文件
        fileName = Dir
    Wend
    Call CloseLog(logFilePath)
End Sub

Sub singleWordDocPrint(pDoc As Document, pStepTime As Byte, pStepPrintNum As Byte, pLogFilePath As String)
    On Error GoTo ErrorHandler ' 添加错误处理
    
    Dim thisDoc As Document
    Dim sec As Section
    Dim totalPages As Long '文档总页数
    Dim startPage As Integer '文档起始页
    Dim endPage As Integer '文档结束页
    Dim totalPrintPages As Long '总打印页数
    Dim startPrintPage As Integer '当前打印起始页
    Dim endPrintPage As Integer '当前打印结束页
    Dim stepPrintNum As Byte '每次打印页数
    Dim startTime As Double '时间搓，用于暂停
    Dim stepTime As Byte '打印间隔时间
    Dim logMessage  As String

    Set thisDoc = pDoc
    Set sec = thisDoc.Sections(1)
    logMessage = ""
    totalPages = thisDoc.ComputeStatistics(wdStatisticPages)
    startPage = sec.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber
    endPage = startPage + totalPages
    stepTime = pStepTime
    stepPrintNum = pStepPrintNum
    
    ' 新增内存优化代码
    Application.ScreenUpdating = False
    thisDoc.Activate
    
    startPrintPage = startPage
    endPrintPage = startPrintPage + stepPrintNum - 1
    totalPrintPages = totalPages
    Do
        If endPrintPage > endPage Then
            endPrintPage = endPage
        End If
        
        'thisDoc.ActiveWindow.PrintOut Range:=wdPrintFromTo, From:=CStr(startPrintPage), To:=CStr(endPrintPage)
        logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & "    " & Replace(thisDoc.FullName, thisDoc.Path, "") & ":" & CStr(startPrintPage) & " - " & CStr(endPrintPage) & "  （共" & CStr(totalPages) & "页)"
        
        Application.StatusBar = "Printing: " & logMessage
        Debug.Print logMessage
        Call WriteLog(pLogFilePath, logMessage) '写日志
        
        startTime = Timer
        Do While Timer < startTime + stepTime
            DoEvents ' 允许其他操作进行
            
            ' 新增强制回收资源
            Set obj = Nothing
            If FreeFile > 255 Then Close
        Loop
        
        startPrintPage = startPrintPage + stepPrintNum
        endPrintPage = startPrintPage + stepPrintNum - 1
    Loop While (startPrintPage <= endPage)

    Application.ScreenUpdating = True
    Set thisDoc = Nothing
    DoEvents
    Exit Sub
ErrorHandler:
        Call WriteLog(pLogFilePath, Format(Now, "yyyy-mm-dd hh:mm:ss") & " 错误 " & Err.Description & " 在文档 " & thisDoc.Name)
        Resume Next
End Sub

Sub InitializeLog(fso As Object, logFolderPath As String, logFilePath As String)
    Dim logFile As Integer
    
    ' Create 日志文件的父级目录
    If Not fso.FolderExists(logFolderPath) Then
        fso.CreateFolder logFolderPath
    End If
    logFile = FreeFile
    Open logFilePath For Append As #logFile
       Print #logFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & "    打印任务开始"
    Close #logFile
End Sub

Sub WriteLog(logFilePath As String, message As String)
    ' Append message to the log file
    Dim logFile As Integer
    logFile = FreeFile
    Open logFilePath For Append As #logFile
    Print #logFile, message
    Close #logFile
End Sub

Sub CloseLog(logFilePath As String)
    Dim logFile As Integer
    logFile = FreeFile
    Open logFilePath For Append As #logFile
       Print #logFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & "    打印结束"
       Print #logFile, ""
    Close #logFile
End Sub