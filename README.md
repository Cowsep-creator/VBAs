# VBAs
For Automation

Public Sub PPT_generate()


Application.ScreenUpdating = False

Dim obPptApp As PowerPoint.Application
Dim OpenPptDialogBox As Object
Dim SavePptDialogBox As Object
Dim PdfFilePath As String
Dim DefaultPPT As String
'Dim ExcelApp As Excel.Application
Dim wbExcel As Excel.Workbook
'Dim PowerPointApp As PowerPoint.Application
'Dim ppPowerPoint As PowerPoint.Presentation
Dim Pic1 As String
Dim FolderPics As String
Dim CurrentMonth As String
Dim PreviousMonth As String
Dim MQPFreeze As Date
Dim Factor As Double
Dim image As Shape



Worksheets("Sheet1").Activate
DestinationPPT = Range("B1").Value
PdfFilePath = Range("B3").Value
excel_save_route = Range("B2").Value
picture_source = "C:\Localdata\"

'Dim SavePptDialogBox As Object


Set obPptApp = CreateObject("PowerPoint.Application")
Set OpenPptDialogBox = obPptApp.FileDialog(msoFileDialogOpen)



'Resize factor: use 72 in case you have cm, 28 in case you have inch
Factor = 72

CurrentMonth = Format(Date, "mmmm yyyy")
PreviousMonth = Format(DateAdd("m", -1, Date), "mmmm yyyy")


supplier_list_count = Worksheets("file name list").Range("B" & Excel.Rows.Count).End(xlUp).Row



For iii = 1 To supplier_list_count

'    Print (Left(Cells(i, 4), 1))
    If Left(Cells(iii, 4), 1) <> "1" Then
        Cells(iii, 3) = Cells(iii, 3) & Cells(iii, 4)
'        Cells(iii, 3) = Cells(iii, 3) & Cells(iii, 4)
        Cells(iii, 4) = Cells(iii, 5)
        Cells(iii, 5) = Cells(iii, 6)
        Cells(iii, 6) = Cells(iii, 7)
        Cells(iii, 7) = Cells(iii, 8)
    End If
Next


For i = 1 To supplier_list_count Step 6
obPptApp.Presentations.Open (DestinationPPT)
    MySupplierName = Worksheets("file name list").Range("C" & i)
    MySupplierCode = Worksheets("file name list").Range("E" & i)
        'definition photo link (load picture)
        
        
        slide2_picture1_path = picture_source & Worksheets("file name list").Range("B" & i).Value
        slide2_picture2_down_path = picture_source & Worksheets("file name list").Range("B" & i + 1)
        slide3_picture1_path = picture_source & Worksheets("file name list").Range("B" & i + 5)
        slide3_picture2_down_path = picture_source & Worksheets("file name list").Range("B" & i + 4)
        slide4_picture1_path = picture_source & Worksheets("file name list").Range("B" & i + 3)
        slide4_picture2_down_path = picture_source & Worksheets("file name list").Range("B" & i + 2)
        
        
        'set to MySupplierName
        For Each sld In obPptApp.ActivePresentation.Slides
        For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, "Supplier_Name", MySupplierName)
            End If
        End If
        Next shp
        Next sld
        
        'set to current month
        For Each sld In obPptApp.ActivePresentation.Slides
        For Each shp In sld.Shapes
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, "Month_Reporting", CurrentMonth)
            End If
        End If
        Next shp
        Next sld
        
        
        

            'Slide 2
        
        Slide2_Pic1_Width = 9.05 * Factor
        Slide2_Pic1_Height = 3.85 * Factor
        Slide2_Pic1_Xpos = 2.14 * Factor      'left
        Slide2_Pic1_Ypos = 1.58 * Factor       'top
        
        Slide2_Pic2_Width = 9.05 * Factor
        Slide2_Pic2_Height = 1.68 * Factor
        Slide2_Pic2_Xpos = 2.14 * Factor      'left
        Slide2_Pic2_Ypos = 5.68 * Factor       'top
        
        'Slide 3
        
        Slide3_Pic1_Width = 9.04 * Factor
        Slide3_Pic1_Height = 3.85 * Factor
        Slide3_Pic1_Xpos = 2.15 * Factor
        Slide3_Pic1_Ypos = 1.61 * Factor
        
        Slide3_Pic2_Width = 9.04 * Factor
        Slide3_Pic2_Height = 1.68 * Factor
        Slide3_Pic2_Xpos = 2.15 * Factor
        Slide3_Pic2_Ypos = 5.69 * Factor

        'Slide 4
        
        Slide4_Pic1_Width = 9.04 * Factor
        Slide4_Pic1_Height = 3.85 * Factor
        Slide4_Pic1_Xpos = 2.15 * Factor
        Slide4_Pic1_Ypos = 1.61 * Factor
        
        Slide4_Pic2_Width = 9.04 * Factor
        Slide4_Pic2_Height = 1.68 * Factor
        Slide4_Pic2_Xpos = 2.15 * Factor
        Slide4_Pic2_Ypos = 5.69 * Factor

'Populate Pics Slide 2
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(2)
        myDocument.Shapes.AddPicture Filename:=slide2_picture1_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide2_Pic1_Xpos, Top:=Slide2_Pic1_Ypos, Width:=Slide2_Pic1_Width, Height:=Slide2_Pic1_Height
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(2)
        myDocument.Shapes.AddPicture Filename:=slide2_picture2_down_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide2_Pic2_Xpos, Top:=Slide2_Pic2_Ypos, Width:=Slide2_Pic2_Width, Height:=Slide2_Pic2_Height

'Populate Pics Slide 3
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(3)
        myDocument.Shapes.AddPicture Filename:=slide3_picture1_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide3_Pic1_Xpos, Top:=Slide3_Pic1_Ypos, Width:=Slide3_Pic1_Width, Height:=Slide3_Pic1_Height
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(3)
        myDocument.Shapes.AddPicture Filename:=slide3_picture2_down_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide3_Pic2_Xpos, Top:=Slide3_Pic2_Ypos, Width:=Slide3_Pic2_Width, Height:=Slide2_Pic2_Height

'Populate Pics Slide 4
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(4)
        myDocument.Shapes.AddPicture Filename:=slide4_picture1_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide4_Pic1_Xpos, Top:=Slide4_Pic1_Ypos, Width:=Slide4_Pic1_Width, Height:=Slide4_Pic1_Height
        
        
        Set myDocument = obPptApp.ActivePresentation.Slides(4)
        myDocument.Shapes.AddPicture Filename:=slide4_picture2_down_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=Slide4_Pic2_Xpos, Top:=Slide4_Pic2_Ypos, Width:=Slide4_Pic2_Width, Height:=Slide2_Pic2_Height



'Display message box to select a location or folder, where pdf file will be saved

        ThisWorkbook.Activate
'MsgBox ("Select a location where PDF file will be saved")


'Get the destination location where you want to save your pdf file

'Set SavePptDialogBox = obPptApp.FileDialog(msoFileDialogFolderPicker)

'If SavePptDialogBox.Show = -1 Then
'Use the active presentation's name, replace is used to change file extension to pdf from pptx
'PdfFilePath = SavePptDialogBox.SelectedItems(1) & "\" & Replace(obPptApp.ActivePresentation.Name, "pptx", "pdf")
'MsgBox (PdfFilePath)
'Range("A1").Value = PdfFilePath
'PdfFilePath = "C:\Users\eryeh\Desktop\Daily Report\SNM PPT micro project\Vendor Code_Vendor Name_Q and L performance report_template.pdf"


'PdfFilePath = Range("B3").Value
'PdfFilePath = Range("B3").Value
       PdfFilePath = "C:\Localdata\Supplier KPI\" & MySupplierName & "_" & MySupplierCode & "_" & "Q and L performance report_" & PreviousMonth & ".pdf"
 
        obPptApp.ActivePresentation.ExportAsFixedFormat PdfFilePath, FixedFormatType:=ppFixedFormatTypePDF
'End If

'excel_save_route = Range("B2").Value
'

''NameOfPowerPoint = "Vendor Code_Vendor Name_Q and L performance report_20211018" & ".pptx"
''Save_As = FolderName & "\" & NameOfPowerPoint

excel_save_route = "C:\Localdata\Supplier KPI\" & MySupplierName & "_" & MySupplierCode & "_" & "Q and L performance report_" & PreviousMonth & ".pptx"
 
obPptApp.ActivePresentation.SaveAs Filename:=excel_save_route
      obPptApp.Windows(1).Close
        Next
'        i = i + 1

'Close the ppt

MsgBox ("All completed,you are a true genius:)")
End Sub

