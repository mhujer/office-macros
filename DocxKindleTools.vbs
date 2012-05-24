' Resizes Word document for Kindle
Sub KindleResize()
    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(0)
        .BottomMargin = CentimetersToPoints(0)
        .LeftMargin = CentimetersToPoints(0)
        .RightMargin = CentimetersToPoints(0)
        .PageWidth = CentimetersToPoints(12)
        .PageHeight = CentimetersToPoints(8.9)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(0)
        .FooterDistance = CentimetersToPoints(0)
    End With
	
	' We need smaller text
    Selection.EndKey Unit:=wdStory
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    Selection.Font.Shrink
    Selection.Font.Shrink
    Selection.Font.Shrink
    
    'Remove header and Footer
    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
    
    ' Magic to hide header and footer
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

' Saves current document as PDF
Sub KindleSavePDF()
    Dim sNewFileName As String
    
    sNewFileName = ActiveDocument.FullName
    sNewFileName = Replace(sNewFileName, ".docx", "_kindle.pdf")
    
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        sNewFileName, ExportFormat:= _
        wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub
