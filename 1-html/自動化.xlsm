' 指定されたソースシートの範囲から指定された宛先シートの範囲へデータを転送します。
'
' @param sourceSheetName      データを転送する元のワークシートの名前。
' @param destinationSheetName データを転送する先のワークシートの名前。
' @param sourceRangeAddress   ソースシート上のセル範囲のアドレス。
' @param destinationRangeAddress 宛先シート上のデータが配置されるセル範囲のアドレス。
'
' @example
'           TransferData "商品一覧", "テンプレート", "B2:D2", "A15:C15"
'           これは、"商品一覧"シートのセルB2:D2から"テンプレート"シートのセルA15:C15へ
'           値を転送する例です。
'
Private Sub TransferData(sourceSheetName As String, destinationSheetName As String, sourceRangeAddress As String, destinationRangeAddress As String)
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim srcRange As Range
    Dim destRange As Range

    ' ソースシートと宛先シートを設定します。
    Set srcSheet = ThisWorkbook.Worksheets(sourceSheetName)
    Set destSheet = ThisWorkbook.Worksheets(destinationSheetName)

    ' データを転送する範囲を設定します。
    Set srcRange = srcSheet.Range(sourceRangeAddress)
    Set destRange = destSheet.Range(destinationRangeAddress)

    ' ソースから宛先へデータを転送します。
    destRange.Cells(1, 1).Value = srcRange.Cells(1, 1).Value
    destRange.Cells(1, 2).Value = srcRange.Cells(1, 2).Value
    destRange.Cells(1, 3).Value = srcRange.Cells(1, 3).Value
End Sub

Private Sub Check1_Click()
    TransferData "商品一覧", "テンプレート", "B2:D2", "A15:C15"
End Sub

Private Sub Check10_Click()
    TransferData "商品一覧", "テンプレート", "B3:D3", "A15:C15"
End Sub

Private Sub Check2_Click()
    TransferData "商品一覧", "テンプレート", "B4:D4", "A15:C15"
End Sub

Private Sub Check3_Click()
    TransferData "商品一覧", "テンプレート", "B5:D5", "A15:C15"
End Sub

Private Sub Check4_Click()
    TransferData "商品一覧", "テンプレート", "B6:D6", "A15:C15"
End Sub

Private Sub Check5_Click()
    TransferData "商品一覧", "テンプレート", "B7:D7", "A15:C15"
End Sub

Private Sub Check6_Click()
    TransferData "商品一覧", "テンプレート", "B8:D8", "A15:C15"
End Sub

Private Sub Check7_Click()
    TransferData "商品一覧", "テンプレート", "B9:D9", "A15:C15"
End Sub

Private Sub Check8_Click()
    TransferData "商品一覧", "テンプレート", "B10:D10", "A15:C15"
End Sub

Private Sub Check9_Click()
    TransferData "商品一覧", "テンプレート", "B11:D11", "A15:C15"
End Sub
