Attribute VB_Name = "MdAnalysis"
Option Explicit
Sub stock_analysis()

Application.ScreenUpdating = False

'Variable Declaration

    Dim sh As Worksheet
    Dim ticker As String
    Dim ticker_row As Integer
    Dim vol As Double
    Dim end_row As Long
    Dim row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim great_inc_v As Double
    Dim great_dec_v As Double
    Dim great_vol_v As Double
    Dim great_inc_t As String
    Dim great_dec_t As String
    Dim great_vol_t As String

'Set the start row
    Const start_row As Integer = 2
    
'Loop through the sheets
    For Each sh In Worksheets

        'Clear analysis area
        sh.Range("I1").CurrentRegion.Interior.ColorIndex = 0
        sh.Range("I1").CurrentRegion.ClearContents
        sh.Range("P1").CurrentRegion.Interior.ColorIndex = 0
        sh.Range("P1").CurrentRegion.ClearContents
        
    'Variables assigment
        
        vol = 0
        
        'find the first ticker
        ticker = sh.Cells(2, 1).Value
        ticker_row = 2
        
        'find the first opening price
        open_price = sh.Cells(2, 3).Value
        
        'Iniate greatest part variables
        great_inc_v = 0
        great_dec_v = 0
        great_vol_v = 0
        
        'Find the last row for each sheet
        end_row = sh.Range("A1").CurrentRegion.Rows.Count
    
    'Calculate total volumes
    
        For row = start_row To end_row
            
            vol = vol + sh.Cells(row, 7).Value
            
            'Verify if current ticker is different of the next ticker
            'If it's different, write summary and calculate greatest values
            If sh.Cells(row, 1).Value <> sh.Cells(row + 1, 1) Then
    
                'Write ticker and volume
                sh.Cells(ticker_row, 9).Value = ticker
                sh.Cells(ticker_row, 12).Value = vol
                
                'Calculate Yearly Change
                close_price = sh.Cells(row, 6).Value
                sh.Cells(ticker_row, 10).Value = close_price - open_price
                If open_price <> 0 Then
                    sh.Cells(ticker_row, 11).Value = (close_price / open_price) - 1
                Else
                    sh.Cells(ticker_row, 11).Value = 0
                End If
                
                'Aplly colors
                If (close_price - open_price) >= 0 Then
                    sh.Cells(ticker_row, 10).Interior.Color = VBA.ColorConstants.vbGreen
                Else
                    sh.Cells(ticker_row, 10).Interior.Color = VBA.ColorConstants.vbRed
                End If 'End colors setting
                 
                'Calculate greatest % increase
                If sh.Cells(ticker_row, 11).Value > great_inc_v Then
                    great_inc_v = sh.Cells(ticker_row, 11).Value
                    great_inc_t = ticker
                End If
        
                'Calculate greatest % decrease
                If sh.Cells(ticker_row, 11).Value < great_dec_v Then
                    great_dec_v = sh.Cells(ticker_row, 11).Value
                    great_dec_t = ticker
                End If
                
                'Calculate greatest volume
                If sh.Cells(ticker_row, 12).Value > great_vol_v Then
                    great_vol_v = sh.Cells(ticker_row, 12).Value
                    great_vol_t = ticker
                End If
                
                'Reset variables
                ticker = sh.Cells(row + 1, 1).Value
                vol = 0
                open_price = sh.Cells(row + 1, 3).Value
                ticker_row = ticker_row + 1
                
            End If 'End change ticker verification
            
        
        Next row 'End loop throug rows
    
        'Write greatest values
        sh.Range("P2").Value = great_inc_t
        sh.Range("Q2").Value = great_inc_v
        sh.Range("P3").Value = great_dec_t
        sh.Range("Q3").Value = great_dec_v
        sh.Range("P4").Value = great_vol_t
        sh.Range("Q4").Value = great_vol_v
        
    'Create results header and format number
    
        sh.Range("I1").Value = "Ticker"
        sh.Range("J1").Value = "Yearly Change"
        sh.Range("K1").Value = "Percent Change"
        sh.Range("L1").Value = "Total Stock Volume"
        sh.Range("I:L").Columns.AutoFit
        sh.Range("P1").Value = "Ticker"
        sh.Range("Q1").Value = "Value"
        sh.Range("O2").Value = "Greatest % Increase"
        sh.Range("O3").Value = "Greatest % Decrease"
        sh.Range("O4").Value = "Greatest Total Volume"
        'Aplly format
        Range("J2:J" & ticker_row).NumberFormat = "0.00"
        Range("K2:K" & ticker_row).NumberFormat = "0.00%"
        sh.Range("Q2").Style = "Percent"
        sh.Range("Q2").NumberFormat = "0.00%"
        sh.Range("Q3").NumberFormat = "0.00%"
        sh.Range("Q4").NumberFormat = "0.00"
        sh.Range("O:Q").Columns.AutoFit
            
    Next sh 'End loop through the sheets

Application.ScreenUpdating = True

End Sub

