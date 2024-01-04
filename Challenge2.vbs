{\rtf1\ansi\ansicpg1252\cocoartf2759
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub Tickers()\
\
For Each ws In Worksheets\
\
    'Set all Dimensions and Parameters for first part of challenge\
    Dim Total_Vol As LongLong\
    Dim RowCount As LongLong\
    Dim i, r As Long\
    Dim LastRow As LongLong\
    Dim Yearly_Change As Double\
    Dim OpenPerc As Double\
    Dim ClosingPerc As Double\
    \
    'Add Titles for Ticker Symbol, Yearly Change, Percentage Change and Total Volume\
    ws.Range("I1").Value = "Ticker Symbol"\
    ws.Range("J1").Value = "Yearly Change"\
    ws.Range("K1").Value = "Percentage Change"\
    ws.Range("L1").Value = "Total Volume"\
    ws.Range("I1:P1").Font.Bold = True\
    \
    'Find last filled row in first column\
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
    \
    'Set RowCount Variable\
    RowCount = 2\
    \
    'Setting a second rowc ounter to 2\
    r = 2\
    \
    'Create a loop for the rows\
    For i = 2 To LastRow\
\
        'Define variables for OpenPerc & ClosingPerc\
        OpenPerc = ws.Cells(r, 3).Value\
        ClosingPerc = ws.Cells(i, 6).Value\
        \
        'Add Total volume per ticker in the loop\
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value\
        \
        'Check each row in column 1 and check if row after has the same value\
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then\
        \
            'If not, add Ticker Symbol to I column\
            ws.Cells(RowCount, 9).Value = ws.Cells(i, 1).Value\
            \
            'Calculate Yearly Change by the equation closing percentave minus open percentage\
            ws.Cells(RowCount, 10).Value = ClosingPerc - OpenPerc\
            Yearly_Change = ws.Cells(RowCount, 10).Value\
            \
            'Calculate percentage change by yearly change divided by open percentage\
            ws.Cells(RowCount, 11).Value = Yearly_Change / OpenPerc\
            \
            'Format K column to Percentage\
            ws.Cells(RowCount, 11).NumberFormat = "0.00%"\
            \
            'Set conditional formatting Yearly Change\
            If ws.Cells(RowCount, 10).Value < 0 Then\
            \
                'Set values that are less than 0 to red\
                ws.Cells(RowCount, 10).Interior.ColorIndex = 3\
            Else\
                'Set values that are greater than 0 to green\
                ws.Cells(RowCount, 10).Interior.ColorIndex = 4\
            End If\
            \
            'Print Total Volume into column 12 in Row as per RowCoutn\
            ws.Cells(RowCount, 12).Value = Total_Vol\
\
            'Reset Total_Vol to 0\
            Total_Vol = 0\
            \
            'Increase RowCount by 1\
            RowCount = RowCount + 1\
            \
            'Start new row from the new ticker range\
            r = i + 1\
            \
         End If\
    Next i\
\
    'Set Dimensions for "bonus" section\
    Dim MaxValue As Double\
    Dim MaxTicker As String\
    Dim MinValue As Double\
    Dim MinTicker As String\
    Dim MaxTotal As LongLong\
    Dim MaxTotalTicker As String\
    Dim LastLast As Long\
    \
    'Set New headings for "Bonus" section\
    ws.Range("O1").Value = "Ticker"\
    ws.Range("P1").Value = "Value"\
    ws.Range("N2").Value = "Greatest % Increase"\
    ws.Range("N3").Value = "Greatest % Decrease"\
    ws.Range("N4").Value = "Greatest Total Volume"\
    ws.Range("N2:N4").Font.Bold = True\
    \
    'Set new counter for LastRow for new second table\
    LastLast = ws.Cells(Rows.Count, 11).End(xlUp).Row\
    \
    'Find the maximum and minimum values\
    MaxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastLast))\
    MinValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastLast))\
    MaxTotal = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastLast))\
\
    'Create a loop to find the Ticker names\
    For i = 2 To LastLast\
        If ws.Cells(i, 11).Value = MaxValue Then\
                MaxTicker = ws.Cells(i, 9).Value\
\
        ElseIf ws.Cells(i, 11).Value = MinValue Then\
                MinTicker = ws.Cells(i, 9).Value\
\
        ElseIf ws.Cells(i, 12).Value = MaxTotal Then\
                MaxTotalTicker = ws.Cells(i, 9).Value\
\
        End If\
\
    Next i\
\
        ' Write for Greates percentage increase in percentage format\
        ws.Range("O2").Value = MaxTicker\
        ws.Range("P2").Value = MaxValue\
        ws.Range("P2").NumberFormat = "0.00%"\
        \
        ' Write the results for greatest percentage decrease in percentave format\
        ws.Range("O3").Value = MinTicker\
        ws.Range("P3").Value = MinValue\
        ws.Range("P3").NumberFormat = "0.00%"\
        \
        ' Write the results for greatest total volume in scientific format\
        ws.Range("O4").Value = MaxTotalTicker\
        ws.Range("P4").Value = MaxTotal\
        ws.Range("P4").NumberFormat = "##0.00E+0"\
        \
        ' Autofit to display data\
        ws.Columns("A:P").AutoFit\
               \
    Next ws\
    \
    \
End Sub}