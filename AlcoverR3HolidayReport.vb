
' ALCOVER RODRIGO MARTIN
' 12/3/2019


' THE CLASSROOM POINT OF VIEW IS:
'   PRACTICE THE IPO PROCESS.
'   NPRACTICE THE DESIGN OF HIERARCHY AND  CONTINUE PRACTICING FLOWCHARTS,
'   DOING CHANGES IN PROCESS RECORD MODULE ADDING RECORD SELECTION.
'   FOLLOW THE STRUCTURE OF THE CODE AND THE FLOWCHART TO FIND ERRORS.
'   PRACTICE WITH VARIABLES AND CONSTANT, RECORD DEFINITIONS, FILE DEFINITION, MODULE DEFINITIONS.
'   PRACTICE ACCUM, AVERAGES.
'   PRACTICE READ, WRITE, AND MOVE RECORDS.
'   IN THIS ASSIGNMENT WE AGAIN USE PAGINATION, 
'   WE OUTPUT 20 DETAIL LINES PER PAGE.
'   WE USE BOOLEAN TO MAKE A RECORD SELECTION, NESTED IF AND CASE, WE CAN ALSO USE COMPLEX IF OR COMBINED.
'   THE MARKUP CODE  IS DETERMINE USING A CASE STRUCTURE. THE DISCOUNT PERCENT IS DETERMINED
'   USING A NESTED IF.
'   I PREFER TO USE A SIMPLE STATMENT FOR COMPARATION IN THE RECORD SELECTION MODULE,
'   FOR MAKE THE WEEKLE CHANGES EASIER.
'   I TRY TO USE DECIMAL FOR MARK UP PERCENTAGE AND DECIMAL FOR THE MARK UP AMOUNT. 

' THE BUSINESS POINT OF VIEW IS:

'   THE PURPOSE OF THIS PROGRAM IS TO PRODUCE WEEKLY SALES REPORT OF

'   VARIOUS ITEMS SOLD, BASED ON THE OWNER'S DESIRE.

'   THE PROGRAM IS ABLE TO REPORT DIFFERENT ITEMS BASED ON THE SELECTION OF THE USER.


' THE PROGRAM DETAIL :
'                       ITEM# 
'                       DESCRIPTION
'                       WHOLESALE PRICE
'                       MARK UP CODE
'                       MARK UP PERCENTAGE --BASED ON THE MARKUP CODE ---
'                       MARK UP AMOUNT                       
'                       RETAIL PRICE                         
'                       QUANTITY SOLD
'                       EXTENDED PRICE
'                       DISCOUNT PERCENTAGE  --BASED ON THE QTY SOLD---
'                       DISCOUNT AMOUNT

' IN THE END OF THE REPORT THE TOTALS ANS AVERAGES:
'                       FINAL AMOUNT DUE
'                       FINAL TOTAL # ITEM SOLD
'                       FINAL TOTAL EXTENDED PRICE
'                       FINAL TOTAL DISCOUNTS
'                       FINAL TOTAL DUE
'                       AVERAGE DISCOUNT PER ITEM SOLD
'                       AVERAGE PRICE PAID ITEM


'  THE REPORT OUTPUT 20 DETAIL LINES PER PAGE AND

'  SPECIFIED THE PAGE NUMBER AT THE TOP OF EACH PAGE.





Module AlcoverR3HolidayReport
    '                                                  START OF PROGRAM


    Private HolidayNmoreReportFile As New Microsoft.
        VisualBasic.FileIO.TextFieldParser("HOLIDAYWK472019.TXT") 'FILE NAME

    Private CurentRecord() As String  ' CURRENT RECORDS
    '                                             NOW WE'LL DECLARE THE FILE WE'LL
    '                                       WE USE IN THE PROGRAM AND ASSICIATE IT 
    '                          WITH THE ACTUAL FILE NAME, WHERE THE DATA IS STORED

    '                                                    INITIALIZE CONSTANTS FIELD:
    Private Const MARKUP_PERCENT_CODE1_Decimal As Decimal = 5.0 '%
    Private Const MARKUP_PERCENT_CODE2_Decimal As Decimal = 10.0 '%
    Private Const MARKUP_PERCENT_CODE3_Decimal As Decimal = 15.0 '%
    Private Const MARKUP_PERCENT_CODE4_Decimal As Decimal = 20.0 '%
    Private Const MARKUP_PERCENT_CODE5_Decimal As Decimal = 25.0 '%

    Private Const QTYSOLD05_DESC_PERCENT_DECIMAL As Decimal = 0.0 '%
    Private Const QTYSOLD615_DESC_PERCENT_DECIMAL As Decimal = 5.5 '%
    Private Const QTYSOLD1630_DESC_PERCENT_DECIMAL As Decimal = 10.0 '%
    Private Const QTYSOLD3150_DESC_PERCENT_DECIMAL As Decimal = 12.5 '%
    Private Const QTYSOLD5175_DESC_PECENT_DECIMAL As Decimal = 20.0 '%
    Private Const QTYSOLD75_DESC_PERCENT_DECIMAL As Decimal = 30.0 '%


    '                                                    INPUT VARIABLES/FLIEDS:
    Private ItemNumberInteger As Integer
    Private DescriptionString As String
    Private WholesalePriceDecimal As Decimal
    Private MarkUpCodeInteger As Integer
    Private QtySoldInteger As Integer
    '                                                         CALCULATED FIELDS:
    Private MarkUpAmountDecimal As Decimal  '                  EACH / RECORD
    Private RetailPriceDecimal As Decimal
    Private ExtendedPriceDecimal As Decimal
    Private DiscountedAmountDecimal As Decimal
    Private FinalAmountDecimal As Decimal

    '                                                          AVERAGES
    Private AverageDiscountDecimal As Decimal
    Private AveragePaidItemDecimal As Decimal

    '                                                         ACUMULATED FIELDS
    Private AcumFinalTotalItemSoldDecimal As Decimal = 0
    Private AcumFinalTotalExtendedPriceDecimal As Decimal = 0
    Private AcumFinalTotalDiscountsDecimal As Decimal = 0
    Private AcumFinalTotalDueDecimal As Decimal = 0


    '                                             PAGINATION VARIABLES:

    Private LineCounterInteger As Integer = 99         '      99 FOR HEADINGS ON FIRST PAGE
    '                                                              
    Private Const PAGE_SIZE_INTEGER As Integer = 20
    '                                                              
    Private PageNumberInteger As Integer = 1 '             PAGE #'S FOR HEADINGS            
    '                                                      FILE RECORD AND FILE NAME DECLARATIONS:
    '                                                      WHEN THE FILE IS READ, 
    '                                                      THE RECORD IS PLACED IN THIS VARIABLE
    '                                                      ASSIGNED FIELDS FROM DESICIONS


    '                                                      BOOLEAN VARIABLE 
    '                                                      WORKING FIELDS
    Private RecordSelectionBoolean As Boolean
    '                                                      ASSIGNED FIELDS FROM DESICIONS
    Private MarkUpPercentDecimal As Decimal
    Private DiscountedPercentDecimal As Decimal

    Sub Main()   '                                         PROGRAM EXECUTION LOGIC STARTS.
        Call HouseKeeping()
        Do While Not HolidayNmoreReportFile.EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Private Sub HouseKeeping()  '                          LEVEL 2 CONTROL MODULES
        Call SetFileDelimiter()

    End Sub

    Private Sub ProcessRecords()
        Call ReadFile()
        Call RecordSelection()
        If RecordSelectionBoolean Then
            Call DetailCalculations()
            Call AccumulateTotals()
            Call WriteDetailLine()
        End If

    End Sub

    Private Sub EndOfJob()
        Call SummaryCalculations()
        Call SummaryOutput()
        Call CloseFile()

    End Sub


    Private Sub SetFileDelimiter()      '         HOUSEKEAPING MODULES
        '                                         DEFINES FILES AS A DELIMITER
        '                                         DEFINES DELIMITER AS A COMMA        
        HolidayNmoreReportFile.TextFieldType = FileIO.FieldType.Delimited

        HolidayNmoreReportFile.SetDelimiters(",")

    End Sub
    Private Sub ReadFile() '  READ WHOLE RECORD AND ASSIGN TO THE CURRENT RECORD VARIABLE
        CurentRecord = HolidayNmoreReportFile.ReadFields()
        ItemNumberInteger = CurentRecord(0)
        DescriptionString = CurentRecord(1) '            PLACE CURRENT RECORDS FIELDS 
        WholesalePriceDecimal = CurentRecord(2) '              THE CURRENT RECORD 1  
        MarkUpCodeInteger = CurentRecord(3) '               IS SKIP BECAUSE NOT USED
        QtySoldInteger = CurentRecord(4)
    End Sub


    Private Sub RecordSelection()
        '                                    RECORD SELECTION
        RecordSelectionBoolean = False
        If ItemNumberInteger >= 1000 And ItemNumberInteger <= 1999 Or
            ItemNumberInteger >= 2000 And ItemNumberInteger <= 2999 Or
            ItemNumberInteger >= 4000 And ItemNumberInteger <= 4999 Or
            ItemNumberInteger >= 6000 And ItemNumberInteger <= 6999 Then
            RecordSelectionBoolean = True
        End If
    End Sub


    Private Sub DetailCalculations()
        '                   CALCULATED DETAIL LINE

        Call DetermineItemCategory()
        MarkUpAmountDecimal = WholesalePriceDecimal * MarkUpPercentDecimal
        MarkUpAmountDecimal = MarkUpAmountDecimal / 100
        RetailPriceDecimal = WholesalePriceDecimal + MarkUpAmountDecimal
        ExtendedPriceDecimal = RetailPriceDecimal * QtySoldInteger

        Call DetermineDiscountPercent()
        DiscountedAmountDecimal = ExtendedPriceDecimal * DiscountedPercentDecimal
        DiscountedAmountDecimal = DiscountedAmountDecimal / 100
        FinalAmountDecimal = ExtendedPriceDecimal - DiscountedAmountDecimal
    End Sub


    Private Sub DetermineItemCategory()
        '                              DETERMINE ITEM CATEGORY BY MARKUPCODE

        Select Case MarkUpCodeInteger

            Case 1
                MarkUpPercentDecimal = MARKUP_PERCENT_CODE1_DECIMAL
            Case 2
                MarkUpPercentDecimal = MARKUP_PERCENT_CODE2_Decimal
            Case 3
                MarkUpPercentDecimal = MARKUP_PERCENT_CODE3_Decimal
            Case 4
                MarkUpPercentDecimal = MARKUP_PERCENT_CODE4_Decimal
            Case Else
                MarkUpPercentDecimal = MARKUP_PERCENT_CODE5_Decimal

        End Select
    End Sub



    Private Sub DetermineDiscountPercent()
        '                                 DETERMINE DISCOUNT PERCENT
        If QtySoldInteger <= 5 Then
            DiscountedPercentDecimal = QTYSOLD05_DESC_PERCENT_DECIMAL
        Else
            If QtySoldInteger <= 15 Then
                DiscountedPercentDecimal = QTYSOLD615_DESC_PERCENT_DECIMAL
            Else
                If QtySoldInteger <= 30 Then
                    DiscountedPercentDecimal = QTYSOLD1630_DESC_PERCENT_DECIMAL
                Else
                    If QtySoldInteger <= 50 Then
                        DiscountedPercentDecimal = QTYSOLD3150_DESC_PERCENT_DECIMAL
                    Else
                        If QtySoldInteger <= 75 Then
                            DiscountedPercentDecimal = QTYSOLD5175_DESC_PECENT_DECIMAL
                        Else
                            DiscountedPercentDecimal = QTYSOLD75_DESC_PERCENT_DECIMAL
                        End If
                    End If
                End If
            End If
        End If
    End Sub


    Private Sub AccumulateTotals()
        '                                       ACCUMULATE FINAL TOTALS 
        AcumFinalTotalItemSoldDecimal = AcumFinalTotalItemSoldDecimal + QtySoldInteger
        AcumFinalTotalExtendedPriceDecimal = AcumFinalTotalExtendedPriceDecimal + ExtendedPriceDecimal
        AcumFinalTotalDiscountsDecimal = AcumFinalTotalDiscountsDecimal + DiscountedAmountDecimal
        AcumFinalTotalDueDecimal = AcumFinalTotalDueDecimal + FinalAmountDecimal

    End Sub


    Private Sub WriteDetailLine()
        '                                          WRITE DETAIL LINE
        If LineCounterInteger >= PAGE_SIZE_INTEGER Then ' HERE WHEN THE CINE COUNTER INTEGER
            ' IS GREATER OR EQUAL TO PAGESIZE CALL WRITE HEADINGS
            Call WriteHeadings()
        End If
        Console.WriteLine(ItemNumberInteger.ToString.PadLeft(4) & Space(2) &
                          DescriptionString.PadRight(10) & Space(1) &
                          WholesalePriceDecimal.ToString("n").PadLeft(5) & Space(3) &
                          MarkUpCodeInteger & Space(2) &
                          MarkUpPercentDecimal.ToString.PadLeft(2) & Space(1) &
                          MarkUpAmountDecimal.ToString("n").PadLeft(5) & Space(2) &
                          RetailPriceDecimal.ToString("n").PadLeft(6) & Space(2) &
                          QtySoldInteger.ToString.PadLeft(2) & Space(3) &
                          ExtendedPriceDecimal.ToString("n").PadLeft(6) & Space(2) &
                          DiscountedPercentDecimal.ToString("n1").PadLeft(4) & Space(1) &
                          DiscountedAmountDecimal.ToString("n").PadLeft(6) & Space(1) &
                          FinalAmountDecimal.ToString("n").PadLeft(8))
        '                                     LineCounterInteger = LineCounterInteger +1    
        '                                     COUNT THE LINE PRINTED
        LineCounterInteger += 1 '             +=  IS A ' COMBINED OPERATOR'
        '                                     SHORTCUT FOR ACCUMULATION
        '                                     OUTPUT 1 LINE FOR EACH PERSON PROCESSED 
        '                                     TEST FOR PAGINATION
    End Sub

    Private Sub WriteHeadings()
        '                             WRITE HEADINGS MODULE IS PART OF PROCESS RECORD MODULES
        '                             AND IS CALL BY WRITE DETAILLINE WEN THE LINE COUNTER 
        '                             IS GREATER OR EQUAL TO 20.
        '                             WRITE REPORTHEADLINES
        Console.WriteLine()
        Console.WriteLine("Page " & PageNumberInteger.ToString("n0".PadLeft(2)) & Space(18) &
                          "Holidays N More Sales Report for")
        Console.WriteLine(Space(29) & "Rodrigo Martin Alcover")
        Console.WriteLine()                          'WRITE COLUMN LEADER LINES
        Console.WriteLine("Item" & Space(11) &
                          "WhlSale" & Space(2) &
                          "-- Markup --" & Space(2) &
                          "Retail" & Space(1) &
                          "Qty" & Space(1) &
                          "Extended" & Space(4) & "Discount" & Space(5) & "Final")
        Console.WriteLine("Num" & Space(3) &
                          "Desc" & Space(7) &
                          "Price" & Space(2) &
                          "Cde" & Space(1) &
                          "PC" & Space(3) &
                          "Amt" & Space(3) &
                          "Price" & Space(1) &
                          "Sld" & Space(4) &
                          "Price" & Space(3) &
                          "PC" & Space(5) &
                          "Amt" & Space(2) &
                          "Amt" & Space(1) & "Due")
        Console.WriteLine()
        LineCounterInteger = 0 '               RESET LINE COUNTER &
        PageNumberInteger += 1       '   ADD TO PAGE#     +=  IS CALLED A  COBINED OPERATOR
    End Sub
    '                                          END OF JOBS MODULES
    '                                          FINAL AVERAGES
    Private Sub SummaryCalculations()
        'STATE CALCULATIONS FOR DISCOUNT PER ITEM SOLD AND PRICE PAID PER ITEM

        AverageDiscountDecimal = AcumFinalTotalDiscountsDecimal / AcumFinalTotalItemSoldDecimal
        AveragePaidItemDecimal = AcumFinalTotalDueDecimal / AcumFinalTotalItemSoldDecimal

    End Sub



    Private Sub SummaryOutput()
        '                                         WRITE TOTAL LINE AND MOVE ACCUM & AVERAGES
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine("TOTALS & AVERAGES:")
        Console.WriteLine()
        Console.WriteLine(Space(5) & "Final Total # Item Sold" & Space(15) &
                          AcumFinalTotalItemSoldDecimal.ToString("n0").PadLeft(5))
        Console.WriteLine(Space(5) & "Final Total Extended Price" & Space(8) &
                          AcumFinalTotalExtendedPriceDecimal.ToString("c").PadLeft(9))
        Console.WriteLine(Space(5) & "Final Total Discounts" & Space(13) &
                          AcumFinalTotalDiscountsDecimal.ToString("c").PadLeft(9))
        Console.WriteLine(Space(5) & "Final Total Due" & Space(19) &
                          AcumFinalTotalDueDecimal.ToString("c").PadLeft(9))
        Console.WriteLine() '                         WRITE AVERAGE LINE AND MOVE AVERAGES
        Console.WriteLine()
        Console.WriteLine(Space(5) & "Average Discount Per Item Sold" & Space(8) &
                          AverageDiscountDecimal.ToString("c").PadLeft(5))
        Console.WriteLine(Space(5) & "Average Price Paid Per Item" & Space(7) &
                          AveragePaidItemDecimal.ToString("c").PadLeft(9))
        Console.WriteLine()
    End Sub

    Private Sub CloseFile()                        ' END OF JOB MODULES
        HolidayNmoreReportFile.Close() '             CLOSING THE FILE
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine("Click ENTER Close Output Window")
        Console.ReadKey() '  WRITE MESSAGE FOR PRESS ENTER AND
        '                    CLOSE THE WINDOW PROMPT 
    End Sub
End Module
