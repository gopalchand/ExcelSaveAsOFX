Option Explicit

Public Sub cbGenerateOFX()
' Generate OFX file from the data in the 'Export to OFX' spreadsheet

On Error Resume Next

Dim SGML_PAD As String

' General variables
Dim iResult As Integer
Dim iTransaction As Integer
Dim dtDatetime As Date

' Output file name
Dim OutputFilename As String

Dim CurrentWB As Workbook, TempWB As Workbook

Set CurrentWB = ActiveWorkbook

' OFX file header
Dim OFX_HEADER As String
Dim OFX_SIGNONMSGSRSV1_HEADER As String
Dim OFX_SIGNONMSGSRSV1_DTSERVER_HEADER As String
Dim OFX_SIGNONMSGSRSV1_DTSERVER_FOOTER As String
Dim OFX_SIGNONMSGSRSV1_FOOTER As String
Dim OFX_BANKMSGSRSV1_HEADER As String
Dim OFX_BANKMSGSRSV1_FOOTER As String

' Bank account information
Dim OFX_BANKACCTFROM_HEADER As String
Dim OFX_BANKID_HEADER As String
Dim OFX_BANKID_FOOTER As String
Dim OFX_ACCTID_HEADER As String
Dim OFX_ACCTID_FOOTER As String
Dim OFX_ACCTTYPE_HEADER As String
Dim OFX_ACCTTYPE_FOOTER As String
Dim OFX_BANKACCTFROM_FOOTER As String

' Transaction list information
Dim OFX_BANKTRANLIST_HEADER As String
Dim OFX_BANKTRANLIST_DTSTART_HEADER As String
Dim OFX_BANKTRANLIST_DTSTART_FOOTER As String
Dim OFX_BANKTRANLIST_DTEND_HEADER As String
Dim OFX_BANKTRANLIST_DTEND_FOOTER As String
Dim OFX_BANKTRANLIST_FOOTER As String

' Transactions information
Dim OFX_STMTTRN_HEADER As String
Dim OFX_STMTTRN_TRNTYPE_HEADER As String
Dim OFX_STMTTRN_TRNTYPE_FOOTER As String
Dim OFX_STMTTRN_DTPOSTED_HEADER As String
Dim OFX_STMTTRN_DTPOSTED_FOOTER As String
Dim OFX_STMTTRN_TRNAMT_HEADER As String
Dim OFX_STMTTRN_TRNAMT_FOOTER As String
Dim OFX_STMTTRN_FITID_HEADER As String
Dim OFX_STMTTRN_FITID_FOOTER As String
Dim OFX_STMTTRN_NAME_HEADER As String
Dim OFX_STMTTRN_NAME_FOOTER As String
Dim OFX_STMTTRN_MEMO_HEADER As String
Dim OFX_STMTTRN_MEMO_FOOTER As String
Dim OFX_STMTTRN_FOOTER As String

' Ledger balance information
Dim OFX_LEDGERBAL_HEADER As String
Dim OFX_LEDGERBAL_BALAMT_HEADER As String
Dim OFX_LEDGERBAL_BALAMT_FOOTER As String
Dim OFX_LEDGERBAL_DTASOF_HEADER As String
Dim OFX_LEDGERBAL_DTASOF_FOOTER As String
Dim OFX_LEDGERBAL_FOOTER As String

' Closing tag
Dim OFX_FOOTER As String
Dim OFX_STMTRS_FOOTER As String
Dim OFX_STMTTRNRS_FOOTER As String

Dim CURRENCY_CODE As String
Dim SORT_CODE As String
Dim ACCOUNT_NUMBER As String

' CUSTOMISE
SGML_PAD = Chr(9) 'Or ""
CURRENCY_CODE = "GBP"
SORT_CODE = "112233"
ACCOUNT_NUMBER = "12345678"

' OFX file header
OFX_HEADER = "OFXHEADER:100" & Chr(13) & _
                "DATA:OFXSGML" & Chr(13) & _
                "VERSION:102" & Chr(13) & _
                "SECURITY:NONE" & Chr(13) & _
                "ENCODING:USASCII" & Chr(13) & _
                "CHARSET:1252" & Chr(13) & _
                "COMPRESSION:NONE" & Chr(13) & _
                "OLDFILEUID:NONE" & Chr(13) & _
                "NEWFILEUID:NONE" & Chr(13) & Chr(13) & _
                "<OFX>"

OFX_SIGNONMSGSRSV1_HEADER = SGML_PAD & "<SIGNONMSGSRSV1>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & "<SONRS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "<STATUS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<CODE>0</CODE>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<SEVERITY>INFO</SEVERITY>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "</STATUS>"
OFX_SIGNONMSGSRSV1_DTSERVER_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & "<DTSERVER>"
OFX_SIGNONMSGSRSV1_DTSERVER_FOOTER = "</DTSERVER>"

OFX_SIGNONMSGSRSV1_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & "<LANGUAGE>ENG</LANGUAGE>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "<INTU.BID>01267</INTU.BID>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & "</SONRS>" & Chr(13) & _
                            SGML_PAD & "</SIGNONMSGSRSV1>"

OFX_BANKMSGSRSV1_HEADER = SGML_PAD & "<BANKMSGSRSV1>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & "<STMTTRNRS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "<TRNUID>1</TRNUID>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "<STATUS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<CODE>0</CODE>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<SEVERITY>INFO</SEVERITY>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "</STATUS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & "<STMTRS>" & Chr(13) & _
                            SGML_PAD & SGML_PAD & SGML_PAD & "<CURDEF>" & CURRENCY_CODE & "</CURDEF>"
                            
OFX_STMTRS_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & "</STMTRS>"
OFX_STMTTRNRS_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & "</STMTTRNRS>"
OFX_BANKMSGSRSV1_FOOTER = SGML_PAD & "</BANKMSGSRSV1>"

' Bank account information
OFX_BANKACCTFROM_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & "<BANKACCTFROM>"
OFX_BANKID_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<BANKID>"
OFX_BANKID_FOOTER = "</BANKID>"
OFX_ACCTID_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<ACCTID>"
OFX_ACCTID_FOOTER = "</ACCTID>"
OFX_ACCTTYPE_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<ACCTTYPE>"
OFX_ACCTTYPE_FOOTER = "</ACCTTYPE>"
OFX_BANKACCTFROM_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & "</BANKACCTFROM>"

' Transaction list information
OFX_BANKTRANLIST_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & "<BANKTRANLIST>"
OFX_BANKTRANLIST_DTSTART_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<DTSTART>"
OFX_BANKTRANLIST_DTSTART_FOOTER = "</DTSTART>"
OFX_BANKTRANLIST_DTEND_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<DTEND>"
OFX_BANKTRANLIST_DTEND_FOOTER = "</DTEND>"
OFX_BANKTRANLIST_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "</BANKTRANLIST>"

' Transactions information
OFX_STMTTRN_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<STMTTRN>"
OFX_STMTTRN_TRNTYPE_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<TRNTYPE>"
OFX_STMTTRN_TRNTYPE_FOOTER = "</TRNTYPE>"
OFX_STMTTRN_DTPOSTED_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<DTPOSTED>"
OFX_STMTTRN_DTPOSTED_FOOTER = "</DTPOSTED>"
OFX_STMTTRN_TRNAMT_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<TRNAMT>"
OFX_STMTTRN_TRNAMT_FOOTER = "</TRNAMT>"
OFX_STMTTRN_FITID_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<FITID>"
OFX_STMTTRN_FITID_FOOTER = "</FITID>"
OFX_STMTTRN_NAME_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<NAME>"
OFX_STMTTRN_NAME_FOOTER = "</NAME>"
OFX_STMTTRN_MEMO_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<MEMO>"
OFX_STMTTRN_MEMO_FOOTER = "</MEMO>"
OFX_STMTTRN_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "</STMTTRN>"

' Ledger balance information
OFX_LEDGERBAL_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<LEDGERBAL>"
OFX_LEDGERBAL_BALAMT_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<BALAMT>"
OFX_LEDGERBAL_BALAMT_FOOTER = "</BALAMT>"
OFX_LEDGERBAL_DTASOF_HEADER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "<DTASOF>"
OFX_LEDGERBAL_DTASOF_FOOTER = "</DTASOF>"
OFX_LEDGERBAL_FOOTER = SGML_PAD & SGML_PAD & SGML_PAD & SGML_PAD & "</LEDGERBAL>"

' Closing tag
OFX_FOOTER = "</OFX>"

' Open output file
'OutputFilename = Worksheets("XLS2OFX").Range("OutputFile")
OutputFilename = CurrentWB.Path & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & ".ofx"

Dim fs
Dim ofxstream

Set fs = CreateObject("Scripting.FileSystemObject")
Set ofxstream = fs.CreateTextFile(OutputFilename, True)

Dim iReturn As Integer
Dim StatementDate As String
Dim StatementStartDate As String
Dim StatementEndDate As String


If Err.Number <> 0 Then
    iReturn = MsgBox(Error(Err.Number), vbCritical, "XLS2OFX Runtime Error")
Else
    ' Write OFX Header
    ofxstream.WriteLine (OFX_HEADER)
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_HEADER)
    'dtDatetime = Worksheets("XLS2OFX").Range("StatementStartDate")
    dtDatetime = Range("A2").Value
    StatementStartDate = Format(dtDatetime, "yyyymmdd") & "000000"
    StatementDate = Format(Now(), "yyyymmddHHMMSS")
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_DTSERVER_HEADER & StatementDate & OFX_SIGNONMSGSRSV1_DTSERVER_FOOTER)
    ofxstream.WriteLine (OFX_SIGNONMSGSRSV1_FOOTER)
    
    Dim AcctCurrency As String
    'AcctCurrency = "GBP" '= Worksheets("XLS2OFX").Range("AcctCurrency")
    ofxstream.WriteLine (OFX_BANKMSGSRSV1_HEADER)
    
    ofxstream.WriteLine (OFX_BANKACCTFROM_HEADER)
    Dim BankId As String
    BankId = SORT_CODE 'Worksheets("XLS2OFX").Range("BankId")
    ofxstream.WriteLine (OFX_BANKID_HEADER & BankId & OFX_BANKID_FOOTER)
    Dim AccountNo As String
    AccountNo = SORT_CODE & ACCOUNT_NUMBER 'Worksheets("XLS2OFX").Range("AccountNo")
    ofxstream.WriteLine (OFX_ACCTID_HEADER & AccountNo & OFX_ACCTID_FOOTER)
    Dim AcctType As String
    AcctType = "CHECKING" 'Worksheets("XLS2OFX").Range("AcctType")
    ofxstream.WriteLine (OFX_ACCTTYPE_HEADER & AcctType & OFX_ACCTTYPE_FOOTER)
    ofxstream.WriteLine (OFX_BANKACCTFROM_FOOTER)
    
    ' Write financial transactions
    ofxstream.WriteLine (OFX_BANKTRANLIST_HEADER)
    ofxstream.WriteLine (OFX_BANKTRANLIST_DTSTART_HEADER & StatementStartDate & OFX_BANKTRANLIST_DTSTART_FOOTER)
    'StatementStartDate = Worksheets("XLS2OFX").Range("StatementEndDate")
    dtDatetime = Cells(Rows.Count, 1).End(xlUp).Value 'Worksheets("XLS2OFX").Range("StatementEndDate ")
    StatementEndDate = Format(dtDatetime, "yyyymmdd") & "000000"
    ofxstream.WriteLine (OFX_BANKTRANLIST_DTEND_HEADER & StatementEndDate & OFX_BANKTRANLIST_DTEND_FOOTER)
    
    Dim PreviousBalance
    PreviousBalance = 0 'Worksheets("XLS2OFX").Range("PreviousBalance")
    Dim FinalBalance
    FinalBalance = PreviousBalance
    iTransaction = 1
    
    Dim rgeTransactionList As Range
    Set rgeTransactionList = CurrentWB.ActiveSheet.Range("A1")

    'Columns are Date, Memo and Amount
    
    While rgeTransactionList.Offset(iTransaction, 0).Value <> ""
        'Get transaction information
        dtDatetime = rgeTransactionList.Offset(iTransaction, 0).Value ' Worksheets("XLS2OFX").Range("TransactionListTop").Offset(iTransaction, 0).Value
        Dim sTranDate
        sTranDate = Format(dtDatetime, "yyyymmdd") & "000000"
        Dim sTranName As String
        sTranName = rgeTransactionList.Offset(iTransaction, 1).Value
        Dim sTranMemo
        sTranMemo = rgeTransactionList.Offset(iTransaction, 1).Value
        
        Dim sTranAmount As Double
        sTranAmount = rgeTransactionList.Offset(iTransaction, 2).Value
        
        'Record transaction in OFX format
        ofxstream.WriteLine (OFX_STMTTRN_HEADER)
        ofxstream.WriteLine (OFX_STMTTRN_TRNTYPE_HEADER & "OTHER" & OFX_STMTTRN_TRNTYPE_FOOTER)
        FinalBalance = FinalBalance + Val(sTranAmount)
                  
        ofxstream.WriteLine (OFX_STMTTRN_DTPOSTED_HEADER & sTranDate & OFX_STMTTRN_DTPOSTED_FOOTER)
        ofxstream.WriteLine (OFX_STMTTRN_TRNAMT_HEADER & sTranAmount & OFX_STMTTRN_TRNAMT_FOOTER)
        'Emulated FTID format is date stamp plus transaction number (should be unique)
        'Example: "200303170001"
        Dim sTranFTID
        sTranFTID = Format(dtDatetime, "yyyymmdd") & Format(iTransaction, "0000")
        ofxstream.WriteLine (OFX_STMTTRN_FITID_HEADER & sTranFTID & OFX_STMTTRN_FITID_FOOTER)
        ofxstream.WriteLine (OFX_STMTTRN_NAME_HEADER & sTranName & OFX_STMTTRN_NAME_FOOTER)
        If Len(sTranMemo) > 0 Then
            ofxstream.WriteLine (OFX_STMTTRN_MEMO_HEADER & sTranMemo & OFX_STMTTRN_MEMO_FOOTER)
        End If
        ofxstream.WriteLine (OFX_STMTTRN_FOOTER)
        
        'Get next transaction
        iTransaction = iTransaction + 1
    Wend
    
    ofxstream.WriteLine (OFX_BANKTRANLIST_FOOTER)
    'Ledger balance
    ofxstream.WriteLine (OFX_LEDGERBAL_HEADER)
    ofxstream.WriteLine (OFX_LEDGERBAL_BALAMT_HEADER & Str(0) & OFX_LEDGERBAL_BALAMT_FOOTER)
    ofxstream.WriteLine (OFX_LEDGERBAL_DTASOF_HEADER & StatementEndDate & OFX_LEDGERBAL_DTASOF_FOOTER)
    ofxstream.WriteLine (OFX_LEDGERBAL_FOOTER)
    
    ' Write OFX Footer
    ofxstream.WriteLine (OFX_STMTRS_FOOTER)
    ofxstream.WriteLine (OFX_STMTTRNRS_FOOTER)
    ofxstream.WriteLine (OFX_BANKMSGSRSV1_FOOTER)
    ofxstream.WriteLine (OFX_FOOTER)
    
    ' Close file
    ofxstream.Close
End If

End Sub

