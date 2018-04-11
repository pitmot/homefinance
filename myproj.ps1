cls

# Declaration of Import File variables

$username                                 = $env:username
$SourceFilesPathName                      = "C:\Users\"+$username+"\Dropbox\_Home Finance\myproj\"
$BankTrasactionsSourceFilesPathName       = $SourceFilesPathName + "BankFiles\"
$ChequeDataSourceFilesPathName            = $SourceFilesPathName + "Chq_Files\"
$CreditCardDataSourceFilesPathName        = $SourceFilesPathName + "CC_Files\"

$BankTransactionsImportFileName           = "Bank Transactions from start of 2015 until present date2.csv"
$BankTransactionsCategoriesImportFileName = "BankTransCategories6.csv"
$ChequeDataImportFileName                 = "ChequeData3.csv"
$ExportFileName                           = "exp19.csv"

$BankTransactionsImportFilePath           = $BankTrasactionsSourceFilesPathName + $BankTransactionsImportFileName
$BankTransactionsCategoriesImportFilePath = $BankTrasactionsSourceFilesPathName + $BankTransactionsCategoriesImportFileName
$ChequeDataImportFilePath                 = $SourceFilesPathName + $ChequeDataImportFileName 
$ExportFilePath                           = $SourceFilesPathName + $ExportFileName 

$BankTransList =@()
$BankTemp = @()
$ChqNum = ""
$ChqDescription = ""
$ChqPayee = ""
$TransCategory =""
$TransDescription =""


# Importing CSV files to arrays
$BankTransactions = import-csv $BankTransactionsImportFilePath -Encoding UTF8
$BankTransactionsCategories = import-csv $BankTransactionsCategoriesImportFilePath -Encoding UTF8
$ChequeData = import-csv $ChequeDataImportFilePath   -Encoding UTF8

$ChqCategory = ""
$ChqDetail = ""

$BankTransactions |
    ForEach-Object { $TransDescription = $_.תאור.Split(" ")
                    $TransDescription2 = $TransDescription[0]+ " " + $_.אסמכתא
                     if ($TransDescription[0] -eq "שיק") 
                            {$ChqNum = $_.אסמכתא;$ChqCategory = $ChequeData | ? {$_."Cheque Number" -eq $ChqNum} | select -ExpandProperty category;$ChqDetail = $ChequeData | ? {$_."Cheque Number" -eq $ChqNum} | select -ExpandProperty detail

                    $BankTempInstance = New-Object -TypeName psobject
                    $BankTempInstance | Add-Member -MemberType NoteProperty -Name Description -Value $TransDescription2
                    $BankTempInstance | Add-Member -MemberType NoteProperty -Name Category -Value $ChqCategory
                    $BankTempInstance | Add-Member -MemberType NoteProperty -Name Detail -Value $ChqDetail

                    $BankTemp += $BankTempInstance
                    }
                    $ChqCategory = ""
                    $ChqDetail = ""
                    $ChqNum = ""
                    }
$BankTransactionsCategories1 = $BankTransactionsCategories + $BankTemp

$BankTransactions |
    ForEach-Object { $TransDescription = $_.תאור
                    $TransDescription1 = $_.תאור.Split(" ")
                    $TransDescription2 = $TransDescription1 + " " + $_.אסמכתא
                    $TransDescription3 = $TransDescription
                    if ($TransDescription1[0] -eq "שיק") 
                            {$ChqNum = $_.אסמכתא;$ChqPayee = $ChequeData | ? {$_."Cheque Number" -eq $ChqNum} | select -ExpandProperty payee;$TransDescription3 = $TransDescription2}

                    $TransCategory = $BankTransactionsCategories1 | ? {$_.Description -eq $TransDescription3} | select -ExpandProperty category

                    $BankTransInstance = New-Object -TypeName psobject
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name Date -Value $_.'תאריך '
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name Description -Value $TransDescription
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name VerifyID -Value $_.אסמכתא
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name Debit -Value $_.חובה
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name Credit -Value $_.זכות
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name Balance -Value $_.'יתרה בש"ח'
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name "Category" -Value $TransCategory
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name "Cheque Number" -Value $ChqNum
                    $BankTransInstance | Add-Member -MemberType NoteProperty -Name "Cheque Payee" -Value $ChqPayee
                    $BankTransList += $BankTransInstance
                    $ChqNum = ""
                    $ChqDescription = ""
                    $ChqPayee = ""
                    $TransCategory =""
                    $TransDescription =""
#                    $BankTransList.count
                    }

$BankTransList | Export-Csv $ExportFilePath -Encoding UTF8 -NoTypeInformation

