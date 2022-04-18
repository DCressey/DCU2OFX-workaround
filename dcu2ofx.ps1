<#
.NOTES
    Script:  dcu2ynab                  Rev:  1.0
    Author:  dcressey                  Date: 4/17/22

 Known Issues:
     negative numbers with parentheses haven't been tested.
     fitids will be generated in the same sequence every run.
     Category field not included.
     $balance has not been tested.


     
.SYNOPSIS
    Converts data as delivered by DCU into format for use with YNAB
.DESCRIPTION
     This script reads a file named checking.csv that contains transaction data
     from a checking account at the DCU, and downloaded by the DCU customer interface.
     It transforms the data from CSV format to OFX format,  a format accepted by YNAB
     and several other PFM packages.

     More documentation is available in dcu2ynab.rtf, a wordpad document.

#>

# This function gets the amount, eliminating currency symbol and commas.

function Get-Amount ($amount) {
$abs  = $amount -replace '[^0-9.]'
If ($amount -match '-|\(') { "-$abs"} else {$abs}
}

# This function distinguishes between credits and debits
# Note that this is the opposite of the way bookkeepers think

function Get-Type ($amount) {
if ($amount -match  '-|\(') {'DEBIT'} else {'CREDIT'}
}



#  This function is just a poor man's template engine.

function Expand-csv {
    [CmdletBinding()]
    Param (
       [Parameter(Mandatory=$true)] [string] $driver,
       [Parameter(Mandatory=$true)] [string] $template
    )
    Process {
       Import-Csv $driver | % {
           $_.psobject.properties | % {Set-variable -name $_.name -value $_.value}
           Get-Content $template | % {$ExecutionContext.InvokeCommand.ExpandString($_)} 
       }
    }
}

# the main script starts here

#Welcome

@'


Welcome to the DCU to OFX data converter!

This script will read the transactions the DCU provides as a csv file and produce
an ofx file that Personal Finance Manager packages can process.  

There are some items of data that are needed in the OFX file that the DCU
does not provide in the CSV file.  The following dialog collects this data
from you.

More documentation in dcu2ofx.txt.

'@





# User dialog to fill in missing data

$acctid = Read-Host 'The account number'
$dt = Read-Host 'The start date for this batch (eg 4/1/2022)'
$dtstart = (get-date $dt).ToString('yyyyMMddHHmmss')
$dt = Read-Host 'The end date for this batch (eg 4/30/2022)'
$dtend = (get-date $dt).ToString('yyyyMMddHHmmss')
$memo = Read-Host 'The memo field of each transaction  (eg NONE)'
$bal = Read-Host 'The balance before these transactions'
$balance = [decimal]$bal

#this is a stopgap.  The fitid is no longer part of the dcu feed.

$fitid = [int]$dtend.Substring(3,8)  + 1




"    Reading input file checking.csv"
"    Reformatting data fields"

# Create a csv file with just one record for the header

[pscustomobject]@{
    'ACCTID' = $acctid
    'DTSTART' = $dtstart
    'DTEND' = $dtend
}  |
   Export-csv header.csv


# Create the main csv file for the transactions

Import-csv checking.csv  | % {

    $balance +=  [decimal]$(Get-Amount $_.AMOUNT)

    [pscustomobject]@{
          'TRNTYPE' = Get-Type $_.AMOUNT        
          'DTPOSTED' = (Get-Date $_.DATE).Tostring('yyyyMMddHHmmss')
          'TRNAMT' = Get-Amount $_.AMOUNT
          'FITID' = $fitid++        #this is a stopgap
          'NAME' = $_.DESCRIPTION
          'MEMO' = $memo

          }
    } |
    Export-csv transactions.csv


# Create a csv file with just one record for the trailer

[pscustomobject]@{
    'BALAMT' = $balance
    'DTASOF' = $dtend

}  |
   Export-csv trailer.csv



"    Reformatting transactions as XML"

Expand-csv header.csv header.tmplt |
    Out-File header.txt

Expand-csv  transactions.csv  transactions.tmplt |
    Out-File transactions.txt

Expand-csv trailer.csv trailer.tmplt |
    Out-file trailer.txt



"    Creating output file checking.ofx"

Get-Content header.txt, transactions.txt, trailer.txt|
    Out-File checking.ofx


"    Done!"


       