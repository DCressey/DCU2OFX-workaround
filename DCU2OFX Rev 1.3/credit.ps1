<#
.NOTES
    Script: credit.ps1           Rev:  1.3
    Author: dcressey             Date: 5/8/22
 
 Known Issues:
    The format of the csv input is only a conjecture.
       
.SYNOPSIS
    Handles credit card transactions for DCU2OFX
           
#>

# This function gets the amount, eliminating currency symbol and commas.
# Note that it also reverses the sign on the amount.

function Get-Amount ($amount) {
$abs  = $amount -replace '[^0-9.]'
If ($amount -match '-|\(') { "$abs"} else {-$abs}
}


# This function distinguishes between credits and debits
# Note that payments are negative numbers

function Get-Type ($amount) {
if ($amount -match  '-|\(') {'CREDIT'} else {'DEBIT'}
}

# The Main script begins here

@'

This script is not yet ready for Prime time.  It has not been tested with 
actual inputs from the new DCU platform.
    In the meantime, you could try running credit card transactions through
the LOAN function, and see if the output file is useable.  This has not been
tested either, but it might work.
'@
exit

# File spec set up

"    Select the input file spec"
$input  = gci *.csv | Out-Gridview -OutputMode Single
$output = $input.name -replace '.csv', '.ofx'




# User dialog to fill in missing data

if (Test-Path  dcu2ofx-config.ps1) {"    Using existing configuration file"}
    else {
    "Generating standard configuration file" 
@'
# These parameters are needed in the .ofx file, but are not supplied
# in the .csv file.

$acctid      = '12345'          #Your Account number to be used by your PFM
$startDate =  '4/1/2022'     #The start date for this batch of transactions
$endDate   = '4/30/2022'      #The end date for this batch of transactions
$memo       = ''                 #The memo field of each transaction
$balance    = '1000.0'        #The balance before these transactions

'@   | Out-File dcu2ofx-config.ps1 
    }

"    Edit the configuration file"

Start-Process -Wait -Filepath notepad -ArgumentList dcu2ofx-config.ps1


. ./dcu2ofx-config.ps1




$dtstart = (get-date $startDate).ToString('yyyyMMddHHmmss')
$dtend = (get-date $endDate).ToString('yyyyMMddHHmmss')
$balance = [decimal]$balance






"    Generating header"


@"
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
  <SIGNONMSGSRSV1>
    <SONRS>
      <STATUS>
        <CODE>0
        <SEVERITY>INFO
      </STATUS>
      <DTSERVER>20220223165700[-5:EST]
      <LANGUAGE>ENG
    </SONRS>
  </SIGNONMSGSRSV1>
   <CREDITCARDMSGSRSV1>
    <CCSTMTTRNRS>
      <TRNUID>0
      <STATUS>
        <CODE>0
        <SEVERITY>INFO
      </STATUS>

      <CCSTMTRS>
        <CURDEF>USD

          <CCACCTFROM>
			  <BANKID>211391825
              <ACCTID>$acctidL141
              <ACCTTYP>CREDITLINE
          </CCACCTFROM>
  
        <BANKTRANLIST>
          <DTSTART>$dtstart[-5:EST]
          <DTEND>$dtend[-5:EST]

"@ | Out-File header.txt




"    Reading input file $input"
"    Generating transactions"

Import-csv  $input  | % {

    $balance +=  [decimal]$(Get-Amount $_.AMOUNT)

    $TRNTYPE = Get-Type $_.AMOUNT        
    $DTPOSTED = (Get-Date $_.DATE).Tostring('yyyyMMddHHmmss')
    $TRNAMT = Get-Amount $_.AMOUNT
    $FITID = $_.ID.GetHash().ToString('X')
    $NAME = $_.DESCRIPTION.PadRight(25).Substring(0, 25)


@"
						<STMTTRN>
							<TRNTYPE>$TRNTYPE
							<DTPOSTED>$DTPOSTED[-5:EST]
							<TRNAMT>$TRNAMT
							<FITID>$FITID
							<NAME>$NAME
						</STMTTRN>


"@} |  Out-file transactions.txt




"    Generating Trailer"

    $balamt = $balance
    $dtasof= $dtend

@"
       </BANKTRANLIST>
       <LEDGERBAL>
          <BALAMT>$balamt
          <DTASOF>$dtasof[-5:EST]
        </LEDGERBAL>

        <AVAILBAL>
          <BALAMT>$balamt
          <DTASOF>$dtasof[-5:EST]
        </AVAILBAL>

      </CCSTMTRS>
    </CCSTMTTRNRS>
  </CREDITCARDMSGSRSV1>

</OFX>
"@   | Out-file Trailer.txt

"    Creating output file $output"

Get-Content header.txt, transactions.txt, trailer.txt|
    Out-File $output -Encoding ASCII


"    Done!"
