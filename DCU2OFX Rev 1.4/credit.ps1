<#
.NOTES
    Script: credit.ps1           Rev:  1.4
    Author: dcressey             Date: 5/19/22
 
 Known Issues:
    The format of the csv input is from sample, provided by
    a member of the DCU community.

       
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


# File spec set up

"    Select the input file spec"
$input  = gci *.csv | Out-Gridview -OutputMode Single
$output = $input.name -replace '.csv', '.ofx'


$acctid = '99999' + 'L141'
@"
    This script uses $acctid as a dummy account number.
    You can use the features of your PFM to modify this
    number, if desired.

"@



$dtstart = $NULL
$dtend = $NULL
$balance = $NULL


"    Reading input file $input"
"    Generating transactions"

Import-csv  $input  | % {


    $TRNTYPE = ($_.DESCRIPTION -split ' ')[0]       
    $DTPOSTED = (Get-Date $_.DATE).Tostring('yyyyMMddHHmmss')
    $TRNAMT = Get-Amount $_.PRINCIPAL
    if ($_.ID) {$fatid = $_.ID}  else {$fatid = $_.MEMO}
    $FITID = $fatid.GetHashCode().ToString('X')
    $NAME = $_.MEMO.PadRight(25).Substring(0, 25)

    
# The following code depends of the order of transactions being
# Reverse chronological order

    if (!$balance) {$balance = (Get-Amount $_.'BALANCE')}
    if (!$dtend)  {$dtend = $DTPOSTED}
    $dtstart = $DTPOSTED



@"
						<STMTTRN>
							<TRNTYPE>$TRNTYPE
							<DTPOSTED>$DTPOSTED[-5:EST]
							<TRNAMT>$TRNAMT
							<FITID>$FITID
							<NAME>$NAME
						</STMTTRN>


"@} |  Out-file transactions.tmp





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
			<ACCTID>$acctid
		  </CCACCTFROM>
  
        <BANKTRANLIST>
          <DTSTART>$dtstart[-5:EST]
          <DTEND>$dtend[-5:EST]

"@ | Out-File header.tmp







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
"@   | Out-file Trailer.tmp

"    Creating output file $output"

Get-Content header.tmp, transactions.tmp, trailer.tmp|
    Out-File $output -Encoding ASCII


"    Done!"
