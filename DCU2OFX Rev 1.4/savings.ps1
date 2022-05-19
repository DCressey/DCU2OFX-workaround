<#
.NOTES
    Script: savings.ps1         Rev:  1.4
    Author: dcressey             Date: 5/8/22
    
.SYNOPSIS
    Handles savings transactions for DCU2OFX
           
#>

# This function gets the amount, eliminating currency symbol and commas.

function Get-Amount ($amount) {
$abs  = $amount -replace '[^0-9.]'
If ($amount -match '-|\(') { "-$abs"} else {$abs}
}


# This function distinguishes between credits and debits
# Note that Withdrawals are negative numbers

function Get-Type ($amount) {
if ($amount -match  '-|\(') {'DEBIT'} else {'CREDIT'}
}

# File spec set up
"    Processing Savings file"
"    Select the input file spec"
$input  = gci *.csv | Out-Gridview -OutputMode Single
$output = $input.name -replace '.csv', '.ofx'

$acctid = '99999' + 'S1'
@"
    This script uses $acctid as a dummy account number.
    You can use the features of your PFM to modify this
    number, if desired.

"@

$dtstart = $null
$dtend = $null
$balance = $null

"    Reading input file $input"
"    Generating transactions"

Import-csv  $input  | % {

 
    $TRNTYPE = Get-Type $_.AMOUNT        
    $DTPOSTED = (Get-Date $_.DATE).Tostring('yyyyMMddHHmmss')
    $TRNAMT = Get-Amount $_.AMOUNT
    if ($_.ID)  {$fatid = $_.ID}  else {$fatid = $_.DESCRIPTION}
    $FITID = $fatid.GetHashcode().Tostring('X')
    $NAME = $_.DESCRIPTION.PadRight(25).Substring(0, 25)

 
 # The following code depends on the order of transactions
 # being reverse chronological

    if (!$balance) {$balance = (Get-Amount $_.'CURRENT BALANCE')}
    if (!$dtend)   {$dtend = $DTPOSTED}
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
			<DTSERVER>20220304092032[-5:EST]
			<LANGUAGE>ENG
		</SONRS>
	</SIGNONMSGSRSV1>
	 <BANKMSGSRSV1>
		<STMTTRNRS>
			<TRNUID>0
			<STATUS>
				<CODE>0
				<SEVERITY>INFO
			</STATUS>

			<STMTRS>
				<CURDEF>USD

					<BANKACCTFROM>
							<BANKID>211391825
							<ACCTID>$ACCTID
							<ACCTTYPE>SAVINGS
					</BANKACCTFROM>
	
				<BANKTRANLIST>
					<DTSTART>$DTSTART[-5:EST]
					<DTEND>$DTEND[-5:EST]


"@ | Out-File header.tmp






"    Generating Trailer"

    $balamt = $balance
    $dtasof = $dtend

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

			</STMTRS>
		</STMTTRNRS>
	</BANKMSGSRSV1>

</OFX>
"@   | Out-file Trailer.tmp

"    Creating output file $output"
Get-Content header.tmp, transactions.tmp, trailer.tmp|
    Out-File $output -Encoding ASCII


"    Done!"
