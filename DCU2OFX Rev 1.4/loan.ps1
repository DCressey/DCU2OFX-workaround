<#
.NOTES
    Script: loan.ps1             Rev:  1.4
    Author: dcressey             Date: 5/10/22
    
.SYNOPSIS
    Handles loan  transactions for DCU2OFX
    ACCTTYPE rejected by OFXAnalyzer
           
#>

# This function gets the amount, eliminating currency symbol and commas.
# Note that it also reverses the sign on the amounts.

function Get-Amount ($amount) {
$abs  = $amount -replace '[^0-9.]'
If ($amount -match '-|\(') { "$abs"} else {-$abs}
}


# This function distinguishes between credits and debits
# Note that payments are negative numbers

function Get-Type ($amount) {
if ($amount -match  '-|\(') {'CREDIT'} else {'DEBIT'}
}

# The main script begins here.

$acctid = '99999' + 'L141'
@"
    This script uses $acctid as a dummy account number.
    You can use the features of your PFM to modify this
    number, if desired.

"@

# File spec set up

"    Select the input file spec"
$input  = gci *.csv | Out-Gridview -OutputMode Single
$output = $input.name -replace '.csv', '.ofx'








$dtstart = $null
$dtend = $null
$balance = $null

"    Reading input file $input"
"    Generating transactions"

Import-csv  $input  | % {


    $TRNTYPE = Get-Type $_.AMOUNT        
    $DTPOSTED = (Get-Date $_.DATE).Tostring('yyyyMMddHHmmss')
    $TRNAMT = Get-Amount $_.AMOUNT
    if ($_.ID) {$fatid = $_.ID}  else {$fatid = $_.DESCRIPTION}
    $FITID = $fatid.GetHashCode().ToString('X')
    $NAME = $_.DESCRIPTION.PadRight(25).SubString(0, 25)

# The following code depends of the order of transactions being
# Reverse chronological order

    if (!$balance) {$balance = (Get-Amount $_.'CURRENT BALANCE')}
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


							<ACCTTYPE>CREDITLINE
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
