<#
.NOTES
    Script:  dcu2ofx                  Rev:  1.4
    Author:  dcressey                 Date: 5/18/22

 Known Issues:
     negative numbers with parentheses haven't been tested.
     fitids are untested hashcodes
     Credit untested
     Questionable sign on Credit and Loan transactions.
 

.SYNOPSIS
    Converts data as delivered by DCU into format for use with a PFM.
.DESCRIPTION
    This script reads a file named checking.csv that contains transaction data
    from a checking account at the DCU, and downloaded by the DCU customer interface.
    It transforms the data from CSV format to OFX format,  a format accepted by many
    Personal Finance Manager packages, such as MoneyDance or YNAB. 
    and several other PFM packages.

    More documentation is available in readme.rtf .

#>






# the main script starts here

#Welcome

@'


Welcome to the DCU to OFX data converter!

This script will read the transactions the DCU provides in a csv file and produce
an ofx file that some Personal Finance Manager packages can process.  

'@

"    Select one of the available functions"
$menu = @{
    CHECKING = 'Process Checking Account transactions'
    SAVINGS  = 'Process Savings Account transactions'
    LOAN     = 'Process Loan transactions'
    CREDIT   = 'Process Credit Card TransactionsS'
    GUIDE    = 'Display the User Guide'
    }
Switch (($menu | Out-GridView -OutputMode Single).name) {
    CHECKING {. .\checking.ps1}
    SAVINGS  {. .\savings.ps1}
    LOAN     {. .\loan.ps1}
    CREDIT   {. .\credit.ps1}
    GUIDE    {start readme.rtf}
    DEFAULT  {'    Sorry, that is not one of the choices'}
    }

