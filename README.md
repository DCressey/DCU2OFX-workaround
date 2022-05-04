# DCU2OFX-workaround came about when DCU  (Digital Federal Credit Union) launched a new platform in March.  Some existing depositor/members found they could not do things the old way anymore, and the new way wasn't working for them.  So I wrote a little script that reads CSV files with the data DCU provides in downloads, and produces OFX files that resemble the ones DCU used to make available.  The script actually consists of five scripts.  One is a master script that invokes one of the other scripts or displays user instructions.  The other four process a particular kind of account: checking, savings, loan, or credit card.  The credit card one is just a stub that indicates that the credit card processing isn't available yet. 

It's built for a windows environment, and it's written in Powershell.  Powershell isn't better than other scripting languiages.  It just happens to be what I learned.  

This tool is expected to have a short lifespan.  As DCU fills in the gaps in thier new offering,  the need for this toll will go away.  

