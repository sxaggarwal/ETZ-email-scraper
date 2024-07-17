### Scraper

## FLOW
maybe add to the email app

HOW: Scrapes emails - all at once when ran. 
#### [ ] install make
#### [ ] fix attribute error in pywin32 when accessing certain emails. Dont know for which emails this is happening

## MODULES
[x] yaml file to save the time of the last scrape
    - OPTIONAL can also save the excel file metadata if we do not want to repeat
[x] excel model - Pydantic (to read and write data from) - to ensure that data is correct in terms of the type - CLASS
[ ] script to get the data - CLASS
    [ ] get emails to "lookout" for
    [ ] save excel sheet
[ ] write data to MT  (from pydantic)

## OPTIONAL
[ ] separate outlook event handler to listen for emails that are incoming
[ ] attach the event handler to the scraper class and run as daemon

## DEPLOYMENT
[ ] How to make a script a daemon and keep it running in the background
[ ] adding background processes for windows 
    - so the user does not have to activate it everytime the computer restarts

## TESTING
[ ] integration test
[ ] pydantic test - part of unit
