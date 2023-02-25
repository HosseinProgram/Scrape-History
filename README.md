# Scrape-History
Scrapping History of tsetmc

This program makes a list of best offers of tsetmc boards of all symbols for active dates in multi Threads fast and furious!
you can modify Config.json file like this:

{"Symbols":["غگیلا","فولاد","فملی"],
"startdate" : [1394,1,1],
"enddate" : [1400,12,29],
"From" : [9,0,0],
"To" : [12,30,0],
"StepTime": 600
}

"Symbols" is list of desierd symbols ; 

"startdate" is first desiered date;

"enddate" is last desiered date;

"From" is start time of day;

"To" is end time of day;

and "Steptime" is distance between steps of scraping board

Project.py is source of that and Execuatable/project.exe is executable version

program returns xlsx file per each symbols like this:

![Capture](https://user-images.githubusercontent.com/104124540/221358136-19fb2865-d035-430e-938f-b66f8ae15ac6.JPG)

lets' Enjoy!
