# excel-tutoring-center-scheduling-automation
###### An automated Microsoft Excel appointment scheduling system for tutoring centers

### How it works
An incoming appointment file for the week is brought in as a CSV. The script empties 
the previous data out of the in tray, and inserts the new data into the in tray. The
new data is compared to the master list, which holds the grades of all the students,
and students are grouped into groups of three or less for their appointments based
on their grade and appointment date/time.  
