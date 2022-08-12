# excel-tutoring-center-scheduling-automation
###### An automated Microsoft Excel appointment scheduling system for tutoring centers

### How it works
An incoming appointment file for the week is brought in as a CSV. The script empties 
the previous data out of the **in tray** and into **long term records**, and inserts the new data into the **in tray**. The
new data is then compared to the **master list**, which holds the grades of all the students,
and students are grouped into groups of three or less for their appointments based
on their grade and appointment date/time.  

### Pages in the workbook
- In Tray: This page holds all of the student's individual appointment data (before grouping).
- Master List: This page holds each student's name and grade.
- Appointments: This page holds group appointment data.
- Long Term Records: This page holds the idividual student appointment data that is emptied out of the In Tray.
