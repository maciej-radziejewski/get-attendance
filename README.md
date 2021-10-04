# get-attendance

This is a simple Python script for teachers using Microsoft Teams, to get readable attendance lists in tabular form. You need a recent Python interpreter to use it. I wrote this script for myself, and I am sharing it so others can possibly benefit from it.

Teams (e.g., version 1.4.00.7175) lets you download:
* an attendance list, during the meeting that you organized (click on the **...** icon on the list of participants, and then download),
* an attendance report, after the meeting (from the meeting chat, which is also available in the chanel associated with the meeting; it looks like attendance reports are only available for scheduled meetings).

What this script can do for you is:
* process a bunch of downloaded lists and reports in one go,
* automatically match the data therin to various classes in your weekly schedule,
* lay this data out nicely in separate CSV files (viewable in any spreadsheet app), one sheet per class.

Suppose you teach Linear algebra (two groups of students) and Topology (one group). Your schedule might looke like this (class start times are given):

    Wednesday, 9:00, Linear algebra group 1
    Wednesday, 11:00, Linear algebra group 2
    Thursday, 9:00, Topology
    Thursday, 11:00, Topology
    Friday, 13:00, Linear algebra group 1
    Friday, 15:00, Linear algebra group 2

You would get three CSV files: **Linear algebra group 1.csv**, **Linear algebra group 2.csv** and **Topology.csv**. The file **Linear algebra group 1.csv** might look like:

    Full name,2021-03-03,2021-03-05,2021-03-12
    Eleanor Someone,1,1,1
    Joan Someoneelse,0,1,0
    Daniel Whoever,0,0,1
    
because the class is on Wednesdays and Fridays, but it did not take place on March 10th. Here 0 means 'absent' and 1 means 'present'. Since the Topology group has two meetings on one day, dates alone would not identify the meetings. Therefore column headers would be supplemented by start times:

    Full name,2021-03-04 09:00,2021-03-04 11:00,2021-03-11 09:00,2021-03-11 11:00
    John Smith,1,1,1,1
    Adam Smith,1,1,1,1
    Jane Unnamed,1,1,1,1

There is also a simplification taking place here: every student present at the meeting counts as present, whether they were present for 1 minute with a muted microphone, or for 90 minutes with the camera on. The reports/lists from Teams contain more information, i.e. who joined when, but the will not tell you had their microphone or camera on.

## Installation and requirements

You need a Python 3 interpreter. The script was tested with Python 3.9, but 3.8 should be fine, too. You might already have Python in your system. You can install it from https://www.python.org or a more versatile system from https://docs.conda.io/en/latest/miniconda.html or any other Python you like. Try:

    python --version

or

    python3 --version

from the Terminal (i.e. command line or command prompt). Make a note of how you invoke Python (i.e. with python3 or python).

Once you have Python, download the script 'attendance.py' from here and place it in a **new, empty directory**, where you are going to store attendance data. Try running it by opening the Terminal, **going to the script directory** and entering the command:

    python attendance.py

or

    python3 attendance.py

whichever works for you. You should always run the script like this: in the current directory. Any file it generates will also go to the same direcory.

You probably got some error message, but we will deal with that in a moment. If the test-run of the script **created some files** in its directory, please delete them for now.

## Setup

#### 1. Get as many attendance lists and reports as you can.

You can download an attendance list during the meeting that you organized (click on the **...** icon on the list of participants, and then download). If you did not, you can probably still access the meeting chat (e.g., go to the chanel associated with the meeting or re-open a scheduled meeting from the calendar). The meeting chat should contain the downloadable attendance report.

All these lists and reports should go to your Downloads folder. They have names like 'meetingAttendanceList (9).csv' or 'meetingAttendanceReport(General) (4).csv'.
Every attendance list should contain 'meetingAttendanceList' in the file name. Every attendance report should contain 'meetingAttendanceReport' in the file name.

#### 2. To make sure that the script finds your Downloads folder, you may need to edit the script in a text editor (Notepad in Windows).

If you open the script in your editor, you should see a line

    dir_downloads = os.path.expanduser('~') + '/Downloads'

This should work on Unix and macOs systems, so, if you run the script, it should find your downloaded attendance lists and reports, and process them (see below). Otherwise you will get an error message telling you to set the path to your Downloads directory. If your Downloads directory (or whatever directory Teams is saving the attendance lists to) is, say, 'C:/some/path', you can replace the above line with:

    dir_downloads = 'C:/some/path'

#### 3. If your system language is other than Polish, you will need to modify the column headers that the script expects to find in your lists/reports.

This is a sort of double check to make sure that we read correct data. If you get an error like:

    Column header ... expected and ... encountered.

you need to change some lines in the script file by putting the actual column headers between the quotes. Specifically, you have to change the lines:

    column_header_full_name = 'Imię i nazwisko' # In attendance lists and reports
    column_header_time_mark = 'Znacznik czasu' # In attendance lists
    column_header_join_time = 'Godzina dołączenia' # In attendance reports
    column_header_leave_time = 'Godzina opuszczenia' # In attendance reports
    column_header_role = 'Rola' # In attendance reports

Please, open the attendance lists/reports with a text editor to see exactly, what headers were used by your version of Teams. The other settings near those are 0-based column numbers where the data can be found. You should not have to change those, unless Teams changes its behaviour.

Attendance reports should identify you explicitly as the meeting organizer. For example, in Polish you would be called "Organizator". As a final step to adapt the script to your language, please change the line

    organizer_role_in_reports = 'Organizator'

to what Teams used in your files. This way your name will be correctly skipped from the generated files.

#### 4. Create the schedule file.

If the test run was not succesful, re-run the script until it finishes without errors. It should create a provisional schedule file for you, named **schedule.csv** (always in the current directory, whicj should be the same directory as the script). You should edit it with a text editor. It may look like this:

    Wednesday, 9:00, class1
    Wednesday, 11:00, class2
    Thursday, 9:00, class3
    Thursday, 11:00, class4

assuming your actual schedule is like the one at the top, and you had attendance lists/reports for some of your classes. You need to:

* give descriptive names to the groups, reflecting which items in the schedule refer to distinct groups, and
* add the missing items.

You should get something like the schedule at the top:

    Wednesday, 9:00, Linear algebra group 1
    Wednesday, 11:00, Linear algebra group 2
    Thursday, 9:00, Topology
    Thursday, 11:00, Topology
    Friday, 13:00, Linear algebra group 1
    Friday, 15:00, Linear algebra group 2

The last column should contain readable names of your classes, uniquely identifying each group that you teach. These names are also used as **file names for the sheets**. Say, you have an attendance list/report in your Downloads directory from a meeting that started on Wednesday at 8:55. The meeting start time will be matched to the closest one in your schedule. Accordingly, the attendance in the list/report will go to the file **Linear algebra group 1.csv** in the current directory. 

Note that:
* the schedule file is a simple text file with comma-separated values,
* do not change the default text encoding of the file,
* the same group can have multiple meetings on the same day or different days,
* you need to use the 24-hour format for times,
* the times in the provisional file are reasonable defaults, but you may need to correct them.
* weekday names should be in the system default language, see below,

#### 5. Run the script again.

It should process all the lists/reports that it can find in the Downloads directory, see the next section.

## Everyday use

Once the setup is complete, you can just run the script whenever you download new attendance lists/reports.

With the **schedule.csv** in the current directory the script should process all the lists/reports that it can find in the Downloads directory. It will tell you which ones. Watch out for any error reports. Attendance records in these files will be merged with those already on the sheets.

Each meeting should be matched to the closest entry in the schedule, on the basis of the meeting start time. The maximum allowed distance is the greater of: 30 minutes, and the actual meeting length. Unmatched meetings are reported. You may need to modify the schedule in such a case (e.g., add a new possible meeting time for some group).

## When the semester is over

When the term is over and the new one starts, you can create a new, empty directory for your new attendance sheets, and move the script file with your settings. You will have to re-do the steps 1, 4 and 5 to get the new schedule.

## What if your classes are irregular?

The script will work with a semi-regular schedule, because you can enter multiple day/time slots for each class. However, if these are too irregular, or you happen to have different classes at the same slot, you need a different solution. The solution is to change the line

    irregular_classes = False

to

    irregular_classes = True

in the script and then:
* prepare separate directories for separate classes,
* copy the script file to each of them and do not trouble yourself with the schedule,
* run the script from the appropriate directory after the class, and then remove the downloaded attendance list/report (or move it to some per class archive).

The 'irregular_classes' settings makes all the attendance data go to one file 'attendance.csv', so you need to make sure yourself not to mix-up the lists for different classes.
