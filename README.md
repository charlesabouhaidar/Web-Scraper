# Summary 
This is a program that downloads word files, reads data from them searching for a pattern, then automatically deletes them.
I initially built it to get a visualization of grades done in a test at my university, but can be modified to download any 
docx file from the internet and then read its content searching for a specific pattern. Haven't implemented the option to read
pdf files... yet

# Steps to follow
Make sure you have python3 installed

1. Run this command to install all libraries used

 `pip3 install -r requirements.txt`

2. Replace all `PATH/TO/.../DIRECTORY` by the actual paths you want on your machine

3. In my case, I had an excel sheet that had student IDs which were used for the links to visit. You can upload your own excel sheet
and add it's location in `PATH/TO/EXCEL/FILE/OF/PREVIOUS/TEST/CONTAINING/STUDENT/IDS`(line 16), and/or just skip to the next step

4. Change the link to any link that has .docx file(s) that can be downloaded(lines 74-75)

5. By default, this program searches for a grade that has been posted followed by `/100`, change that to what you'd like (line 57)

6. If you don't want to automatically delete the files downloaded, just remove the last 2 lines of the program

7. Be sure you are in the correct directory and run the program with this command:

`python3 scrape.py`
