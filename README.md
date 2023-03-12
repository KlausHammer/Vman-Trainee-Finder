Vman-Trainee-Finder
======
A python script to search though all avaliable trainers in the webgame Vman and save them as an excel sheet


Things to fix:
* Some links are broken (also it does not take the full name this way if more then 2 names are present)
* Cannot always read the age and wage (might have something to do with the 2 names limit)
* Very slow (7 seconds per page)


How to run
------
Install followin python libraries:
* Selenium
* ChromeDriverManager
* openpyxl  

In line 26 change "knham" to your windows user  
In line 27 set the chrome profile to the profile that is logged into Vman (might be "Profile 0" if its the default chrome profile)  
In line 83 to 86 the trainers split into age group, run the script once for each group and make sure to empty the excel sheet between each run as it can corrupt  because of its size.
