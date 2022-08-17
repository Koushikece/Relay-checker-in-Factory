# Relay-checker-in-Factory
This app provides a GUI interface to users to track of relay devices. Any employee can easily take relays from Siemens Goa Factory for testing purpose. This app also provides an interface to submit their name, email, phone number, expected return date of the device . Except this, realys can be returned and more devices can be added using this desktop app.
This is completely built using python language and tkinter library. 

## Authors
Kanchan Kumar Kaity   and   Koushik Dey
- [Kanchan Kumar Kaity](https://github.com/Kanchan1396)
- [Koushik Dey](https://github.com/Koushikece)
## Installation
Just download these three files named as 'relaychecker.py', 'schedule.py' , 'tkinterfinals.xlsx' and keep them in a single folder. Make sure your internet connection is ON . Just open the 'relaychecker.exe' file . Here you can perform these following actions --- 
'Take', 'Return', 'Add new items', and check availability of the relay devices at a glance .

Then , open task scheduler from search bar. After this----
1. Click on 'Create a Basic task' 
2. Provide a name and click on Next
3. Select 'daily' then press 'Next'
4. Adjust the time of your own wish. Try to give between office hours. 
5. Select 'start a program' then Next
6. Browse the .py file , named as 'schedule' from your directory
7. Copy the original folder location of the files as 'Add arguments'

Now Relay Checker is ready. 
Actually , in the given time 'schedule.py' file will be executed automatically and will check the expected return date of every device . If expected return date is that present day , it will send mail to the customer.

    
## Documentation

[Tkinter](https://docs.python.org/3.9/library/tkinter.html)
[OpenPyXl](https://openpyxl.readthedocs.io/en/stable/)
