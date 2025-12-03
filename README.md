1\. Preparation
Project Folder Location



Copy the entire project folder to the following path:
C:\\praktinis assignment\\  (must match this: "C\praktinis assignment" (no underscore or minus signs))
The folder structure must match this location exactly for the automation logic to function correctly.



Python Requirement
Ensure that your computer has Python 3.13.7 installed.
The solution may also work on slightly older or newer Python versions, but 3.13.7 is recommended for compatibility.



2\. Power Automate Login
Exporting or sharing flows is not supported on personal Power Automate accounts.
Organization or school accounts allow exporting only with a Premium license.


To run the flows included in this task, log in to Power Automate using the following credentials:
Email: anonim46547@gmail.com
Password: rpauzduotis
{Additional Microsoft account password (if prompted): 
rpauzduotis1}


If Power Automate Desktop Login Fails
Inside the project directory, there is a folder named:
Power Automate
This folder contains all flow definitions in text form.
If login is unsuccessful, manually recreate the flows:
Create new flows/subflows in Power Automate Desktop.
Copy the corresponding code from the folder.
Name each subflow according to the folder names.
Once the project folder is placed correctly and the Power Automate account is logged in, run the Parabank RPA flow.



3\. Gross overview of Flow Logic (not too detailed):

Main → clean\_folders
Subflow: clean\_folders
Terminates Notepad and Microsoft Excel processes.
Deletes old report files from:
C:\\praktinis assignment\\report



Main → read\_data
Subflow: read\_data
Supports processing multiple data files.
Reads each file and assigns required values into a data table.
After processing, the sequence continues:
read\_data → register\_to\_parabank



Registration Logic
During the registration sequence:
The system attempts to register a new client.
If the account already exists, the system logs in.
After login, the next step is:
register\_to\_parabank → loan\_submit
If no accounts were registered in this iteration:
register\_to\_parabank → customer\_report → next client



Subflow: loan\_submit
Opens a bank account for the client.
Submits a loan request.
Moves to the next client.

