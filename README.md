Practical Automation Task

Preparation



Copy the entire project folder to the following path:

C:\\praktinis assignment\\...

The folder structure must match this location exactly.



Ensure that your computer has Python 3.13.7 installed.

The project may also work on slightly older or newer Python versions, but 3.13.7 is recommended.



Power Automate Login

Exporting or sharing flows is not allowed on personal Power Automate accounts.

School or organization accounts allow exporting only with a Premium license.

To run the flows, log in to Power Automate using the following credentials:



Email: anonim46547@gmail.com

Password: rpauzduotis



Additional Information for Login:

MS password: rpauzduotis1



If logging into the Power Automate Desktop app (available on Microsoft Store) is unsuccessful, there is a folder named "Power Automate" inside the project directory.

This folder contains all flow definitions. You can manually copy and paste the code into new subflows in Power Automate Desktop.

Name each subflow according to the folder names provided.

Once the project folder is placed correctly and the Power Automate account is logged in, run the Parabank RPA flow.



Overview of Flow Logic:

Main → clean\_folders



Subflow: clean\_folders

Terminates Notepad and Microsoft Excel processes.

Deletes old report files from:

C:\\praktinis assignment\\report



Main → read\_data



Subflow: read\_data

Supports handling multiple data files.

Reads the file and assigns required values into a data table.

After processing the file, the next step is:



read\_data → register\_to\_parabank



During registration:

The system registers the client; if the account already exists, it attempts to log in.

After logging into the client account, the next step is:



register\_to\_parabank → loan\_submit

If no accounts were registered, the sequence becomes:

register\_to\_parabank → customer\_report → next client



Subflow: loan\_submit

Opens a new bank account for the client.

Sends a loan request.

Moves on to the next client.

