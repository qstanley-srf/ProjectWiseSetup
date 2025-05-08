# ProjectWiseSetup
This project contains all of the cmdlets used for creating a ProjectWise Project.

***

To make the cmdlets available for use:
* Copy the PWConnectedProjectCmdlets folder
* Go to this location: C:\Program Files\WindowsPowerShell\Modules
* Paste the folder in the above location
* Use the following command (already included in the setup script)

> Import-Module PWConnectedProjectCmdlets

***

Once the steps above are completed, all of the inclueded cmdlets will be available for use.
These cmdlets are requred to run the project setup script for ProjectWise, but they can also be used individually if some part fails to run.

A list of commands is available either in the Commands panel on the right of the Powershill or by running the command:

> Get-Command -Module PWConnectedProjectCmdlets

To get a list of required parameters, run the command:

> Get-Help {command name}

***

If you run into errors or have questions, please reachout to <qstanley@srfconsulting.com>