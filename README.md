# Export2Dataflow
This repository contains everything needed to run the Export2Dataflow PowerShell script as an External Tool in the Power BI Desktop and export the Power Query to a Dataflow JSON file or publishes it to the Power BI service.

## Disclaimer
Please know, that everything I have created and shared on the blogpost and on GitHub is based on best effort. No rights can be derived, as well as I am not liable for the use or misuse of the solution or possible damage resulting from this. Use of the solutions and execution of the scripts is all on your own risk and your own responsibility.

## Getting Started
### Download everything you need from this repository.
* External Tool integration file: _export2dataflow.pbitool.json_
* External Tool integration file: _publish2dataflow.pbitool.json_
* The PowerShell script: _Export2Dataflow.ps1_

### Copy the External Tool integration file
The External Tool integration file is needed to get the button in the Power BI ribbon. In order to achieve this, you need to copy the _export2dataflow.pbitool.json_ and _publish2dataflow.pbitool.json_ file in the External Tools folder. For me this location was:
_C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools_

While copying the files, it can be that Windows asks you to login with admin privileges before you can continue. This is mandatory to copy the files. If you cannot do this yourself, please contact your administrator.

### Copy the PowerShell script
Create a subfolder _Export2Dataflow_ in the _C:\Program Files\\_ folder and copy the PowerShell script _Export2Dataflow.ps1_ file into it.

Identical to the previous step, Windows may ask you to authenticate with administrator privileges before you can proceed.

### Restart Power BI Desktop
You have applied all required steps by now. The new buttons will appear in the Power BI Desktop top ribbon for External Tools. In case you had Power BI running already, please restart Power BI desktop first.

Did something not workout as expected for you, kindly check the FAQ to see if your question is already listed there. If not, please let me know.

## Usage
I want to shortly describe how this tool works and what you can expect.
1. Open a Power BI Desktop file whose Power Query transformations you want to export to a Dataflow.
2. In the Top Ribbon under External Tools, click _Export to Dataflow_.
3. After the click a PowerShell window opens which converts the Power Query transformations into the Dataflow JSON format. 
During the first execution it may be necessary to agree to the installation of _nuget_ and the _package source_ for the installation of the latest Microsoft.AnalysisServices.Tabular.dll version.
4. Then the script asks for the location where the Dataflow JSON file should be exported.
5. The exported Dataflow JSON file can then be uploaded to the Power BI Portal as Dataflow. To do this, create a new dataflow in the workspace with the Import Datamodel option.

## Author
Marcus Wegener 

[Website](https://thinkbi.de) - 
[twitter](https://twitter.com/PowerBIler) - 
[LinkedIn](https://www.linkedin.com/in/marcuswegener/) - 
[Xing](https://www.xing.com/profile/Marcus_Wegener3/cv)

## Acknowledgements
Marc Lelijveld, for his tutorial on creating external tools for Power BI and his GitHub repository [External-Tools-Model-Documentation](https://github.com/marclelijveld/External-Tools-Model-Documentation), from which I was able to adopt a lot.

[Website](https://data-marc.com/) - 
[Blogpost](https://data-marc.com/2020/07/28/external-tools-document-your-power-bi-model/) - 
[twitter](https://twitter.com/PowerBIler) - 
[LinkedIn](https://www.linkedin.com/in/marclelijveld/) - 
[Github](https://github.com/marclelijveld/External-Tools-Model-Documentation)

Julian Kaiser, for reviewing, testing and adding some features to the script.

[LinkedIn](https://www.linkedin.com/in/julian-kaiser-5b849519a/) 

## FAQ
### The Document model button does not appear in Power BI Desktop
There are a few things that can cause this issue. There are a few things you can check up front:

1. Do you run the latest version of Power BI Desktop? If not, please update first.
2. Do you have other External Tools, such as DAX Studio, Tabular Editor or ALM Toolkit successfully running?
3. Did you put the export2dataflow.pbitool.json file in the correct location? Please check the blogpost to find the correct location.

If still nothing happens, or you don't have the External Tools folder on your PC, I advise you to install any of the above mentioned external tools that will generate this location for you during installation.

### The buttons in the External Tools section in Power BI Desktop are greyed out
External tools require Enhanced Meta Data to be enabled. You can enable this in the preview features of Power BI Desktop.

1. Open Power BI Desktop
2. Go to File
3. Options & Settings
4. Options
5. In the Global Settings go to Preview Features
6. Close all Power BI Desktop instances and re-open Power BI.

### I clicked the button, saw the PowerShell window flickering on the screen, but nothing happened
Most likely, this is caused by the execution policies for PowerShell configured on the computer. As far as I know, this is not something I can change in my script, but has everything to do with the setup of your PC or how your company configured it. This results in the fact that PowerShell.exe application did start, but it did not execute the script, because this was prevented by the policy. The solution to fix this, is changing the execution policy (which is a register thing on your computer. Please know, that this is all at your own risk! I cannot take any responsibility for this.

The following steps might help you

1. Open PowerShell.exe manually
2. Execute Get-ExecutionPolicy -List
3. As a result, you see the current configuration of your execution policies. Most probably everything is set to undefined at this moment.
4. The easiest way to fix this, is by changing the execution policy for the current user. The ensures that this does not happen again. We need to do this by setting it to Unrestricted by executing this task: Set-ExecutionPolicy Unrestricted -Scope CurrentUser (all on your own responsibility and risk!)
5. Confirm that you want to change this.
6. Close PowerShell and Power BI Desktop
Try again if it works now. More information about Execution Policies can be found [in this article](https://winaero.com/change-powershell-execution-policy-windows-10/)

## License
All code in this repository is licenses as specified by the [LICENSE](https://github.com/MarcusWegener/Export2Dataflow/blob/main/LICENSE) file.