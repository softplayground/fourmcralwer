As the name implies, this is a crawler for a fourm. 

To run it, you should have windowws/mac/linux operation system, (not tested except windows), download .net8.0 sdk, and run the command below to compile the app:

dotnet publish -r win-x64 -c Release --self-contained /p:PublishSingleFile=true

Once done, you should go to the publish folder to get the two compiled .exe files. 

Run the app and select the area and page, it runs a while and will generate the excel file on the desktop folder.
