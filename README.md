# design.automation-extract-fabpart
Extract the data of FabricationPart of Revit

## Demo Video:
https://www.youtube.com/watch?v=c8Szl7Jtrlk

# Steps
1.	Follow [Readme of LearnForge](https://github.com/Autodesk-Forge/learn.forge.designautomation/tree/master/forgesample) to setup the environment 
2.	Build Revit plugin (Revit 2019), ensure the zip of the bundle is output to \forgesample\wwwroot\bundles. The post build event uses 7zip. 
3.	Build web app
4.	Start ngrok
5.	Input ngrok link to properties of web app
6.	Launch web app
7.	Click [Configuration] 
8.	Select the bundle, select the enginer (Revit 2019), click [Create/Update]
9.	An appbundle and activity will be created.
10.	go to web page, ignore [Width] and [Height]. 
11.	Upload your soure Revit file, which contains the Fabrification Parts.
12.	Click [Start WorkItem]. Wait a moment
13.	It will output the csv file, and upload to Forge bucket, and send web page with the download link of the csv

Note: do not define the output file as *.txt or *.html if you want to use this LearnForge tutorial because it upload to Forge bucket, but bucket does not accept *.txt or *.html now. 

