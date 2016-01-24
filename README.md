# Automated Software Deployment Test Suite

Automated deployment test suite for LANDesk software deployments.

This file was created during extra time while I was employed at the University of Ontario Institute of Technology. The university uses LANDesk to deploy software titles specific to each program. For example, you may need MATLAB if you are in engineering. You would log into the system, select MATLAB and click download. Then a script created by myself would run, and perform a custom installation of MATLAB. During the course of my employment, many architectural updates were made to the underlying systems leading to broken scripts and other problems. The current solution was to wait for support tickets from users, or manually deploy software and check where the break occured and why. This file is a solution to this problem.

## What it does

The program is built to run software deployments and monitor the computer to determine if the installation occured. The program begins by reading configuration variables from a file. Then it queries a database for a list of titles available to the user specified in the config file. The program then starts a virtual machine, either on the local computer or a specified datacenter. The program logs the user in and runs a series of tests. As the deployment happens, the program takes screenshots of progress which are later converted into a gif file. After all the tests are completed, the results are inserted into a database. My supervisor Zander Kidd then wrote a front end graphical view which takes the data from the tests and displays it for easy monitoring. 

## Interesting Details

There were several interesting and educational challenges that had to be overcome in this project. First, to automate the deployments a GUI macro could not be used as the application did not have any buttons detectable by the AutoIT3 system. Through reverse engineering and tinkering, it was determined that the deployments were being performed by a program which had an xml file in the appdata folder on it's command line. So to automate this, we simply queried the folder and got all the xml files, and fed them to the program one at a time.

As we discovered more about how the program worked, we realized that the program must be called by the SYSTEM user. The VMWare VIX API does not allow calls by the system user, and thus a workaround was created. To achieve SYSTEM priviliges, the Hstart application was called with the ```/UAC``` flag to run with elevated rights. The program specified to run in elevated (admin) rights was PsExec, which has the ability to run programs as SYSTEM. This hacky method of running programs was made into a global method to be called with ease.

### Credits

Created by Connor Sheehan, Jamie Turner and Zander Kidd
