# Powershell
Welcome to the PowerShell GitHub Community! PowerShell Core is a cross-platform (Windows, Linux, and macOS) automation and configuration tool/framework that works well with your existing tools and is optimized for dealing with structured data (e.g. JSON, CSV, XML, etc.), REST APIs, and object models. It includes a command-line shell and scripts developed in powershell to automate different tasks (Windows Administration, File System Management, Excel Automation etc)

Version: Powershell 2.0
-----------------------------

Script Name - Automation of Excel using Powershell
--------------------------------------------------

Background:
-----------

There are number of sources from which Incident data is from different sources (Remedy, ServiceNow, Jira etc). The management wanted to look for a common dashboard presenting the current status of incidents SLA across multiple accounts. 

I automated the entire process of consolidation of data coming in excel sheet, combining these multiple sheets into common dashboard. 


Script Details:
--------------

This script is developed in powershell version 4.1 and reads the data from excel sheet across different files. These files are merged into one to generate a common dashboard. 


Function Provided in the scripts as follows:
--------------------------------------------

1. Read data from different files (CSV and Excel)
2. Create Data-Table and perform different opterations using .Net DataTable functions (Select, compute etc)
3. Merge different files
4. Perform Sorting, Auto-filtering, Computing functions

You can download the excel files and powershell script and run in your env.
