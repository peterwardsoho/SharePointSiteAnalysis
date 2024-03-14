SharePoint Online Site Analysis Script Documentation 

Purpose 

This PowerShell script is tailored for SharePoint Online administrators to aid in the analysis of SharePoint sites. It facilitates three core analyses: 

1.	Identifying stale sites by their last content modification date. 

2.	Identifying stale site owners by their last sign-in activity. 

4.	Fetching classic sites within specified site title ranges. 

These functionalities support administrators in efficiently managing SharePoint Online environments by providing insights into site usage and configuration. 

Prerequisites 

•	PowerShell version 5.1 or later. 

•	SharePoint Online Management Shell. 

•	PnP PowerShell module. 

•	Microsoft Graph PowerShell SDK. 

Usage 

1.	Global Admin Prompt:  Asks if user is global admin of the tenant, they want to run scripts for. 

2.	Tenant Name Prompt: Asks for the Azure Tenant Name that user is global admin of. 

3.	Script Option Prompt: The script prompts the user to choose from one of three analysis options: 

•	(A) Stale Sites 

•	(B) Stale Owners 

•	(C) Fetch Classic Sites 

 

4. 	The detailed step after user selects one of the scripts is given below:  
 

 

Option A: Stale Sites Analysis 

1.	User Authentication: The script authenticates the user and connects to SharePoint Online. 

2.	Time Period Selection: Users specify the analysis time (e.g., Quarterly, Semi-Annual, Annual, or Custom). 

3.	Batch Processing: Allows for processing sites in batches according to their titles (A-J, K-S, T-Z, or All). 

4.	Site Fetching and Analysis: The script retrieves all site collections and filters them based on the last content modification date and the specified time. 

5.	Data Exporting: Users have the option to export the analysis results to a CSV file. 

 

Option B: Stale Owners Analysis 

1.	Service Connection: Establishes connections to SharePoint Online and Microsoft Graph. 

2.	Time Period Selection: Users specify the analysis time (e.g., Quarterly, Semi-Annual, Annual, or Custom). 

3.	Batch Processing: Allows for processing sites in batches according to their titles (A-J, K-S, T-Z, or All). 

4.	Site Owner Fetching: Retrieves site collections and identifies the owners, checking their last sign-in activity via Microsoft Graph. 

5.	Filtering and Exporting: Filters sites based on owner activity and offers CSV export functionality. 

 

Option C: Fetch Classic Sites 

1.	User Authentication: Authenticates the user with SharePoint Online using PnP login. 

2.	Site Collection Retrieval: Fetches all site collections from SharePoint Online. 

3.	Classic Sites Filtering: Filters and identifies classic sites based on the template and optional title ranges. 

4.	Batch Processing: Allows for processing sites in batches according to their titles (A-J, K-S, T-Z, or All). 

5.	Data Exporting: Offers the option to export the filtered list of classic sites to a CSV file or display it in the console. 

 

 

 

 
