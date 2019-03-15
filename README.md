# Assembly Builder
Enter non-existent Item Assemblies and modify existing Item Assemblies with up-to-date information. Built using Microsoft's Visual Studio Assembly Builder, works hand in hand with QuickBooks to modify Company Files with new and current intelligence. 


## Table of contents
* [Features](#features)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installing](#installing)
  * [Running](#running)
  	* [Importer](#importer)
	* [BOM Viewer](#bom)
* [Parameters](#parameters)
  * [Checked Boxes](#checked)
  * [Error Messages](#errors)
* [Versions](#versions)
* [Built With](#built-with)
* [Authors](#authors)
* [Copyright and License](#copyright-and-license)


## Features
Few things you can do using Assembly Builder:
* Open ready to upload Item Assemblies in a user-friendly view
* View complete BOMs (if available) of Item Assemblies and their sub-Assemblies
* Identify errors that might exist in pre-built assemblies
  * *i.e. wrong ItemCode as a sub assembly/part*
* Add Assemblies with approriate sub Assemblies/Part regardless if sub Items exist
* Modify pre-existsing Assemblies with up-to-date sub-Items


## Getting Started


### Prerequisites
* [QuickBooks](https://quickbooks.intuit.com/)
* [SDK Installer](https://developer.intuit.com/downloads/restricted?filename=qbsdk130.exe&download=true)


### Installing 
1. [Download](https://github.com/esanch/AssemblyBuilder/archive/v2.0.zip) Assembly Builder project<br />
2. Extract ZIP files and save on your computer<br />


### Running
*\*Refer to [Installing](#installing) prior to running program\**
1. Once program files are downloaded, double-click and run **StartUp.exe** file <br />
	*\*An interactive dialog window will appear*<br />


#### Importer
1. Ensure step 1 in [Running](#running) is completed prior to moving on to step 2<br />
2. On the interactive window locate and click **"Open File"**<br />
	*\*This will open a new dialog box*
3. Select the **.xml** file of the Assembly you wish to add/modify
4. Once the Assembly loads in the pane check the top right textbox if an error exists in the Assembly
5. If no errors exist send the Assembly to QuickBooks by clicking **"Send"** <br />
	*\*Note: This might take longer than a couple seconds depending on the size of your assembly*<br />
	*\**Steps taken during implementation will be displayed on the bottom right pane*
6. Repeat steps 2-5 if you wish to add/modify another Assembly
7. To close the program simply click **"Exit"**


#### BOM
1. Ensure step 1 in [Running](#running) is completed prior to moving on to step 2 <br />
2. On the interactive window locate and click **"Show BOM"**<br />
	*\*This will open a new dialog box*<br />
3. Select the **.xml** file of the BOM Assembly you wish to view<br />
	*\*Note: You will be unable to import a complete BOM in **Show BOM** mode.*<br />
	*\**You must use the [importer](#importer)'s **Open File** method to correctly add/modify an Assembly*<br />
4. If an item contains a BOM (with sub assemblies/parts) it will be displayed on the center left pane<br />
	*\*If no complete BOM exists an empty table with headers will appear*<br />
5. Repeat steps 2-4 if you wish to view another Item's complete BOM Assembly<br />
6. To close the program simply click **"Exit"**<br />


## Parameters 
* **Top Level:** **|** Indicates the ItemCode of the item being viewed/modified/added. Acts as the top level of the table being displayed <br />


### Checked
* **Item already exists?**<br />
	* Yes 
		* **|** Item trying to import already exists in QuickBooks <br />
		*\*Note: If top level Item already exists importer will try to modify the Item Assembly*
	* No  
		* **|** Item trying to import needs to be added as a new item part/assembly 


* **Is Part or Assembly?**<br />
	* Part     
		* **|** Top level Item is an Item Part <br />
		*\*Note: A part item cannot be added as an assembly*
	* Assembly 
		* **|** Top level Item is an Item Assembly and thus can be added/modified as a top level assembly


* **Showing BOM?**<br />
	* Yes 
		* **|** Complete BOM is being displayed with its sub-Assemblies and all sub-Assemblies beyond those<br />
		*\*Note: This is an indication of whether the user has chosen to view a BOM regardless of its existence*<br />
		*\*Complete BOMs cannot be sent to be added/modifies as Assemblies
	* No 
		* **|** An Item Assembly is being displayed along with only its immediate sub-Assemblies<br />
		*\*Note: This is typically done when a user intends to add/modify an Item Assembly into Quickbook*

----

* **ICodes match?**<br />
	* Yes                
		* **|** The ItemCode provided in the **.xml** file exists in the database and thus can be added/modified
	* No                 
		* **|** There exists an error in the input of the ItemCode. *Check to see whether the **.xml** file is wrong.*<br />
		*\*This poses a problem since importer cannot detrmine which Item you are trying to add/modify into QuickBooks*
	* ItemCode not found 
		* **|** The ItemCode provided does not exist in the company database<br />
		*\*This poses a problem since importer cannot detrmine which Item you are trying to add/modify into QuickBooks*


* **Columns in correct order:**<br />
	*\*Note: This will **not** cause an error. Program will automatically fix wrong row ordering*
	* Item Number 
		* **|** The **.xml** file contains the column "NO." in the first (correct) column
	* ItemCode    
		* **|** The **.xml** file contains the column "ITEMCODE" in the second (correct) column
	* Part Number 
		* **|** The **.xml** file contains the column "PARTNUMBER" in the third (correct) column
	* Description 
		* **|** The **.xml** file contains the column "DESCRIPTION" in the fourth (correct) column
	* Quantity    
		* **|** The **.xml** file contains the column "QTY" in the fifth (correct) column


* **Correct input?**<br />
	* Yes
		* **|** All the values in in table are valid and able to be added/modified
	* No  
		* **|** A/Some value(s) in the table are invalid and must be changed in order to add/modify<br />
			*\*Possible Errors:*<br />
			- *Column contians value of **one** instead of **1** or vice versa*
			- *Invalid ItemCode is provided i.e. 14HW8900* <br />
		*Note: The top right error pane will povide the line number in which the error occurs*<br />


* **Row order correct?**<br />
	* Yes
		* **|** All rows are sequentially correct and thus program can continue with implementation
	* No  
		* **|** Rows provided contain an error<br />
		*\*Possible Errors:*<br />
			- *Rows are not in ascending sequential order starting from 1 i.e. 2, 1, 3, 4, 5*<br />
			- *Rows are missing i.e. 1, 3, 4, 5*


### Errors


* **Table contains errors!** <br />
	* Possible Errors:<br />
		* Values in table are wrong i.e. **one** vs **1**
		* Row sequence is wrong/missing 


* **Error in the following ItemNo provided:**<br />
	* The provided value will point to the row \# that contains the error<br />
	*\*Note: If no ItemNo is provided the value is missing from the table*<br />


* **Error in the following ItemCode provided in line #:**
	* \# will be changed according to the row number that contains an error followed by the value raising the error <br />
	*\*Note: If no ItemCode is provided the value is missing from the table*<br />


* **Error in the following Quantity provided in line #:**
	* \# will be changed according to the row number that contains an error followed by the value raising the error <br />
	*\*Note: If no Quantity is provided the value is missing from the table*<br />


* **There is no row at position 0.**
	* Top level ItemCode is missing <br />
	*\*This value is needed to alert QuickBooks of the Item being added/modified*


* **Cannot import a table with errors**<br />
	**or a complete BOM with sub BOMs!**<br />
	* Possible Errors: <br />
		* User attempted to send an Item Assembly with errors to QuickBooks
		* User was viewing a BOM and attempted to send the complete BOM to Quickbooks


## Versions
>**Assembly Builder** 
>>[v1.0](https://github.com/esanch/AssemblyBuilder/tree/v1.0)  (01/22/2019) <br />
>>[v1.0.1](https://github.com/esanch/AssemblyBuilder/archive/v1.0.1.zip) (01/30/2019) <br />
>>[v2.0](https://github.com/esanch/AssemblyBuilder/archive/v2.0.zip) (02/06/2019) <br />

## Built With
* IDE - [Visual Studio](https://visualstudio.microsoft.com/vs/)

## Authors
**Elizabeth Earl**
* [Stack Overflow](https://stackoverflow.com/users/10267094/e-earl)
* [GitHub](https://github.com/esanch/)

## Copyright and License
Code and documentation rights reserved 2018-2019 [Elizabeth Earl](mailto:elizabeth.earl@ubeusa.com) 

Code released under supervision of [United Bakery Equipment](http://ubeusa.com/)
