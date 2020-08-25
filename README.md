# Data Access Layer in VBA for connecting to SQL Server

## Purpose

This project is Data Access layer that can be used to connect from VBA to SQL Server.

The purpose of the layer is to keep VBA code within projects tidy and core database commands centralized within a project. 

The layer handles all actions involved with reading and writing data to SQL Server.


## Requirements

You must have the following VBA Project references installed in the VBE Editor reference window (minimum versions stated):

- Visual Basic For Applications
- Microsoft Excel 14.0 Object Library
- OLE Automation
- Microsoft Office 14.0 Object Library
- Microsoft Active X Data Objects 6.1 Library
- Microsoft Active X Data Objects Recordset Library 6.0
- Microsoft ADO Ext. 6.0 for DDL and Security


## Overview

The layer is organized as follows:

<img src=screenshots/DataAccessLayer.png width=500>


## Configuration

You must specify the name of your server and the name of your database in the class DBInstance. 

These are the 2 lines to amend:

    objDBCredentials.SetServer = "MY_SERVER_NAME"
    objDBCredentials.SetDBName = "MY_DATABASE_NAME"


## Methods

The layer has three main methods:

- ExecuteStoredProc - this is for running stored procs on the database for CRUD operations.
- GetRecordsetFromStoredProc - this is for retrieving data from the database via stored procs.
- GetDataFromSQLStatement - this is for retrieving data from the database via a raw SQL query.

All the methods can accept variables being passed back and forth between the layer and the database.


## Implementation

There are various examples of how to use the Data Access Layer in the module 'mdTestDataAccessLayer'. These examples include executing
stored procs in the database to retrieve data, and for executing common CRUD operations. To give you a brief idea of how the layer works, here is an
example of getting data from SQL Server using a simple SQL statement:

    Dim rstData     As New ADODB.Recordset
    Dim oDBInstance As New clsDBInstance
    Dim oDatabase   As clsDatabase
    
    Set oDatabase = oDBInstance.GetSharedDatabase()
    Set rstData = oDatabase.GetDataFromSQLStatement("SELECT country_name, country_region FROM [dbo].[country]")
    rstData.MoveFirst
    Do While Not rstData.EOF
        Debug.Print rstData.Fields("country_name").Value & " " & rstData.Fields("country_region").Value
        rstData.MoveNext
    Loop
    
    oDBInstance.CloseSharedDatabase


## Further information

- https://datapluscode.com/general/building-a-data-access-layer-in-vba-for-sql-server-part-1/
- https://datapluscode.com/general/how-to-use-the-data-access-layer-part-2/


