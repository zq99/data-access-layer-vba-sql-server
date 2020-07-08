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

## Implementation

There are various examples of how to use the Data Access Layer in the module 'mdTestDataAccessLayer'. These examples include executing
stored procs on the database to retrieve information as well as executing common CRUD operations. As a glimpse of how the layer works, here is a brief 
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

https://datapluscode.com/general/building-a-data-access-layer-in-vba-for-sql-server-part-1/


