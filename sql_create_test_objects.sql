
-- **********************************************************************************************************
-- PURPOSE : This creates the tables and stored procedures required to run the tests for the Data Access Layer
-- WARNING : This script DROPS and RECREATES the following named database objects:
--           country (TABLE)
--			 spGetItems (STORED PROC)
--           spDeleteItem (STORED PROC)
-- **********************************************************************************************************

SET NOCOUNT ON; 


-- ================================================================
-- Creating the data tables and populating with test data
-- ================================================================


IF (EXISTS (SELECT * 
                 FROM INFORMATION_SCHEMA.TABLES 
                 WHERE TABLE_SCHEMA = 'dbo' 
                 AND  TABLE_NAME = 'country'))
BEGIN
	DROP TABLE [dbo].[country]
END
GO

CREATE TABLE [dbo].[country]
		(
			[country_id] [int] IDENTITY(1,1) PRIMARY KEY,
			[country_name]   [nvarchar] (200) NULL,
			[country_region] [nvarchar] (50) NOT NULL
		) 
GO

BEGIN
insert into [dbo].[country] ([country_name],[country_region]) values ('United Kingdom','EUROPE')
insert into [dbo].[country] ([country_name],[country_region]) values ('France','EUROPE')
insert into [dbo].[country] ([country_name],[country_region]) values ('Germany','EUROPE')
insert into [dbo].[country] ([country_name],[country_region]) values ('Spain','EUROPE')
insert into [dbo].[country] ([country_name],[country_region]) values ('Japan','ASIA')
insert into [dbo].[country] ([country_name],[country_region]) values ('India','ASIA')
END
GO


-- ================================================================
-- Creating the stored procedures to write and read from the tables
-- ================================================================

IF EXISTS (SELECT * 
            FROM   sysobjects 
            WHERE  id = object_id(N'[dbo].[spGetItems]') 
                   and OBJECTPROPERTY(id, N'IsProcedure') = 1 )
BEGIN
    DROP PROCEDURE [dbo].[spGetItems]
END
GO

-- This proc demonstrates passing 1 argument to the database and returning the results of a query

CREATE PROCEDURE [dbo].[spGetItems] @region varchar(50) = null
AS
	BEGIN
		SET NOCOUNT ON; 
		IF @region is null
			SELECT country_name, country_region
			FROM dbo.country;
		ELSE
			SELECT country_name, country_region
			FROM dbo.Country
			WHERE country_region = @region;
	END
GO


-- This proc demonstrates executing a proc on the database that accepts 1 variable and returning a
-- a result indicating its success.

IF EXISTS (SELECT * 
            FROM   sysobjects 
            WHERE  id = object_id(N'[dbo].[spDeleteItem]') 
                   and OBJECTPROPERTY(id, N'IsProcedure') = 1 )
BEGIN
    DROP PROCEDURE [dbo].[spDeleteItem]
END
GO


CREATE PROCEDURE [dbo].[spDeleteItem]
	@name varchar(50),
	@rowcount int OUTPUT,
	@err int OUTPUT
	AS
	BEGIN
		DELETE FROM dbo.country
		WHERE country_name = @name
		SELECT @err=@@error, @rowcount = @@ROWCOUNT
	End
GO


