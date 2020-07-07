
-- **********************************************************************************************************
-- PURPOSE : This creates the tables and stored procedures required to run the tests for the Data Access Layer
-- WARNING : This script DROPS and RECREATES the following named database objects:
--           country
--	     spGetItems
--           spDeleteItem
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
	DELETE FROM [dbo].[country]
END
ELSE
BEGIN
	CREATE TABLE [dbo].[country]
		(
			[country_id] [int] IDENTITY(1,1) PRIMARY KEY,
			[country_name]   [nvarchar] (200) NULL,
			[country_region] [nvarchar] (50) NOT NULL
		) 

END
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

-- This proc demonstrates passing one argument and returns the results of a query

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


-- This proc demonstrates passing one argument executing and action on the database and return a
-- value based on whether that action was successful or not

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


