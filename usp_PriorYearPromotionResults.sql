--PASSES 8/22
USE [Inventory-Promotions]
GO
/****** Object:  StoredProcedure [dbo].[uspPriorYearPromotionResults]    Script Date: 08/22/2013 12:00:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--IF OBJECT_ID ( 'uspPromotions', 'P' ) IS NOT NULL 
  --  DROP PROCEDURE uspPromotions;
--GO

ALTER PROCEDURE [dbo].[uspPriorYearPromotionResults] AS
SET NOCOUNT ON

--GATHER THE PARAMETERS INTO THE TEMP_TABLE TO RUN QUERIES

--DECLARE VARIABLES
--CREATE @Iter TO LOOP THROUGH TEMP TABLE
 --

DECLARE @Iter int,
		 @MaxRownum int,  
		 @ID varchar(10),
		 @PROMO_ID varchar(255),
		 @PROPERTY_ID varchar(30),
		 @CHANNEL_ID varchar(255),
		 @TRAVEL_START_DATE varchar(255),
		 @TRAVEL_END_DATE varchar(255),
		 @PROMO_START_DATE varchar(255),
		 @PROMO_END_DATE varchar(255),
		 @MIN_NIGHTS varchar(255),
		 @PRIOR_YEAR_TOTAL_RESERVATIONS int,
		 @PRIOR_YEAR_TOTAL_NIGHTS int,
		 @PRIOR_YEAR_TRAVEL_START_DATE datetime,
		 @PRIOR_YEAR_TRAVEL_END_DATE datetime,
		 @PRIOR_YEAR_PROMO_START_DATE datetime,
		 @PRIOR_YEAR_PROMO_END_DATE datetime

                SET @MaxRownum = (SELECT MAX(RowNum) FROM [Inventory-Promotions].dbo.TEMP_TABLE);
				SET @Iter = (SELECT MIN(RowNum) FROM [Inventory-Promotions].dbo.TEMP_TABLE);
				
                --Start the Loop    
                WHILE @Iter <= @MaxRownum
					BEGIN 
						--Query the SAP_MASTER_GRID for current campaign parameters
						--Set the contents of the variables
						SET @CHANNEL_ID = 			(SELECT CHANNEL_ID FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE RowNum = @Iter);
						SET @PROMO_ID = 			(SELECT PROMO_ID FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE RowNum = @Iter);
						SET @PROPERTY_ID = 			(SELECT PROPERTY_ID FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE RowNum = @Iter);
						SET @TRAVEL_START_DATE = 	(SELECT TRAVEL_START_DATE FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE Rownum = @Iter);
						SET @TRAVEL_END_DATE = 		(SELECT TRAVEL_END_DATE FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE Rownum = @Iter);
						SET @PROMO_START_DATE = 	(SELECT PROMO_START_DATE FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE Rownum = @Iter);
						SET @PROMO_END_DATE = 		getdate();
						SET @MIN_NIGHTS = 			(SELECT MIN_NIGHTS FROM [Inventory-Promotions].dbo.TEMP_TABLE WHERE Rownum = @Iter);
						SET @PRIOR_YEAR_TRAVEL_START_DATE = dateadd(yyyy, -1, @TRAVEL_START_DATE);
						SET @PRIOR_YEAR_TRAVEL_END_DATE = dateadd(yyyy, -1, @TRAVEL_END_DATE);
						SET @PRIOR_YEAR_PROMO_START_DATE = dateadd(yyyy, -1, @PROMO_START_DATE);
						SET @PRIOR_YEAR_PROMO_END_DATE = dateadd(yyyy, -1, @PROMO_END_DATE);
						
					----BEGIN
						-----IF @CHANNEL_ID = 'PPR'
							

						---ELSE
							BEGIN
								TRUNCATE TABLE PROMOTION;
								
								INSERT INTO [Inventory-Promotions].dbo.[PROMOTION]
								---------------------------------------------------------------------------------------------
								--Query without promo id
								--Now check for total res for current year
								---------------------------------------------------------------------------------------------
								SELECT DISTINCT 
									PROP.POPROP,
									CLUBINV.START_DATE,
									CLUBINV.END_DATE,
									CLUBINV.SITE_REFERENCE, 
									CLUBINV.NIGHTS,
									CLUBINV.TRAN_DATE,
									CLUBINV.UNIT_TYPE,
									CLUBRUX.UNIT_SIZE,
									PERFI.PRNAME,
									CLUBINV.CLUB

								FROM BBX.dbo.CLUBINV

								LEFT JOIN BBX.dbo.CLUBSIT ON CLUBSIT.SITE_ID=CLUBINV.SITE
								LEFT JOIN BBX.dbo.PROP ON PROP.POPROP=CLUBSIT.SITE_TYPE
								LEFT JOIN BBX.dbo.CLUBRUX ON CLUBRUX.PROP=PROP.POPROP AND CLUBRUX.ROOM_TYPE=CLUBINV.UNIT_TYPE
								LEFT JOIN BBX.dbo.PERFI ON PERFI.PRNUMB=SUBSTRING(CLUBINV.AUTH_USE,4,5)
								LEFT JOIN BBX.dbo.AGTSUP ON AGTSUP.AGT=SUBSTRING(CLUBINV.AUTH_USE,4,5) 

								AND AGTSUP.FROM_DATE<=CLUBINV.TRAN_DATE 
								AND (AGTSUP.THRU_DATE IS NULL OR AGTSUP.THRU_DATE>=CLUBINV.TRAN_DATE) 

								LEFT JOIN BBX.dbo.PERFI P ON P.PRNUMB=AGTSUP.SUP
								LEFT JOIN BBX.dbo.RSRV R on R.RSRV_NUM = CLUBINV.SITE_REFERENCE
								WHERE CLUBINV.CLUB='SVN' AND PROP.POPROP IN (@PROPERTY_ID) 
								AND CLUBINV.SITE_REFERENCE IS NOT NULL 
								AND CLUBINV.DISPOSITION = 2

								--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
								--Modify these dates for daily
								AND CLUBINV.TRAN_DATE BETWEEN @PRIOR_YEAR_PROMO_START_DATE AND @PRIOR_YEAR_PROMO_END_DATE
								AND CLUBINV.START_DATE BETWEEN @PRIOR_YEAR_TRAVEL_START_DATE AND @PRIOR_YEAR_TRAVEL_END_DATE AND CLUBINV.NIGHTS>=@MIN_NIGHTS
								--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx	

								
								---------------------------------------------------------------------------------------------
								--INSERT CALCULATED DATA INTO SAP_MASTER GRID
								---------------------------------------------------------------------------------------------
								SELECT
									@PRIOR_YEAR_TOTAL_RESERVATIONS = COUNT(NIGHTS)
								FROM 
									PROMOTION 
								WHERE  SITE_TYPE=@PROPERTY_ID
								---------------------------------------------------------------------------------------------
								SELECT   
					
										@PRIOR_YEAR_TOTAL_NIGHTS = SUM(CASE WHEN B.DEP_CAMPAIGN_NIGHTS IS NULL and C.ARR_CAMPAIGN_NIGHTS IS NULL THEN A.NIGHTS 
										WHEN B.DEP_CAMPAIGN_NIGHTS IS NOT NULL and C.ARR_CAMPAIGN_NIGHTS IS NULL THEN B.DEP_CAMPAIGN_NIGHTS
										WHEN B.DEP_CAMPAIGN_NIGHTS IS NULL and C.ARR_CAMPAIGN_NIGHTS IS NOT NULL THEN C.ARR_CAMPAIGN_NIGHTS
										ELSE A.NIGHTS END) --AS CAMPAIGN_NIGHTS
										
									FROM 
										PROMOTION 
								A

								LEFT OUTER JOIN (
										SELECT  RSRV_NUM, NULL AS TOTAL_NIGHTS, 
										NIGHTS - DATEDIFF(d,@PRIOR_YEAR_TRAVEL_END_DATE,DEPARTURE) AS DEP_CAMPAIGN_NIGHTS,  
										DATEDIFF(d,@PRIOR_YEAR_TRAVEL_END_DATE,DEPARTURE) AS OUT_DEP_CAMPAIGN_NIGHTS
									FROM 
										PROMOTION 
									WHERE 
										DEPARTURE > @PRIOR_YEAR_TRAVEL_END_DATE
									) 
								B	
								ON A.RSRV_NUM = B.RSRV_NUM

								LEFT OUTER JOIN (
										SELECT  RSRV_NUM, NULL AS TOTAL_NIGHTS, 
										NIGHTS - DATEDIFF(d,ARRIVAL, @PRIOR_YEAR_TRAVEL_START_DATE) AS ARR_CAMPAIGN_NIGHTS,  
										DATEDIFF(d,ARRIVAL, @PRIOR_YEAR_TRAVEL_START_DATE) AS OUT_ARR_CAMPAIGN_NIGHTS
									FROM 
										PROMOTION 
									WHERE 
										 ARRIVAL < @PRIOR_YEAR_TRAVEL_START_DATE
									) 
								C
									ON A.RSRV_NUM = C.RSRV_NUM
									
									GROUP BY 
									A.SITE_TYPE
								
								---------------------------------------------------------------------------------------------
								UPDATE [Inventory-Promotions].dbo.[SAP_MASTER_GRID]
								SET	PRIOR_YEAR_TOTAL_NIGHTS = @PRIOR_YEAR_TOTAL_NIGHTS,
										PRIOR_YEAR_TOTAL_RESERVATIONS = @PRIOR_YEAR_TOTAL_RESERVATIONS
								WHERE  ID=@Iter;
								---------------------------------------------------------------------------------------------
								
								END
					-----END
					--Increment the @Iter variable to loop to the next row in the TEMP_TABLE
					SET @Iter = @Iter + 1;
					--PRINT @Iter
				END
				--xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx	
GO