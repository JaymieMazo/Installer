

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS WAITINGPARTS
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID = 2 
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= '2018/01/01'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
	
--============================================================================================
	

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS FORSCHEDULE
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID = 4
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= '2018/01/01'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY 	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
	
--============================================================================================
	

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS ONGOING
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID = 1
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= '2018/01/01'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
	
--============================================================================================
	

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS FINISHEDREPAIR
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) = '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NOT NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
	
--============================================================================================
	

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS NEWBREAKDOWN
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) = '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
	
--============================================================================================
	

	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS PREVIOUSPENDING
	FROM WorkOrderDetails
	
	INNER JOIN Types	
	ON WorkOrderDetails.CompanyID	= Types.CompanyID	
	AND WorkOrderDetails.TypeId	= Types.TypeID
	
	INNER JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= '2018/02/12'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '2019/02/13'
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo