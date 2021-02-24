

ALTER FUNCTION [dbo].[GetBreakDownDetails_ForkLift] 
(@startDate as DateTime, @endDate as DATETIME, @searchDate as DATETIME, @previousDate as DATETIME)
 RETURNS Table 
 AS RETURN
 
SELECT
WorkOrderDetails.CompanyID,
Companies.CompanyName,
WorkOrderItems.ItemName,
ISNULL(B.WAITINGPARTS,0) AS WAITINGPARTS,
ISNULL(C.FORSCHEDULE,0)  AS FORSCHEDULE,
ISNULL(D.ONGOING,0)  AS ONGOING,
ISNULL(E.FINISHEDREPAIR,0) AS FINISHEDREPAIR,
ISNULL(F.NEWBREAKDOWN,0) AS NEWBREAKDOWN,
ISNULL(G.PREVIOUSPENDING,0) AS PREVIOUSPENDING

FROM WorkOrderDetails

INNER JOIN Companies
ON WorkOrderDetails.CompanyID	= Companies.CompanyID

INNER JOIN WorkOrderItems
ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

LEFT JOIN
(
	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS WAITINGPARTS
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE WorkOrderDetails.StatusID = 2 
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) BETWEEN @startDate AND @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName, WorkOrderDetails.WorkOrderControlNo
)B
ON WorkOrderDetails.CompanyID = B.CompanyID
AND WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo = B.ControlNo

LEFT JOIN
(
	SELECT DISTINCT	
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS FORSCHEDULE
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE WorkOrderDetails.StatusID = 4 
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) BETWEEN @startDate AND @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
--	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @startDate
--	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @endDate
	
	GROUP BY 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName, WorkOrderDetails.WorkOrderControlNo
)C
ON WorkOrderDetails.CompanyID = C.CompanyID
AND WorkOrderDetails.AbbreviatedName + ' ' + WorkOrderDetails.WorkOrderControlNo = C.ControlNo

LEFT JOIN
(
	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS ONGOING
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE WorkOrderDetails.StatusID = 1
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) BETWEEN @startDate AND @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
)D
ON WorkOrderDetails.CompanyID = D.CompanyID
AND WorkOrderDetails.AbbreviatedName + ' ' + WorkOrderDetails.WorkOrderControlNo = D.ControlNo

LEFT JOIN
(
	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS FINISHEDREPAIR
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) = @searchDate
	AND WorkOrderDetails.FinishedDate IS NOT NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
)E
ON WorkOrderDetails.CompanyID = E.CompanyID
AND WorkOrderDetails.AbbreviatedName + ' ' + WorkOrderDetails.WorkOrderControlNo = E.ControlNo

LEFT JOIN
(
	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS NEWBREAKDOWN
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) = @searchDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
)F
ON WorkOrderDetails.CompanyID = F.CompanyID
AND WorkOrderDetails.AbbreviatedName + ' ' + WorkOrderDetails.WorkOrderControlNo = F.ControlNo

LEFT JOIN
(
	SELECT 
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName	+ ' ' + WorkOrderDetails.WorkOrderControlNo AS ControlNo, 
	count(*) AS PREVIOUSPENDING
	FROM WorkOrderDetails
	
	LEFT JOIN Types	
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID

	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode

	LEFT JOIN PrsHeaders
	On  WorkOrderDetails.AbbreviatedName+' '+WorkOrderDetails.WorkOrderControlNo = PrsHeaders.WorkOrderControlNo
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName NOT LIKE '%PREVENTIVE%'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) BETWEEN @previousDate AND @searchDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
	
	GROUP BY
	WorkOrderDetails.CompanyID, 
	WorkOrderDetails.AbbreviatedName,WorkOrderDetails.WorkOrderControlNo
)G
ON WorkOrderDetails.CompanyID = G.CompanyID
AND WorkOrderDetails.AbbreviatedName + ' ' + WorkOrderDetails.WorkOrderControlNo = G.ControlNo


