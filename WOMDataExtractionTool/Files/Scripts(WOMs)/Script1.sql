
ALTER FUNCTION [dbo].[GetBreakDownDetails_ForkLift] 
(@startDate as DateTime, @endDate as DATETIME, @searchDate as DATETIME, @previousdate as DATETIME)
 RETURNS Table 
 AS RETURN


SELECT DISTINCT	
A.CompanyID,
Companies.CompanyName,
WorkOrderItemsView.ItemName,

ISNULL(B.WAITINGPARTS,0) AS WAITINGPARTS,
ISNULL(C.FORSCHEDULE,0)  AS FORSCHEDULE,
ISNULL(D.ONGOING,0)  AS ONGOING,
ISNULL(E.FINISHEDREPAIR,0) AS FINISHEDREPAIR,
ISNULL(F.NEWBREAKDOWN,0) AS NEWBREAKDOWN,
ISNULL(G.PREVIOUSPENDING,0) AS PREVIOUSPENDING

FROM WorkOrderDetails A


INNER JOIN Companies
ON A.CompanyID = Companies.CompanyID

INNER JOIN Types
ON A.CompanyID = Types.CompanyID
AND A.TypeId = Types.TypeID

LEFT JOIN MainCategories
ON A.CompanyID = MainCategories.CompanyID
AND A.MainCategoryID = MainCategories.MainCategoryID


LEFT JOIN WorkOrderItemsView
ON  A.ItemCode = WorkOrderItemsView.ItemCode
AND A.CompanyID = WorkOrderItemsView.CompanyID


--=======================================WAITING PARTS================================================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS WAITINGPARTS
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID

	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	WHERE WorkOrderDetails.StatusID = 2
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @startDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) B
ON A.CompanyID = B.CompanyID
AND A.ItemCode = B.ItemCode


--=======================================FOR SCHEDULE================================================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS FORSCHEDULE
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID = 4
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @startDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) C
ON A.CompanyID = C.CompanyID
AND A.ItemCode = C.ItemCode


--=======================================ON GOING================================================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS ONGOING
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID = 1
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @startDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @endDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) D
ON A.CompanyID = D.CompanyID
AND A.ItemCode = D.ItemCode


--=======================================FINISHED REPAIR OF THE DAY==============================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS FINISHEDREPAIR
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @searchDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @searchDate
	AND WorkOrderDetails.FinishedDate IS NOT NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) E
ON A.CompanyID = E.CompanyID
AND A.ItemCode = E.ItemCode


--=======================================NEW BREAKDOWN UNIT OF THE DAY==============================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS NEWBREAKDOWN
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @searchDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @searchDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) F
ON A.CompanyID = F.CompanyID
AND A.ItemCode = F.ItemCode


--=======================================PREVIOUS PENDING==============================

LEFT JOIN
(
	SELECT WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode, count(*) AS PREVIOUSPENDING
	FROM WorkOrderDetails
	
	INNER JOIN Types
	ON WorkOrderDetails.CompanyID = Types.CompanyID
	AND WorkOrderDetails.TypeId = Types.TypeID
	
	INNER JOIN WorkOrderItems
	ON WorkOrderDetails.CompanyID = WorkOrderItems.CompanyID
	AND WorkOrderDetails.ItemCode = WorkOrderItems.ItemCode	
	
	LEFT JOIN MainCategories
	ON WorkOrderDetails.CompanyID = MainCategories.CompanyID
	AND WorkOrderDetails.MainCategoryID = MainCategories.MainCategoryID
	
	WHERE WorkOrderDetails.StatusID IN (1,2,4,5)
	AND Types.TypeName LIKE '%FORKLIFT%'
	AND MainCategories.MainCategoryName <> 'PREVENTIVE'
	AND CONVERT(VARCHAR(20),ReceivedDate,111) >= @searchDate
	AND CONVERT(VARCHAR(20),ReceivedDate,111) <= @previousDate
	AND WorkOrderDetails.FinishedDate IS NULL
	AND WorkOrderDetails.DeletedDate IS NULL
		
	GROUP BY WorkOrderDetails.CompanyID, WorkOrderItems.ItemCode
) G
ON A.CompanyID = G.CompanyID
AND A.ItemCode = G.ItemCode


WHERE Types.TypeName LIKE '%FORKLIFT%'
AND MainCategories.MainCategoryName <> 'PREVENTIVE'
AND A.FinishedDate IS NULL
AND A.DeletedDate IS NULL

GROUP BY 
A.CompanyID,
Companies.CompanyName,
A.ItemCode,
WorkOrderItemsView.ItemName,
B.WAITINGPARTS,
C.FORSCHEDULE,
D.ONGOING,
E.FINISHEDREPAIR,
F.NEWBREAKDOWN,
G.PREVIOUSPENDING


GO

