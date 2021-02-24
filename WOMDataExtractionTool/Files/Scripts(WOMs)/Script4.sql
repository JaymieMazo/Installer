


 SELECT  
 
 CompanyID,
 CompanyName,
 sum(forschedule) AS [FOR SCHEDULE], 
 sum(Waitingparts) AS [WAITING PARTS], 
 sum(Ongoing) As [ON GOING], 
 sum(FinishedRepair) As [FINISHED REPAIR], 
 sum(NEWBREAKDOWN) As [NEW BREAKDOWN], 
 sum(PREVIOUSPENDING) As [PREVIOUS PENDING] 
 
 FROM  GetBreakDownDetails_ForkLift('2018/01/01','2019/02/13','2019/02/13','2019/02/12') 
 
 
 WHERE ItemName LIKE '%DIESEL%' 
 GROUP BY CompanyName,CompanyID

--SELECT * FROM GetBreakDownDetails_ForkLift('2018/01/01','2019/02/13','2019/02/13','2019/02/12') 

	

	