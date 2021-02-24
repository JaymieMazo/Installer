


SET DATEFIRST 1;
SELECT  datepart(weekday,ReceivedDate) AS 'Day' ,* FROM GetDetails('2016/10/05','2016/12/05')

INNER JOIN 
		(SELECT  datename(weekday,ReceivedDate) AS DayName,
		datepart(day,ReceivedDate) AS DayMismo,ReceivedDate, *
		FROM GetDetails('2016/10/05','2016/12/05')) DayName
ON dayname.ReceivedDate = Dayview.ReceivedDate




SELECT * FROM GetDetails('2016/10/05','2016/12/05')