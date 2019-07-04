DECLARE @date_start date;
DECLARE @phone_number nvarchar(200);
DECLARE @start_id int;
SET @date_start = CAST('2019-04-30' AS date);
SET @phone_number = '8007750774';
SET @start_id = 60;

-- total number of 'Hold'
SELECT COUNT(*) as quantity
FROM dbo.A_Stat_FailedCalls
WHERE ANumberdialed = '*7400'
AND TimeStart > @date_start
AND ReasonFailed = 27

-- fast disconnection
SELECT TimeStart, TimeStop, AOutNumber,
       DATEDIFF(s, TimeStart, TimeStop) as duration
FROM dbo.A_Stat_FailedCalls
WHERE TimeStart > @date_start
AND ANumberdialed = '*7400'
AND ReasonFailed = 27
AND DATEDIFF(s, TimeStart, TimeStop) <= 10

-- waiting more than 10 seconds
SELECT TimeStart, TimeStop, AOutNumber,
       DATEDIFF(s, TimeStart, TimeStop) as duration
FROM dbo.A_Stat_FailedCalls
WHERE TimeStart > @date_start
AND ANumberdialed = '*7400'
AND ReasonFailed = 27
AND DATEDIFF(s, TimeStart, TimeStop) > 10

-- duration MAX
SELECT DATEDIFF(s, DateTimeStart, DateTimeStop) as maximum
FROM dbo.Stat_CallCenter
WHERE VhodNomer = @phone_number
AND Id > @start_id
AND DATEDIFF(s, DateTimeStart, DateTimeStop) IN
(SELECT MAX(DATEDIFF(s, DateTimeStart, DateTimeStop))
 FROM dbo.Stat_CallCenter
 WHERE VhodNomer = @phone_number
 AND Id > @start_id
 AND DateTimeStart > @date_start);

-- duration MIN
SELECT DISTINCT DATEDIFF(s, DateTimeStart, DateTimeStop) as minimum
FROM dbo.Stat_CallCenter
WHERE VhodNomer = @phone_number
AND Id > @start_id
AND DATEDIFF(s, DateTimeStart, DateTimeStop) IN
(SELECT MIN(DATEDIFF(s, DateTimeStart, DateTimeStop))
 FROM dbo.Stat_CallCenter
 WHERE VhodNomer = @phone_number
 AND Id > @start_id
 AND DateTimeStart > @date_start);

-- duration MEAN
SELECT AVG(DATEDIFF(s, DateTimeStart, DateTimeStop)) as mean
FROM dbo.Stat_CallCenter
WHERE VhodNomer = @phone_number
AND Id > @start_id
AND DateTimeStart > @date_start;

-- number of rated calls, mean rate
SELECT COUNT(*) as rated_calls, AVG(CAST(Ball AS int)) as mean_rate
FROM dbo.Stat_CallCenter
WHERE VhodNomer = @phone_number
AND Id > @start_id
AND DateTimeStart > @date_start
AND Ball <> '<Без оценки>';