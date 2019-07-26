Select WORKORDERID
				, [VALUE1] = MAX(CASE WHEN rn = 1 THEN Answer END)
				, [VALUE2] = MAX(CASE WHEN rn = 2 THEN Answer END)
				, [VALUE3] = MAX(CASE WHEN rn = 3 THEN Answer END)
				, [VALUE4] = MAX(CASE WHEN rn = 4 THEN Answer END)
				, [VALUE5] = MAX(CASE WHEN rn = 5 THEN Answer END)
From 
(
Select w.WORKORDERID
, rn = ROW_NUMBER() OVER (PARTITION BY WORKORDERID ORDER BY WORKORDERID)
, y.Answer
FROM
[servicedesk].[dbo].[WORKORDERSTATES] w
LEFT JOIN [servicedesk].[dbo].[WO_RESOURCES] x
ON  w.workorderid = x.woid
Full Outer Join [servicedesk].[dbo].ResourcesQAMapping y
ON x.UID = y.MAPPINGID

WHERE w.APPR_STATUSID = '2'
AND w.STATUSID = '1'
AND w.REOPENED = 'false'
)q
GROUP BY WorkOrderID