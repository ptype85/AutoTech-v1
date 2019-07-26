Select WORKORDERID, LAST_TECH_UPDATE
									, [VALUE1] = MAX(CASE WHEN rn = 1 THEN Answer END)
From 
(
Select w.WORKORDERID, w.LAST_TECH_UPDATE
, rn = ROW_NUMBER() OVER (PARTITION BY WORKORDERID ORDER BY WORKORDERID)
, y.Answer
FROM
[servicedesk].[dbo].[WORKORDERSTATES] w
LEFT JOIN [servicedesk].[dbo].[WO_RESOURCES] x
ON  w.workorderid = x.woid
Full Outer Join [servicedesk].[dbo].ResourcesQAMapping y
ON x.UID = y.MAPPINGID

WHERE w.APPR_STATUSID = '2'
AND w.STATUSID = '2'
AND w.REOPENED = 'false'
)q
GROUP BY WorkOrderID, LAST_TECH_UPDATE