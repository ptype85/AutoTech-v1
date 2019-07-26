Select WORKORDERID, SUBCATEGORYID, ITEMID
				, [VALUE1] = MAX(CASE WHEN rn = 1 THEN Answer END)
From 
(
Select w.WORKORDERID,w.SUBCATEGORYID, w.ITEMID
, rn = ROW_NUMBER() OVER (PARTITION BY WORKORDERID ORDER BY WORKORDERID)
, y.Answer
FROM
[servicedesk].[dbo].[WORKORDERSTATES] w
LEFT JOIN [servicedesk].[dbo].[WO_RESOURCES] x
ON  w.workorderid = x.woid
Full Outer Join [servicedesk].[dbo].ResourcesQAMapping y
ON x.UID = y.MAPPINGID

WHERE w.STATUSID = '1'
AND w.REOPENED = 'false'
AND w.ITEMID = '3301'
)q
GROUP BY WorkOrderID, SUBCATEGORYID, ITEMID