-- query: office_transfers_ytd
SELECT DISTINCT
    p.EMPLOYEE_CODE,
    p.EMPLOYEE_NAME,
    p.OFFC                          AS Current_Office,
    td_hist.Office_Code             AS Historical_Office,
    MIN(td_hist.[Date])             AS First_Historical_Date,
    MAX(td_hist.[Date])             AS Last_Historical_Date,
    COUNT(DISTINCT td_hist.[Date])  AS Days_In_Old_Office,
    SUM(ISNULL(td_hist.Base_hrs,0)) AS Hrs_In_Old_Office
FROM CMSOPEN..HBM_PERSNL p (NOLOCK)
INNER JOIN ReportDevl..TargetDaily td_hist
    ON td_hist.EMPL_UNO = p.EMPL_UNO
    AND td_hist.[Date] BETWEEN ? AND ?
INNER JOIN CMSOPEN..HBL_OFFICE o (NOLOCK)
    ON p.OFFC = o.OFFC_CODE
WHERE
    td_hist.Office_Code <> p.OFFC
    AND p.EMPLOYEE_CODE LIKE '[0-9]%'
    AND p.Employee_Type_Code NOT IN ('TMP', 'INT', 'SIT', 'WIT')
    AND o.OFFC_CODE <> 'COR'
    AND td_hist.Base_hrs > 0
GROUP BY
    p.EMPLOYEE_CODE,
    p.EMPLOYEE_NAME,
    p.OFFC,
    td_hist.Office_Code
HAVING SUM(ISNULL(td_hist.Base_hrs,0)) > 0
ORDER BY p.EMPLOYEE_CODE
