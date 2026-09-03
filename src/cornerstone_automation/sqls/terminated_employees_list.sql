-- query: terminated_employees_list
SELECT DISTINCT
    wd.LegacyID                          AS EMPLOYEE_CODE,
    wd.FirstName + ' ' + wd.LastName     AS EMPLOYEE_NAME,
    wd.TerminationDate                   AS TERMINATE_DATE,
    p.OFFC,
    SUM(TT.BASE_HRS)                     AS Total_Hrs
FROM sqlt4costagedw.[HumanResources].[dbo].[Master_Users_WD] wd
INNER JOIN CMSOPEN..HBM_PERSNL p (NOLOCK)
    ON p.EMPLOYEE_CODE = wd.LegacyID
INNER JOIN CMSOPEN..TAT_TIME TT (NOLOCK)
    ON TT.TK_EMPL_UNO = p.EMPL_UNO
INNER JOIN CMSOPEN..HBM_MATTER HM (NOLOCK)
    ON TT.MATTER_UNO = HM.MATTER_UNO
INNER JOIN CMSOPEN..hbm_client HC (NOLOCK)
    ON HM.CLIENT_UNO = HC.CLIENT_UNO
INNER JOIN CMSOPEN..HBL_OFFICE o (NOLOCK)
    ON p.OFFC = o.OFFC_CODE
WHERE
    -- Terminated during the report year
    wd.TerminationDate BETWEEN ? AND ?
    -- Had actual hours during the report period
    AND TT.TRAN_DATE BETWEEN ? AND ?
    -- Only regular permanent employees — Workday source of truth
    AND wd.EmployeeType = 'Regular'
    AND wd.EmploymentType IN ('Full time', 'Part time')
    -- Only valid time entries
    AND TT.WIP_STATUS IN ('W','P','B')
    -- Exclude internal time
    AND HC.CLIENT_CODE <> '99008'
    -- Only numeric employee codes
    AND wd.LegacyID LIKE '[0-9]%'
    -- Exclude corporate office
    AND o.OFFC_CODE <> 'COR'
GROUP BY
    wd.LegacyID,
    wd.FirstName + ' ' + wd.LastName,
    wd.TerminationDate,
    p.OFFC
HAVING SUM(TT.BASE_HRS) > 0