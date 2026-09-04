-- query: active_employees_in_scope
SELECT
    wd.LegacyID                        AS EMPLOYEE_CODE,
    wd.FirstName + ' ' + wd.LastName   AS EMPLOYEE_NAME,
    wd.Office,
    wd.Department_Code,
    wd.EmployeeType,
    wd.EmploymentType,
    wd.CurrentStatus,
    wd.HireDate,
    wd.TerminationDate,
    wd.Level,
    wd.JobTitle
FROM sqlt4costagedw.[HumanResources].[dbo].[Master_Users_WD] wd
INNER JOIN CMSOPEN..HBM_PERSNL p (NOLOCK)
    ON p.EMPLOYEE_CODE = wd.LegacyID
INNER JOIN CMSOPEN..HBL_OFFICE o (NOLOCK)
    ON p.OFFC = o.OFFC_CODE
WHERE
    wd.EmployeeType = 'Regular'
    AND wd.EmploymentType IN ('Full time', 'Part time')
    AND wd.Department_Code IN ('1000', '1010', '3010')
    AND wd.HireDate <= ?
    AND (wd.TerminationDate IS NULL OR wd.TerminationDate >= ?)
    AND wd.LegacyID LIKE '[0-9]%'
    AND o.OFFC_CODE <> 'COR'