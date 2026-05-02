-- query: get_empl_uno_by_employee_code
SELECT Empl_uno
FROM   cmsopen..hbm_persnl
WHERE  Employee_code = :employee_code

-- query: employee_details_by_name
SELECT *
FROM   cmsopen..hbm_persnl
WHERE  Employee_name LIKE :employee_name

-- query: employee_details_by_empno
SELECT *
FROM   cmsopen..tbm_persnl
WHERE  Empl_uno = :empno
