-- query: employee_billable_hours_by_office
SELECT   hp.Employee_name
       , hp.OFFC
       , SUM(Base_hrs) AS Total_Billable_Hours
FROM     cmsopen..tat_time tt
INNER JOIN CMSOPEN..HBM_PERSNL hp ON hp.Empl_uno = tt.TK_EMPL_UNO
WHERE    TRAN_DATE BETWEEN :start_date AND :end_date
  AND    hp.Empl_uno = :empno
  AND    BILLABLE_FLAG = 'B'
GROUP BY hp.Employee_name
       , hp.OFFC
ORDER BY hp.OFFC
