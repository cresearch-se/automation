-- query: employee_target_vs_actual_hours
SELECT SUM(VA_Target) AS Target_Hrs
     , SUM(Base_hrs)  AS Actual_Hrs
FROM   TargetDaily
WHERE  Empl_uno = :empno
  AND  Period BETWEEN :period_start AND :period_end
