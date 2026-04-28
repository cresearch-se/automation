import sys, os
from datetime import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "src")))

from cornerstone_automation.utils.api_utils import get_request
from cornerstone_automation.consultingcomp.pojos.employeeDetails import EmployeeDetails

EMPLOYEES_API_URL = "https://appstaging.cornerstone.com/CompWebAPI/api/comp/GetAllConsultantDetails/2025-01-01"


def test_all_employees_have_name():
    status, data = get_request(EMPLOYEES_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    employees = [EmployeeDetails(emp) for emp in data]
    for emp in employees:
        assert emp.employeeName is not None and emp.employeeName.strip() != "", f"Employee name missing for id: {emp.id}"

# Test case: https://cornerstoneresearch.atlassian.net/browse/QA-156
def test_perf_rating_rules():
    status, data = get_request(EMPLOYEES_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    employees = [EmployeeDetails(emp) for emp in data]
    failures = []

    for emp in employees:
        if emp.startDate:
            start_date = datetime.strptime(emp.startDate[:10], "%Y-%m-%d")
            perf_2h = emp.priorYear2HPerfRating
            final_2025 = emp.finalCYPerfRating

            # If Start date >= 01/01/2025 then 2H 2024 Perf rating = 0
            if start_date >= datetime(2025, 1, 1):
                if perf_2h != 0:
                    failures.append(
                        f"2H 2024 Perf rating should be 0 for {emp.employeeName} (startDate: {emp.startDate}, got: {perf_2h})"
                    )

            # If 01/07/2024 <= Start date <= 12/31/2024
            elif datetime(2024, 7, 1) <= start_date <= datetime(2024, 12, 31):
                # If Final 2025 Perf rating = 1, 1.25, 2.5, or 2.25 then 2H 2024 Perf rating = Final 2025 Perf rating
                if final_2025 in {1, 1.25, 2.5, 2.25}:
                    if perf_2h != final_2025:
                        failures.append(
                            f"2H 2024 Perf rating ({perf_2h}) should match Final 2025 Perf rating ({final_2025}) for {emp.employeeName} (startDate: {emp.startDate})"
                        )
                # If Final 2025 Perf rating = 1.5, 1.75, or 2.0 then 2H 2024 Perf rating defaults to 2.0
                elif final_2025 in {1.5, 1.75, 2.0}:
                    if perf_2h != 2.0:
                        failures.append(
                            f"2H 2024 Perf rating ({perf_2h}) should default to 2.0 for {emp.employeeName} (startDate: {emp.startDate}, Final 2025 Perf rating: {final_2025})"
                        )

            # For startdate before June 30, 2024 then 2H 2024 Perf rating = Final 2025 Perf rating
            elif start_date < datetime(2024, 6, 30):
                if perf_2h != final_2025:
                    failures.append(
                        f"2H 2024 Perf rating ({perf_2h}) should match Final 2025 Perf rating ({final_2025}) for {emp.employeeName} (startDate: {emp.startDate})"
                    )

    assert not failures, "Failures:\n" + "\n".join(failures)

# Test case: https://cornerstoneresearch.atlassian.net/browse/QA-153
def test_job_level_offcycle_rules():
    status, data = get_request(EMPLOYEES_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    employees = [EmployeeDetails(emp) for emp in data]
    failures = []

    for emp in employees:
        # Only check if both job levels are present
        if emp.cyJanJobLevel is not None and emp.priorYearJulyJobLevel is not None:
            if emp.cyJanJobLevel == emp.priorYearJulyJobLevel:
                if emp.offCycle is not False:
                    failures.append(
                        f"offCycle should be False when cyJanJobLevel == priorYearJulyJobLevel for {emp.employeeName} (id: {emp.id}, cyJanJobLevel: {emp.cyJanJobLevel}, priorYearJulyJobLevel: {emp.priorYearJulyJobLevel}, offCycle: {emp.offCycle})"
                    )
            else:
                if emp.offCycle is not True:
                    failures.append(
                        f"offCycle should be True when cyJanJobLevel != priorYearJulyJobLevel for {emp.employeeName} (id: {emp.id}, cyJanJobLevel: {emp.cyJanJobLevel}, priorYearJulyJobLevel: {emp.priorYearJulyJobLevel}, offCycle: {emp.offCycle})"
                    )

    assert not failures, "Job level/offCycle rule failures:\n" + "\n".join(failures)
    