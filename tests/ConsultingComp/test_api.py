import sys, os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

from cornerstone_automation.utils.api_utils import get_request

LOC_API_URL = "https://appstaging.cornerstone.com/CompWebAPI/api/comp/GetAllLocations"
BASE_SALARIES_API_URL = "https://appstaging.cornerstone.com/CompWebAPI/api/comp/GetAllBaseSalaries"
CONFIG_API_URL = "https://appstaging.cornerstone.com/CompWebAPI/api/comp/GetAllConfigurations"

################################################## Location API related Tests ##################################################

def test_get_all_locations():
    status, data = get_request(LOC_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    # Check structure of each location
    for location in data:
        assert isinstance(location, dict)
        assert "locationID" in location
        assert "countryCode" in location
        assert "name" in location
        assert "ampMain" in location
        assert isinstance(location["locationID"], int)
        assert isinstance(location["countryCode"], str)
        assert isinstance(location["name"], str)
        assert isinstance(location["ampMain"], list)

def test_location_ids_unique():
    status, data = get_request(LOC_API_URL)
    assert status == 200
    location_ids = [loc["locationID"] for loc in data]
    assert len(location_ids) == len(set(location_ids))

def test_country_codes_present():
    status, data = get_request(LOC_API_URL)
    assert status == 200
    country_codes = {loc["countryCode"] for loc in data}
    expected_codes = {"BE", "UK", "US"}
    assert expected_codes.issubset(country_codes)

################################################## Base Salaries API related Tests ##################################################

def test_get_all_base_salaries():
    status, data = get_request(BASE_SALARIES_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    # Check structure of each base salary entry
    for entry in data:
        assert isinstance(entry, dict)
        assert "compStructureID" in entry
        assert "jobLevel" in entry
        assert "targetBonus" in entry
        assert "salary" in entry
        assert "prpYear" in entry
        assert "location" in entry
        assert "jobOrder" in entry
        assert isinstance(entry["compStructureID"], int)
        assert isinstance(entry["jobLevel"], str)
        assert isinstance(entry["targetBonus"], (int, float))
        assert entry["finance"] is None or isinstance(entry["finance"], (int, float))
        assert entry["econ"] is None or isinstance(entry["econ"], (int, float))
        assert entry["mba"] is None or isinstance(entry["mba"], (int, float))
        assert isinstance(entry["salary"], (int, float))
        assert isinstance(entry["prpYear"], int)
        assert isinstance(entry["location"], str)
        assert isinstance(entry["jobOrder"], int)
        # createdBy, dateCreated, modifiedBy, dateModified can be None or str/int/datetime, so not strictly checked

def test_comp_structure_ids_unique():
    status, data = get_request(BASE_SALARIES_API_URL)
    assert status == 200
    comp_structure_ids = [entry["compStructureID"] for entry in data]
    assert len(comp_structure_ids) == len(set(comp_structure_ids))

def test_job_levels_present():
    status, data = get_request(BASE_SALARIES_API_URL)
    assert status == 200
    job_levels = {entry["jobLevel"] for entry in data}
    expected_levels = {"P7"}  # Add more if needed
    assert expected_levels.issubset(job_levels)

 ################################################## Configurations API related Tests ##################################################

def test_get_all_configurations():
    status, data = get_request(CONFIG_API_URL)
    assert status == 200
    assert isinstance(data, list)
    assert len(data) > 0

    # Expected compType and compTitle values from your sample response
    expected_comp_types = {
        "CompEditable",
        "CompanyPerformance1H",
        "CompanyPerformance2H",
        "AssociatesBonusCutoffdate",
        "AnalystsCutoffdate",
        "AssociatesCutoffdate",
        "PRPYear",
        "Workdays"
    }
    expected_comp_titles = {
        "Comp is Editable",
        "Company Performance 1H 2025",
        "Company Performance 2H 2024",
        "Departed Employees Bonus Cutoff Date",
        "Departed Employees Bonus Cutoff Date (Analysts)",
        "Departed Employees Cutoff Date (Associates)",
        "PRP Year (07/01/2024-06/30/2025)",
        "Workdays in PRP Year 2024"
    }

    comp_types = {entry["compType"] for entry in data}
    comp_titles = {entry["compTitle"] for entry in data}

    # Assert all expected compTypes and compTitles are present
    assert expected_comp_types.issubset(comp_types)
    assert expected_comp_titles.issubset(comp_titles)

    # Assert compValue is not None for all entries
    for entry in data:
        assert entry["compValue"] is not None, f"compValue is None for compType: {entry.get('compType')}"