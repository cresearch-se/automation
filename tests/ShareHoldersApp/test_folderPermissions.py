# -*- coding: latin-1 -*-

import os
import re
import json
import argparse
from typing import Dict, List

# Run this using: py .\test_folderPermissions.py
# run this from code folder as: py .\tests\ShareHoldersApp\test_folderPermissions.py
class FolderPermissionsParser:
    """
    Parse a text file containing folder paths and permissions in the format:
      Folder Path || Folder Permissions

    Produces a dictionary: { folder_path: [perm1, perm2, ...], ... }

    This implementation is resilient to:
    - header and summary lines like "Folder Path || Folder Permissions" or "Total folders||10"
    - files where records are on separate lines or the whole file is a single long line
    """

    _LINE_PATTERN = re.compile(r'(?m)(?P<path>[^|]+?)\s*\|\|\s*(?P<perms>[^\r\n]+)')

    def __init__(self, filepath: str, encoding: str = 'latin-1'): #utf-8
        if not os.path.isfile(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        self.filepath = filepath
        self.encoding = encoding
        self._parsed: Dict[str, List[str]] = {}

    def parse(self) -> Dict[str, List[str]]:
        """
        Parse the file and return a dictionary mapping folder paths to lists of permissions.
        Calling parse() multiple times will return the cached result.
        """
        if self._parsed:
            return self._parsed

        with open(self.filepath, 'r', encoding=self.encoding) as fh:
            content = fh.read()

        parsed: Dict[str, List[str]] = {}
        for m in self._LINE_PATTERN.finditer(content):
            path = m.group('path').strip()
            perms_raw = m.group('perms').strip()

            # skip header / summary lines
            low = path.lower()
            if low.startswith('folder path') or low.startswith('total folders'):
                continue

            perms = [p.strip() for p in perms_raw.split(';') if p.strip()]
            parsed[path] = perms

        self._parsed = parsed
        return parsed

    @classmethod
    def from_text(cls, text: str) -> Dict[str, List[str]]:
        """
        Convenience method to parse from an in-memory string (useful for tests).
        """
        parsed: Dict[str, List[str]] = {}
        for m in cls._LINE_PATTERN.finditer(text):
            path = m.group('path').strip()
            perms_raw = m.group('perms').strip()
            low = path.lower()
            if low.startswith('folder path') or low.startswith('total folders'):
                continue
            perms = [p.strip() for p in perms_raw.split(';') if p.strip()]
            parsed[path] = perms
        return parsed


# Role -> permissions mapping requested by the developer/user.
# Note: normalized "CR_HumanResources" to match the sample data format.
role_to_permissions: Dict[str, List[str]] = {
    "Analysts": [
        "CR_AssociatesA2", "CR_AssociatesAFE2", "CR_AssociatesA3", "CR_AssociatesA4",
        "CR_ManagersEco", "CR_ManagersM1", "CR_ManagersM2", "CR_ManagersM3",
        "CR_SeniorManagersM1", "CR_SeniorManagersM2", "CR_SeniorManagersM3",
        "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3", "CR_PrincipalP4",
        "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Senior Analysts": [
        "CR_AssociatesA2", "CR_AssociatesAFE2", "CR_AssociatesA3", "CR_AssociatesA4",
        "CR_ManagersEco", "CR_ManagersM1", "CR_ManagersM2", "CR_ManagersM3",
        "CR_SeniorManagersM1", "CR_SeniorManagersM2", "CR_SeniorManagersM3",
        "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3", "CR_PrincipalP4",
        "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Research Associates": [
        "CR_AssociatesA2", "CR_AssociatesAFE2", "CR_AssociatesA3", "CR_AssociatesA4",
        "CR_ManagersEco", "CR_ManagersM1", "CR_ManagersM2", "CR_ManagersM3",
        "CR_SeniorManagersM1", "CR_SeniorManagersM2", "CR_SeniorManagersM3",
        "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3", "CR_PrincipalP4",
        "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Associates": [
        "CR_SeniorManagersM2", "CR_SeniorManagersM3", "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3",
        "CR_PrincipalP4", "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Managers": [
        "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3", "CR_PrincipalP4",
        "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Sr.Managers": [
        "CR_PrincipalP1", "CR_PrincipalP2", "CR_PrincipalP3", "CR_PrincipalP4",
        "CR_PrincipalP5", "CR_PrincipalP6", "CR_PrincipalP7",
        "CR_SeniorAdvisor", "CR_VicePresident", "CR_SeniorVicePresident",
        "CR_HumanResources"
    ],
    "Principals": [
        "CR_VicePresident", "CR_SeniorVicePresident","CR_HumanResources"
    ],
    "Officers": [
        "CR_HumanResources"
    ],
    "Inactive": [
        "CR_HumanResources"
    ]
}

def test_parse_and_print_folder_permissions(role_to_permissions: Dict[str, List[str]]):
    """
    Tests the FolderPermissionsParser, including a structural check for the
    employee's name in the expected permissions list.
    """
    
    #filepath = "fixtures/PMSPermissionsReport_20251025104458.txt"
    #filepath = "fixtures/sptest-all.txt"
    filepath = "tests/ShareHoldersApp/fixtures/PRPPermissionsReport.txt"
    abs_filepath = os.path.abspath(filepath)
    error_count = 0  
    
    print(f"Testing FolderPermissionsParser with file: {abs_filepath}")

    if not os.path.exists(abs_filepath):
        raise FileNotFoundError(f"Test failed: Fixture file not found at {abs_filepath}. Cannot proceed.")

    try:
        parser = FolderPermissionsParser(abs_filepath)
        result = parser.parse()
    except Exception as e:
        raise Exception(f"Test failed: Critical parsing error occurred: {e}")
    
    print("\n--- Starting Validation ---")
    
    # A. Check the primary result type and size
    if not isinstance(result, dict):
        print("FAIL: Validation failed: The parser result must be a dictionary.")
        error_count += 1
    elif len(result) == 0:
        print("FAIL: Validation failed: The result dictionary is empty, meaning no folder paths were parsed.")
        error_count += 1
    
    # Check if extracted_name is in inactive_users
    inactive_file_path = "tests/ShareHoldersApp/fixtures/InactiveEmployees.txt"
    abs_inactive_file_path = os.path.abspath(inactive_file_path)
    if not os.path.exists(abs_inactive_file_path):
        raise FileNotFoundError(f"Inactive employees file not found at {abs_inactive_file_path}. Cannot proceed.")
    with open(abs_inactive_file_path, "r", encoding="latin-1") as f:
        inactive_names = [line.strip() for line in f if line.strip()]

    # B. Loop through ALL entries
    for folder_path, permissions_list in result.items():
        
        # Structural Asserts (omitted for space, but they should remain)

        # --- POSITION AND NAME EXTRACTION ---
        path_parts = folder_path.split('/')
        
        if len(path_parts) >= 4:
            extracted_position = path_parts[2].strip() 
            raw_name_component = path_parts[3].strip()
            extracted_name = raw_name_component.split('_')[0].strip()

           
            if extracted_name in inactive_names:
                role_to_check = "Inactive"
            else:
                role_to_check = extracted_position
            
            # --- FALLBACK LOGIC ---
            if role_to_check in role_to_permissions:
                perms_expected_base = role_to_permissions[role_to_check]
                active_role = role_to_check
            
            elif 'Analysts' in role_to_permissions:
                perms_expected_base = role_to_permissions['Analysts']
                active_role = f"{role_to_check} (FALLBACK to Analysts)"
                print(
                    f"WARNING: Role '{role_to_check}' not found in map. "
                    f"Using permissions for 'Analysts' for path: {folder_path}"
                )
            
            else:
                print(
                    f"WARNING: Role '{role_to_check}' not found and 'Analysts' fallback role is missing. "
                    f"Skipping permission check for path: {folder_path}"
                )
                continue
                
            # --- NEW LOGIC: ADD NAME TO EXPECTED PERMISSIONS ---
            # Create a new list/copy to safely add the person's name without modifying the global map
            perms_expected = list(perms_expected_base)
            if extracted_name not in inactive_names:
                perms_expected.append(extracted_name) 

            # --- PERMISSION COMPARISON ---
            set_expected = set(perms_expected)
            set_actual = set(permissions_list)

            # 2.1. Check for permissions MISSING FROM THE FILE (Under-permissioned)
            missing_in_file = list(set_expected - set_actual)

            if missing_in_file:
                #if extracted_name in missing_in_file:
                if any(extracted_name in perm for perm in missing_in_file):
                    '''
                    print(
                        f"FAIL: MISSING NAME for role '{active_role}' in path '{folder_path}' (Name: {extracted_name}): "
                        f"File is missing the expected name permission: {extracted_name}"
                    )
                    '''
                else:
                    print(
                        f"FAIL: FILE is missing permissions for role '{active_role}' in path '{folder_path}' (Name: {extracted_name}): "
                        f"File is missing expected permissions: {missing_in_file}"
                    )
                error_count += 1


            
            # 2.2. Check for permissions MISSING FROM THE MAP (Over-permissioned/Unexpected)
            missing_in_map = list(set_actual - set_expected)
            if missing_in_map:
    
                # Check if the extracted name is a substring of ANY missing permission.
                # This covers cases where 'extracted_name' is, say, "GroupA" and a missing
                # permission ('perm') is "GroupA_FullControl" or "GroupNameA".
                if any(extracted_name in perm for perm in missing_in_map):
        
                    # Scenario: Missing permission CONTAINS the extracted name.
                    # Action: Print "missing name permission"
                    '''
                    print(
                        f"FAIL: MISSING NAME PERMISSION for role '{active_role}' in path '{folder_path}' (Name: {extracted_name}): "
                        f"File contains unexpected name permission(s) containing the extracted name: "
                        f"{[perm for perm in missing_in_map if extracted_name in perm]}"
                    )
                    '''
                else:
                    # Scenario: Missing permission does NOT contain the extracted name.
                    # Action: Print "miss as some permissions" (e.g., completely unrelated permissions are missing)
                    print(
                        f"FAIL: MAP is missing permissions for role '{active_role}' in path '{folder_path}' (Name: {extracted_name}): "
                        f"Map contains unexpected permissions: {missing_in_map}"
                    )
        
                error_count += 1
             
       
        else:
            print(f"WARNING: Folder path '{folder_path}' does not match the expected structure. Skipping role check.")


    # 4. Final Assert: Fail the test if ANY errors were counted
    print("\n--- Finished Validation ---")
    if error_count > 0:
        assert error_count == 0, f"Validation failed: {error_count} error(s) found. See details above."
    else:
        print("Validation Successful! No errors found.")

# --- Execution Entry Point ---
if __name__ == "__main__":
    test_parse_and_print_folder_permissions(role_to_permissions)

