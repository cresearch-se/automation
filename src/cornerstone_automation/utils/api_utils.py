import os, requests
from requests_ntlm import HttpNtlmAuth
from dotenv import load_dotenv

# Load environment variables from config/api.env
env_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'config', 'api.env')
load_dotenv(env_path)

USERNAME = os.getenv("API_USERNAME", "")
PASSWORD = os.getenv("API_PASSWORD", "")

#USERNAME = 'sowjikarumuri'
#PASSWORD = 'Thankgoodness$123'

# Create NTLM auth object if credentials are provided
auth = HttpNtlmAuth(USERNAME, PASSWORD) if USERNAME and PASSWORD else None

def get_request(url: str, headers: dict = None):
    """Perform a GET request with optional headers and NTLM authentication."""
    try:
        response = requests.get(url, headers=headers, auth=auth, timeout=10, verify=False)
        response.raise_for_status()
        # print(f"[get_request] Response Status Code: {response.status_code}")
        # print(f"[get_request] Response Content: {response.text}")
        return response.status_code, response.json()
    except requests.exceptions.Timeout as e:
        print(f"[get_request] Timeout error: {e}")
        return 408, {"error": "Request timed out"}
    except requests.exceptions.RequestException as e:
        print(f"[get_request] RequestException: {e}")
        return getattr(e.response, "status_code", 500), {"error": str(e)}
    except ValueError as e:
        print(f"[get_request] ValueError: {e}")
        return response.status_code, {"error": "Invalid JSON response"}

def post_request(url: str, data: dict, headers: dict = None):
    """Perform a POST request with JSON body, optional headers, and NTLM authentication."""
    try:
        response = requests.post(url, json=data, headers=headers, auth=auth, timeout=10, verify=False)
        response.raise_for_status()
        return response.status_code, response.json()
    except requests.exceptions.Timeout as e:
        print(f"[post_request] Timeout error: {e}")
        return 408, {"error": "Request timed out"}
    except requests.exceptions.RequestException as e:
        print(f"[post_request] RequestException: {e}")
        return getattr(e.response, "status_code", 500), {"error": str(e)}
    except ValueError as e:
        print(f"[post_request] ValueError: {e}")
        return response.status_code, {"error": "Invalid JSON response"}

def put_request(url: str, data: dict, headers: dict = None):
    """Perform a PUT request with optional headers and NTLM authentication."""
    try:
        response = requests.put(url, json=data, headers=headers, auth=auth, timeout=10, verify=False)
        response.raise_for_status()
        return response.status_code, response.json()
    except requests.exceptions.Timeout as e:
        print(f"[put_request] Timeout error: {e}")
        return 408, {"error": "Request timed out"}
    except requests.exceptions.RequestException as e:
        print(f"[put_request] RequestException: {e}")
        return getattr(e.response, "status_code", 500), {"error": str(e)}
    except ValueError as e:
        print(f"[put_request] ValueError: {e}")
        return response.status_code, {"error": "Invalid JSON response"}

def delete_request(url: str, headers: dict = None):
    """Perform a DELETE request with optional headers and NTLM authentication."""
    try:
        response = requests.delete(url, headers=headers, auth=auth, timeout=10, verify=False)
        response.raise_for_status()
        try:
            return response.status_code, response.json()
        except ValueError as e:
            print(f"[delete_request] ValueError: {e}")
            return response.status_code, {"message": "No content"}
    except requests.exceptions.Timeout as e:
        print(f"[delete_request] Timeout error: {e}")
        return 408, {"error": "Request timed out"}
    except requests.exceptions.RequestException as e:
        print(f"[delete_request] RequestException: {e}")
        return getattr(e.response, "status_code", 500), {"error": str(e)}