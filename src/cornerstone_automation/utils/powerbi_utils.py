"""Power BI paginated (RDL) report export utilities.

Exports a Power BI *paginated* report to a file (e.g. XLSX) via the Power BI
REST API, authenticating as a service principal with the OAuth2
client-credentials flow. No browser or UI automation needed.

Credentials are read from ``config/creds/powerbi.env`` (gitignored). A template
with the required keys lives at ``config/powerbi.env.example``.

Required env keys (auth only):
    PBI_TENANT_ID, PBI_CLIENT_ID, PBI_CLIENT_SECRET

The report to export is identified by its Power BI URL (passed in by the caller),
not by env vars -- so this module is reusable for any workspace report.

Only ``requests`` is needed (already a project dependency) -- the
client-credentials token is a plain OAuth POST, so ``msal`` is not required.
"""

import os
import re
import time

import requests
from dotenv import load_dotenv

# config/creds/powerbi.env  (this file is src/cornerstone_automation/utils/powerbi_utils.py)
_CREDS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
    "config", "creds",
)
load_dotenv(dotenv_path=os.path.join(_CREDS_DIR, "powerbi.env"))

_AUTHORITY = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
_PBI_API = "https://api.powerbi.com/v1.0/myorg"


def _require(name):
    """Return env var ``name`` or raise a clear error pointing at the config file."""
    value = os.getenv(name)
    if not value:
        raise ValueError(
            f"{name} is not set. Add it to config/creds/powerbi.env "
            f"(see config/powerbi.env.example for the template)."
        )
    return value


def parse_report_url(url):
    """Extract ``(workspace_id, report_id)`` from a Power BI workspace report URL.

    Expects the workspace form::

        https://app.powerbi.com/groups/<workspace-id>/rdlreports/<report-id>?...

    Raises ValueError for app-style URLs (``groups/me/apps/...``) or anything
    that doesn't match, so unsupported links fail loudly instead of silently.
    """
    m = re.search(r"/groups/([0-9a-fA-F-]{36})/(?:rdl)?reports/([0-9a-fA-F-]{36})", url)
    if not m:
        raise ValueError(
            "Unsupported Power BI report URL. Expected a workspace report URL like "
            "'https://app.powerbi.com/groups/<workspace-id>/rdlreports/<report-id>'. "
            f"Got: {url!r}"
        )
    return m.group(1), m.group(2)


def get_access_token():
    """Acquire an Azure AD access token for the Power BI API (service principal)."""
    tenant = _require("PBI_TENANT_ID")
    resp = requests.post(
        _AUTHORITY.format(tenant=tenant),
        data={
            "grant_type": "client_credentials",
            "client_id": _require("PBI_CLIENT_ID"),
            "client_secret": _require("PBI_CLIENT_SECRET"),
            "scope": _SCOPE,
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Failed to get Power BI token ({resp.status_code}): {resp.text[:400]}"
        )
    return resp.json()["access_token"]


def export_report_to_file(
    report_url,
    out_path,
    fmt="XLSX",
    parameter_values=None,
    poll_interval=5,
    timeout=600,
    token=None,
):
    """Export a paginated report (by its Power BI URL) and write bytes to ``out_path``.

    Args:
        report_url: the Power BI workspace report URL (workspace + report IDs
            are parsed from it).
        out_path: where to save the exported file.
        fmt: export format ("XLSX", "PDF", "CSV", ...). Default "XLSX".
        parameter_values: optional list of ``{"name": ..., "value": ...}`` dicts.
            If ``None`` (default), the report's *default* parameter values are
            used -- i.e. the latest period the report's dropdown shows by
            default. So the normal monthly run needs no parameter name.
        poll_interval: seconds between status polls.
        timeout: max seconds to wait for the export to finish.
        token: reuse an existing access token; otherwise one is acquired.

    Returns:
        ``out_path`` on success.
    """
    token = token or get_access_token()
    workspace, report = parse_report_url(report_url)
    headers = {"Authorization": f"Bearer {token}"}
    base = f"{_PBI_API}/groups/{workspace}/reports/{report}"

    body = {"format": fmt}
    if parameter_values:
        body["paginatedReportConfiguration"] = {"parameterValues": parameter_values}

    # 1) kick off the export
    resp = requests.post(f"{base}/ExportTo", headers=headers, json=body, timeout=60)
    if resp.status_code not in (200, 202):
        raise RuntimeError(
            f"ExportTo failed ({resp.status_code}): {resp.text[:400]}"
        )
    export_id = resp.json()["id"]

    # 2) poll until the render finishes
    waited = 0
    status = None
    while waited < timeout:
        s = requests.get(f"{base}/exports/{export_id}", headers=headers, timeout=30)
        s.raise_for_status()
        data = s.json()
        status = data.get("status")
        if status == "Succeeded":
            break
        if status == "Failed":
            raise RuntimeError(f"Power BI export failed: {data}")
        time.sleep(poll_interval)
        waited += poll_interval
    else:
        raise TimeoutError(
            f"Export did not finish within {timeout}s (last status: {status})."
        )

    # 3) download the rendered file
    f = requests.get(f"{base}/exports/{export_id}/file", headers=headers, timeout=120)
    f.raise_for_status()
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "wb") as fh:
        fh.write(f.content)
    return out_path


def check_access(report_url):
    """Verify the credentials work: acquire a token and reach the report.

    Prints a short report and returns the token. Raises with a clear message if
    a credential is missing or the API rejects the request.
    """
    print("Checking Power BI access...")
    token = get_access_token()
    print("  [OK] Acquired access token.")

    workspace, report = parse_report_url(report_url)
    r = requests.get(
        f"{_PBI_API}/groups/{workspace}/reports/{report}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if r.status_code == 200:
        info = r.json()
        print(f"  [OK] Report reachable: {info.get('name')!r} (type={info.get('reportType')})")
    else:
        print(f"  [WARN] Report fetch returned {r.status_code}: {r.text[:300]}")
        print("        Check that the service principal is added to the workspace,")
        print("        and that the workspace has paginated-export capacity (Premium/PPU/Fabric).")
    return token


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        sys.exit("Usage: python powerbi_utils.py <power-bi-report-url>")
    check_access(sys.argv[1])
