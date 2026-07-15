"""
SafetyCulture (iAuditor) API client.
Docs: https://developer.safetyculture.com/reference
"""

import os
import time
import requests
from datetime import datetime, timezone

BASE_URL = "https://api.safetyculture.io"


class SCClient:
    def __init__(self, token: str = None):
        self.token = token or os.environ["SC_API_TOKEN"]
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        })

    def _get(self, path: str, params: dict = None) -> dict:
        resp = self.session.get(f"{BASE_URL}{path}", params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    # ── Sites ────────────────────────────────────────────────────────────────

    def list_sites(self) -> list[dict]:
        """Return all sites in the org."""
        results = []
        after = None
        while True:
            params = {"limit": 100}
            if after:
                params["after"] = after
            data = self._get("/sites/v1/sites", params)
            results.extend(data.get("sites", []))
            if not data.get("next_page_token"):
                break
            after = data["next_page_token"]
        return results

    def get_site_map(self) -> dict[str, str]:
        """Return {site_name_lower: site_id} for fuzzy matching."""
        sites = self.list_sites()
        return {s["name"].lower(): s["id"] for s in sites}

    # ── Templates ────────────────────────────────────────────────────────────

    def list_templates(self) -> list[dict]:
        results = []
        after = None
        while True:
            params = {"limit": 100}
            if after:
                params["after"] = after
            data = self._get("/templates/v1/templates", params)
            results.extend(data.get("templates", []))
            if not data.get("next_page_token"):
                break
            after = data["next_page_token"]
        return results

    # ── Audits ───────────────────────────────────────────────────────────────

    def search_audits(
        self,
        date_from: datetime,
        date_to: datetime,
        template_ids: list[str] = None,
        site_ids: list[str] = None,
    ) -> list[dict]:
        """
        Return all audits (started or modified) within the date window.
        date_from / date_to are timezone-aware datetimes.
        """
        params = {
            "field": [
                "audit_id", "name", "date_started", "date_completed",
                "date_modified", "score", "total_score", "score_percentage",
                "duration", "site", "template_id", "template_name",
                "created_by", "owner_name", "status",
            ],
            "modified_after": date_from.isoformat(),
            "modified_before": date_to.isoformat(),
            "limit": 100,
        }
        if template_ids:
            params["template_id"] = template_ids
        if site_ids:
            params["site_id"] = site_ids

        results = []
        after = None
        while True:
            if after:
                params["after"] = after
            data = self._get("/audits/search", params)
            audits = data.get("audits", [])
            results.extend(audits)
            if not data.get("next_page_token") or not audits:
                break
            after = data["next_page_token"]
            time.sleep(0.2)  # be polite
        return results

    def get_audit(self, audit_id: str) -> dict:
        return self._get(f"/audits/{audit_id}")
