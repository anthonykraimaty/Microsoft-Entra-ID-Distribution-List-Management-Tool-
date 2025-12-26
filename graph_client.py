"""Microsoft Graph API client for Entra ID operations."""

import msal
import requests
import logging
from datetime import datetime
from typing import Optional
from config import Config

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)


class GraphClient:
    """Client for Microsoft Graph API operations."""

    def __init__(self):
        logger.info("Initializing Microsoft Graph client...")
        Config.validate()
        logger.info(f"Tenant ID: {Config.TENANT_ID[:8]}...{Config.TENANT_ID[-4:]}")
        logger.info(f"Client ID: {Config.CLIENT_ID[:8]}...{Config.CLIENT_ID[-4:]}")

        self._token: Optional[str] = None
        self._app = msal.ConfidentialClientApplication(
            Config.CLIENT_ID,
            authority=Config.AUTHORITY,
            client_credential=Config.CLIENT_SECRET,
        )
        logger.info("MSAL application created")

    def _get_token(self) -> str:
        """Acquire access token for Microsoft Graph API."""
        logger.info("Acquiring access token from Azure AD...")
        result = self._app.acquire_token_for_client(scopes=Config.SCOPE)

        if "access_token" in result:
            self._token = result["access_token"]
            logger.info("Access token acquired successfully")
            return self._token
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Failed to acquire token: {error}")
            raise Exception(f"Failed to acquire token: {error}")

    @property
    def token(self) -> str:
        """Get current token or acquire a new one."""
        if not self._token:
            self._get_token()
        return self._token

    @property
    def headers(self) -> dict:
        """Get headers for API requests."""
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

    def get(self, endpoint: str, params: Optional[dict] = None) -> dict:
        """Make GET request to Graph API."""
        url = f"{Config.GRAPH_BASE_URL}{endpoint}"
        logger.debug(f"GET {endpoint}")
        response = requests.get(url, headers=self.headers, params=params)

        if response.status_code == 401:
            logger.warning("Token expired, refreshing...")
            self._get_token()
            response = requests.get(url, headers=self.headers, params=params)

        if not response.ok:
            logger.error(f"GET {endpoint} failed: {response.status_code} - {response.text}")
        response.raise_for_status()
        return response.json()

    def post(self, endpoint: str, data: dict) -> dict:
        """Make POST request to Graph API."""
        url = f"{Config.GRAPH_BASE_URL}{endpoint}"
        logger.debug(f"POST {endpoint}")
        response = requests.post(url, headers=self.headers, json=data)

        if response.status_code == 401:
            logger.warning("Token expired, refreshing...")
            self._get_token()
            response = requests.post(url, headers=self.headers, json=data)

        if not response.ok:
            error_text = response.text
            logger.error(f"POST {endpoint} failed: {response.status_code} - {error_text}")
            raise Exception(f"Graph API Error {response.status_code}: {error_text}")
        return response.json() if response.content else {}

    def patch(self, endpoint: str, data: dict) -> dict:
        """Make PATCH request to Graph API."""
        url = f"{Config.GRAPH_BASE_URL}{endpoint}"
        logger.debug(f"PATCH {endpoint}")
        response = requests.patch(url, headers=self.headers, json=data)

        if response.status_code == 401:
            logger.warning("Token expired, refreshing...")
            self._get_token()
            response = requests.patch(url, headers=self.headers, json=data)

        if not response.ok:
            logger.error(f"PATCH {endpoint} failed: {response.status_code} - {response.text}")
        response.raise_for_status()
        return response.json() if response.content else {}

    def delete(self, endpoint: str) -> bool:
        """Make DELETE request to Graph API."""
        url = f"{Config.GRAPH_BASE_URL}{endpoint}"
        logger.debug(f"DELETE {endpoint}")
        response = requests.delete(url, headers=self.headers)

        if response.status_code == 401:
            logger.warning("Token expired, refreshing...")
            self._get_token()
            response = requests.delete(url, headers=self.headers)

        if not response.ok:
            error_text = response.text
            logger.error(f"DELETE {endpoint} failed: {response.status_code} - {error_text}")
            # Raise with full error message for proper detection
            raise Exception(f"Graph API Error {response.status_code}: {error_text}")
        return True

    def get_all_pages(self, endpoint: str, params: Optional[dict] = None) -> list:
        """Get all pages of results from a paginated endpoint."""
        logger.info(f"Fetching: {endpoint}")
        results = []
        url = f"{Config.GRAPH_BASE_URL}{endpoint}"
        page = 1

        while url:
            if url.startswith(Config.GRAPH_BASE_URL):
                response = requests.get(url, headers=self.headers, params=params)
            else:
                response = requests.get(url, headers=self.headers)

            if response.status_code == 401:
                logger.warning("Token expired, refreshing...")
                self._get_token()
                continue

            if not response.ok:
                logger.error(f"Request failed: {response.status_code} - {response.text}")
            response.raise_for_status()
            data = response.json()
            items = data.get("value", [])
            results.extend(items)
            logger.debug(f"Page {page}: {len(items)} items")
            url = data.get("@odata.nextLink")
            params = None  # params only for first request
            page += 1

        logger.info(f"Fetched {len(results)} total items")
        return results
