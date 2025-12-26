"""Configuration management for Distribution List Manager."""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class Config:
    """Application configuration from environment variables."""

    TENANT_ID: str = os.getenv("AZURE_TENANT_ID", "")
    CLIENT_ID: str = os.getenv("AZURE_CLIENT_ID", "")
    CLIENT_SECRET: str = os.getenv("AZURE_CLIENT_SECRET", "")

    # Microsoft Graph API endpoints
    GRAPH_BASE_URL: str = "https://graph.microsoft.com/v1.0"
    AUTHORITY: str = f"https://login.microsoftonline.com/{TENANT_ID}"
    SCOPE: list = ["https://graph.microsoft.com/.default"]

    # Export directory
    EXPORT_DIR: Path = Path("exports")

    @classmethod
    def validate(cls) -> bool:
        """Validate that all required configuration is present."""
        missing = []
        if not cls.TENANT_ID:
            missing.append("AZURE_TENANT_ID")
        if not cls.CLIENT_ID:
            missing.append("AZURE_CLIENT_ID")
        if not cls.CLIENT_SECRET:
            missing.append("AZURE_CLIENT_SECRET")

        if missing:
            raise ValueError(
                f"Missing required environment variables: {', '.join(missing)}\n"
                "Please copy .env.example to .env and fill in your Azure AD values."
            )
        return True
