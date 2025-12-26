"""Exchange Online PowerShell client for Distribution List management."""

import subprocess
import json
import logging
import os
from typing import Optional
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger(__name__)

# Get config
from config import Config


@dataclass
class ExchangeDistributionList:
    """Represents an Exchange distribution list."""
    identity: str
    display_name: str
    primary_smtp: str
    member_count: int = 0


@dataclass
class ExchangeMember:
    """Represents a member of an Exchange distribution list."""
    name: str
    email: str


class ExchangeClient:
    """Client for Exchange Online PowerShell operations with app-only auth."""

    def __init__(self):
        self._connected = False
        self._cert_thumbprint = os.getenv("EXCHANGE_CERT_THUMBPRINT", "")
        self._organization = os.getenv("EXCHANGE_ORGANIZATION", "")  # e.g., contoso.onmicrosoft.com

    def _run_powershell(self, command: str, connect: bool = True) -> str:
        """Run a PowerShell command with Exchange connection."""

        # Build connection command if needed
        if connect and self._cert_thumbprint and self._organization:
            connect_cmd = (
                f"Connect-ExchangeOnline -AppId '{Config.CLIENT_ID}' "
                f"-CertificateThumbprint '{self._cert_thumbprint}' "
                f"-Organization '{self._organization}' "
                f"-ShowBanner:$false -ErrorAction Stop"
            )
        else:
            connect_cmd = ""

        script = f'''
$ErrorActionPreference = "Stop"
Import-Module ExchangeOnlineManagement -ErrorAction Stop
{connect_cmd}
{command}
if ($?) {{ Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }}
'''
        logger.info(f"Running Exchange command...")
        logger.debug(f"Command: {command[:100]}...")

        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script],
            capture_output=True,
            text=True
        )

        if result.returncode != 0 or "error" in result.stderr.lower():
            error_msg = result.stderr.strip() or result.stdout.strip()

            # Check for specific errors
            if "not installed" in error_msg.lower() or "not recognized" in error_msg.lower():
                raise RuntimeError(
                    "ExchangeOnlineManagement module not installed.\n"
                    "Run in PowerShell as Admin: Install-Module ExchangeOnlineManagement -Force"
                )
            if "certificate" in error_msg.lower() or "thumbprint" in error_msg.lower():
                raise RuntimeError(
                    "Certificate authentication not configured.\n"
                    "See setup instructions in the app."
                )

            logger.error(f"PowerShell error: {error_msg}")
            raise RuntimeError(error_msg or "PowerShell command failed")

        return result.stdout

    def check_module_installed(self) -> bool:
        """Check if ExchangeOnlineManagement is installed."""
        try:
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command',
                 'Get-Module -ListAvailable ExchangeOnlineManagement | Select-Object -First 1'],
                capture_output=True,
                text=True
            )
            return bool(result.stdout.strip())
        except:
            return False

    def connect(self, upn: Optional[str] = None):
        """Connect to Exchange Online. Will prompt for login if needed."""
        logger.info("Connecting to Exchange Online...")

        if upn:
            cmd = f"Connect-ExchangeOnline -UserPrincipalName '{upn}' -ShowBanner:$false"
        else:
            cmd = "Connect-ExchangeOnline -ShowBanner:$false"

        try:
            self._run_powershell(cmd)
            logger.info("Connected to Exchange Online")
        except RuntimeError as e:
            if "Please call Connect-ExchangeOnline" in str(e):
                raise RuntimeError("Not connected. Please run Connect-ExchangeOnline manually first.")
            raise

    def disconnect(self):
        """Disconnect from Exchange Online."""
        self._run_powershell("Disconnect-ExchangeOnline -Confirm:$false")
        logger.info("Disconnected from Exchange Online")

    def list_distribution_groups(self) -> list[ExchangeDistributionList]:
        """Get all distribution groups."""
        logger.info("Fetching distribution groups...")

        cmd = (
            "Get-DistributionGroup -ResultSize Unlimited | "
            "Select-Object Identity, DisplayName, PrimarySmtpAddress | "
            "ConvertTo-Json -Compress"
        )

        output = self._run_powershell(cmd)
        if not output.strip():
            return []

        data = json.loads(output)
        # Handle single result (not a list)
        if isinstance(data, dict):
            data = [data]

        groups = []
        for item in data:
            groups.append(ExchangeDistributionList(
                identity=item.get("Identity", ""),
                display_name=item.get("DisplayName", ""),
                primary_smtp=item.get("PrimarySmtpAddress", ""),
            ))

        logger.info(f"Found {len(groups)} distribution groups")
        return groups

    def get_members(self, identity: str) -> list[ExchangeMember]:
        """Get members of a distribution group."""
        logger.info(f"Fetching members of: {identity}")

        cmd = (
            f"Get-DistributionGroupMember -Identity '{identity}' -ResultSize Unlimited | "
            "Select-Object Name, PrimarySmtpAddress | "
            "ConvertTo-Json -Compress"
        )

        output = self._run_powershell(cmd)
        if not output.strip():
            return []

        data = json.loads(output)
        if isinstance(data, dict):
            data = [data]

        members = []
        for item in data:
            members.append(ExchangeMember(
                name=item.get("Name", ""),
                email=item.get("PrimarySmtpAddress", ""),
            ))

        logger.info(f"Found {len(members)} members")
        return members

    def add_member(self, identity: str, member_email: str) -> bool:
        """Add a member to a distribution group."""
        logger.info(f"Adding {member_email} to {identity}")

        cmd = f"Add-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
        self._run_powershell(cmd)

        logger.info(f"Successfully added {member_email}")
        return True

    def remove_member(self, identity: str, member_email: str) -> bool:
        """Remove a member from a distribution group."""
        logger.info(f"Removing {member_email} from {identity}")

        cmd = f"Remove-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
        self._run_powershell(cmd)

        logger.info(f"Successfully removed {member_email}")
        return True

    def add_members_bulk(self, identity: str, emails: list[str]) -> dict:
        """Add multiple members."""
        results = {"success": [], "failed": []}

        for email in emails:
            try:
                self.add_member(identity, email)
                results["success"].append(email)
            except Exception as e:
                results["failed"].append({"email": email, "error": str(e)})

        return results

    def remove_members_bulk(self, identity: str, emails: list[str]) -> dict:
        """Remove multiple members."""
        results = {"success": [], "failed": []}

        for email in emails:
            try:
                self.remove_member(identity, email)
                results["success"].append(email)
            except Exception as e:
                results["failed"].append({"email": email, "error": str(e)})

        return results
