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
        """Add a member to a distribution group. Creates mail contact for external users if needed."""
        logger.info(f"Adding {member_email} to {identity}")

        # Try to add directly first
        cmd = f"Add-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
        try:
            self._run_powershell(cmd)
            logger.info(f"Successfully added {member_email}")
            return True
        except RuntimeError as e:
            error_str = str(e)
            # If user not found, try creating a mail contact first
            if "Couldn't find object" in error_str or "couldn't be found" in error_str.lower():
                logger.info(f"User not found, creating mail contact for {member_email}...")
                return self._add_external_member(identity, member_email)
            raise

    def _add_external_member(self, identity: str, member_email: str) -> bool:
        """Create a mail contact for external user and add to distribution group.
        First checks if it's an existing distribution group or recipient."""

        # First, check if this is an existing distribution group
        check_dg_cmd = f"Get-DistributionGroup -Identity '{member_email}' -ErrorAction SilentlyContinue"
        try:
            result = self._run_powershell(check_dg_cmd)
            if result.strip():
                # It's a distribution group - add it directly by its identity
                logger.info(f"{member_email} is a distribution group, adding directly")
                add_cmd = f"Add-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
                self._run_powershell(add_cmd)
                logger.info(f"Successfully added distribution group {member_email}")
                return True
        except RuntimeError:
            pass  # Not a distribution group, continue to check other types

        # Check if it's any existing recipient (mailbox, mail user, etc.)
        check_recipient_cmd = f"Get-Recipient -Identity '{member_email}' -ErrorAction SilentlyContinue"
        try:
            result = self._run_powershell(check_recipient_cmd)
            if result.strip():
                # Found as a recipient - might need time to sync, try adding
                logger.info(f"{member_email} found as recipient, attempting to add")
                add_cmd = f"Add-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
                self._run_powershell(add_cmd)
                logger.info(f"Successfully added recipient {member_email}")
                return True
        except RuntimeError:
            pass  # Not found as recipient either

        # Use full email as unique name to avoid conflicts
        display_name = member_email  # Use full email as name for uniqueness

        # Create alias from email, removing invalid chars and making unique
        alias = member_email.replace("@", "_at_").replace(".", "_")[:64]

        # Check if contact already exists by external email address
        check_cmd = f"Get-MailContact -Filter \"ExternalEmailAddress -eq '{member_email}'\" -ErrorAction SilentlyContinue"
        try:
            result = self._run_powershell(check_cmd)
            if result.strip():
                # Contact already exists, just add to group
                logger.info(f"Mail contact already exists for {member_email}")
            else:
                # Contact doesn't exist, create it
                create_cmd = (
                    f"New-MailContact -Name '{display_name}' -ExternalEmailAddress '{member_email}' "
                    f"-Alias '{alias}' -ErrorAction Stop"
                )
                self._run_powershell(create_cmd)
                logger.info(f"Created mail contact for {member_email}")
        except RuntimeError as e:
            # If check failed, try to create anyway
            if "already exists" not in str(e).lower():
                try:
                    create_cmd = (
                        f"New-MailContact -Name '{display_name}' -ExternalEmailAddress '{member_email}' "
                        f"-Alias '{alias}' -ErrorAction Stop"
                    )
                    self._run_powershell(create_cmd)
                    logger.info(f"Created mail contact for {member_email}")
                except RuntimeError as create_error:
                    if "already exists" not in str(create_error).lower():
                        raise

        # Now add the contact to the distribution group
        add_cmd = f"Add-DistributionGroupMember -Identity '{identity}' -Member '{member_email}' -Confirm:$false"
        self._run_powershell(add_cmd)
        logger.info(f"Successfully added external member {member_email}")
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

    def create_distribution_group(self, name: str, alias: str, primary_smtp: str) -> bool:
        """Create a new distribution group."""
        logger.info(f"Creating distribution group: {name} ({primary_smtp})")

        cmd = (
            f"New-DistributionGroup -Name '{name}' -Alias '{alias}' "
            f"-PrimarySmtpAddress '{primary_smtp}' -Type 'Distribution' "
            f"-MemberDepartRestriction 'Closed' -MemberJoinRestriction 'Closed'"
        )
        self._run_powershell(cmd)

        logger.info(f"Successfully created distribution group: {primary_smtp}")
        return True

    def update_distribution_group(
        self,
        identity: str,
        display_name: str = None,
        primary_smtp: str = None,
        alias: str = None
    ) -> bool:
        """Update distribution group properties via Exchange PowerShell."""
        logger.info(f"Updating distribution group: {identity}")

        # If changing email, check for conflicting mail contacts and remove them
        if primary_smtp:
            self._remove_conflicting_contact(primary_smtp)

        # Build Set-DistributionGroup command with only provided parameters
        params = []
        if display_name:
            params.append(f"-DisplayName '{display_name}'")
        if primary_smtp:
            params.append(f"-PrimarySmtpAddress '{primary_smtp}'")
        if alias:
            params.append(f"-Alias '{alias}'")

        if not params:
            logger.info("No parameters to update")
            return True

        cmd = f"Set-DistributionGroup -Identity '{identity}' {' '.join(params)}"
        self._run_powershell(cmd)

        logger.info(f"Successfully updated distribution group: {identity}")
        return True

    def _remove_conflicting_contact(self, email: str) -> bool:
        """Remove a mail contact that conflicts with an email address."""
        logger.info(f"Checking for conflicting mail contact: {email}")

        # Check if a mail contact exists with this email
        check_cmd = f"Get-MailContact -Identity '{email}' -ErrorAction SilentlyContinue"
        try:
            result = self._run_powershell(check_cmd)
            if result.strip():
                # Mail contact exists, remove it
                logger.info(f"Found conflicting mail contact, removing: {email}")
                remove_cmd = f"Remove-MailContact -Identity '{email}' -Confirm:$false"
                self._run_powershell(remove_cmd)
                logger.info(f"Removed conflicting mail contact: {email}")
                return True
        except RuntimeError:
            pass  # No contact found or error checking

        return False

    def delete_distribution_group(self, identity: str) -> bool:
        """Delete a distribution group."""
        logger.info(f"Deleting distribution group: {identity}")

        cmd = f"Remove-DistributionGroup -Identity '{identity}' -Confirm:$false"
        self._run_powershell(cmd)

        logger.info(f"Successfully deleted distribution group: {identity}")
        return True
