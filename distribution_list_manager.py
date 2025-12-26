"""Distribution List Manager for Microsoft 365 / Entra ID."""

import logging
import os
from typing import Optional
from dataclasses import dataclass
from graph_client import GraphClient

logger = logging.getLogger(__name__)

# Flag to track if we should use Exchange PowerShell
USE_EXCHANGE_POWERSHELL = False


@dataclass
class DistributionList:
    """Represents a distribution list (mail-enabled group)."""

    id: str
    display_name: str
    mail: str
    description: Optional[str] = None
    member_count: int = 0

    @classmethod
    def from_graph(cls, data: dict) -> "DistributionList":
        """Create DistributionList from Graph API response."""
        return cls(
            id=data.get("id", ""),
            display_name=data.get("displayName", ""),
            mail=data.get("mail", ""),
            description=data.get("description"),
        )


@dataclass
class Member:
    """Represents a member of a distribution list."""

    id: str
    display_name: str
    email: str
    user_type: str

    @classmethod
    def from_graph(cls, data: dict) -> "Member":
        """Create Member from Graph API response."""
        return cls(
            id=data.get("id", ""),
            display_name=data.get("displayName", ""),
            email=data.get("mail", "") or data.get("userPrincipalName", ""),
            user_type=data.get("@odata.type", "").replace("#microsoft.graph.", ""),
        )


class DistributionListManager:
    """Manager for distribution list operations."""

    def __init__(self):
        self.client = GraphClient()

    def list_all(self, include_members: bool = False) -> list[DistributionList]:
        """
        List all distribution lists (mail-enabled security groups and distribution groups).

        Distribution lists in M365 can be:
        - Mail-enabled security groups (mailEnabled=true, securityEnabled=true)
        - Distribution groups (mailEnabled=true, securityEnabled=false, no "Unified" in groupTypes)

        NOT included:
        - Microsoft 365 Groups (have "Unified" in groupTypes - these are team mailboxes)
        """
        # Get groups that have mail enabled
        filter_query = "mailEnabled eq true"
        select_fields = "id,displayName,mail,description,groupTypes,securityEnabled"

        groups = self.client.get_all_pages(
            "/groups", params={"$filter": filter_query, "$select": select_fields}
        )

        distribution_lists = []
        for group in groups:
            group_types = group.get("groupTypes", [])
            security_enabled = group.get("securityEnabled", False)

            # Skip Microsoft 365 Groups (they have "Unified" in groupTypes)
            if "Unified" in group_types:
                logger.debug(f"Skipping M365 Group: {group.get('displayName')}")
                continue

            # Skip Mail-enabled Security Groups (securityEnabled=true)
            if security_enabled:
                logger.debug(f"Skipping Security Group: {group.get('displayName')}")
                continue

            # Include only pure Distribution Lists:
            # - mailEnabled=true
            # - securityEnabled=false
            # - no "Unified" in groupTypes
            dl = DistributionList.from_graph(group)
            if include_members:
                members = self.get_members(dl.id)
                dl.member_count = len(members)
            distribution_lists.append(dl)

        logger.info(f"Found {len(distribution_lists)} distribution lists")
        return distribution_lists

    def get_by_id(self, list_id: str) -> DistributionList:
        """Get a distribution list by its ID."""
        data = self.client.get(f"/groups/{list_id}")
        return DistributionList.from_graph(data)

    def get_by_email(self, email: str) -> Optional[DistributionList]:
        """Get a distribution list by its email address."""
        filter_query = f"mail eq '{email}'"
        result = self.client.get("/groups", params={"$filter": filter_query})

        groups = result.get("value", [])
        if groups:
            return DistributionList.from_graph(groups[0])
        return None

    def search(self, query: str) -> list[DistributionList]:
        """Search distribution lists by name or email."""
        filter_query = (
            f"mailEnabled eq true and "
            f"(startswith(displayName, '{query}') or startswith(mail, '{query}'))"
        )
        groups = self.client.get_all_pages("/groups", params={"$filter": filter_query})
        return [DistributionList.from_graph(g) for g in groups]

    def get_members(self, list_id: str) -> list[Member]:
        """Get all members of a distribution list."""
        members_data = self.client.get_all_pages(f"/groups/{list_id}/members")
        return [Member.from_graph(m) for m in members_data]

    def add_member(self, list_id: str, user_email: str) -> bool:
        """Add a member to a distribution list by email."""
        # Get the list info for Exchange fallback
        dl = self.get_by_id(list_id)

        # First, find the user by email
        user = self._find_user_by_email(user_email)
        if not user:
            # User not in directory - try Exchange directly (external contacts)
            logger.warning(f"User {user_email} not in directory, trying Exchange...")
            return self._add_via_exchange(dl.mail, user_email)

        # Add member to group via Graph API
        try:
            data = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user['id']}"}
            self.client.post(f"/groups/{list_id}/members/$ref", data)
            return True
        except Exception as e:
            if "Cannot Update a mail-enabled" in str(e) or "Request_BadRequest" in str(e):
                logger.warning("Graph API cannot modify this list. Trying Exchange PowerShell...")
                return self._add_via_exchange(dl.mail, user_email)
            raise

    def _add_via_exchange(self, list_email: str, member_email: str) -> bool:
        """Add member using Exchange Online PowerShell."""
        logger.info(f"Using Exchange PowerShell to add {member_email}...")

        try:
            from exchange_client import ExchangeClient
            exchange = ExchangeClient()
            exchange.add_member(list_email, member_email)
            logger.info(f"Successfully added {member_email} via Exchange")
            return True
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Exchange error: {error_msg}")

            if "not installed" in error_msg.lower():
                raise ValueError(
                    f"ExchangeOnlineManagement module not installed.\n\n"
                    f"Run setup_exchange.bat to configure Exchange access."
                )
            elif "EXCHANGE_CERT_THUMBPRINT" not in os.environ or not os.environ.get("EXCHANGE_CERT_THUMBPRINT"):
                raise ValueError(
                    f"Exchange not configured for this app.\n\n"
                    f"Run setup_exchange.bat to set up Exchange access,\n"
                    f"then you can manage distribution lists from the app."
                )
            else:
                raise ValueError(f"Exchange error: {error_msg}")

    def add_members_bulk(self, list_id: str, emails: list[str]) -> dict:
        """Add multiple members to a distribution list."""
        results = {"success": [], "failed": []}

        for email in emails:
            try:
                self.add_member(list_id, email)
                results["success"].append(email)
            except Exception as e:
                results["failed"].append({"email": email, "error": str(e)})

        return results

    def remove_member(self, list_id: str, user_email: str) -> bool:
        """Remove a member from a distribution list by email."""
        logger.info(f"Removing member: {user_email} from list: {list_id[:8]}...")

        # First, try to find the member directly in the group's member list
        members = self.get_members(list_id)
        logger.info(f"Found {len(members)} members in list")

        member = next(
            (m for m in members if m.email.lower() == user_email.lower()),
            None
        )

        # Get the list info for Exchange fallback
        dl = self.get_by_id(list_id)

        if member:
            logger.info(f"Found member by email match: {member.display_name} (ID: {member.id[:8]}...)")
            try:
                # Use the member's ID from the group
                self.client.delete(f"/groups/{list_id}/members/{member.id}/$ref")
                logger.info(f"Successfully removed {user_email}")
                return True
            except Exception as e:
                if "Cannot Update a mail-enabled" in str(e) or "Request_BadRequest" in str(e):
                    logger.warning("Graph API cannot modify this list. Trying Exchange PowerShell...")
                    return self._remove_via_exchange(dl.mail, user_email)
                raise

        # Log all member emails for debugging
        logger.warning(f"Member {user_email} not found by exact email match")
        logger.info(f"Members in list: {[m.email for m in members]}")

        # Fallback: try to find user in directory
        user = self._find_user_by_email(user_email)
        if not user:
            raise ValueError(f"Member not found in list: {user_email}")

        logger.info(f"Found user in directory: {user.get('displayName')} (ID: {user['id'][:8]}...)")
        try:
            # Remove from group using directory user ID
            self.client.delete(f"/groups/{list_id}/members/{user['id']}/$ref")
            logger.info(f"Successfully removed {user_email}")
            return True
        except Exception as e:
            if "Cannot Update a mail-enabled" in str(e) or "Request_BadRequest" in str(e):
                logger.warning("Graph API cannot modify this list. Trying Exchange PowerShell...")
                return self._remove_via_exchange(dl.mail, user_email)
            raise

    def _remove_via_exchange(self, list_email: str, member_email: str) -> bool:
        """Remove member using Exchange Online PowerShell."""
        logger.info(f"Using Exchange PowerShell to remove {member_email}...")

        try:
            from exchange_client import ExchangeClient
            exchange = ExchangeClient()
            exchange.remove_member(list_email, member_email)
            logger.info(f"Successfully removed {member_email} via Exchange")
            return True
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Exchange error: {error_msg}")

            # Check if Exchange is not configured
            if "not installed" in error_msg.lower():
                raise ValueError(
                    f"ExchangeOnlineManagement module not installed.\n\n"
                    f"Run setup_exchange.bat to configure Exchange access."
                )
            elif "EXCHANGE_CERT_THUMBPRINT" not in os.environ or not os.environ.get("EXCHANGE_CERT_THUMBPRINT"):
                raise ValueError(
                    f"Exchange not configured for this app.\n\n"
                    f"Run setup_exchange.bat to set up Exchange access,\n"
                    f"then you can manage distribution lists from the app."
                )
            else:
                raise ValueError(f"Exchange error: {error_msg}")

    def remove_members_bulk(self, list_id: str, emails: list[str]) -> dict:
        """Remove multiple members from a distribution list."""
        results = {"success": [], "failed": []}

        for email in emails:
            try:
                self.remove_member(list_id, email)
                results["success"].append(email)
            except Exception as e:
                results["failed"].append({"email": email, "error": str(e)})

        return results

    def create_list(
        self,
        display_name: str,
        mail_nickname: str,
        description: Optional[str] = None,
    ) -> DistributionList:
        """
        Create a new distribution list (mail-enabled group).

        Args:
            display_name: Display name of the list
            mail_nickname: Email alias (the part before @domain.com)
            description: Optional description
        """
        data = {
            "displayName": display_name,
            "mailEnabled": True,
            "mailNickname": mail_nickname,
            "securityEnabled": False,
            "groupTypes": [],
        }
        if description:
            data["description"] = description

        result = self.client.post("/groups", data)
        return DistributionList.from_graph(result)

    def update_list(
        self,
        list_id: str,
        display_name: Optional[str] = None,
        description: Optional[str] = None,
        mail_nickname: Optional[str] = None,
    ) -> bool:
        """Update distribution list properties."""
        data = {}
        if display_name:
            data["displayName"] = display_name
        if description is not None:
            data["description"] = description
        if mail_nickname:
            data["mailNickname"] = mail_nickname

        if data:
            self.client.patch(f"/groups/{list_id}", data)
        return True

    def delete_list(self, list_id: str) -> bool:
        """Delete a distribution list."""
        return self.client.delete(f"/groups/{list_id}")

    def _find_user_by_email(self, email: str) -> Optional[dict]:
        """Find a user or contact by email address."""
        # Try finding by mail
        filter_query = f"mail eq '{email}'"
        result = self.client.get("/users", params={"$filter": filter_query})

        users = result.get("value", [])
        if users:
            return users[0]

        # Try by userPrincipalName
        filter_query = f"userPrincipalName eq '{email}'"
        result = self.client.get("/users", params={"$filter": filter_query})

        users = result.get("value", [])
        if users:
            return users[0]

        return None

    def get_user_memberships(self, user_email: str) -> list[DistributionList]:
        """Get all distribution lists a user is a member of."""
        # First try the fast method using Graph API memberOf
        user = self._find_user_by_email(user_email)
        if user:
            try:
                groups = self.client.get_all_pages(
                    f"/users/{user['id']}/memberOf",
                    params={"$filter": "mailEnabled eq true"},
                )
                return [DistributionList.from_graph(g) for g in groups if g.get("mail")]
            except Exception:
                pass  # Fall through to manual search

        # Fallback: search through all distribution lists manually
        # This works for external contacts, guests, or when Graph memberOf fails
        return self.find_email_in_all_lists(user_email)

    def find_email_in_all_lists(self, search_term: str, progress_callback=None,
                                  result_callback=None, partial_match: bool = True) -> list[DistributionList]:
        """
        Search all distribution lists to find which ones contain the given email.

        Args:
            search_term: Email or partial email to search for
            progress_callback: Called with (current, total, list_name) for progress updates
            result_callback: Called with (dl, matching_email) when a match is found
            partial_match: If True, matches if search_term is contained in email
        """
        search_lower = search_term.lower()
        results = []

        # Get all distribution lists
        all_lists = self.list_all()

        for i, dl in enumerate(all_lists):
            if progress_callback:
                progress_callback(i + 1, len(all_lists), dl.display_name)

            try:
                members = self.get_members(dl.id)
                for member in members:
                    member_email_lower = member.email.lower()

                    # Check for match (partial or exact)
                    if partial_match:
                        is_match = search_lower in member_email_lower
                    else:
                        is_match = member_email_lower == search_lower

                    if is_match:
                        results.append((dl, member.email))
                        if result_callback:
                            result_callback(dl, member.email)
                        break
            except Exception as e:
                logger.warning(f"Failed to get members for {dl.display_name}: {e}")

        return results
