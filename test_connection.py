#!/usr/bin/env python3
"""Test script to verify Azure AD / Microsoft Graph connection."""

from rich.console import Console
from rich.panel import Panel

console = Console()


def main():
    console.print(Panel("Distribution List Manager - Connection Test", style="bold blue"))

    # Test 1: Configuration
    console.print("\n[bold]1. Checking configuration...[/bold]")
    try:
        from config import Config
        Config.validate()
        console.print("   [green]✓[/green] Configuration loaded successfully")
        console.print(f"   [dim]Tenant ID: {Config.TENANT_ID[:8]}...[/dim]")
        console.print(f"   [dim]Client ID: {Config.CLIENT_ID[:8]}...[/dim]")
    except ValueError as e:
        console.print(f"   [red]✗[/red] Configuration error: {e}")
        return False

    # Test 2: Authentication
    console.print("\n[bold]2. Testing authentication...[/bold]")
    try:
        from graph_client import GraphClient
        client = GraphClient()
        token = client.token
        console.print("   [green]✓[/green] Successfully acquired access token")
    except Exception as e:
        console.print(f"   [red]✗[/red] Authentication failed: {e}")
        return False

    # Test 3: API Access
    console.print("\n[bold]3. Testing API access...[/bold]")
    try:
        # Try to list groups (just first one to verify access)
        result = client.get("/groups", params={"$top": 1, "$select": "id,displayName"})
        console.print("   [green]✓[/green] Successfully connected to Microsoft Graph API")
    except Exception as e:
        console.print(f"   [red]✗[/red] API access failed: {e}")
        console.print("   [yellow]Make sure you have granted admin consent for the required permissions.[/yellow]")
        return False

    # Test 4: Distribution lists
    console.print("\n[bold]4. Testing distribution list access...[/bold]")
    try:
        from distribution_list_manager import DistributionListManager
        manager = DistributionListManager()
        lists = manager.list_all()
        console.print(f"   [green]✓[/green] Found {len(lists)} distribution list(s)")
    except Exception as e:
        console.print(f"   [red]✗[/red] Failed to list distribution lists: {e}")
        return False

    console.print("\n[bold green]All tests passed! Your configuration is working correctly.[/bold green]")
    console.print("\nYou can now use the CLI:")
    console.print("  [cyan]python cli.py list[/cyan]          - List all distribution lists")
    console.print("  [cyan]python cli.py --help[/cyan]        - See all available commands")

    return True


if __name__ == "__main__":
    main()
