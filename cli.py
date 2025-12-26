#!/usr/bin/env python3
"""Command-line interface for Distribution List Manager."""

import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.prompt import Confirm

from distribution_list_manager import DistributionListManager

app = typer.Typer(
    name="dlm",
    help="Distribution List Manager for Microsoft 365 / Entra ID",
    no_args_is_help=True,
)
console = Console()


def get_manager() -> DistributionListManager:
    """Get initialized manager instance."""
    try:
        return DistributionListManager()
    except ValueError as e:
        console.print(f"[red]Configuration Error:[/red] {e}")
        raise typer.Exit(1)


# ============================================================================
# List Commands
# ============================================================================


@app.command("list")
def list_distribution_lists(
    show_members: bool = typer.Option(False, "--members", "-m", help="Show member count"),
    search: Optional[str] = typer.Option(None, "--search", "-s", help="Search by name or email"),
):
    """List all distribution lists."""
    manager = get_manager()

    with console.status("Fetching distribution lists..."):
        if search:
            lists = manager.search(search)
        else:
            lists = manager.list_all(include_members=show_members)

    if not lists:
        console.print("[yellow]No distribution lists found.[/yellow]")
        return

    table = Table(title="Distribution Lists")
    table.add_column("Display Name", style="cyan")
    table.add_column("Email", style="green")
    table.add_column("Description")
    if show_members:
        table.add_column("Members", justify="right")
    table.add_column("ID", style="dim")

    for dl in lists:
        row = [dl.display_name, dl.mail, dl.description or "-"]
        if show_members:
            row.append(str(dl.member_count))
        row.append(dl.id[:8] + "...")
        table.add_row(*row)

    console.print(table)
    console.print(f"\n[dim]Total: {len(lists)} distribution list(s)[/dim]")


@app.command("show")
def show_distribution_list(
    identifier: str = typer.Argument(..., help="List ID or email address"),
):
    """Show details of a distribution list including all members."""
    manager = get_manager()

    with console.status("Fetching distribution list..."):
        # Try by email first, then by ID
        if "@" in identifier:
            dl = manager.get_by_email(identifier)
        else:
            try:
                dl = manager.get_by_id(identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {identifier}")
            raise typer.Exit(1)

        members = manager.get_members(dl.id)

    # Show list info
    console.print(Panel(
        f"[bold]{dl.display_name}[/bold]\n\n"
        f"Email: [green]{dl.mail}[/green]\n"
        f"Description: {dl.description or 'N/A'}\n"
        f"ID: [dim]{dl.id}[/dim]",
        title="Distribution List Details",
    ))

    # Show members
    if members:
        table = Table(title=f"Members ({len(members)})")
        table.add_column("Name", style="cyan")
        table.add_column("Email", style="green")
        table.add_column("Type")

        for member in members:
            table.add_row(member.display_name, member.email, member.user_type)

        console.print(table)
    else:
        console.print("[yellow]No members in this list.[/yellow]")


# ============================================================================
# Member Commands
# ============================================================================


@app.command("add")
def add_member(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    user_email: str = typer.Argument(..., help="Email of user to add"),
):
    """Add a member to a distribution list."""
    manager = get_manager()

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier and not list_identifier.startswith("@"):
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    with console.status(f"Adding {user_email} to {dl.display_name}..."):
        try:
            manager.add_member(dl.id, user_email)
            console.print(f"[green]Successfully added[/green] {user_email} to [cyan]{dl.display_name}[/cyan]")
        except Exception as e:
            console.print(f"[red]Failed to add member:[/red] {e}")
            raise typer.Exit(1)


@app.command("remove")
def remove_member(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    user_email: str = typer.Argument(..., help="Email of user to remove"),
    force: bool = typer.Option(False, "--force", "-f", help="Skip confirmation"),
):
    """Remove a member from a distribution list."""
    manager = get_manager()

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier and not list_identifier.startswith("@"):
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    if not force:
        if not Confirm.ask(f"Remove [bold]{user_email}[/bold] from [cyan]{dl.display_name}[/cyan]?"):
            console.print("[yellow]Cancelled.[/yellow]")
            raise typer.Exit(0)

    with console.status(f"Removing {user_email} from {dl.display_name}..."):
        try:
            manager.remove_member(dl.id, user_email)
            console.print(f"[green]Successfully removed[/green] {user_email} from [cyan]{dl.display_name}[/cyan]")
        except Exception as e:
            console.print(f"[red]Failed to remove member:[/red] {e}")
            raise typer.Exit(1)


# ============================================================================
# Bulk Operations
# ============================================================================


@app.command("import")
def import_members(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    file_path: Path = typer.Argument(..., help="CSV or TXT file with emails (one per line)"),
    column: str = typer.Option("email", "--column", "-c", help="Column name for CSV files"),
):
    """Import members from a file (CSV or TXT with one email per line)."""
    manager = get_manager()

    if not file_path.exists():
        console.print(f"[red]File not found:[/red] {file_path}")
        raise typer.Exit(1)

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier:
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    # Read emails from file
    emails = []
    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        import pandas as pd
        df = pd.read_csv(file_path)
        if column not in df.columns:
            console.print(f"[red]Column '{column}' not found in CSV.[/red]")
            console.print(f"Available columns: {', '.join(df.columns)}")
            raise typer.Exit(1)
        emails = df[column].dropna().tolist()
    elif suffix in [".xlsx", ".xls"]:
        import pandas as pd
        df = pd.read_excel(file_path)
        if column not in df.columns:
            console.print(f"[red]Column '{column}' not found in Excel file.[/red]")
            console.print(f"Available columns: {', '.join(df.columns)}")
            raise typer.Exit(1)
        emails = df[column].dropna().tolist()
    else:
        # Plain text, one email per line
        with open(file_path, "r") as f:
            emails = [line.strip() for line in f if line.strip() and "@" in line]

    if not emails:
        console.print("[yellow]No valid emails found in file.[/yellow]")
        raise typer.Exit(0)

    console.print(f"Found [bold]{len(emails)}[/bold] emails to import into [cyan]{dl.display_name}[/cyan]")

    if not Confirm.ask("Proceed with import?"):
        console.print("[yellow]Cancelled.[/yellow]")
        raise typer.Exit(0)

    with console.status("Importing members..."):
        results = manager.add_members_bulk(dl.id, emails)

    console.print(f"\n[green]Successfully added:[/green] {len(results['success'])} members")

    if results["failed"]:
        console.print(f"[red]Failed:[/red] {len(results['failed'])} members")
        for fail in results["failed"]:
            console.print(f"  - {fail['email']}: {fail['error']}")


@app.command("export")
def export_members(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    output: Optional[Path] = typer.Option(None, "--output", "-o", help="Output file path"),
    format: str = typer.Option("csv", "--format", "-f", help="Output format: csv, txt, xlsx"),
):
    """Export members of a distribution list to a file."""
    manager = get_manager()

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier:
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    with console.status("Fetching members..."):
        members = manager.get_members(dl.id)

    if not members:
        console.print("[yellow]No members to export.[/yellow]")
        raise typer.Exit(0)

    # Determine output path
    from config import Config
    Config.EXPORT_DIR.mkdir(exist_ok=True)

    if output is None:
        safe_name = dl.mail.split("@")[0].replace(".", "_")
        output = Config.EXPORT_DIR / f"{safe_name}_members.{format}"

    # Export based on format
    if format == "csv":
        import pandas as pd
        df = pd.DataFrame([
            {"name": m.display_name, "email": m.email, "type": m.user_type}
            for m in members
        ])
        df.to_csv(output, index=False)
    elif format == "xlsx":
        import pandas as pd
        df = pd.DataFrame([
            {"name": m.display_name, "email": m.email, "type": m.user_type}
            for m in members
        ])
        df.to_excel(output, index=False)
    else:  # txt
        with open(output, "w") as f:
            for m in members:
                f.write(f"{m.email}\n")

    console.print(f"[green]Exported {len(members)} members to:[/green] {output}")


# ============================================================================
# List Management
# ============================================================================


@app.command("create")
def create_distribution_list(
    name: str = typer.Argument(..., help="Display name for the list"),
    alias: str = typer.Argument(..., help="Email alias (part before @domain.com)"),
    description: Optional[str] = typer.Option(None, "--description", "-d", help="Description"),
):
    """Create a new distribution list."""
    manager = get_manager()

    with console.status(f"Creating distribution list '{name}'..."):
        try:
            dl = manager.create_list(name, alias, description)
            console.print(f"[green]Successfully created distribution list:[/green]")
            console.print(f"  Name: [cyan]{dl.display_name}[/cyan]")
            console.print(f"  Email: [green]{dl.mail}[/green]")
            console.print(f"  ID: [dim]{dl.id}[/dim]")
        except Exception as e:
            console.print(f"[red]Failed to create list:[/red] {e}")
            raise typer.Exit(1)


@app.command("update")
def update_distribution_list(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    name: Optional[str] = typer.Option(None, "--name", "-n", help="New display name"),
    description: Optional[str] = typer.Option(None, "--description", "-d", help="New description"),
):
    """Update a distribution list's properties."""
    manager = get_manager()

    if not name and description is None:
        console.print("[yellow]No changes specified. Use --name or --description.[/yellow]")
        raise typer.Exit(0)

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier:
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    with console.status(f"Updating {dl.display_name}..."):
        try:
            manager.update_list(dl.id, display_name=name, description=description)
            console.print(f"[green]Successfully updated[/green] [cyan]{dl.display_name}[/cyan]")
        except Exception as e:
            console.print(f"[red]Failed to update list:[/red] {e}")
            raise typer.Exit(1)


@app.command("delete")
def delete_distribution_list(
    list_identifier: str = typer.Argument(..., help="Distribution list ID or email"),
    force: bool = typer.Option(False, "--force", "-f", help="Skip confirmation"),
):
    """Delete a distribution list."""
    manager = get_manager()

    # Resolve list
    with console.status("Finding distribution list..."):
        if "@" in list_identifier:
            dl = manager.get_by_email(list_identifier)
        else:
            try:
                dl = manager.get_by_id(list_identifier)
            except Exception:
                dl = None

        if not dl:
            console.print(f"[red]Distribution list not found:[/red] {list_identifier}")
            raise typer.Exit(1)

    if not force:
        console.print(f"[bold red]WARNING:[/bold red] This will permanently delete the distribution list:")
        console.print(f"  Name: [cyan]{dl.display_name}[/cyan]")
        console.print(f"  Email: [green]{dl.mail}[/green]")
        if not Confirm.ask("Are you sure you want to delete this list?"):
            console.print("[yellow]Cancelled.[/yellow]")
            raise typer.Exit(0)

    with console.status(f"Deleting {dl.display_name}..."):
        try:
            manager.delete_list(dl.id)
            console.print(f"[green]Successfully deleted[/green] [cyan]{dl.display_name}[/cyan]")
        except Exception as e:
            console.print(f"[red]Failed to delete list:[/red] {e}")
            raise typer.Exit(1)


# ============================================================================
# User Lookup
# ============================================================================


@app.command("user-lists")
def user_lists(
    user_email: str = typer.Argument(..., help="User's email address"),
):
    """Show all distribution lists a user belongs to."""
    manager = get_manager()

    with console.status(f"Finding lists for {user_email}..."):
        try:
            lists = manager.get_user_memberships(user_email)
        except ValueError as e:
            console.print(f"[red]Error:[/red] {e}")
            raise typer.Exit(1)

    if not lists:
        console.print(f"[yellow]{user_email} is not a member of any distribution lists.[/yellow]")
        return

    table = Table(title=f"Distribution Lists for {user_email}")
    table.add_column("Display Name", style="cyan")
    table.add_column("Email", style="green")

    for dl in lists:
        table.add_row(dl.display_name, dl.mail)

    console.print(table)
    console.print(f"\n[dim]Total: {len(lists)} list(s)[/dim]")


if __name__ == "__main__":
    app()
