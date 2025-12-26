# Distribution List Manager for Microsoft 365 / Entra ID

A graphical and command-line tool to manage distribution lists in your Microsoft 365 organization.

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Configure (copy and edit .env)
copy .env.example .env

# Run the GUI
python gui.py

# Or use the CLI
python cli.py --help
```

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Register Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com) > **Entra ID** > **App registrations** > **New registration**
2. Name: `Distribution List Manager`
3. Account type: **Single tenant**
4. Redirect URI: **Public client/native** > `http://localhost`

### 3. Add API Permissions

In your app registration, go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**:

- `Group.ReadWrite.All` - Manage distribution lists
- `User.Read.All` - Look up users by email
- `Directory.Read.All` - Read directory data

Then click **Grant admin consent**.

### 4. Create Client Secret

Go to **Certificates & secrets** > **New client secret** and copy the value.

### 5. Configure Environment

```bash
cp .env.example .env
```

Edit `.env` with your values:
```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

## Graphical Interface (GUI)

Launch the GUI with:
```bash
python gui.py
```

**Features:**
- View all distribution lists in a searchable list
- Select a list to view all its members
- Add single or multiple members (bulk import from CSV/Excel/TXT)
- Remove selected members
- Edit list properties (name, description)
- Export members to CSV, Excel, or TXT
- Dark theme with modern UI

## Command Line Interface (CLI)

### List all distribution lists
```bash
python cli.py list
python cli.py list --members    # Include member count
python cli.py list --search "sales"
```

### View distribution list details
```bash
python cli.py show sales@company.com
python cli.py show <list-id>
```

### Add members
```bash
python cli.py add sales@company.com john@company.com
```

### Remove members
```bash
python cli.py remove sales@company.com john@company.com
python cli.py remove sales@company.com john@company.com --force
```

### Import members from file
```bash
# From text file (one email per line)
python cli.py import sales@company.com members.txt

# From CSV
python cli.py import sales@company.com members.csv --column email

# From Excel
python cli.py import sales@company.com members.xlsx --column email
```

### Export members
```bash
python cli.py export sales@company.com
python cli.py export sales@company.com --output members.csv
python cli.py export sales@company.com --format xlsx
python cli.py export sales@company.com --format txt
```

### Create distribution list
```bash
python cli.py create "Sales Team" sales --description "Sales department"
```

### Update distribution list
```bash
python cli.py update sales@company.com --name "Sales Department"
python cli.py update sales@company.com --description "New description"
```

### Delete distribution list
```bash
python cli.py delete sales@company.com
python cli.py delete sales@company.com --force
```

### View user's memberships
```bash
python cli.py user-lists john@company.com
```

## Required Permissions Summary

| Permission | Type | Purpose |
|------------|------|---------|
| Group.ReadWrite.All | Application | Create, read, update, delete groups |
| User.Read.All | Application | Find users by email |
| Directory.Read.All | Application | Read directory objects |

**Note:** Admin consent is required for application permissions.
