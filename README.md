# Microsoft Entra ID Distribution List Manager

A GUI and CLI tool for managing Microsoft 365 distribution lists via Microsoft Graph API.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## Features

- **GUI Application** - Modern dark-themed interface for easy management
- **CLI Support** - Full command-line interface for automation and scripting
- **Bulk Operations** - Import/export members from CSV, Excel, or TXT files
- **Search** - Find distribution lists and search email memberships across all lists

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Configure credentials
copy .env.example .env   # Then edit .env with your Azure credentials

# 3. Run
python gui.py            # GUI mode
python cli.py --help     # CLI mode
```

## Azure Setup

### 1. Register an App in Azure Portal

1. Go to [Azure Portal](https://portal.azure.com) → **Entra ID** → **App registrations** → **New registration**
2. Name: `Distribution List Manager`
3. Account type: **Single tenant**
4. Redirect URI: Leave blank

### 2. Add API Permissions

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**:

| Permission | Purpose |
|------------|---------|
| `Group.ReadWrite.All` | Manage distribution lists |
| `User.Read.All` | Look up users by email |
| `Directory.Read.All` | Read directory data |

Click **Grant admin consent** after adding permissions.

### 3. Create Client Secret

Go to **Certificates & secrets** → **New client secret** → Copy the value.

### 4. Configure Environment

Edit `.env` with your values from the Azure app registration:

```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

## CLI Usage

```bash
# List all distribution lists
python cli.py list

# Show members of a list
python cli.py show sales@company.com

# Add/remove members
python cli.py add sales@company.com user@company.com
python cli.py remove sales@company.com user@company.com

# Import from file
python cli.py import sales@company.com members.csv --column email

# Export members
python cli.py export sales@company.com --format csv

# Create/delete lists
python cli.py create "Sales Team" sales-team --description "Sales department"
python cli.py delete sales@company.com --force

# Find user's memberships
python cli.py user-lists user@company.com
```

## Requirements

- Python 3.8+
- Microsoft 365 tenant with admin access
- Azure AD app registration with appropriate permissions

## License

MIT
