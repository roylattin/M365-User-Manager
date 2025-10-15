# M365 Copilot User Management Tool

A Windows PowerShell GUI application for provisioning and managing Microsoft 365 users with Copilot licensing in your Azure tenant.

## 🚀 Features

- **Create Users**: Simple form to create new M365 users with automatic Copilot licensing
- **Manage Users**: View all tenant users and bulk deprovision selected accounts
- **Real-time Logging**: See all operations in real-time with detailed status updates
- **Secure Authentication**: Uses Microsoft Graph PowerShell SDK with proper scopes
- **Clean Interface**: Windows Forms GUI with tabbed interface for easy navigation

## 📋 Prerequisites

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or later
- Internet connection
- Microsoft 365 tenant with admin privileges
- Available M365 E5 and Copilot licenses

## 🛠️ Quick Setup

### 1. Clone the Repository
```powershell
git clone https://github.com/your-username/M365-Copilot-User-Management.git
cd M365-Copilot-User-Management
```

### 2. Run the Setup Script
```powershell
.\Setup-Environment.ps1
```

This will:
- Install required PowerShell modules (Microsoft.Graph)
- Create necessary directories
- Set up configuration template
- Validate your environment

### 3. Configure Your Tenant
```powershell
.\Setup-Configuration.ps1
```

This will guide you through:
- Setting your tenant domain
- Configuring license SKUs
- Testing connectivity

### 4. Launch the Application
```powershell
.\M365UserManager.ps1
```

## 📁 Project Structure

```
M365-Copilot-User-Management/
├── M365UserManager.ps1                         # Main GUI application
├── Setup-Environment.ps1                       # Environment setup script
├── Setup-Configuration.ps1                     # Tenant configuration script
├── Quick-Start.ps1                             # One-click setup script
├── UserProvisioning.ps1                        # Command-line interface (optional)
├── Config/
│   ├── settings.json                          # Tenant configuration
│   └── settings.template.json                 # Configuration template
├── Logs/                                       # Application logs (auto-created)
├── README.md                                   # This file
├── QUICK-SETUP.md                             # 5-minute team guide
├── LICENSE                                     # MIT license
└── .gitignore                                 # Git ignore rules
```

## 🔧 Configuration

The `Config/settings.json` file contains your tenant-specific settings:

```json
{
  "tenant": {
    "tenantId": "your-tenant-id",
    "domain": "yourdomain.onmicrosoft.com"
  },
  "licensing": {
    "m365E5Sku": "Microsoft_365_E5_(no_Teams)",
    "copilotSku": "Microsoft_365_Copilot"
  }
}
```

## 🎯 Usage

### Creating Users
1. Click **"Connect to Microsoft Graph"** on the "Create Users" tab
2. Fill in the user details (First Name, Last Name, Username)
3. Click **"Provision User"** to create the account with M365 + Copilot licenses
4. View the real-time log for status updates

### Managing Users
1. Switch to the **"Manage Users"** tab
2. Click **"Connect to Microsoft Graph"** and **"Refresh User List"**
3. Check the boxes next to users you want to deprovision
4. Click **"Deprovision Selected"** to remove accounts and licenses

## 🔐 Security & Permissions

The application requires the following Microsoft Graph permissions:
- `User.ReadWrite.All` - Create and delete users
- `Directory.ReadWrite.All` - Access directory information
- `Directory.Read.All` - Read directory data
- `User.Read.All` - Read all user profiles
- `Organization.Read.All` - Read organization info

These permissions are requested during the initial connection and require admin consent.

## 📝 Logging

All operations are logged to:
- **GUI Output**: Real-time log display in the application
- **Log Files**: `Logs/YYYY-MM-DD-HHMMSS_Complete_UserManagement.log`

## 🐛 Troubleshooting

### Common Issues

**"Cannot connect to Microsoft Graph"**
- Ensure you have admin privileges in the tenant
- Check your internet connection
- Verify tenant ID in settings.json

**"Insufficient licenses available"**
- Check available M365 E5 licenses in admin center
- Verify Copilot licenses are available
- Update license SKUs in settings.json if needed

**"User ID is missing or empty"**
- Refresh the user list
- Ensure proper connection to Microsoft Graph
- Check the detailed logs for more information

### Getting Help

1. Check the log files in the `Logs/` directory
2. Enable detailed logging by running with `-Verbose`
3. Review the troubleshooting section in this README
4. Contact your Azure administrator for tenant-specific issues

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This tool creates and deletes user accounts in your Microsoft 365 tenant. Always test in a development environment first. The authors are not responsible for any data loss or unintended changes to your production environment.

## 🆘 Support

For issues and questions:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review existing [GitHub Issues](https://github.com/your-username/M365-Copilot-User-Management/issues)
3. Create a new issue with detailed information about your problem

---

**Made with ❤️ for Microsoft 365 administrators**