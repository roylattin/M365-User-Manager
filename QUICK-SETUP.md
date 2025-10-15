# Quick Setup Guide for Team Members

## 🎯 What is this?
This tool provides a Windows GUI for creating and managing Microsoft 365 users with Copilot licensing in your Azure tenant.

## 🚀 Super Quick Start (5 minutes)

1. **Clone the repository:**
   ```powershell
   git clone https://github.com/your-username/M365-Copilot-User-Management.git
   cd M365-Copilot-User-Management
   ```

2. **Run the quick setup:**
   ```powershell
   .\Quick-Start.ps1
   ```

That's it! The script will:
- Install required modules
- Walk you through tenant configuration  
- Launch the GUI application

## 📋 What You Need

- Windows machine with PowerShell 5.1+
- Microsoft 365 tenant admin access
- Your tenant ID and domain name
- Available M365 E5 and Copilot licenses

## 🎯 Using the Application

### Creating Users
1. **"Create Users" tab** → Connect to Graph → Fill form → Provision User
2. User gets M365 E5 + Copilot licenses automatically
3. Temporary password is displayed (save it!)

### Managing Users  
1. **"Manage Users" tab** → Connect to Graph → Refresh User List
2. Check boxes next to users you want to remove
3. Click "Deprovision Selected" → Confirm deletion

## 🔧 Manual Setup (if Quick-Start fails)

```powershell
# Step 1: Environment setup
.\Setup-Environment.ps1

# Step 2: Configure your tenant
.\Setup-Configuration.ps1

# Step 3: Launch the app
.\M365UserManager.ps1
```

## 🆘 Need Help?

1. **Check the logs:** `Logs/` directory has detailed information
2. **Common issues:** See the main README.md troubleshooting section
3. **Error messages:** The GUI shows real-time status in the output log

## ⚠️ Important Notes

- **Test first!** Always test with non-production users initially
- **Admin required:** You need M365 admin privileges  
- **License check:** Verify you have available licenses before creating users
- **Security:** Never commit your `Config/settings.json` file to Git

---

**Questions?** Check the full [README.md](README.md) or contact your Azure administrator.