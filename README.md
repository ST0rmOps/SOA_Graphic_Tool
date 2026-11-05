### Universal SOA Management Tool for Microsoft Entra ID

I developed a universal PowerShell-based tool with a graphical interface to simplify and automate Source of Authority (SOA) management in Microsoft Entra ID.

The tool enables administrators to transition synchronized users and groups from on-premises Active Directory management to cloud management, either individually or in bulk.

**Key capabilities:**
- Convert individual users, groups, or entire sets of domain users to Cloud Managed.
- Perform bulk conversions using CSV files.
- Discover and export users and groups from Microsoft Entra ID.
- Check SOA status and validate Microsoft Graph permissions.
- Manage multiple domains through a graphical interface.
- Automate Microsoft Graph API batch operations with throttling handling, retries, progress tracking, and CSV reporting.
- Enable or disable user accounts individually or in bulk.

**Built with:** PowerShell, Windows Forms, Microsoft Graph API, Microsoft Entra ID.

My goal was to make complex hybrid identity migration and SOA management tasks easier, faster, and more accessible to administrators without requiring them to manually execute Graph API requests.
