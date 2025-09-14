# MFA Migration Assessment Tool

A comprehensive PowerShell toolkit for assessing and planning Microsoft 365 MFA migrations from legacy per-user settings to Authentication Methods Policy.

## 📋 Overview

Microsoft is deprecating legacy per-user MFA settings on **September 30, 2025**. This tool provides data-driven assessment and migration planning to ensure zero disruption while identifying security enhancement opportunities.

### Key Features

- **Zero Disruption Migration Planning** - Ensures all current MFA users continue working
- **Dual Risk Assessment** - Separates migration compliance from security improvements  
- **Comprehensive Reporting** - Professional Word documents and Excel spreadsheets
- **Privileged User Analysis** - Special focus on administrator account security
- **FIDO2 Deployment Planning** - Roadmap for phishing-resistant authentication
- **Conditional Access Guidance** - Pragmatic policy deployment recommendations

## 🚀 Quick Start

### Prerequisites

1. **PowerShell 5.1** or later
2. **Microsoft Graph PowerShell SDK**
3. **Required Modules** (auto-installed):
   - `ImportExcel` - For Excel report generation
   - `PSWriteWord` - For Word document creation

### Installation

```powershell
# Clone the repository
git clone https://github.com/yourusername/mfa-migration-assessment.git
cd mfa-migration-assessment

# Install required modules
Install-Module -Name ImportExcel -Scope CurrentUser -Force
Install-Module -Name PSWriteWord -Scope CurrentUser -Force

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Policy.Read.All", "User.Read.All", "Directory.Read.All"
```

### Basic Usage

```powershell
# Run the complete assessment
.\Run-MfaAssessment.ps1

# Or run individual components
.\Get-CurrentMfaStatus.ps1 | .\Generate-MfaReports.ps1
```

## 📊 What You Get

### Generated Reports

1. **Migration Report** (`MFA_Migration_Report_[timestamp].docx`)
   - Executive summary with dual timeline assessment
   - Phase-by-phase implementation plan
   - Security compliance analysis
   - Next steps and recommendations

2. **User Methods Analysis** (`MFA_User_Methods_[timestamp].xlsx`)
   - Complete user inventory with current methods
   - Migration impact assessment per user
   - Security compliance status
   - Phase 2 action planning

3. **Privileged Users Security** (`MFA_Privileged_Users_[timestamp].xlsx`)
   - Administrator account analysis
   - FIDO2 deployment recommendations
   - Risk assessment and compliance status

### Sample Executive Summary

```
MIGRATION READINESS ASSESSMENT:
✅ Phase 1 Ready: Zero disruption expected - All current MFA users will continue working
✅ September 30th Deadline: Achievable without service interruption

SECURITY COMPLIANCE ASSESSMENT:
- Users with MFA (Compliant): 105
- Users without MFA: 15 users [WARNING - Security Policy Gap]  
- Privileged users without MFA: 0 [COMPLIANT]

KEY DISTINCTION:
• Migration Timeline: No urgency - zero disruption expected for September 30th deadline
• Security Compliance: Ongoing concern requiring attention per organizational security policy
```

## 🎯 Migration Philosophy

This tool embodies a **pragmatic approach** to security consulting:

- **Meet customers where they are** in their security journey
- **Perfect is the enemy of better** - incremental improvement over paralysis
- **Every bit helps** - basic protection today > perfect protection someday
- **Data-driven conversations** rather than fear-based emergency responses

### Two-Phase Approach

**Phase 1: Meet the Deadline (1-2 days)**
- Enable all currently used authentication methods
- Ensure zero user disruption  
- Achieve compliance with Microsoft deadline

**Phase 2: Security Enhancement (4-6 weeks)**
- Remove insecure authentication methods
- Register MFA for unprotected users
- Deploy FIDO2 for administrators
- Implement Conditional Access policies

## 🔧 Technical Details

### Core Assessment Logic

The tool performs comprehensive analysis across multiple dimensions:

```powershell
# User categorization logic
if ($userStatus.Status -eq "Has Current Methods") {
    $migrationImpact = "Protected - Will continue working"
    $phase1Action = "Enable existing methods in policy"
} 
elseif ($userStatus.Status -eq "Password Only - Needs MFA") {
    $migrationImpact = "Unaffected - No change in access"  
    $securityCompliance = if ($isPrivileged) { "CRITICAL" } else { "WARNING" }
}
```

### Privileged User Analysis

- **Automatic detection** of administrative roles
- **Phishing-resistant method assessment** (FIDO2, Certificate, Windows Hello)
- **Break-glass account management** 
- **FIDO2 deployment prioritization**

### Report Generation

- **Professional Word documents** using PSWriteWord with advanced formatting
- **Excel workbooks** with tables, filtering, and conditional formatting
- **Fallback to CSV/TXT** if modules unavailable
- **Long path handling** for complex directory structures

## 📚 Use Cases

### For MSPs and IT Consultants

- **Multi-tenant assessment** across customer base
- **Professional reports** for stakeholder communication  
- **Implementation planning** with clear timelines
- **Risk communication** separating urgent from important

### For Internal IT Teams

- **Current state assessment** of MFA deployment
- **Migration planning** with zero disruption guarantee
- **Security gap identification** and remediation planning
- **Executive reporting** with clear recommendations

### For Security Teams

- **Privileged user compliance** assessment
- **FIDO2 deployment planning** for administrators
- **Conditional Access roadmap** development
- **Risk-based security improvement** prioritization

## 🛡️ Security Best Practices

### Conditional Access Deployment

The tool recommends a **pragmatic phased approach**:

**Phase 1: Foundation Policies (No Dependencies)**
- Admin protection (require MFA for administrative roles)
- Guest user controls
- Location-based protection  
- Legacy authentication blocking

**Phase 2: Enhanced Policies (When Ready)**
- Device compliance requirements
- Application protection policies
- Risk-based conditional access

### FIDO2 Security Keys

- **Addresses phone enrollment resistance** - no personal device required
- **Phishing-resistant protection** - stronger than SMS/authenticator apps  
- **Cost-effective security** - $20-50 per administrator for significant improvement
- **User-friendly experience** - simple tap-to-authenticate

## 🤝 Contributing

This tool was developed based on real-world experience with 1,200+ customer tenants. Contributions welcome!

### Areas for Enhancement

- Additional authentication method support
- Conditional Access policy templates
- Automated policy deployment
- Integration with ITSM systems
- Multi-language support

### Development Setup

```powershell
# Clone and setup development environment
git clone https://github.com/yourusername/mfa-migration-assessment.git
cd mfa-migration-assessment

# Install development dependencies
Install-Module -Name Pester -Scope CurrentUser  # For testing
Install-Module -Name PSScriptAnalyzer -Scope CurrentUser  # For code analysis
```

## 📝 Blog Post

Read the full story behind this tool: [Meeting Customers Where They Are: A Pragmatic Approach to MFA Migration](link-to-blog-post)

## 📞 Support

- **Issues**: Please use GitHub Issues for bug reports and feature requests
- **Discussions**: Use GitHub Discussions for questions and community support
- **Professional Services**: Contact for enterprise consulting and customisation

---

## 📄 **License:** Apache 2.0 (see LICENSE)  
**Additional restriction:** Commons Clause (see COMMONS-CLAUSE.txt)

**SPDX headers**
- Each source file includes:  
  `SPDX-License-Identifier: Apache-2.0 WITH Commons-Clause`

---

### FAQ: MSP and Consulting Use

**Q: Can an MSP or consultant use this tool in a paid engagement?**  
**A:** It depends on how the tool is used:  
- **Allowed:** If the tool is used internally by the end customer (e.g., installed in their tenant) and the consultant is simply assisting, this is generally acceptable.  
- **Not allowed without a commercial licence:** If the MSP or consultant provides a managed service where the tool runs in their own environment (e.g., their tenant or infrastructure) or if the value of the service substantially derives from the tool’s functionality, this falls under the definition of “Sell” in the Commons Clause and requires a commercial licence.

**Q: Why is this restricted?**  
The Commons Clause removes the right to “Sell,” which includes providing a service for a fee where the value derives from the software. This ensures fair use and prevents competitors from monetising the tool without contributing back.

**Q: How do I get a commercial licence?**  
Contact Global Micro Solutions (Pty) Ltd at:  
📧 licensing@globalmicro.co.za

---

## ⚠️ Warranty Disclaimer

Distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. Please review the Apache-2.0 WITH Commons-Clause License for the specific language governing permissions and limitations under the License.

---

## Author

**JJ Milner**  
Blog: https://jjrmilner.substack.com
Github: https://github.com/jjrmilner

## 🙏 Acknowledgments

- **Microsoft Graph Team** - For comprehensive authentication APIs
- **PSWriteWord Community** - For excellent Word document generation
- **ImportExcel Community** - For powerful Excel manipulation capabilities
- **1,200+ Customer Tenants** - For providing real-world testing and validation

---

**Remember: Better security today is worth more than perfect security someday.**

*Meeting customers where they are in their security journey, one practical improvement at a time.*
