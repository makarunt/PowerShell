# Exchange Connectors and Certificates Information Script

A comprehensive PowerShell script for gathering and exporting information about Exchange Server connectors and certificates.

## Features

- **Receive Connectors**: Collects detailed configuration and settings
- **Send Connectors**: Gathers connector information and routing details
- **Exchange Certificates**: Retrieves certificate information with expiration tracking
- **Flexible Target Selection**: Query specific server or all Exchange servers
- **Multiple Export Formats**: CSV and/or HTML reports
- **Certificate Monitoring**: Automatic detection of expired and expiring certificates
- **Professional HTML Reports**: Interactive, styled reports with summary statistics

## Requirements

- Exchange Management Shell
- Exchange Server Administrator permissions
- PowerShell 5.1 or later
- Exchange Server 2013, 2016, 2019, or later

## Installation

1. Copy the script to your desired location
2. Open Exchange Management Shell (run as Administrator)
3. Navigate to the script location
4. Run the script with desired parameters

## Usage

### Basic Syntax

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 [-Server <ServerName>] [-OutputPath <Path>] [-ExportCSV] [-ExportHTML] [-ExportAll]
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-Server` | String | No | Specific Exchange server name. If omitted, all servers are queried |
| `-OutputPath` | String | No | Output directory for reports. Default is current directory |
| `-ExportCSV` | Switch | No | Export data to CSV files |
| `-ExportHTML` | Switch | No | Generate HTML report |
| `-ExportAll` | Switch | No | Export both CSV and HTML (default if no export option specified) |

### Examples

#### Example 1: Gather from all servers, export both formats
```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll
```

#### Example 2: Specific server, CSV only
```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EXCH01" -ExportCSV
```

#### Example 3: All servers, HTML report only
```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportHTML
```

#### Example 4: Custom output path
```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EXCH01" -OutputPath "C:\Reports\Exchange" -ExportAll
```

#### Example 5: Default behavior (exports both formats)
```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1
```

## Output Files

The script generates timestamped files to prevent overwrites:

### CSV Files
- `ReceiveConnectors_YYYYMMDD_HHMMSS.csv` - Receive connector details
- `SendConnectors_YYYYMMDD_HHMMSS.csv` - Send connector details
- `Certificates_YYYYMMDD_HHMMSS.csv` - Certificate information

### HTML Report
- `ExchangeConnectorsAndCertificates_YYYYMMDD_HHMMSS.html` - Comprehensive report with all data

## Data Collected

### Receive Connectors
- Server name
- Connector name and identity
- Bindings and remote IP ranges
- FQDN and banner
- Authentication mechanisms
- TLS settings
- Message size limits
- Protocol logging level
- Connection timeouts
- And more...

### Send Connectors
- Connector name and identity
- Enabled status
- Address spaces
- Source transport servers
- Smart hosts
- DNS routing configuration
- TLS settings
- Authentication mechanisms
- Message size limits
- And more...

### Certificates
- Server assignment
- Subject and issuer
- Friendly name
- Associated services
- Domain names (SANs)
- Validity period
- Days until expiration
- Expiration status (Valid/Warning/Expiring Soon/Expired)
- Self-signed status
- Public key size
- Signature algorithm
- And more...

## Certificate Status Indicators

The script automatically categorizes certificates:

| Status | Description |
|--------|-------------|
| **Expired** | Certificate has already expired |
| **Expiring Soon** | Certificate expires within 30 days |
| **Warning** | Certificate expires within 31-90 days |
| **Valid** | Certificate expires in more than 90 days |

## HTML Report Features

- **Summary Dashboard**: Quick overview with counts and alerts
- **Color-Coded Status**: Visual indicators for enabled/disabled states
- **Certificate Warnings**: Highlighted expired and expiring certificates
- **Sortable Data**: Organized tables for easy review
- **Professional Styling**: Clean, modern interface
- **Timestamp**: Report generation date and time
- **Target Information**: Shows which servers were queried

## Troubleshooting

### "Exchange Management Shell not detected"
- Ensure you're running the script from Exchange Management Shell
- Verify Exchange Management Tools are installed

### "Access Denied" or Permission Errors
- Run Exchange Management Shell as Administrator
- Verify you have Exchange Organization Management or Exchange View-Only Administrator rights

### No Data Returned
- Verify the server name is correct (use `Get-ExchangeServer` to list servers)
- Check that the target server is accessible
- Ensure Exchange services are running

### Output Path Errors
- Verify the path exists or the script has permissions to create it
- Use absolute paths for clarity (e.g., `C:\Reports` instead of `..\Reports`)

## Best Practices

1. **Regular Monitoring**: Schedule script execution to monitor certificate expiration
2. **Archive Reports**: Keep historical reports for compliance and auditing
3. **Certificate Alerts**: Review "Expiring Soon" certificates immediately
4. **Security**: Store reports in secure locations as they contain configuration details
5. **Documentation**: Use reports for disaster recovery documentation

## Script Workflow

1. Validates Exchange Management Shell availability
2. Determines target servers (specific or all)
3. Collects receive connector information
4. Collects send connector information
5. Collects certificate information
6. Generates summary statistics
7. Exports to requested format(s)
8. Displays completion status and warnings

## Performance Considerations

- Processing time depends on the number of servers and connectors
- Large environments may take several minutes
- Network latency affects remote server queries
- Consider targeting specific servers for faster execution

## Security Notes

- Script is read-only and makes no changes to Exchange configuration
- Reports may contain sensitive configuration information
- Protect output files with appropriate file system permissions
- Follow your organization's data handling policies

## Version History

- **1.0** - Initial release
  - Receive connector collection
  - Send connector collection
  - Certificate collection
  - CSV export
  - HTML report generation
  - Certificate expiration monitoring

## Support

For issues, questions, or feature requests, please refer to your organization's Exchange administrator documentation or Microsoft Exchange documentation.

## License

Use this script in accordance with your organization's policies and Microsoft Exchange licensing terms.
