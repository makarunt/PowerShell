# Invoke-SoaConversion.ps1

Tool for changing the Source of Authority (SOA) in a hybrid Microsoft 365 environment
(on-premises AD + Entra Connect Sync + Exchange Online + Cloud Sync EXO→AD writeback).

The script changes **only the SOA flag**:

| Category | Change | Check |
|---|---|---|
| `ExchangeAttributes` | `Set-Mailbox -Identity <ExternalDirectoryObjectId> -IsExchangeCloudManaged $true/$false` | `Get-Mailbox` → `IsExchangeCloudManaged` |
| `User` | `PATCH /users/{id}/onPremisesSyncBehavior` `{ "isCloudManaged": true/false }` | `GET /users/{id}/onPremisesSyncBehavior` |
| `Group` | `PATCH /groups/{id}/onPremisesSyncBehavior` | `GET /groups/{id}/onPremisesSyncBehavior` |
| `Contact` | `PATCH /contacts/{id}/onPremisesSyncBehavior` | `GET /contacts/{id}/onPremisesSyncBehavior` |

The script does **not** change tenant-wide settings (for example
`Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault` or
`blockCloudObjectTakeoverThroughHardMatchEnabled`), does not touch on-premises AD, Entra Connect or Cloud Sync,
does not change any other object attribute, and does not use app-only authentication (if an existing EXO or
Graph connection is app-only, the script stops).

---

## Prerequisites

| Item | Version / note |
|---|---|
| PowerShell | **7.2 or later**. Windows PowerShell 5.1 is not supported and not needed. |
| `ExchangeOnlineManagement` | 3.2.0+ (needed for `Get-ConnectionInformation`; `Set-Mailbox` must have `-IsExchangeCloudManaged`, which the script checks). |
| `Microsoft.Graph.Authentication` | 2.x. All Graph calls go through `Invoke-MgGraphRequest`. |
| Entra Connect Sync | **2.5.190.0+** for ExchangeAttributes (older versions try to push Exchange attributes of already transferred mailboxes to Entra ID and fail). At least 2.5.76.0 for User/Group SOA and **2.5.79.0 for Contact SOA**. Check: `(Get-ADSyncGlobalSettings).Parameters['Microsoft.Synchronize.ServerConfigurationVersion']` |
| Cloud Sync agent | At least 1.1.1370.0 for User/Group SOA; **1.1.1107.0+** for Exchange attribute writeback (configuration *EXO to AD attribute sync*). Cloud Sync runs next to Connect Sync, not instead of it. |
| Sign-in | Interactive (delegated). App-only / certificate authentication is not supported for `IsExchangeCloudManaged`. |

Checks at start-up:
- **PowerShell:** if the version is older than 7.2 (for example Windows PowerShell 5.1), the script prints the
  required version and how to install PowerShell 7 (`winget install --id Microsoft.PowerShell --source winget`),
  then exits with code 3 without any change and without creating an output folder. Run the script from `pwsh`,
  not from `powershell`.
- **Modules:** only the modules needed for the selected category are checked. If a module is missing or too old,
  the script asks `Install '<module>' from the PowerShell Gallery for the current user now? [Y/N]`.
  - `Y` installs the module (`Install-Module -Scope CurrentUser -Repository PSGallery -Force`, side by side with
    older versions if present) and continues.
  - Any other answer stops the script and prints the command for a manual installation.
  - If an older version of the module is already loaded in the same session, the script asks for a new
    PowerShell session.
  - The installation also runs when the script is started with `-WhatIf`, because it is a local change that
    you have just confirmed.

Manual installation, if you prefer to do it in advance:

```powershell
Install-Module ExchangeOnlineManagement -MinimumVersion 3.2.0 -Scope CurrentUser
```
```powershell
Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Scope CurrentUser
```

The script connects only to the services the category needs, EXO first and then Graph (this order reduces the
risk of MSAL library conflicts between the two modules). When both connections are active, it checks that they
point to the same tenant.

| Category | EXO | Graph |
|---|---|---|
| ExchangeAttributes | yes | no |
| User | no | yes |
| Group | for Inventory/Convert/Rollback when `-GroupType` is not `Security` (EXO properties of mail-enabled groups in the backup) | yes |
| Contact | for Inventory/Convert/Rollback (EXO properties of the contact in the backup) | yes |

## Roles and permissions

| Category | Entra / EXO role (minimum) | Graph delegated scopes requested by the script |
|---|---|---|
| ExchangeAttributes | Exchange Administrator (recommended by Microsoft), or Hybrid Identity Administrator, or Global Administrator | – |
| User | **Hybrid Administrator** | `User.Read.All`, `User-OnPremisesSyncBehavior.ReadWrite.All`, `RoleManagement.Read.Directory` (reads the members of the *Directory Synchronization Accounts* role, see *Protected system objects*) |
| Group | **Hybrid Administrator** (+ EXO read access for mail-enabled groups, for example View-Only Recipients) | `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All`, `Group-OnPremisesSyncBehavior.ReadWrite.All` |
| Contact | **Hybrid Administrator** (+ EXO read access to contacts) | `OrgContact.Read.All`, `Contacts-OnPremisesSyncBehavior.ReadWrite.All` |

- Both `GET` and `PATCH` on `onPremisesSyncBehavior` require `*-OnPremisesSyncBehavior.ReadWrite.All`, so
  Inventory requests it too.
- The first time, someone with the **Application Administrator** or **Cloud Application Administrator** role
  must grant consent for these scopes to the *Microsoft Graph Command Line Tools* application (Entra admin
  center → Enterprise applications → Microsoft Graph Command Line Tools → Permissions → Grant admin consent).
  If a scope has not been granted, the script stops before reading or changing anything.

## Parameters

```
-Action        Inventory | Convert | Verify | Rollback       (required)
-Category      ExchangeAttributes | User | Group | Contact   (required)
-InputCsv      required for Convert/Verify/Rollback; optional for Inventory
-GroupType     Security | MailEnabledSecurity | Distribution | All   (Category=Group only, default All)
-OutputFolder  default .\SoaOutput
-MaxObjects    process at most N objects (first N rows of the CSV / first N eligible objects for Inventory)
-Delimiter     default ';'
-ExpectedState CloudManaged (default) | OnPremManaged   (Verify only)
-WhatIf / -Confirm
```

`-ExpectedState` was not part of the original specification. It was added so that a stand-alone Verify can also
be used after a rollback.

Parameter values are case-insensitive (`-Category user` works) and are normalized to the canonical names used in
folder names and CSV columns.

For Group, `-GroupType` applies to every action: a group of another type in the CSV is skipped with the note
`GroupTypeNotSelected`. Microsoft 365 (Unified) groups are never eligible.

## Protected system objects

Some objects are never changed by the script, neither by Convert nor by Rollback. Such an object gets the value
`ProtectedSystemObject: <reason>` in the `Note` column. Inventory without a CSV leaves them out and reports how
many it found; with a CSV they are shown with the reason, and Convert and Rollback skip them.

| Rule | What it catches |
|---|---|
| Members of the Entra role **Directory Synchronization Accounts** (template `d29b2b05-8046-44ba-8758-1e26182fcf32`) | Entra Connect / Cloud Sync service accounts, whatever their name. Category User only. If the role members cannot be read, the script stops (fail closed). |
| Name patterns (UPN, alias/mailNickname, display name, on-premises SAM/UPN) | `Sync_*`, `ADToAADSyncServiceAccount*`, *On-Premises Directory Synchronization Service Account*, `MSOL_*`, `AAD_*`, Administrator, Guest, krbtgt, DefaultAccount, `SUPPORT_388945a0*`, Exchange system mailboxes (`HealthMailbox*`, `SystemMailbox{*`, `DiscoverySearchMailbox*`, `FederatedEmail.*`, `Migration.8f3e7716*`, `SM_*`, `CAS_{*`) |
| Built-in groups (exact name) | Exchange RBAC groups (*Organization Management*, *Recipient Management*, *Exchange Servers*, *Exchange Trusted Subsystem*, …), *DnsAdmins*, *DnsUpdateProxy* |
| On-premises SID with a well-known RID (< 1000) | Built-in users and groups (Administrator 500, krbtgt 502, Domain Admins 512, …), even when renamed. |

The *Directory Synchronization Accounts* role is Microsoft's recommended way to find sync accounts
(*Entra Connect: Accounts and permissions*). A false match only skips an object, so the lists are deliberately
broad. The lists are at the top of the script (`$ProtectedNamePatterns`, `$ProtectedGroupNames`) and can be
extended if you find another system account.

## Which scenario to use

| Scenario | Convert | Do NOT convert |
|---|---|---|
| **Remove on-premises Exchange** (AD stays the source for identities) | Exchange attributes of mailboxes, distribution groups (DL), mail-enabled security groups (MESG), mail contacts | users, security groups (not mail-enabled) |
| **Cloud only** (the customer wants to remove AD completely) | users, security groups, DL, MESG, mail contacts | Exchange attributes of mailboxes, because the User conversion covers them |

Users and security groups are converted **only** in the cloud-only scenario. A User conversion moves the whole
user object to the cloud, including the Exchange attributes of the mailbox (Microsoft: User SOA *"clears the
user's IsExchangeCloudManaged state"*), so the ExchangeAttributes conversion is not needed in that scenario. Just
removing Exchange does not require a User conversion; it also affects sign-in and passwords, because AD no longer
manages the user.

In the cloud-only scenario, shared, room and equipment mailboxes must also be in the User CSV, because they have a
user object in AD as well. A phased approach is fine too: if Exchange is removed first (ExchangeAttributes) and the
customer moves to cloud only later (User), the User conversion simply clears `IsExchangeCloudManaged`.

### Commands for removing Exchange

For groups, **always specify `-GroupType`**. The default is `All`, which includes security groups.

Inventory:

```powershell
.\Invoke-SoaConversion.ps1 -Action Inventory -Category ExchangeAttributes
.\Invoke-SoaConversion.ps1 -Action Inventory -Category Group -GroupType Distribution
.\Invoke-SoaConversion.ps1 -Action Inventory -Category Group -GroupType MailEnabledSecurity
.\Invoke-SoaConversion.ps1 -Action Inventory -Category Contact
```

Review the CSV files in the output folders and delete the rows you do not want to convert.

Convert, in the same order for every category: `-WhatIf`, then one pilot object, then the rest:

```powershell
.\Invoke-SoaConversion.ps1 -Action Convert -Category ExchangeAttributes -InputCsv .\mailboxes.csv -WhatIf
.\Invoke-SoaConversion.ps1 -Action Convert -Category ExchangeAttributes -InputCsv .\mailboxes.csv -MaxObjects 1
.\Invoke-SoaConversion.ps1 -Action Convert -Category ExchangeAttributes -InputCsv .\mailboxes.csv
```
```powershell
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType Distribution -InputCsv .\dl.csv -WhatIf
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType Distribution -InputCsv .\dl.csv -MaxObjects 1
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType Distribution -InputCsv .\dl.csv
```
```powershell
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType MailEnabledSecurity -InputCsv .\mesg.csv -WhatIf
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType MailEnabledSecurity -InputCsv .\mesg.csv -MaxObjects 1
.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -GroupType MailEnabledSecurity -InputCsv .\mesg.csv
```
```powershell
.\Invoke-SoaConversion.ps1 -Action Convert -Category Contact -InputCsv .\contacts.csv -WhatIf
.\Invoke-SoaConversion.ps1 -Action Convert -Category Contact -InputCsv .\contacts.csv -MaxObjects 1
.\Invoke-SoaConversion.ps1 -Action Convert -Category Contact -InputCsv .\contacts.csv
```

Notes:
- For mailboxes that were migrated or changed on-premises (`Set-RemoteMailbox`) recently, Microsoft requires a sync
  cycle plus 24 hours before the conversion. This does not apply to mailboxes migrated several days ago.
- After the conversion, DLs, MESGs and contacts are managed only in EXO and have no writeback to AD.
- Writeback of mailbox Exchange attributes to AD requires Cloud Sync (*EXO to AD attribute sync*).

## Recommended order of work

1. **Inventory**: `.\Invoke-SoaConversion.ps1 -Action Inventory -Category Group -GroupType Security`
2. **Review the CSV**: check the columns `CurrentSoaState`, `EligibleForConvert`, `Note`, `SoaNotes` (groups).
   Delete the rows you do not want to convert (see *CSV and Excel* below, because of the encoding).
3. **Convert with `-WhatIf`**: validates the CSV, takes a fresh backup and shows what would change. Nothing is
   changed.
   `.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -InputCsv .\groups.csv -WhatIf`
4. **Convert with `-MaxObjects 1`**: one pilot object.
5. **Post-check the pilot**: the automatic Verify in the same run, then manually:
   - Entra admin center: the object's source is no longer "Windows Server AD DS"; Audit log activity
     *Change Source of Authority from AD to cloud*.
   - After the next Connect Sync cycle: Event ID **6956** on the Connect server for that object (the change from
     AD is blocked).
   - User: a test sign-in of the pilot user (see the warnings below).
   - ExchangeAttributes: change an Exchange attribute in EXO and check the writeback to AD.
6. **Convert a small batch** (for example `-MaxObjects 10`), then post-check again.
7. **Convert the rest** of the same CSV. Objects that are already converted are skipped as `AlreadyCloudManaged`.
8. **Verify** later as needed: `-Action Verify -InputCsv <any CSV with ObjectId>`.
9. **Rollback** only when needed (see below).

Nested groups: SOA is not inherited by child groups, so each group is converted separately. Microsoft recommends
starting at the lowest level. The script processes the groups in the order of the CSV, so put child groups
above their parents if you want to follow that recommendation. The columns `IsNested`, `ParentGroups` and
`ChildGroups` in the inventory show the direct nesting (one level up and down).

## What happens in each action

**Inventory**
- Without `-InputCsv`: Graph query `$filter=onPremisesSyncEnabled eq true` with `ConsistencyLevel: eventual` and
  `$count=true`, with paging via `@odata.nextLink`. For mailboxes `Get-Mailbox -Filter 'IsDirSynced -eq $true'`.
  The report contains the eligible objects plus any object that could not be read (`Note` starts with
  `Error:`); the objects that were left out are summarized by reason in the log.
- With `-InputCsv`: one row per ObjectId; ineligible and missing objects get the reason in the `Note` column.
- `isCloudManaged` is read **per object** (`GET .../onPremisesSyncBehavior`), because it cannot be retrieved with
  a list query (see *Notes from the Microsoft documentation*). In large tenants this means one extra call per
  object (two for users, because of the manager).

**Convert / Rollback**
1. The CSV must have the columns `ObjectId` and `Category`, and `Category` must equal `-Category` in **every**
   row. Otherwise the script stops. An invalid GUID in any row also stops the script. Duplicate and empty rows
   are ignored (with a warning).
2. A fresh inventory serves as the backup (`Backup_Inventory_<Category>.csv`, for groups also
   `Backup_GroupMemberships.csv`). If the backup cannot be written, the script stops before any change.
3. Eligibility is checked again:
   - Convert skips `AlreadyCloudManaged`, `NotFound`, `NotSynced`, `GroupTypeNotSelected`, protected objects and
     read errors.
   - Rollback changes only objects that are currently `CloudManaged`. It skips groups with cloud members
     (`CloudMemberCount > 0`), because Microsoft requires cloud references to be removed before a rollback.
     With nested groups this means one run per level, with a sync cycle in between.
4. Category warning, summary and confirmation by typing `YES` (case-sensitive). **User Convert** asks for two
   confirmations. With `-WhatIf` there is no confirmation and nothing is changed.
5. Changes are made one object at a time (`ShouldProcess`; with `-Confirm` it asks for every object). An error on
   one object does not stop the others.
6. Graph 429 and 500/502/503/504: up to 6 attempts, using `Retry-After` when the service returns it, otherwise
   exponential backoff (at most 60 s).
7. Automatic Verify: up to 5 checks with a 20 s pause (eventual consistency).

**Verify** (stand-alone): reads the current state of every ObjectId in the CSV, without retries. `StateBefore` is
taken from the `CurrentSoaState` or `StateAfter` column of the input CSV, if present.

## Output files

Every run creates `OutputFolder\yyyyMMdd_HHmmss_<Action>_<Category>` (with a suffix `_1`, `_2`, … if the folder
already exists). Existing files are never overwritten.

| File | When |
|---|---|
| `Transcript.log` | always |
| `Inventory_<Category>.csv` | Inventory |
| `GroupMemberships.csv` | Inventory, Group |
| `Backup_Inventory_<Category>.csv` (+ `Backup_GroupMemberships.csv`) | Convert, Rollback |
| `Results_<Action>_<Category>.csv` | Convert, Rollback, Verify |

- The key in every CSV is `ObjectId` (Entra object ID; `ExternalDirectoryObjectId` for mailboxes).
- Every CSV has `Category` and `CurrentSoaState`. The results CSV has the columns
  `ObjectId; DisplayName; Category; StateBefore; StateAfter; Status; ErrorMessage; Timestamp` and, at the end,
  `CurrentSoaState` (= `StateAfter`), so it can be used directly as input for Verify or Rollback.
- Multiple values in one cell are separated by `|`. Complex objects (for example contact addresses) are written as
  compact JSON, contact phones as `type:number`.

`CurrentSoaState` values:

| Value | Meaning |
|---|---|
| `OnPremManaged` | Graph: `onPremisesSyncEnabled=true`, `isCloudManaged=false`. EXO: `IsDirSynced=true`, `IsExchangeCloudManaged=false`. Eligible for Convert. |
| `CloudManaged` | `isCloudManaged=true` / `IsExchangeCloudManaged=true`. |
| `NotSynced` | Graph: `isCloudManaged=false` and `onPremisesSyncEnabled` is not `true` (a cloud-only object **or** a rollback waiting for the next sync cycle). EXO: `IsDirSynced=false`. |
| `NotFound` | No object with this ObjectId exists. |
| `Unknown` | Read error (the reason is in `Note` / `ErrorMessage`). |

Backup columns per category:
- **ExchangeAttributes**: all requested attributes, `UserCertificate`/`UserSMimeCertificate` only as
  `HasUserCertificate`/`HasUserSMimeCertificate` (True/False), plus `Guid` and `ExchangeGuid`. It also contains
  the other attributes that Microsoft's table lists as editable in EXO after the transfer (ArchiveName, Audit*,
  SingleItemRecoveryEnabled, LitigationHoldDate/Owner, retention hold dates, RetentionComment/Url,
  MailTipTranslations, moderation, Resource*, ImmutableId) and `AssistantName` from `Get-User`. **Not in the
  backup:** mailbox permissions (`Get-MailboxPermission`) and junk e-mail settings; export them separately if you
  change them in the cloud. The on-premises DN, SAM and ImmutableId are not available in EXO, and the script does
  not connect to Graph for this category. Find the object in AD by UPN or PrimarySmtpAddress.
- **User**: all requested attributes, `managerUserPrincipalName`, `extensionAttribute1-15`, all `onPremises*`.
- **Group**: the requested Graph attributes, `extensionAttribute1-15`, counts of members, cloud members, child and
  parent groups, owners, `SoaNotes`, and `Exo_*` columns for mail-enabled groups. Graph v1.0 has no
  `onPremisesDistinguishedName` or `onPremisesImmutableId` for groups; use `onPremisesSamAccountName` /
  `onPremisesSecurityIdentifier` to find the group in AD.
- **Contact**: the requested attributes plus `Exo_ExternalEmailAddress`, `Exo_HiddenFromAddressListsEnabled`,
  `Exo_EmailAddresses`. Graph has no on-premises DN for contacts.

## CSV and Excel

CSV files are written with `-Encoding Unicode` (UTF-16 LE with BOM), and import reads the same encoding. The
script rejects an input CSV without a UTF-16 LE BOM (bytes `FF FE`) instead of silently reading wrong characters.
The same format is called *UTF-16 LE* in Notepad and VS Code, *UTF-16 LE BOM* (older versions: *UCS-2 LE BOM*) in
Notepad++, and *Unicode Text* in Excel.

**Known Excel behaviour (confirmed):** Excel opens UTF-16 files as *Unicode Text*, which for Excel means **tab**
as the separator. Diacritics are displayed correctly, but on double-click all columns end up in column A.
Recommendations:
- To view: Excel → *Data → From Text/CSV*, File Origin *Unicode (UTF-16 LE)*, Delimiter *Semicolon*, or split the
  column with *Text to Columns*.
- To delete rows before a Convert, use a text editor (Notepad, VS Code) and save as UTF-16 LE, or filter with
  `Import-Csv` / `Export-Csv -Encoding Unicode -Delimiter ';'` in PowerShell. A file saved from Excel as
  *Unicode Text* is tab-delimited; run the script with ``-Delimiter "`t"`` in that case. A file saved as
  *CSV UTF-8* or *CSV (comma delimited)* is rejected because it is not UTF-16. The script reports clearly when it
  cannot find the `ObjectId` column or the encoding is wrong.
- Never save backup CSV files from Excel: Excel can change values (leading zeros, long numbers, dates).
- If UTF-8 with BOM suits you better (Excel with Croatian regional settings opens it correctly on double-click,
  with `;`), change one variable at the top of the script: `$CsvEncoding = 'utf8BOM'`. Import must then use the
  same encoding, and the BOM check in `Read-InputCsv` has to be adapted.

## Warnings shown by the script

**ExchangeAttributes Convert** (according to Microsoft):
- After the last on-premises change (`Set-RemoteMailbox`) or mailbox move, a full sync cycle plus 24 hours must
  pass (for example a 6-hour sync means 30 hours in total).
- After a mailbox move: first confirm that the move has completed and that the on-premises object is a
  RemoteMailbox. Then wait until the post-move attributes have synchronized to Entra ID, and only then start the
  24 hours. Converting too early can let writeback overwrite the correct remote-mailbox values in AD with stale
  cloud recipient-type values.
- The script cannot verify this.
- Writeback to AD (Cloud Sync) exists only for CustomAttribute1-15, ExtensionCustomAttribute1-5, mail
  (WindowsEmailAddress), proxyAddresses and the recipient type. Everything else (for example
  HiddenFromAddressListsEnabled, forwarding, moderation) stays in the cloud only and reverts to the AD value after
  a rollback.
- Before moving a mailbox back on-premises (offboarding), run a Rollback first (`IsExchangeCloudManaged = $false`);
  otherwise synchronization breaks during offboarding.

**User Convert**:
- The user becomes cloud authoritative and changes from AD are no longer applied (Connect Sync logs event 6956).
- The conversion clears the `IsExchangeCloudManaged` state.
- Sign-in and password, according to Microsoft's documentation:
  - users federated through **AD FS are not supported**,
  - users must not depend on on-premises password authentication (LDAP bind, Kerberos with a password),
  - passwordless authentication is recommended (Windows Hello for Business with **Cloud Kerberos Trust**, FIDO2),
  - with third-party federation, the AD account and password must be maintained manually.
- All mailboxes must be in EXO before User SOA.
- The `onPremises*` attributes stay on the cloud object and must be maintained manually if access to on-premises
  resources is still needed.

**Group Convert**: SOA is not inherited by nested groups. Distribution groups and mail-enabled security groups are
then managed only in EXO and have no writeback to AD. `extensionAttribute1-15` are not supported on converted
groups.

**Contact Convert**: requires Connect Sync 2.5.79.0+.

**Rollback (all categories)**: the next sync overwrites the cloud values with the values from AD. Values worth
keeping are in the backup CSV (the path is printed). The script does not restore attributes automatically. The
rollback is complete only after the next Connect Sync cycle.

**User Rollback, additionally**: according to Microsoft, the tenant feature
`blockCloudObjectTakeoverThroughHardMatchEnabled` must be disabled before a user rollback
(`PATCH /beta/directory/onPremisesSynchronization/{id}`) and re-enabled after Connect Sync has taken over the
object. This is a tenant-wide setting, so the script does **not** change it; it only warns. Cloud references
(memberships in converted groups, access packages) must also be removed.

## Notes from the Microsoft documentation

Checked on learn.microsoft.com on 29 October 2026:
- *Cloud-based management of Exchange attributes for Remote Mailboxes in hybrid environments* (updated
  2026-09-28),
- *Configure user SOA* (2026-04-15), *Configure Group SOA* (2025-11-19), *Transfer user SOA* (2026-09-24),
- Graph *Get/Update onPremisesSyncBehavior* and *onPremisesSyncBehavior resource type*,
- *Decommission the last Exchange Server after transferring SOA to cloud*,
- *Microsoft Entra built-in roles* and *Entra Connect: Accounts and permissions* (Directory Synchronization
  Accounts role).

The Microsoft blog on techcommunity.microsoft.com was not reachable, but the official Exchange page covers
everything that was needed.

1. **v1.0 or beta.** The Entra how-to articles (User/Group/Contact SOA) use `https://graph.microsoft.com/v1.0/...`.
   The Graph API reference documents `onPremisesSyncBehavior` (get/update/resource) **under beta only**; even when
   v1.0 is selected on learn.microsoft.com, the page shows the beta content. The script uses v1.0 and switches to
   beta for the SOA endpoint only if v1.0 returns *"Resource not found for the segment"*. The log shows which
   version was used.
2. **State after conversion.** `onPremisesSyncEnabled` does **not** stay `true` after the conversion; it becomes
   **`null`**. Microsoft's table:

   | Step | isCloudManaged | onPremisesSyncEnabled |
   |---|---|---|
   | Object synchronized from AD | false | true |
   | SOA transferred to the cloud | true | null |
   | Rollback | false | null (until the sync client takes over the object) |
   | Cloud-native object | false | null |

   Consequences: Inventory without a CSV does not see converted objects, Verify relies only on `isCloudManaged`,
   and a rollback looks the same as a cloud-native object until sync takes over the object (state `NotSynced`).
3. **`$select` and `$expand`.** According to the API reference, GET `onPremisesSyncBehavior` **does not support
   query parameters** (the Entra articles still use `?$select=isCloudManaged`). The script therefore does not send
   `$select`; the response contains only `id` and `isCloudManaged` anyway. `$expand` on list queries is not
   documented, and advanced queries do not support `$expand`. Beta documents only
   `$filter=onPremisesSyncBehavior/isCloudManaged eq true` with an advanced query, which does not return the value
   per object. That is why `isCloudManaged` is read per object.
4. **Password and sign-in after User SOA.** The documentation does not explicitly describe what happens to
   password hash sync. It lists the prerequisites: no AD FS federation, no applications that authenticate with a
   password, passwordless with Cloud Kerberos Trust. The script's warning repeats exactly that and recommends a
   test with a pilot user.
5. **User rollback** requires disabling `blockCloudObjectTakeoverThroughHardMatchEnabled` (tenant-wide). The script
   does not touch that setting.
6. **Groups:** cloud members must be removed before a rollback; the script skips such groups. `extensionAttribute1-15`
   are not supported on converted groups. Distribution groups and mail-enabled security groups cannot be edited
   through Graph or the Entra admin center after the conversion, only in EXO.
7. **Connect Sync versions:** at least 2.5.76.0 (User/Group) and 2.5.79.0 (Contact).
8. **ExchangeAttributes.** The documentation confirms:
   - the sync cycle + 24 h rule (with extra rules for mailbox moves, see *Warnings*),
   - that `-IsExchangeCloudManaged` *"should not be used together with any other parameters"*,
   - that certificate-based app-only is not supported (*"an app-only session doesn't provide the delegated user
     authorization required to update the IsExchangeCloudManaged property in Microsoft Entra ID"*),
   - that User SOA *"clears the user's IsExchangeCloudManaged state"*.

   Also from the documentation:
   - the flag applies only to mailboxes of users with `IsDirSynced = true`; mail-enabled groups and mail contacts
     use Group/Contact SOA,
   - after User SOA `IsDirSynced` becomes `False` (the script shows such a mailbox as `NotSynced`),
   - writeback to AD covers only some attributes (see *Warnings*).

   Microsoft's examples use `-Identity <User>` / `Alias`; the documentation does not say explicitly that
   `-Identity` accepts `ExternalDirectoryObjectId`. The script uses it as specified, and it worked in testing
   against a real tenant.

## Testing

- **Stub modules:** Inventory (all categories), Convert, Rollback, Verify, `-WhatIf`, `-MaxObjects`, category
  mismatch, non-UTF-16 CSV, cancellation, retry on 429, the v1.0 → beta fallback, an error on one object while
  the others continue, a failed verify, the PowerShell version check, module installation, and protected system
  objects.
- **Real tenant (in progress):** inventory for all categories, input validation, `-WhatIf`, Convert and Rollback
  for a test mailbox (ExchangeAttributes, including writeback), nested security groups, a distribution group and a
  mail-enabled security group. Graph v1.0 works for `onPremisesSyncBehavior` in that tenant.
- **Not yet tested:** Contact and User against a real tenant, Windows PowerShell 5.1 (the version message), and
  PSScriptAnalyzer. Run PSScriptAnalyzer before production and follow the order above, starting with one pilot
  object.
