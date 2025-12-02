# CLAUDE.md - PowerShell Repository Guide for AI Assistants

**Last Updated:** 2025-12-02
**Repository:** PowerShell
**Purpose:** Comprehensive guide for AI assistants working with PowerShell codebases

---

## Table of Contents

1. [Repository Overview](#repository-overview)
2. [Codebase Structure](#codebase-structure)
3. [Technology Stack](#technology-stack)
4. [Development Workflows](#development-workflows)
5. [Build System](#build-system)
6. [Testing Guidelines](#testing-guidelines)
7. [Code Conventions](#code-conventions)
8. [Git Workflow](#git-workflow)
9. [Key Files and Locations](#key-files-and-locations)
10. [Common Tasks](#common-tasks)
11. [AI Assistant Guidelines](#ai-assistant-guidelines)

---

## Repository Overview

### About PowerShell

PowerShell is a cross-platform task automation solution made up of a command-line shell, a scripting language, and a configuration management framework. This repository typically contains:

- PowerShell Core engine (C#/.NET)
- Built-in cmdlets and modules
- PowerShell Standard Library
- Test infrastructure
- Documentation and examples

### Repository State

**Current Status:** Repository initialization phase
**Primary Language:** C# (.NET)
**Secondary Language:** PowerShell
**License:** Typically MIT (verify LICENSE file when present)

---

## Codebase Structure

### Expected Directory Layout

```
PowerShell/
├── src/                          # Source code
│   ├── Microsoft.PowerShell.*   # PowerShell assemblies
│   ├── System.Management.Automation/  # Core engine
│   ├── powershell-native/       # Native code components
│   └── Modules/                 # Built-in modules
├── test/                        # Test suites
│   ├── powershell/             # PowerShell (Pester) tests
│   ├── xUnit/                  # xUnit tests for C#
│   └── common/                 # Common test utilities
├── tools/                       # Build and development tools
│   ├── packaging/              # Packaging scripts
│   ├── releaseBuild/           # Release build automation
│   └── install-powershell.ps1  # Installation scripts
├── docs/                        # Documentation
│   ├── building/               # Build instructions
│   ├── cmdlet-example/         # Cmdlet development examples
│   └── maintainers/            # Maintainer guides
├── assets/                      # Images, icons, resources
├── demos/                       # Demo scripts and examples
├── docker/                      # Docker configurations
└── scripts/                     # Utility scripts

### Core Components

#### Engine (System.Management.Automation)
- **Location:** `src/System.Management.Automation/`
- **Purpose:** Core PowerShell runtime engine
- **Key Areas:**
  - Parser and language components
  - Runspace management
  - Cmdlet infrastructure
  - Pipeline execution
  - Remoting infrastructure

#### Host Executables
- **pwsh:** Main PowerShell executable
- **Location:** `src/Microsoft.PowerShell.ConsoleHost/`
- Platform-specific hosting

#### Native Components
- **Location:** `src/powershell-native/`
- **Languages:** C/C++
- **Purpose:** Platform-specific native code

---

## Technology Stack

### Primary Technologies

| Technology | Usage | Version |
|------------|-------|---------|
| .NET | Runtime platform | .NET 6.0+ |
| C# | Primary language | C# 10+ |
| PowerShell | Scripting/Testing | 7.x |
| CMake | Native build system | 3.x+ |
| MSBuild | .NET build system | Latest |

### Development Tools

- **Build:** dotnet CLI, MSBuild, CMake
- **Testing:** xUnit, Pester
- **Package Management:** NuGet, PowerShellGet
- **CI/CD:** GitHub Actions, Azure Pipelines
- **Code Analysis:** PSScriptAnalyzer, Roslyn analyzers

### Dependencies

- .NET runtime and SDK
- Platform-specific dependencies (see docs/building/)
- NuGet packages (defined in .csproj files)

---

## Development Workflows

### Setting Up Development Environment

```bash
# 1. Clone repository (when populated)
git clone <repository-url>
cd PowerShell

# 2. Install prerequisites
# - .NET SDK 6.0 or higher
# - CMake 3.15 or higher
# - Platform-specific compilers

# 3. Bootstrap the build
./build.ps1 -Bootstrap

# 4. Build PowerShell
./build.ps1

# 5. Run tests
./build.ps1 -Test
```

### Development Cycle

1. **Create Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Changes**
   - Edit C# code in `src/`
   - Update tests in `test/`
   - Update documentation in `docs/`

3. **Build and Test Locally**
   ```bash
   ./build.ps1
   ./build.ps1 -Test
   ```

4. **Run Code Analysis**
   ```bash
   ./build.ps1 -PSScriptAnalyzer
   ```

5. **Commit and Push**
   ```bash
   git add .
   git commit -m "Descriptive message"
   git push origin feature/your-feature-name
   ```

---

## Build System

### Build Scripts

#### Primary Build Script: `build.ps1`

```powershell
# Full build
./build.ps1

# Build with specific configuration
./build.ps1 -Configuration Release

# Build specific targets
./build.ps1 -Clean
./build.ps1 -Restore
./build.ps1 -Build

# Run tests
./build.ps1 -Test

# Create packages
./build.ps1 -Package
```

#### Common Build Parameters

- `-Configuration`: Debug or Release
- `-Runtime`: Target runtime (win-x64, linux-x64, osx-arm64, etc.)
- `-Clean`: Clean build artifacts
- `-Restore`: Restore NuGet packages
- `-PSModuleRestore`: Restore PowerShell modules
- `-CI`: CI build mode

### Build Outputs

- **Binaries:** `src/powershell-*/bin/`
- **Packages:** `bin/packages/`
- **Test Results:** `test/results/`

### Project Files

- **Solution:** `PowerShell.sln`
- **Projects:** `src/**/*.csproj`
- **Native:** `src/powershell-native/CMakeLists.txt`

---

## Testing Guidelines

### Test Framework Organization

#### xUnit Tests (C#)
- **Location:** `test/xUnit/`
- **Purpose:** Unit tests for C# code
- **Run:** `dotnet test`

```csharp
// Example xUnit test
[Fact]
public void TestCmdletExecution()
{
    using (PowerShell ps = PowerShell.Create())
    {
        var results = ps.AddCommand("Get-Process").Invoke();
        Assert.NotEmpty(results);
    }
}
```

#### Pester Tests (PowerShell)
- **Location:** `test/powershell/`
- **Purpose:** Integration and cmdlet tests
- **Run:** `Invoke-Pester`

```powershell
# Example Pester test
Describe "Get-Process Tests" {
    It "Should return processes" {
        $processes = Get-Process
        $processes | Should -Not -BeNullOrEmpty
    }
}
```

### Test Categories

- **CI:** Fast tests run on every commit
- **Feature:** Feature-specific tests
- **Scenario:** End-to-end scenario tests
- **Slow:** Performance or long-running tests

### Running Tests

```bash
# All tests
./build.ps1 -Test

# Specific test suite
dotnet test test/xUnit/csharp/test_Runspace.csproj

# Pester tests
pwsh -c "Invoke-Pester test/powershell"

# Specific tag
pwsh -c "Invoke-Pester -Tag CI"
```

---

## Code Conventions

### C# Coding Standards

#### Naming Conventions

```csharp
// Classes: PascalCase
public class RunspaceManager { }

// Methods: PascalCase
public void ExecuteCommand() { }

// Private fields: camelCase with underscore
private readonly string _commandName;

// Properties: PascalCase
public string CommandName { get; set; }

// Constants: PascalCase
public const int MaxRetries = 3;

// Namespaces: Microsoft.PowerShell.*
namespace Microsoft.PowerShell.Commands { }
```

#### Code Style

- **Indentation:** 4 spaces (no tabs)
- **Braces:** Allman style (opening brace on new line)
- **Line Length:** Prefer < 120 characters
- **Using Directives:** Sort alphabetically, System.* first

```csharp
// Good example
public class CmdletExample : PSCmdlet
{
    protected override void ProcessRecord()
    {
        if (ShouldProcess(Target, Action))
        {
            WriteObject(result);
        }
    }
}
```

### PowerShell Coding Standards

#### Script Style

```powershell
# Function names: Verb-Noun (approved verbs only)
function Get-ProcessInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    # PascalCase for parameters
    # Use proper indentation (4 spaces)
    # Add comment-based help
}

# Variables: camelCase or PascalCase
$processName = "pwsh"
$ProcessId = $PID

# Constants: PascalCase
$MaximumRetryCount = 5
```

#### PSScriptAnalyzer

All PowerShell scripts must pass PSScriptAnalyzer checks:

```powershell
Invoke-ScriptAnalyzer -Path . -Recurse -Settings PSGallery
```

### EditorConfig

Follow `.editorconfig` settings when present:
- Encoding: UTF-8
- Line endings: LF (Unix-style)
- Trim trailing whitespace
- Insert final newline

---

## Git Workflow

### Branch Strategy

- **main/master:** Stable release branch
- **feature/*:** New features
- **fix/*:** Bug fixes
- **refactor/*:** Code refactoring
- **docs/*:** Documentation updates
- **claude/*:** AI assistant working branches

### Commit Message Format

```
<type>: <short summary>

<optional detailed description>

<optional footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting)
- `refactor`: Code refactoring
- `test`: Adding or updating tests
- `chore`: Build system or tooling changes
- `perf`: Performance improvements

**Examples:**

```
feat: Add support for ARM64 architecture

Implements ARM64 compilation and runtime support for macOS and Linux.
Includes platform detection and native library loading.

Fixes #12345
```

```
fix: Correct null reference in pipeline execution

Adds null check before accessing pipeline object properties to prevent
NullReferenceException in edge cases.
```

### Pull Request Guidelines

1. **Title:** Clear, descriptive summary
2. **Description:** What, why, and how
3. **Tests:** Include relevant tests
4. **Documentation:** Update if needed
5. **Changelog:** Add entry if user-facing
6. **Review:** Address feedback promptly

---

## Key Files and Locations

### Configuration Files

| File | Purpose |
|------|---------|
| `.editorconfig` | Editor configuration |
| `.gitignore` | Git ignore patterns |
| `.gitattributes` | Git attributes (line endings, etc.) |
| `global.json` | .NET SDK version |
| `Directory.Build.props` | MSBuild common properties |
| `Directory.Build.targets` | MSBuild common targets |
| `nuget.config` | NuGet package sources |

### Build Configuration

| File | Purpose |
|------|---------|
| `PowerShell.sln` | Visual Studio solution |
| `build.psm1` | Build module |
| `tools/packaging/packaging.psd1` | Packaging configuration |
| `tools/install-powershell.ps1` | Installation script |

### Documentation

| Location | Content |
|----------|---------|
| `README.md` | Repository overview |
| `CONTRIBUTING.md` | Contribution guidelines |
| `CODE_OF_CONDUCT.md` | Code of conduct |
| `LICENSE.txt` | License information |
| `docs/building/` | Build instructions |
| `docs/cmdlet-example/` | Cmdlet development guide |

---

## Common Tasks

### Adding a New Cmdlet

```bash
# 1. Create cmdlet class in appropriate assembly
# Location: src/Microsoft.PowerShell.Commands.*/

# 2. Implement cmdlet
[Cmdlet(VerbsCommon.Get, "Example")]
public class GetExampleCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string Name { get; set; }

    protected override void ProcessRecord()
    {
        WriteObject($"Example: {Name}");
    }
}

# 3. Add tests
# - xUnit: test/xUnit/csharp/test_Example.cs
# - Pester: test/powershell/Modules/Example.Tests.ps1

# 4. Update documentation
# - Help XML in src/
# - External help in docs/
```

### Debugging PowerShell

```bash
# Debug build
./build.ps1 -Configuration Debug

# Launch with debugger
# - Visual Studio: Open PowerShell.sln, F5
# - VS Code: Use launch.json configuration
# - Command line: dotnet run --project src/powershell-*/
```

### Updating Dependencies

```bash
# Update NuGet packages
dotnet restore
dotnet list package --outdated

# Update specific package
dotnet add package <PackageName> --version <Version>

# Update PowerShell modules
./build.ps1 -PSModuleRestore
```

### Running Specific Tests

```bash
# Single xUnit test class
dotnet test --filter FullyQualifiedName~TestClassName

# Single Pester test file
Invoke-Pester -Path test/powershell/Modules/Example.Tests.ps1

# Tests matching pattern
Invoke-Pester -Tag "CI" -ExcludeTag "Slow"
```

---

## AI Assistant Guidelines

### General Principles

1. **Read Before Modifying**
   - Always read existing code before making changes
   - Understand the current implementation and patterns
   - Check for related tests and documentation

2. **Follow Existing Patterns**
   - Match coding style of surrounding code
   - Use established patterns for cmdlets, error handling, etc.
   - Maintain consistency with the codebase

3. **Test-Driven Development**
   - Write or update tests for all changes
   - Run tests before committing
   - Ensure both xUnit and Pester tests pass

4. **Minimal Changes**
   - Make focused, targeted changes
   - Avoid unnecessary refactoring
   - Don't modify unrelated code

### PowerShell-Specific Guidelines

#### Cmdlet Development

```csharp
// Always use approved verbs (Get, Set, New, Remove, etc.)
[Cmdlet(VerbsCommon.Get, "Example")]
[OutputType(typeof(ExampleObject))]
public class GetExampleCommand : PSCmdlet
{
    // Use proper parameter attributes
    [Parameter(
        Mandatory = true,
        Position = 0,
        ValueFromPipeline = true,
        ValueFromPipelineByPropertyName = true)]
    [ValidateNotNullOrEmpty()]
    public string Name { get; set; }

    // Implement ShouldProcess for -WhatIf/-Confirm
    protected override void ProcessRecord()
    {
        if (ShouldProcess(Name, "Get example"))
        {
            WriteObject(new ExampleObject(Name));
        }
    }

    // Use appropriate Write methods
    // WriteObject, WriteError, WriteWarning, WriteVerbose, WriteDebug
}
```

#### Error Handling

```csharp
// Use ErrorRecord for proper error reporting
catch (Exception ex)
{
    var errorRecord = new ErrorRecord(
        ex,
        "ExampleError",
        ErrorCategory.InvalidOperation,
        targetObject);

    // Terminating error
    ThrowTerminatingError(errorRecord);

    // OR non-terminating error
    WriteError(errorRecord);
}
```

#### Pipeline Support

```csharp
// Proper pipeline implementation
protected override void BeginProcessing()
{
    // Initialize once
}

protected override void ProcessRecord()
{
    // Process each pipeline object
}

protected override void EndProcessing()
{
    // Cleanup and final output
}
```

### Code Review Checklist

Before committing changes, verify:

- [ ] Code builds successfully (`./build.ps1`)
- [ ] All tests pass (`./build.ps1 -Test`)
- [ ] New tests added for new functionality
- [ ] PSScriptAnalyzer passes (for .ps1 files)
- [ ] Code follows naming conventions
- [ ] Comments added for complex logic
- [ ] Documentation updated if needed
- [ ] No sensitive information in code
- [ ] Error handling is appropriate
- [ ] Performance considerations addressed

### Common Pitfalls to Avoid

1. **Don't use Write-Host in cmdlets** → Use WriteObject, WriteVerbose, etc.
2. **Don't catch generic exceptions** → Catch specific exception types
3. **Don't modify global state** → Keep cmdlets isolated and testable
4. **Don't use non-approved verbs** → Check approved verb list
5. **Don't ignore pipeline input** → Support ValueFromPipeline where appropriate
6. **Don't skip ShouldProcess** → Implement for state-changing operations
7. **Don't hardcode paths** → Use Path.Combine, Environment variables
8. **Don't ignore platform differences** → Test cross-platform compatibility

### File Modification Patterns

#### When editing C# files:

```csharp
// Add using statements alphabetically
using System;
using System.Collections.Generic;
using System.Management.Automation;

// Match existing indentation (4 spaces)
// Use Allman brace style
// Add XML documentation comments for public APIs

/// <summary>
/// Gets the example object.
/// </summary>
/// <param name="name">The name of the example.</param>
/// <returns>An ExampleObject instance.</returns>
public ExampleObject GetExample(string name)
{
    // Implementation
}
```

#### When editing PowerShell files:

```powershell
# Use proper function structure
function Verb-Noun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ParameterName
    )

    begin {
        # Initialization
    }

    process {
        # Main logic
    }

    end {
        # Cleanup
    }
}
```

### Documentation Requirements

1. **Code Comments:**
   - XML docs for public C# APIs
   - Comment-based help for PowerShell functions
   - Inline comments for complex algorithms

2. **External Documentation:**
   - Update README.md if adding features
   - Add examples to docs/
   - Update CHANGELOG.md for user-facing changes

3. **Help Content:**
   - Cmdlet help XML
   - About topics for new concepts
   - Example usage in help

### Performance Considerations

- **Avoid N+1 queries:** Batch operations when possible
- **Use StringBuilder:** For string concatenation in loops
- **Dispose resources:** Use `using` statements
- **Lazy loading:** Don't load data until needed
- **Cache results:** For expensive operations

### Security Considerations

- **Validate input:** Check parameters for malicious content
- **Sanitize paths:** Prevent path traversal attacks
- **No credential logging:** Never log passwords or secrets
- **Use SecureString:** For sensitive data
- **Code injection:** Prevent script injection in dynamic code

---

## Version-Specific Notes

### PowerShell 7.x Features

- Pipeline parallelization (`ForEach-Object -Parallel`)
- Ternary operator (`$condition ? $true : $false`)
- Null coalescing (`$var ?? $default`)
- Chain operators (`&&`, `||`)
- Error view improvements

### .NET 6+ Features

- Record types
- Init-only properties
- Pattern matching enhancements
- Top-level statements
- File-scoped namespaces

---

## Useful Resources

### Documentation
- **PowerShell Docs:** https://docs.microsoft.com/powershell
- **PowerShell GitHub:** https://github.com/PowerShell/PowerShell
- **Cmdlet Development:** docs/cmdlet-example/
- **Building Guide:** docs/building/

### Tools
- **PSScriptAnalyzer:** Code analysis for PowerShell
- **Pester:** PowerShell testing framework
- **platyPS:** PowerShell help documentation
- **Visual Studio:** Full IDE support
- **VS Code:** Lightweight editor with PowerShell extension

### Community
- **GitHub Issues:** Bug reports and feature requests
- **GitHub Discussions:** General questions
- **PowerShell.org:** Community forum
- **PowerShell Summit:** Annual conference

---

## Maintenance Notes

### For Repository Maintainers

This CLAUDE.md file should be updated when:

- Major architectural changes occur
- Build system changes
- New conventions are adopted
- Directory structure changes
- New development tools are introduced
- Testing framework changes

### Update Frequency

- **After major releases:** Review and update all sections
- **After structural changes:** Update relevant sections immediately
- **Quarterly:** General review for accuracy
- **On request:** When AI assistants report issues

### Feedback

If you encounter issues or inaccuracies in this guide:
1. Create a GitHub issue with label `documentation`
2. Propose specific changes
3. Reference specific sections that need updates

---

## Quick Reference Card

### Essential Commands

```bash
# Build
./build.ps1                          # Build PowerShell
./build.ps1 -Configuration Release   # Release build
./build.ps1 -Clean                   # Clean build

# Test
./build.ps1 -Test                    # Run all tests
dotnet test <project>                # Run specific xUnit tests
Invoke-Pester <path>                 # Run specific Pester tests

# Run
./src/powershell-*/pwsh              # Run built PowerShell

# Package
./build.ps1 -Package                 # Create packages

# Analysis
Invoke-ScriptAnalyzer -Path .        # Analyze PowerShell scripts
```

### Common Locations

```
src/System.Management.Automation/    → Core engine
src/Microsoft.PowerShell.Commands.* → Built-in cmdlets
test/xUnit/                          → C# unit tests
test/powershell/                     → PowerShell tests
docs/                                → Documentation
tools/                               → Build tools
```

### Key Patterns

```csharp
// Cmdlet template
[Cmdlet(Verb, Noun)]
public class VerbNounCommand : PSCmdlet
{
    [Parameter()]
    public string Property { get; set; }

    protected override void ProcessRecord()
    {
        WriteObject(result);
    }
}
```

---

**Document Version:** 1.0
**Created:** 2025-12-02
**For:** AI Assistant guidance on PowerShell codebase development

*This document is maintained as part of the PowerShell repository to assist AI assistants in understanding and contributing to the codebase effectively.*
