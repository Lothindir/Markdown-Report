#requires -version 4
<#
.SYNOPSIS
  Creates a new rapport in markdown, to be used with Pandoc.

.DESCRIPTION
  Creates and populates a new rapport with the given parameters. Can work in two ways : 
  1) Provide all the parameters via the cli when calling the script
  2) Use a YAML file to provide parameters
  3) Enter the interactive mode to provide the parameters one by one

  If you choose to use a YAML configuration file, you must specify it with the -ConfigFile param.

.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.10
  Author:         Francesco Monti
  Creation Date:  29.02.20
  Purpose/Change: Added CLI and Interactive mode

.EXAMPLE
  <Example explanation goes here>
  
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

[CmdletBinding(DefaultParameterSetName = 'ConfigFile')]
Param (
  # Destination path
  [Parameter(Mandatory = $true,
    ParameterSetName = 'ConfigFile')]
  [Parameter(Mandatory = $true,
    ParameterSetName = 'CLI')]
  [string]
  $DestinationPath,

  # File name
  [Parameter(Mandatory = $true,
    ParameterSetName = 'ConfigFile')]
  [Parameter(Mandatory = $true,
    ParameterSetName = 'CLI')]
  [String]
  $FileName,

  # Model name
  [Parameter()]
  [String]
  $ModelName = 'Rapport',

  #--- Interactive mode ---

  [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
  [Switch]
  $Interactive,

  #--- CLI mode ---

  [Parameter(Mandatory = $true,
    ParameterSetName = 'CLI')]
  [String]
  $Title,

  [Parameter(Mandatory = $true,
    ParameterSetName = 'CLI')]
  [String]
  $Author,

  [Parameter(ParameterSetName = 'CLI')]
  [Switch]
  $TitlePage,

  [Parameter(ParameterSetName = 'CLI')]
  [Switch]
  $TOC,

  [Parameter(ParameterSetName = 'CLI')]
  [Switch]
  $TOCSinglePage,

  [Parameter(ParameterSetName = 'CLI')]
  [Switch]
  $CustomHeaderFooter,

  [Parameter(ParameterSetName = 'CLI')]
  [Switch]
  $Book,

  [Parameter(ParameterSetName = 'CLI')]
  [Hashtable]
  $CustomParams,

  #--- YAML mode ---

  [Parameter(Mandatory = $true, 
    ParameterSetName = 'ConfigFile',
    HelpMessage = 'Enter the yaml config file path')]
  [Alias('ConfigFilePath')]
  [String]
  $ConfigFile
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'Continue'

# Models folder path and list of models
$ModelsPath = './Models/'

#Import Modules & Snap-ins
Import-Module './PSYaml/PSYaml'

#----------------------------------------------------------[Declarations]----------------------------------------------------------


#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#

Function <FunctionName> {
  Param ()

  Begin {
    Write-Host '<description of what is going on>...'
  }

  Process {
    Try {
      <code goes here>
    }

    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception.Message)"
      Break
    }
  }

  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}

#>

Function Get-MarkdownModel {
  Param (
    # Path of Models folder
    [Parameter(Mandatory = $true)]
    [String]
    $Path,

    # Name of the Model
    [Parameter(Mandatory = $true)]
    [String]
    $ModelName
  )

  Begin {
    Write-Host 'Loading Markdown model...'
  }

  Process {
    Try {
      $ModelFile = Join-Path -Path $Path -ChildPath ($ModelName + '.md')
      if (Test-Path -Path $ModelFile) {
        return $ModelFile
      }
      else {
        $ModelName = if ($ModelName.Contains('.md')) { $ModelName } else { $ModelName + '.md' }
        if (Test-Path -Path $ModelName) {
          return $ModelName
        }
        else {
          throw 'Model Not Found : ' + $ModelName
        }
      }
    }

    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception.Message)"
      Break
    }
  }

  End {
    If ($?) {
      Write-Host 'Model loaded successfully.'
      Write-Host ' '
    }
  }
}

Function Verify-ConfigFile {
  Param (
    # Config file
    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.OrderedDictionary]
    $YamlObject
  )

  Begin {
    Write-Host 'Verifying config file...'
    $requiredEntries = @(
      'title',
      'author'
    )
  }

  Process {
    Try {
      foreach ($requiredEntry in $requiredEntries) {
        if (!$YamlObject[$requiredEntry]) {
          throw 'The entry "' + $requiredEntry + '" is required in the config file'
        }
      }

      # Setting graphics path if not defined
      if (!$YamlObject['graphics_path']) {
        $YamlObject['graphics_path'] = $DestinationPath | Resolve-Path | Join-Path -ChildPath '\'
      }
      else {
        if (!(Test-Path -Path $YamlObject['graphics_path'] -PathType Container)) {
          New-Item -Path $YamlObject['graphics_path'] -ItemType Directory >> $null
        }
        $YamlObject['graphics_path'] = $YamlObject['graphics_path'] | Resolve-Path | Join-Path -ChildPath '\'
      }
    }

    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception.Message)"
      Break
    }
  }

  End {
    If ($?) {
      Write-Host ' '
    }
  }
}

Function Fill-YAML {
  Param (
    # Config file
    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.OrderedDictionary]
    $YamlObject
  )

  Begin {
    
  }

  Process {
    Try {
      if ($YamlObject['header-includes']) {
        $YamlObject['header-includes'] = $YamlObject['header-includes'] -replace '(?:##(?<Param>.{1,20})##)', { $YamlObject[$_.Groups[1].Value.ToLower()] } # Replaces all the text contained between two '##' with the value found in the config file
      }

      $YamlFile = ConvertTo-Yaml -inputObject $YamlObject
      $YamlFile = $YamlFile -replace ">`r`n", "|`r`n"
      $YamlFile = $YamlFile -replace '(?im:^.+[:]\s+$)', ''
      $YamlFile = $YamlFile -replace "((`r)*`n){2,}", "`r`n"
      return $YamlFile
    }

    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }

  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}

Function Fill-Markdown {
  Param (
    # Config file
    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.OrderedDictionary]
    $YamlObject,

    # Model file contents
    [Parameter(Mandatory=$true)]
    [System.Object]
    $Model
  )

  Begin {
    
  }

  Process {
    Try {
      
        $Model = $Model -replace '(?:<>(?<Param>.{1,20})<>)', { $YamlObject[$_.Groups[1].Value.ToLower()] } # Replaces all the text contained between two '<>' with the value found in the config file
        return $Model
    }

    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }

  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}

<#
.SYNOPSIS
Prompts the user for an input. Supports a default value.

.DESCRIPTION
This functions prompts the user for an input.
If a default value as been provided and the user responds with an empty string (i.e. pressing enter without entering any value) the result will be the default value.
If no default value as been provided the user is prompted until it enters something.

.PARAMETER Prompt
Specifies the text of the prompt. Type a string. If the string includes spaces, enclose it in quotation marks. PowerShell appends a colon (:) to the text that you enter.

.PARAMETER DefaultValue
Specifies the default value of the prompt. Can be null.

.PARAMETER ValidationSet
Specifies the possible values for the user's response.

.EXAMPLE
Read-HostWithDefaultValue -Prompt 'Enter your name'

Read-HostWithDefaultValue -Prompt 'Do you want to shutdown ? (y/N)' -DefaultValue 'y'

Read-HostWithDefaultValue 'Remove the file ?' 'No'

.NOTES
General notes
#>

Function Read-HostWithDefaultValue {
  [CmdletBinding(DefaultParametersetName = "Default")]
  Param (
    # Prompt
    [Parameter(Mandatory = $true,
      Position = 0)]
    [String]
    $Prompt,

    # Default value
    [Parameter(Position = 1)]
    [String]
    $DefaultValue,

    # Validation set
    [Parameter(Mandatory = $true,
      ParameterSetName = 'Set')]
    [String[]]
    $ValidationSet,

    # Validation hash
    [Parameter(Mandatory = $true,
      ParameterSetName = 'Hash')]
    [HashTable]
    $ValidationHash,

    # Is the response nullable
    [Parameter()]
    [Switch]
    $Nullable,

    # Yes - No shortcut
    [Parameter(Mandatory = $true,
      ParameterSetName = 'YesNo')]
    [Switch]
    $YesNo,

    # Default no
    [Parameter(ParameterSetName = 'YesNo')]
    [Switch]
    $No
  )

  if ($YesNo) {
    $ValidationHash = @{'y' = $true; 'n' = $false; 'yes' = $true; 'no' = $false }
    $Prompt += ' (y/n)'
    $DefaultValue = if ($No) { 'no' } else { 'yes' }
  }

  $DefaultValueString = if ($DefaultValue) { "[$($DefaultValue)]" } else { '' }
  
  do {
    do {
      $Response = Read-Host -Prompt "$($Prompt) $($DefaultValueString)"
    } while (!$Response -and !$DefaultValue -and !$Nullable)
    $Response = if ($null -eq $Response) { $DefaultValue } else { $Response }
    
  } while (($ValidationSet -and !$ValidationSet.Contains($Response)) -or
    $ValidationHash -and !$ValidationHash.ContainsKey($Response))
  
  $Response = if ($ValidationHash) { $ValidationHash[$Response] } else { $Response }
  return $Response
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# --------- Initialiazion ---------
clear
if ($PSCmdlet.ParameterSetName -eq 'Interactive' -or $PSCmdlet.ParameterSetName -eq 'CLI') {
  if ($Interactive) {
    Write-Host 'Interactive mode enabled'

    $DestinationPath = Read-HostWithDefaultValue -Prompt 'Where do you want to create your file ?' '.'

    $FileName = Read-HostWithDefaultValue -Prompt 'What is the filename ?' 'Rapport'

    $Title = Read-HostWithDefaultValue -Prompt 'Enter Title'

    $Author = Read-HostWithDefaultValue 'Enter Author(s)' 'Author'

    $TitlePage = Read-HostWithDefaultValue 'Has a title page ?' -YesNo

    $TOC = Read-HostWithDefaultValue 'Has table of contents ?' -YesNo

    $TOCSinglePage = Read-HostWithDefaultValue 'Is the table of contents a single page' -YesNo

    $CustomHeaderFooter = Read-HostWithDefaultValue 'Add custom header and footer ?' -YesNo

    $Book = Read-HostWithDefaultValue 'Is a book ?' -YesNo -No

    $HasOtherParam = Read-HostWithDefaultValue 'Other config ? (Write config name or press enter to continue)' -Nullable

    $CustomParams = @{ }

    while ($HasOtherParam) {
      $OtherParamValue = Read-HostWithDefaultValue "Custom config -> $($HasOtherParam)"

      if ($CustomParams.ContainsKey($HasOtherParam)) {
        $Overwrite = Read-HostWithDefaultValue "Do you want to overwrite the value $($CustomParams[$HasOtherParam]) ?" -YesNo -No

        if ($Overwrite) {
          $CustomParams[$HasOtherParam] = $OtherParamValue
        }
      }
      else {
        $CustomParams.Add($HasOtherParam, $OtherParamValue)
      }

      Write-Host ' '

      $HasOtherParam = $null

      $HasOtherParam = Read-HostWithDefaultValue 'Other config ? (Write config name or press enter to continue)' -Nullable
    }
  }

  $YAMLConfig = [ordered]@{ }

  $YAMLConfig.title = $Title
  $YAMLConfig.author = $Author
  $YAMLConfig.titlepage = $TitlePage.ToString().ToLower()
  $YAMLConfig.toc = $TOC.ToString().ToLower()
  $YAMLConfig['toc-own-page'] = $TOCSinglePage.ToString().ToLower()
  $YAMLConfig.book = $Book.ToString().ToLower()
  if ($CustomHeaderFooter) {
    $YAMLConfig['disable-header-and-footer'] = 'true'
    $YAMLConfig['header-includes'] = '
    \usepackage{graphicx}
    \usepackage{fancyhdr}
    \pagestyle{fancy}
    \graphicspath{{##GRAPHICS_PATH##}}
    \fancyhead{}
    \fancyfoot{}
    \fancyhead[LO,LE]{##TITLE##}
    \rhead{\includegraphics[width=2cm]{logo.png}}
    \fancyfoot[LO,LE]{##AUTHOR##}
    \fancyfoot[CE,CO]{\thepage}'
  }

  foreach ($customParam in $CustomParams.Keys) {
    $YAMLConfig[$customParam] = $CustomParams[$customParam]
  }

  $YamlFile = ConvertTo-Yaml -inputObject $YAMLConfig
  $YamlFile = $YamlFile -replace '>\r\n', "|`r`n"

  $ConfigFileName = 'config_i.yml'
  $ConfigFile = Join-Path -Path $DestinationPath -ChildPath $ConfigFileName
  while (Test-Path -Path $ConfigFile) {
    $Overwrite = Read-HostWithDefaultValue "File '$($ConfigFileName)' already exists. Overwrite ?" -YesNo -No
    if ($Overwrite) {
      Remove-Item -Path $ConfigFile
    }
    else {
      $ConfigFileName = Read-HostWithDefaultValue -Prompt 'Enter new config filename'
      if (!$ConfigFileName.Contains('.yml')) {
        $ConfigFileName += '.yml'
      }
      $ConfigFile = Join-Path -Path $DestinationPath -ChildPath $ConfigFileName
    }
  }
  New-Item -Path $ConfigFile -ItemType File >> $null
  Set-Content -Path $ConfigFile -Value $YamlFile

  Write-Host ' '
}

$Model = Get-MarkdownModel -Path $ModelsPath -ModelName $ModelName

$YamlObject = ConvertFrom-Yaml -Path $ConfigFile
Verify-ConfigFile -YamlObject $YamlObject
$YamlFile = Fill-YAML $YamlObject
$MarkdownModel = Get-Content -Path $Model
$MarkdownModel = Fill-Markdown -YamlObject $YamlObject -Model $MarkdownModel

$FileName = if ($FileName.Contains('.md')) { $FileName } else { $FileName + '.md' }

$RapportFilePath = Join-Path -Path $DestinationPath -ChildPath $FileName

while (Test-Path -Path $RapportFilePath) {
  $Overwrite = Read-HostWithDefaultValue "File '$($FileName)' already exists. Overwrite ?" -YesNo -No
  if ($Overwrite) {
    Remove-Item -Path $RapportFilePath
  }
  else {
    $FileName = Read-HostWithDefaultValue -Prompt 'Enter new config filename'
    if (!$FileName.Contains('.md')) {
      $FileName += '.md'
    }
    $RapportFilePath = Join-Path -Path $DestinationPath -ChildPath $FileName
  }
}

New-Item -Path $RapportFilePath -ItemType File >> $null
Add-Content -Path $RapportFilePath -Value $YamlFile
Add-Content -Path $RapportFilePath -Value "---`r`n"
Add-Content -Path $RapportFilePath -Value $MarkdownModel
