# import environment variables
$envPath = Join-Path -Path $PSScriptRoot -ChildPath ".\envs.ps1"
. $envPath

Function Get-PhoneNumber {
    param(
        [Parameter()]
        $User
    )

    $phoneProperties = @('mobile', 'MobilePhone', 'HomePhone', 'OfficePhone', 'telephoneNumber', 'msRTCSIP-Line')
    $phone = ""

    foreach ($prop in $phoneProperties) {
        if (![string]::IsNullOrEmpty($phone)) {
            break
        }

        if ([string]::IsNullOrEmpty($User.$prop)) {
            continue
        }

        if ($prop -eq "msRTCSIP-Line") {
            $phone = $User.$prop.Replace("tel:", "")
        }
        else {
            $phone = $User.$prop
            Write-Log -Message "  - Using property '$prop' for phone number"
        }
    }

    if ([string]::IsNullOrEmpty($phone)) {
        Write-Log -Message "No phone number registered on '$($user.mail)'" -Level WARNING
        return $phone
    }

    $phone = $phone.Replace(" ", "").Replace("+47", "")
    if ($phone.StartsWith("0047") -and $phone.Length -eq 12) {
        $phone = $phone.Substring(4, 8)
    }

    return $phone
}

Import-Module Logger
if ($papertrailConfig) {
    Add-LogTarget -Name Papertrail -Configuration $papertrailConfig
}

# file paths
$xmlFolder = Get-LogDir
$xmlPath = "$xmlFolder\export_$((Get-Date).DayOfWeek.ToString().ToLower()).xml"
$logPath = "$xmlFolder\tqm_$((Get-Date).DayOfWeek.ToString().ToLower()).log"

# make sure logpath exists
if (!(Test-Path -Path $xmlFolder)) {
    New-Item -Path $xmlFolder -ItemType Directory -Force -Confirm:$False | Out-Null
}

# remove last weeks xml file
if ((Test-Path -Path $xmlPath) -and ((Get-Date)-(Get-ChildItem -Path $xmlPath | Select-Object -ExpandProperty LastWriteTime)).Days -gt 0)
{
    try { Remove-Item -Path $xmlPath -Force -Confirm:$False -ErrorAction Stop } catch { }
}

# remove last weeks log file
if ((Test-Path -Path $logPath) -and ((Get-Date)-(Get-ChildItem -Path $logPath | Select-Object -ExpandProperty LastWriteTime)).Days -gt 0)
{
    try { Remove-Item -Path $logPath -Force -Confirm:$False -ErrorAction Stop } catch { }
}

Add-LogTarget -Name CMTrace -Configuration @{ Path = $logPath }
Write-Log -Message "Exporting to '$xmlPath'"

# get ad users to export
#$users = Get-ADUser -Filter "samaccountname -eq 'joh1904' -or samaccountname -eq 'run0805'" -Properties displayName,samAccountName,mail,mobile,MobilePhone,HomePhone,OfficePhone,telephoneNumber,msRTCSIP-Line,department,title
$users = D:\Scripts\VTFK-Toolbox\AD\Get-VTFKADUser.ps1 -Domain login.top.no -Properties displayName,samAccountName,mail,mobile,MobilePhone,HomePhone,OfficePhone,telephoneNumber,msRTCSIP-Line,department,title -OnlyAutoUsers | Where-Object { $_.Enabled -eq $True } | Sort-Object displayName
Write-Log -Message "Exporting $($users.Count) users"

# create xml document
$xmlSettings = New-Object System.Xml.XmlWriterSettings
$xmlSettings.Indent = $true
$xmlSettings.IndentChars = "`t"
$xml = [System.Xml.XmlWriter]::Create($xmlPath, $xmlSettings)

try {
    # create userimport root node
    $xml.WriteStartElement("UserImport")
    $xml.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
    $xml.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")

        # create gendate node
        $xml.WriteStartElement("GenDate")
        $xml.WriteValue((Get-Date -Format o))
        $xml.WriteEndElement()

        # create ol1 node
        $xml.WriteStartElement("OL1")

            # write ol1_key node
            $xml.WriteStartElement("OL1_Key")
            $xml.WriteValue("tfk")
            $xml.WriteEndElement()

            # write userinfo node
            $xml.WriteStartElement("UserInfo")

                ########## export users to nodes here ##########
                $users | ForEach-Object {
                    # create user node
                    $xml.WriteStartElement("User")

                        $user = $_
                        try {
                            Write-Log -Message "Adding user '$($user.mail)'"

                            # write name node
                            $xml.WriteStartElement("Name")
                            $xml.WriteValue($user.displayName)
                            $xml.WriteEndElement()

                            # write login node
                            $xml.WriteStartElement("Login")
                            $xml.WriteValue($user.samAccountName)
                            $xml.WriteEndElement()

                            # write employeeNumber node
                            $xml.WriteStartElement("EmployeeNumber")
                            $xml.WriteEndElement()

                            # write email node
                            $xml.WriteStartElement("Email")
                            $xml.WriteValue($user.mail)
                            $xml.WriteEndElement()

                            # write phone node
                            $xml.WriteStartElement("Phone")
                            $xml.WriteValue((Get-PhoneNumber -User $user))
                            $xml.WriteEndElement()

                            # write external node
                            $xml.WriteStartElement("External")
                            $xml.WriteValue($False)
                            $xml.WriteEndElement()

                            # write datedeleted node
                            $xml.WriteStartElement("DateDeleted")
                            $xml.WriteValue("BLANK")
                            $xml.WriteEndElement()

                            # write inactivefrom node
                            $xml.WriteStartElement("InactiveFrom")
                            $xml.WriteValue("BLANK")
                            $xml.WriteEndElement()

                            # write inactiveto node
                            $xml.WriteStartElement("InactiveTo")
                            $xml.WriteValue("BLANK")
                            $xml.WriteEndElement()

                            # write smsuser node
                            $xml.WriteStartElement("SMSUser")
                            $xml.WriteValue($False)
                            $xml.WriteEndElement()

                            # write defaultol3 node
                            $xml.WriteStartElement("DefaultOL3")
                            $xml.WriteEndElement()

                            # write defaultpl node
                            $xml.WriteStartElement("DefaultPL")
                            $xml.WriteValue($user.department)
                            $xml.WriteEndElement()

                            # write caseregsetting node
                            $xml.WriteStartElement("CaseRegSetting")

                                # write caseregistration_ol3 node
                                $xml.WriteStartElement("CaseRegistration_OL3")
                                $xml.WriteEndElement()

                                # write caseregistration_pl node
                                $xml.WriteStartElement("CaseRegistration_PL")
                                $xml.WriteValue($user.department)
                                $xml.WriteEndElement()

                                # write casetype node
                                $xml.WriteStartElement("CaseType")
                                $xml.WriteValue("Observasjon")
                                $xml.WriteEndElement()

                                # write severitydegree node
                                $xml.WriteStartElement("SeverityDegree")
                                $xml.WriteValue("Middels")
                                $xml.WriteEndElement()

                            # close caseregsetting node
                            $xml.WriteEndElement()
                        }
                        catch {
                            Write-Log -Message "Failed : $_" -Level ERROR -Exception $_
                        }
                
                    # close user node
                    $xml.WriteEndElement()
                }

            # close userinfo node
            $xml.WriteEndElement()

        # close ol1 node
        $xml.WriteEndElement()

    # close userimport root node
    $xml.WriteEndElement()
}
catch {
    Write-Log -Message "Failed in document creation : $_" -Level ERROR -Exception $_
}

# close xml document
$xml.Close()
