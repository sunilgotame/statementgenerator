param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("statement", "certificate")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

$ErrorActionPreference = "Stop"

function Read-Payload {
    return Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json
}

function Convert-ToText($Value) {
    if ($null -eq $Value) {
        return ""
    }
    return [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-CellText($Cell) {
    if ($null -eq $Cell) {
        return ""
    }
    return Convert-ToText $Cell.Text
}

function Set-BlankOrValue($Cell, $Value) {
    $text = Convert-ToText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        $Cell.ClearContents() | Out-Null
        return
    }
    $Cell.Value = $text
}

function Set-TextValue($Cell, $Value) {
    $text = Convert-ToText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        $Cell.ClearContents() | Out-Null
        return
    }
    $Cell.NumberFormat = "@"
    $Cell.Value = $text
}

function Set-IntegerValue($Cell, $Value) {
    $text = Convert-ToText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        $Cell.ClearContents() | Out-Null
        return
    }
    $number = 0
    if (-not [int]::TryParse($text, [ref]$number)) {
        Set-TextValue $Cell $text
        return
    }
    $Cell.NumberFormat = "0"
    $Cell.Value2 = $number
}

function Set-DateValue($Cell, $IsoDate) {
    $isoText = Convert-ToText $IsoDate
    if ([string]::IsNullOrWhiteSpace($isoText)) {
        $Cell.ClearContents() | Out-Null
        return
    }
    $cultureInvariant = [System.Globalization.CultureInfo]::InvariantCulture
    $styles = [System.Globalization.DateTimeStyles]::None
    $parsed = [datetime]::ParseExact($isoText, "yyyy-MM-dd", $cultureInvariant, $styles)
    $Cell.NumberFormat = "yyyy-mm-dd"
    $Cell.Value2 = $parsed.ToOADate()
}

function Convert-ToNullableDouble($Value) {
    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        return [double]$Value
    }

    $text = Convert-ToText $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $styles = [System.Globalization.NumberStyles]::Any
    $cultureInvariant = [System.Globalization.CultureInfo]::InvariantCulture
    $cultureCurrent = [System.Globalization.CultureInfo]::CurrentCulture
    $number = 0.0

    if ([double]::TryParse($text, $styles, $cultureInvariant, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($text, $styles, $cultureCurrent, [ref]$number)) {
        return $number
    }

    $normalized = $text.Replace(",", "")
    if ([double]::TryParse($normalized, $styles, $cultureInvariant, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($normalized, $styles, $cultureCurrent, [ref]$number)) {
        return $number
    }

    throw "Could not convert value '$text' to a numeric amount."
}

function Replace-ValueKeepingLabel($Text, $Replacement) {
    $textValue = Convert-ToText $Text
    $replacementValue = Convert-ToText $Replacement
    if ([string]::IsNullOrWhiteSpace($textValue)) {
        return $replacementValue
    }
    if ($textValue -match "^(.*?)(\s*:-\s*|\s*:\s*).*$") {
        $label = $matches[1].TrimEnd()
        $separator = if ($textValue -match ":-") { " :- " } else { ": " }
        return "$label$separator$replacementValue"
    }
    return $replacementValue
}

function Find-HeaderInfo($Worksheet) {
    $used = $Worksheet.UsedRange
    $startRow = $used.Row
    $endRow = [Math]::Min($startRow + [Math]::Min($used.Rows.Count, 40) - 1, 60)
    $startCol = $used.Column
    $endCol = [Math]::Min($startCol + $used.Columns.Count - 1, 12)

    for ($row = $startRow; $row -le $endRow; $row++) {
        $map = @{}
        for ($col = $startCol; $col -le $endCol; $col++) {
            $text = (Get-CellText $Worksheet.Cells.Item($row, $col)).Trim()
            $low = $text.ToLowerInvariant()
            if ($low -eq "date" -or $low -eq "value date") {
                $map["date"] = $col
            } elseif ($low -like "*description*" -or $low -like "*particular*") {
                $map["description"] = $col
            } elseif ($low -like "*cheque*" -or $low -like "*chq*") {
                $map["cheque"] = $col
            } elseif ($low -like "*debit*") {
                $map["debit"] = $col
            } elseif ($low -like "*credit*") {
                $map["credit"] = $col
            } elseif ($low -like "*balance*") {
                $map["balance"] = $col
            }
        }

        if ($map.ContainsKey("date") -and $map.ContainsKey("description") -and $map.ContainsKey("debit") -and $map.ContainsKey("credit") -and $map.ContainsKey("balance")) {
            return @{
                Row = $row
                Map = $map
            }
        }
    }

    return $null
}

function Find-DataRange($Worksheet, $HeaderInfo) {
    $used = $Worksheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $startRow = $HeaderInfo.Row + 1
    $endRow = $startRow - 1
    $seenData = $false

    for ($row = $startRow; $row -le $lastRow; $row++) {
        $desc = (Get-CellText $Worksheet.Cells.Item($row, $HeaderInfo.Map["description"])).Trim()
        $balance = (Get-CellText $Worksheet.Cells.Item($row, $HeaderInfo.Map["balance"])).Trim()
        $date = (Get-CellText $Worksheet.Cells.Item($row, $HeaderInfo.Map["date"])).Trim()
        $debit = (Get-CellText $Worksheet.Cells.Item($row, $HeaderInfo.Map["debit"])).Trim()
        $credit = (Get-CellText $Worksheet.Cells.Item($row, $HeaderInfo.Map["credit"])).Trim()
        $isDataRow = -not [string]::IsNullOrWhiteSpace($desc) -or -not [string]::IsNullOrWhiteSpace($balance) -or -not [string]::IsNullOrWhiteSpace($date) -or -not [string]::IsNullOrWhiteSpace($debit) -or -not [string]::IsNullOrWhiteSpace($credit)
        if ($isDataRow) {
            $seenData = $true
            $endRow = $row
            continue
        }
        if ($seenData) {
            break
        }
    }

    if ($endRow -lt $startRow) {
        $endRow = $lastRow
    }

    return @{
        Start = $startRow
        End = $endRow
        Capacity = ($endRow - $startRow + 1)
        MaxColumn = ($used.Column + $used.Columns.Count - 1)
    }
}

function Update-StatementHeaderCell($Cell, $Payload) {
    $text = (Get-CellText $Cell).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return
    }

    $low = $text.ToLowerInvariant()
    $replacement = $null

    if ($low -like "account holder*name*" -or $low -like "account holder*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.customer_name
    } elseif ($low -match "^name\s*:") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.customer_name
    } elseif ($low -like "address*" -or $low -like "permanent address*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.customer_address
    } elseif ($low -like "*a/c no*" -and $low -like "*member id*") {
        $replacement = "A/C No.: $($Payload.account.account_number)                    Member ID: $($Payload.account.member_id)"
    } elseif ($low -like "*account no*" -or $low -like "*a/c no*" -or $low -like "*a/c no.*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.account_number
    } elseif ($low -like "*account type*" -or $low -like "*a/c type*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.account_type
    } elseif ($low -like "*currency*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.currency
    } elseif ($low -like "*a/c opening date*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.account.opening_date_iso
    } elseif ($low -like "*interest rate*" -and $low -like "*tax*") {
        $replacement = "Interest Rate : $($Payload.rates.interest_rate)%, Tax : $($Payload.rates.tax_rate)%"
    } elseif ($low -like "*interest rate*") {
        $replacement = Replace-ValueKeepingLabel $text "$($Payload.rates.interest_rate)%"
    } elseif ($low -like "*tax rate*" -or ($low -eq "tax" -or $low -like "tax*")) {
        $replacement = Replace-ValueKeepingLabel $text "$($Payload.rates.tax_rate)%"
    } elseif ($low -like "from * to *") {
        $replacement = "FROM $($Payload.statement.period_from_iso) TO $($Payload.statement.period_to_iso)"
    } elseif ($low -like "*statement of account from*") {
        $replacement = "Statement of Account From $($Payload.statement.period_from_slash) to $($Payload.statement.period_to_slash)"
    } elseif ($low -like "*statement from*") {
        $replacement = "STATEMENT FROM $($Payload.statement.period_from_slash) TO $($Payload.statement.period_to_slash)"
    } elseif ($low -like "*account statement period*") {
        $replacement = "Account Statement Period: $($Payload.statement.period_from_iso) to $($Payload.statement.period_to_iso)"
    } elseif ($low -like "date*:-*" -or $low -like "date*:*") {
        $replacement = Replace-ValueKeepingLabel $text $Payload.statement.issue_date_slash
    }

    if ($null -ne $replacement -and $replacement -ne $text) {
        Set-TextValue $Cell $replacement
    }
}

function Update-StatementHeaderPair($Worksheet, $Row, $Col, $MaxColumn, $Payload) {
    if ($Col -ge $MaxColumn) {
        return
    }

    $labelCell = $Worksheet.Cells.Item($Row, $Col)
    $labelText = (Get-CellText $labelCell).Trim()
    if ([string]::IsNullOrWhiteSpace($labelText)) {
        return
    }

    $valueCell = $Worksheet.Cells.Item($Row, $Col + 1)
    $valueText = (Get-CellText $valueCell).Trim()
    if ([string]::IsNullOrWhiteSpace($valueText)) {
        return
    }

    $low = $labelText.ToLowerInvariant()
    if ($low -match "^name\s*:?" -or $low -like "account holder*") {
        Set-TextValue $valueCell $Payload.account.customer_name
        return
    }
    if ($low -match "^address\s*:?" -or $low -like "permanent address*") {
        Set-TextValue $valueCell $Payload.account.customer_address
        return
    }
    if ($low -match "^(a/c no|account no)") {
        Set-TextValue $valueCell $Payload.account.account_number
        return
    }
    if ($low -match "^(a/c type|account type)") {
        Set-TextValue $valueCell $Payload.account.account_type
        return
    }
    if ($low -match "^interest") {
        Set-BlankOrValue $valueCell ([string]([double]$Payload.rates.interest_rate / 100.0))
        return
    }
    if ($low -match "^tax") {
        Set-BlankOrValue $valueCell ([string]([double]$Payload.rates.tax_rate / 100.0))
    }
}

function Compose-Description($BaseDescription, $ChequeNo, $SampleDescription, $HasChequeColumn) {
    $baseText = Convert-ToText $BaseDescription
    $chequeText = Convert-ToText $ChequeNo
    if ($HasChequeColumn -or [string]::IsNullOrWhiteSpace($chequeText)) {
        return $baseText
    }
    $sampleSource = Convert-ToText $SampleDescription
    $sampleLower = $sampleSource.ToLowerInvariant()
    if ($sampleLower.Contains("chq")) {
        return "$baseText CHQ. No. $chequeText"
    }
    if ($sampleLower.Contains("cheque")) {
        return "$baseText Cheque No. $chequeText"
    }
    return "$baseText CHQ. No. $chequeText"
}

function Update-StatementWorkbook($Payload) {
    Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($OutputPath)
        $selectedSheet = $null
        $headerInfo = $null
        foreach ($worksheet in $workbook.Worksheets) {
            $candidate = Find-HeaderInfo $worksheet
            if ($null -ne $candidate) {
                $selectedSheet = $worksheet
                $headerInfo = $candidate
                break
            }
        }

        if ($null -eq $selectedSheet) {
            throw "Could not find a statement table header in the selected Excel template."
        }

        $dataRange = Find-DataRange $selectedSheet $headerInfo
        if ($Payload.statement_rows.Count -gt $dataRange.Capacity) {
            throw "The generated statement has $($Payload.statement_rows.Count) rows, but the template only has space for $($dataRange.Capacity)."
        }

        $headerMaxRow = [Math]::Min($headerInfo.Row - 1, 15)
        if ($headerMaxRow -gt 0) {
            for ($row = 1; $row -le $headerMaxRow; $row++) {
                for ($col = 1; $col -le [Math]::Min($dataRange.MaxColumn, 10); $col++) {
                    Update-StatementHeaderPair $selectedSheet $row $col ([Math]::Min($dataRange.MaxColumn, 10)) $Payload
                    Update-StatementHeaderCell $selectedSheet.Cells.Item($row, $col) $Payload
                }
            }
        }

        $sampleDescription = (Get-CellText $selectedSheet.Cells.Item($dataRange.Start, $headerInfo.Map["description"])).Trim()
        $hasChequeColumn = $headerInfo.Map.ContainsKey("cheque")

        for ($index = 0; $index -lt $Payload.statement_rows.Count; $index++) {
            $row = $Payload.statement_rows[$index]
            $targetRow = $dataRange.Start + $index
            $debitValue = Convert-ToNullableDouble $row.debit
            $creditValue = Convert-ToNullableDouble $row.credit
            Set-DateValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["date"]) $row.date
            $description = Compose-Description $row.description $row.cheque_no $sampleDescription $hasChequeColumn
            Set-TextValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["description"]) $description
            if ($hasChequeColumn) {
                Set-IntegerValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["cheque"]) $row.cheque_no
            }
            if ($null -ne $debitValue -and $debitValue -gt 0) {
                Set-BlankOrValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["debit"]) $row.debit
            } else {
                $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["debit"]).ClearContents() | Out-Null
            }
            if ($null -ne $creditValue -and $creditValue -gt 0) {
                Set-BlankOrValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["credit"]) $row.credit
            } else {
                $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["credit"]).ClearContents() | Out-Null
            }
            if ([string]::IsNullOrWhiteSpace((Convert-ToText $row.balance))) {
                throw "Balance value is missing for statement row $($index + 1)."
            }
            Set-BlankOrValue $selectedSheet.Cells.Item($targetRow, $headerInfo.Map["balance"]) $row.balance
        }

        for ($row = $dataRange.Start + $Payload.statement_rows.Count; $row -le $dataRange.End; $row++) {
            for ($col = 1; $col -le $dataRange.MaxColumn; $col++) {
                $selectedSheet.Cells.Item($row, $col).ClearContents() | Out-Null
            }
        }

        $workbook.Save()
        $workbook.Close($true)
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}

function Get-UpdatedCertificateText($Text, $Payload) {
    $originalText = Convert-ToText $Text
    $trimmed = $originalText.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $originalText
    }

    $lower = $trimmed.ToLowerInvariant()
    if ($lower -like "ref. no*") {
        if ([string]::IsNullOrWhiteSpace((Convert-ToText $Payload.account.reference_no))) {
            return [regex]::Replace($trimmed, "(?i)date\s*:\s*.*$", "Date: $($Payload.statement.issue_date_slash)")
        }
        return "Ref. No.: $($Payload.account.reference_no)    Date: $($Payload.statement.issue_date_slash)"
    }
    if ($lower -match "^issue date\s*:") { return "Issue Date: $($Payload.statement.issue_date_slash)" }
    if ($lower -match "^date\s*:") { return "Date: $($Payload.statement.issue_date_slash)" }
    if ($lower -like "this is to certify*" -and $lower.Contains("under mentioned account holder")) {
        return "This is to certify that the balance in the credit of the under mentioned Account Holder as on $($Payload.statement.as_of_ordinal) is mentioned below."
    }
    if ($lower -match "^name\s*:") { return "Name: $($Payload.account.customer_name)" }
    if ($lower -like "account holder:*") { return "Account Holder: $($Payload.account.customer_name)" }
    if ($lower -match "^address\s*:") { return "Address: $($Payload.account.customer_address)" }
    if ($lower -match "^permanent address\s*:") { return "Permanent Address: $($Payload.account.customer_address)" }
    if ($lower -like "a/c no*" -and $lower -like "*member id*") {
        return "A/C No.: $($Payload.account.account_number)                    Member ID: $($Payload.account.member_id)"
    }
    if ($lower -match "^a/c no") { return "A/C No.: $($Payload.account.account_number)" }
    if ($lower -match "^account no") { return "Account No.: $($Payload.account.account_number)" }
    if ($lower -like "member id*") { return "Member ID: $($Payload.account.member_id)" }
    if ($lower -match "^a/c type") { return "A/C Type: $($Payload.account.account_type)" }
    if ($lower -match "^account type") { return "Account Type: $($Payload.account.account_type)" }
    if ($lower -like "interest rate*") { return "Interest Rate: $($Payload.rates.interest_rate) %" }
    if ($lower -like "currency*") { return "Currency: $($Payload.account.currency)" }
    if ($lower -like "total balance npr*") { return "Total Balance NPR: $($Payload.certificate.total_balance_npr_text)" }
    if ($lower -like "total balance:*") { return "Total Balance: NPR $($Payload.certificate.total_balance_npr_text)" }
    if ($lower -like "has a balance of*") { return "Has a balance of: NPR $($Payload.certificate.total_balance_npr_text)" }
    if ($lower -like "equivalent to usd*") { return "Equivalent to USD: $($Payload.certificate.equivalent_usd_text)" }
    if ($lower -like "usd:*") { return "USD: $($Payload.certificate.equivalent_usd_text)" }
    if ($lower -like "which is equivalent to*") { return "Which is equivalent to: USD $($Payload.certificate.equivalent_usd_text)" }
    if ($lower -like "in words usd*" -or $lower -like "(in words) usd*") {
        $label = ($trimmed -split ":", 2)[0]
        return "${label}: $($Payload.certificate.balance_words_usd)"
    }
    if ($lower -like "(in words:*") {
        return "(In Words: $($Payload.certificate.balance_words_npr))"
    }
    if ($lower.Contains("exchange rate")) {
        if ($lower.Contains("today")) {
            return "Note: Conversion has been done as per issue day exchange rate 1 USD = NPR $($Payload.rates.usd_npr_text)"
        }
        if ($lower.Contains("as of")) {
            return "The exchange rate as of $($Payload.statement.issue_date_ordinal) is 1 USD = $($Payload.rates.usd_npr_text) NPR"
        }
        if ($lower.Contains("source:")) {
            return "At prevailing exchange rate of USD 1 = NPR $($Payload.rates.usd_npr_text) (Source: Nepal Rastra Bank)"
        }
        return "At prevailing exchange rate of USD 1 = NPR $($Payload.rates.usd_npr_text)"
    }

    return $originalText
}

function Apply-ParagraphUpdate($Paragraph, $Payload) {
    $original = Convert-ToText $Paragraph.Range.Text
    $clean = $original.Replace([string][char]13, "").Replace([string][char]7, "")
    $updated = Get-UpdatedCertificateText $clean $Payload
    if ($updated -ne $clean) {
        $Paragraph.Range.Text = $updated + [char]13
    }
}

function Update-CertificateDocument($Payload) {
    Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath -Force
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try {
        $document = $word.Documents.Open($OutputPath, $false, $false)
        foreach ($paragraph in $document.Paragraphs) {
            Apply-ParagraphUpdate $paragraph $Payload
        }
        foreach ($table in $document.Tables) {
            foreach ($row in $table.Rows) {
                foreach ($cell in $row.Cells) {
                    foreach ($paragraph in $cell.Range.Paragraphs) {
                        Apply-ParagraphUpdate $paragraph $Payload
                    }
                }
            }
        }
        $document.Save()
        $document.Close($true)
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
    }
}

$payload = Read-Payload

if ($Mode -eq "statement") {
    Update-StatementWorkbook $payload
} else {
    Update-CertificateDocument $payload
}
