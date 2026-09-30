"""Build the Microsoft 365 kit (SharePoint lists + Power BI + Power Apps) from a register workbook.

Usage: python3 build_m365_kit.py <Batch_register.xlsx> <out_dir> [batch-code]
Writes: data/*.csv (list rows), Create-NavjyotiLists.ps1, and copies the guides next to it.
"""
import csv, os, shutil, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from convert_register import convert
from build_cloud_workbook import registers, LISTS

NOTE_COLS = {"Remarks", "Notes", "Observations", "CultureDose", "Description", "Note", "Remark"}
TEXT_CHOICE = {"Supplier", "Customer"}   # free text so new names can be typed
CALC = {  # SharePoint calculated columns (same-row only); cross-list ones live in Power BI / Power Apps
    ("RawMaterial", "Amount"): ("Currency", "=IF(OR(ISBLANK([QtyKg]),ISBLANK([RatePerKg])),\"\",[QtyKg]*[RatePerKg])"),
    ("Sales", "Revenue"): ("Currency", "=IF(OR(ISBLANK([QtyKg]),ISBLANK([PricePerKg])),\"\",[QtyKg]*[PricePerKg])"),
    ("Harvest", "RecoveryPct"): ("Number", "=IF(OR(ISBLANK([RawKg]),ISBLANK([NetKg]),[RawKg]=0),\"\",[NetKg]/[RawKg])"),
    ("PreCompost", "DaysToTransfer"): ("Number", "=IF(OR(ISBLANK([StartDate]),ISBLANK([TransferDate])),\"\",[TransferDate]-[StartDate])"),
}


def field_type(t, h, kind, extra):
    if kind.startswith("calc"):
        return CALC.get((t, h), (None,))[0] and "Calculated"
    if kind == "date":
        return "DateTime"
    if kind in ("kg", "num", "inr"):
        return "Number"
    if extra and extra not in TEXT_CHOICE:
        return "Choice"
    return "Note" if h in NOTE_COLS else "Text"


def main(src, out, batch):
    data = convert(src, batch)
    data["batches"].append({"code": "B5", "name": "Batch 5", "site": "Akola", "start": "", "status": "In beds", "notes": "In process"})
    regs = registers()
    os.makedirs(os.path.join(out, "data"), exist_ok=True)
    choices = dict(LISTS, Batch=sorted({r["code"] for r in data["batches"]}))
    ps = [PS_HEAD]
    for t, (key, cols) in regs.items():
        rows = data.get(key, [])
        keep = [(h, k, kind, extra) for h, k, kind, extra in cols if field_type(t, h, kind, extra)]
        ps.append(f'\n# ---------- {t} ----------\nNew-NavList "{t}"')
        for h, k, kind, extra in keep:
            ft = field_type(t, h, kind, extra)
            if ft == "Calculated":
                rt, formula = CALC[(t, h)]
                ps.append(f"Add-NavCalc \"{t}\" \"{h}\" \"{rt}\" '{formula}'")
            elif ft == "Choice":
                opts = ",".join(f'"{o}"' for o in choices[extra])
                ps.append(f'Add-NavField "{t}" "{h}" "Choice" @({opts})')
            else:
                ps.append(f'Add-NavField "{t}" "{h}" "{ft}"')
        # CSV of input columns (turns flattened for PreCompost)
        inputs = [(h, k, kind) for h, k, kind, extra in keep if not kind.startswith("calc")]
        with open(os.path.join(out, "data", f"{t}.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([h for h, *_ in inputs])
            for r in rows:
                turns = r.get("turns", [])
                line = []
                for h, k, kind in inputs:
                    if h.startswith("Turn") and key == "precompost":
                        i = int(h[4]) - 1
                        tk = {"Date": "date", "TempC": "temp", "MoisturePct": "moisture", "WateringDate": "watering"}[h[5:]]
                        v = turns[i].get(tk) if i < len(turns) else None
                    else:
                        v = r.get(k)
                    line.append("" if v is None else v)
                w.writerow(line)
        ps.append(f'Import-NavCsv "{t}"')
    ps.append(PS_TAIL)
    with open(os.path.join(out, "Create-NavjyotiLists.ps1"), "w", encoding="utf-8-sig") as f:
        f.write("\n".join(ps))
    here = os.path.dirname(__file__)
    for g in ("README.md", "PowerBI-Measures.dax", "PowerApps-Formulas.md", "PowerAutomate-Flows.md"):
        shutil.copy(os.path.join(here, g), os.path.join(out, g))
    print("kit written to", out)


PS_HEAD = r'''<#
  Navjyoti Vermicompost - create the 10 SharePoint lists and load the register data.
  Run in PowerShell 7 with PnP.PowerShell:   Install-Module PnP.PowerShell -Scope CurrentUser
  Your admin registers a PnP app once:       Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Navjyoti" -Tenant yourcompany.onmicrosoft.com
  Then run from this folder:                 ./Create-NavjyotiLists.ps1 -SiteUrl https://<yourtenant>.sharepoint.com/sites/NavjyotiProduction -ClientId PASTE-CLIENT-ID-HERE
  Safe to re-run: existing lists and columns are kept; add -SkipData to create the structure without loading rows.
#>
param(
  [Parameter(Mandatory)] [string] $SiteUrl,
  [Parameter(Mandatory)] [string] $ClientId,
  [switch] $SkipData
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
$types = @{}

function New-NavList($name) {
  if (-not (Get-PnPList -Identity $name -ErrorAction SilentlyContinue)) {
    New-PnPList -Title $name -Template GenericList -OnQuickLaunch | Out-Null
    Write-Host "Created list $name" -ForegroundColor Green
  }
  # Title is not used; make it optional and hide it from forms
  Set-PnPField -List $name -Identity "Title" -Values @{ Required = $false } | Out-Null
  $types[$name] = @{}
}

function Add-NavField($list, $name, $type, [string[]] $choices) {
  $types[$list][$name] = $type
  if (Get-PnPField -List $list -Identity $name -ErrorAction SilentlyContinue) { return }
  if ($type -eq "Choice") {
    Add-PnPField -List $list -DisplayName $name -InternalName $name -Type Choice -Choices $choices -AddToDefaultView | Out-Null
  } else {
    Add-PnPField -List $list -DisplayName $name -InternalName $name -Type $type -AddToDefaultView | Out-Null
  }
  if ($name -eq "Batch") { Set-PnPField -List $list -Identity $name -Values @{ FillInChoice = $true } | Out-Null }  # new batch codes can be typed
  if ($type -eq "DateTime") { Set-PnPField -List $list -Identity $name -Values @{ DisplayFormat = 0 } | Out-Null }  # date only
}

function Add-NavCalc($list, $name, $resultType, $formula) {
  if (Get-PnPField -List $list -Identity $name -ErrorAction SilentlyContinue) { return }
  $xml = "<Field Type='Calculated' DisplayName='$name' Name='$name' StaticName='$name' ResultType='$resultType' ReadOnly='TRUE'><Formula>$([System.Security.SecurityElement]::Escape($formula))</Formula></Field>"
  Add-PnPFieldFromXml -List $list -FieldXml $xml | Out-Null
  $v = Get-PnPView -List $list -Identity "All Items" -Includes ViewFields
  $v.ViewFields.Add($name); $v.Update(); Invoke-PnPQuery
}

function Import-NavCsv($list) {
  if ($SkipData) { return }
  $path = Join-Path $PSScriptRoot "data/$list.csv"
  if (-not (Test-Path $path)) { return }
  if ((Get-PnPListItem -List $list -PageSize 1 | Measure-Object).Count -gt 0) {
    Write-Host "  $list already has rows - skipping import" -ForegroundColor Yellow; return
  }
  $rows = Import-Csv $path; $n = 0
  $batch = New-PnPBatch
  foreach ($r in $rows) {
    $vals = @{}
    foreach ($p in $r.PSObject.Properties) {
      if ([string]::IsNullOrWhiteSpace($p.Value)) { continue }
      switch ($types[$list][$p.Name]) {
        "Number"   { $vals[$p.Name] = [double]$p.Value }
        "DateTime" { $vals[$p.Name] = [datetime]::ParseExact($p.Value, "yyyy-MM-dd", $null) }
        default    { $vals[$p.Name] = $p.Value }
      }
    }
    Add-PnPListItem -List $list -Values $vals -Batch $batch | Out-Null; $n++
  }
  Invoke-PnPBatch -Batch $batch
  Write-Host "  $list : $n rows loaded" -ForegroundColor Green
}
'''

PS_TAIL = r'''
Write-Host "`nDone. Open $SiteUrl/_layouts/15/viewlsts.aspx to see the lists." -ForegroundColor Cyan
'''

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "B4")
