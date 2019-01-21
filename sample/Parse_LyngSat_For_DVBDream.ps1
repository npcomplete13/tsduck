#-----------------------------------------------------------------------------
#
#  TSDuck - The MPEG Transport Stream Toolkit
#  Copyright (c) 2005-2018, Thierry Lelegard
#  All rights reserved.
#
#  Redistribution and use in source and binary forms, with or without
#  modification, are permitted provided that the following conditions are met:
#
#  1. Redistributions of source code must retain the above copyright notice,
#     this list of conditions and the following disclaimer.
#  2. Redistributions in binary form must reproduce the above copyright
#     notice, this list of conditions and the following disclaimer in the
#     documentation and/or other materials provided with the distribution.
#
#  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
#  AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
#  IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
#  ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
#  LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
#  CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
#  SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
#  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
#  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
#  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
#  THE POSSIBILITY OF SUCH DAMAGE.
#
#-----------------------------------------------------------------------------

<#
 .SYNOPSIS

  Sample Windows PowerShell script which gets the descriptions of all
  transponders for a few satellites on LyngSat. Create one text file
  per satellite with TSDuck tuning options, one line per transponder.

 .PARAMETER NoPause

  Do not wait for the user to press <enter> at end of execution. By default,
  execute a "pause" instruction at the end of execution, which is useful
  when the script was run from Windows Explorer.
#>
param([switch]$NoPause = $false)


#----------------------------------------------------------------------------
# Get first child of an HTML element ($null safe) with specified tag name
# (case insensitive).
#----------------------------------------------------------------------------

function GetFirstChild($node, [string] $tag)
{
    $ltag = $tag.ToLower()
    if ($node -ne $null) {
        for ($e = $node.firstChild; $e -ne $null; $e = $e.nextSibling) {
            if ($e.tagName.ToLower() -eq $ltag) {
                return $e
            }
        }
    }
    return $null
}


#----------------------------------------------------------------------------
# Get next sibling of an HTML element ($null safe) with specified tag name
# (case insensitive).
#----------------------------------------------------------------------------

function GetNextChild($child, [string] $tag)
{
    $ltag = $tag.ToLower()
    if ($child -ne $null) {
        for ($e = $child.nextSibling; $e -ne $null; $e = $e.nextSibling) {
            if ($e.tagName.ToLower() -eq $ltag) {
                return $e
            }
        }
    }
    return $null
}


#----------------------------------------------------------------------------
# Get the packed plain text of an element, $null safe.
#----------------------------------------------------------------------------

function ElementText($elem)
{
    if ($elem -eq $null) {
        return ""
    }
    else {
        return $elem.innerText -replace '&nbsp;',' ' -replace '\s+',' ' -replace '^ *','' -replace ' *$',''
    }
}


#----------------------------------------------------------------------------
# Get a LyngSat page and parse all transponders.
#----------------------------------------------------------------------------

function ParseLyngSat([string] $url, [string] $outFile)
{
    # An array containing all lines of the output file.
    $output = @()
    $set = @{}

    # Fetch the Web page.
    Write-Output "Fetching $url"
    $page = Invoke-WebRequest $url
    Write-Output "Status: $($page.StatusCode), $($page.StatusDescription), size: $($page.RawContentLength)"

    # Loop on all rows in all tables
    foreach ($row in $page.ParsedHtml.getElementsByTagName("tr")) {

        # Get second column in the row
        $col = GetNextChild (GetFirstChild $row "td") "td"
        $desc = ElementText $col

        # Get if the text matches "frequence polarity"
        if ($desc -match '\d+ [HLRV] .*') {

            $fields = -split $desc
            $freq = $fields[0]
            $system = "DVBS"
            $modulation = "QPSK"
            $symbols = $null
            $fec = $null

            # Search other parameters in subsequent columns
            for ($e = (GetNextChild $col "td"); ($e -ne $null) -and (-not $symbols -or -not $fec); $e = (GetNextChild $e "td")) {
                $text = ElementText $e
                if ($text -ilike '*DVB-S2*') {
                    $system = "DVB-S2"
                }
                if ($fields[1] -ilike 'H') {
                    $polarity = "H"
                }
                if ($fields[1] -ilike 'L') {
                    $polarity = "L"
                }
                if ($fields[1] -ilike 'R') {
                    $polarity = "R"
                }
                if ($fields[1] -ilike 'V') {
                    $polarity = "V"
                }
                if ($text -ilike '*8PSK*') {
                    $system = "DVBS"
                    $modulation = "PSK/8"
                }
                if ($text -match '^\d+-\d+/\d+.*') {
                    $fields = $text -split '[\s-]'
                    $symbols = $fields[0]
                    $fec = $fields[1].split('/')[0] + [System.String]::Format("{0:X}",$fields[1].split('/')[1])
                }
            }

            if ($symbols -and $fec) {
                if (!$set.Contains(${freq})) {
                    $line = "--frequency ${freq}000000 --polarity $polarity --symbol-rate ${symbols}000 --fec $fec --delivery $system --modulation $modulation"
                    #$output += "[CHANNEL]"
                    #$output += "        DELIVERY_SYSTEM = DVBS"
                    #$output += "        FREQUENCY = ${freq}000"
                    #$output += "        POLARIZATION = $polarity"
                    #$output += "        SYMBOL_RATE = ${symbols}000"
                    #$output += "        INNER_FEC = $fec"
                    #$output += "        MODULATION = $modulation"
                    #$output += "        INVERSION = AUTO"
                    $output += "$($output.Count+1)=${freq},$polarity,${symbols},${fec}"
                    Write-Output "${desc}: $line"
                    $set.Add(${freq},${freq})
                }
            }
        }
    }

    # Now create the file.
    Write-Output "Found $($output.Count) transponders, writing $outFile"
    $output | Set-Content $outFile
}


#-----------------------------------------------------------------------------
# Main code.
#-----------------------------------------------------------------------------

# Get the description of a few satellites.
ParseLyngSat "https://www.lyngsat.com/Anik-F3.html" "AnikF3-DVBDream"
ParseLyngSat "https://www.lyngsat.com/QuetzSat-1.html" "QuetzSat1-DVBDream"
ParseLyngSat "https://www.lyngsat.com/Nimiq-5.html" "Nimiq5-DVBDream"
ParseLyngSat "https://www.lyngsat.com/EchoStar-16.html" "EchoStar16-DVBDream"

# Exit script.
if (-not $NoPause) {
    pause
}
