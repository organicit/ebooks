
function ConvertTo-EnhancedHTML {
<#
.SYNOPSIS
Provides an enhanced version of the ConvertTo-HTML command that includes
inserting an embedded CSS style sheet, JQuery, and JQuery Data Tables for
interactivity. Intended to be used with HTML fragments that are produced
by ConvertTo-EnhancedHTMLFragment. This command does not accept pipeline
input.
.PARAMETER jQueryURI
A Uniform Resource Indicator (URI) pointing to the location of the 
jQuery script file. You can download jQuery from www.jquery.com; you should
host the script file on a local intranet Web server and provide a URI
that starts with http:// or https://. Alternately, you can also provide
a file system path to the script file, although this may create security
issues for the Web browser in some configurations.

Tested with v1.8.2.

Defaults to http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js, which
will pull the file from Microsoft's ASP.NET Content Delivery Network.
.PARAMETER jQueryDataTableURI
A Uniform Resource Indicator (URI) pointing to the location of the 
jQuery Data Table script file. You can download this from www.datatables.net;
you should host the script file on a local intranet Web server and provide a URI
that starts with http:// or https://. Alternately, you can also provide
a file system path to the script file, although this may create security
issues for the Web browser in some configurations.

Tested with jQuery DataTable v1.9.4

Defaults to http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js,
which will pull the file from Microsoft's ASP.NET Content Delivery Network.
.PARAMETER CssStyleSheet
The CSS style sheet content - not a file name. If you have a CSS file,
you can load it into this parameter as follows:

    -CSSStyleSheet (Get-Content MyCSSFile.css)

Alternately, you may link to a Web server-hosted CSS file by using the
CssUri parameter.
.PARAMETER CssUri
A Uniform Resource Indicator (URI) to a Web server-hosted CSS file.
Must start with either http:// or https://.
.PARAMETER Title
A plain-text title that will be displayed in the Web browser's window
title bar. Note that not all browsers will display this.
.PARAMETER PreContent
Raw HTML to insert before all HTML fragments. Use this to specify a main
title for the report:

    -PreContent "<H1>My HTML Report</H1>"
.PARAMETER PostContent
Raw HTML to insert after all HTML fragments. Use this to specify a 
report footer:

    -PostContent "Created on $(Get-Date)"
.PARAMETER HTMLFragments
One or more HTML fragments, as produced by ConvertTo-EnhancedHTMLFragment

    -HTMLFragments $part1,$part2,$part3
.PARAMETERS CssIDsToMakeDataTables
A list of CSS IDs (corresponding to the table CSS ID names you set using
ConvertTo-EnhancedHTMLFragment) to make into interactive data tables.

    -CssIDsToMakeDataTables 'mytable1','mytable2'

Note that you SHOULD NOT include any table IDs which are to be turned into
charts.
#>
    [CmdletBinding()]
    param(
        [string]$jQueryURI = 'http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js',
        [string]$jQueryDataTableURI = 'http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js',
        [Parameter(ParameterSetName='CSSContent')][string[]]$CssStyleSheet,
        [Parameter(ParameterSetName='CSSURI')][string[]]$CssUri,
        [string]$Title = 'Report',
        [string]$PreContent,
        [string]$PostContent,
        [Parameter(Mandatory=$True)][string[]]$HTMLFragments,
        [string[]]$CssIDsToMakeDataTables
    )

    Write-Verbose "Making CSS style sheet"
    $stylesheet = ""
    if ($PSBoundParameters.ContainsKey('CssUri')) {
        $stylesheet = "<link rel=`"stylesheet`" href=`"$CssUri`" type=`"text/css`" />"
    }
    if ($PSBoundParameters.ContainsKey('CssStyleSheet')) {
        $stylesheet = $CssStyleSheet | Out-String
    }

    Write-Verbose "Creating <TITLE> and <SCRIPT> tags"
    $titletag = ""
    if ($PSBoundParameters.ContainsKey('title')) {
        $titletag = "<title>$title</title>"
    }
    $script += "<script type=`"text/javascript`" src=`"$jQueryURI`"></script>`n<script type=`"text/javascript`" src=`"$jQueryDataTableURI`"></script>"

    Write-Verbose "Combining HTML fragments"
    $body = $HTMLFragments | Out-String

    Write-Verbose "Adding Pre and Post content"
    if ($PSBoundParameters.ContainsKey('precontent')) {
        $body = "$PreContent`n$body"
    }
    if ($PSBoundParameters.ContainsKey('postcontent')) {
        $body = "$PostContent`n$body"
    }

    Write-Verbose "Adding interactivity calls"
    $datatable = ""
    if ($PSBoundParameters.ContainsKey('CssIDsToMakeDataTables')) {
        $datatable = "<script type=`"text/javascript`">"
        $datatable += '$(document).ready(function () {'
        $datatable += "`n"
        foreach ($id in $CssIDsToMakeDataTables) {
            $datatable += "`$('#$id').dataTable();`n"
        }
        $datatable += '} );'
        $datatable += "</script>"
    }

    Write-Verbose "Fixing table HTML"
    $body = $body -replace '<tr><th>','<thead><tr><th>'
    $body = $body -replace '</th></tr>','</th></tr></thead>'

    Write-Verbose "Producing final HTML"
    ConvertTo-HTML -Head "$stylesheet`n$titletag`n$script`n$datatable" -Body $body  
    Write-Debug "Finished producing final HTML"

}

function ConvertTo-EnhancedHTMLFragment {
<#
.SYNOPSIS
Creates an HTML fragment (much like ConvertTo-HTML with the -Fragment switch
that includes CSS class names for table rows, CSS class and ID names for the
table, and wraps the table in a <DIV> tag that has a CSS class and ID name.
.PARAMETER InputObject
The object to be converted to HTML. You cannot select properties using this
command; precede this command with Select-Object if you need a subset of
the objects' properties.
.PARAMETER EvenRowCssClass
The CSS class name applied to even-numbered <TR> tags. Optional, but if you
use it you must also include -OddRowCssClass.
.PARAMETER OddRowCssClass
The CSS class name applied to odd-numbered <TR> tags. Optional, but if you 
use it you must also include -EvenRowCssClass.
.PARAMETER TableCssID
The CSS ID name applied to the <TABLE> tag.
.PARAMETER DivCssID
The CSS ID name applied to the <DIV> tag which is wrapped around the table.
.PARAMETER TableCssClass
Optional. The CSS class name to apply to the <TABLE> tag.
.PARAMETER DivCssClass
Optional. The CSS class name to apply to the wrapping <DIV> tag.
.PARAMETER As
Must be 'List' or 'Table.' Defaults to Table. Actually produces an HTML
table either way; with Table the output is a grid-like display. With
List the output is a two-column table with properties in the left column
and values in the right column.
.PARAMETER Properties
A comma-separated list of properties to include in the HTML fragment.
This can be * (which is the default) to include all properties of the
piped-in object(s). In addition to property names, you can also use a
hashtable similar to that used with Select-Object. For example:

 Get-Process | ConvertTo-EnhancedHTMLFragment -As Table `
               -Properties Name,ID,@{n='VM';e={$_.VM};css={ 
                 if ($_.VM -gt 100) { 'red' }
                 else { 'green' }
                }}

This will create table cell rows with the calculated CSS class names.
E.g., for a process with a VM greater than 100, you'd get:

  <TD class="red">475858</TD>
  
You can use this feature to specify a CSS class for each table cell
based upon the contents of that cell. Valid keys in the hashtable are:

  n, name, l, or label: The table column header
  e or expression: The table cell contents
  css or csslcass: The CSS class name to apply to the <TD> tag   

.PARAMETER PreContent
Raw HTML content to be placed before the wrapping <DIV> tag.
.PARAMETER PostContent
Raw HTML content to be placed after the wrapping <DIV> tag.
.PARAMETER MakeHiddenSection
Used in conjunction with -PreContent. Adding this switch, which
needs no value, turns your -PreContent into  clickable report
section header. The section will be hidden by default, and clicking
the header will toggle its visibility.

When using this parameter, consider adding a symbol to your -PreContent
that helps indicate this is an expandable section. For example:

    -PreContent '<h2>&diams; My Section</h2>'
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$InputObject,

        [string]$EvenRowCssClass,
        [string]$OddRowCssClass,

        [Parameter(Mandatory=$True)]
        [string]$TableCssID,

        [Parameter(Mandatory=$True)]
        [string]$DivCssID,

        [ValidateSet('List','Table')]
        [string]$As = 'Table',

        [string]$DivCssClass,
        [string]$TableCssClass,

        [object[]]$Properties = '*',

        [string]$PreContent,

        [switch]$MakeHiddenSection,

        [string]$PostContent
    )
    BEGIN {
        Write-Verbose "Precontent"
        if ($PSBoundParameters.ContainsKey('PreContent')) {
            if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
                Write-Output "<span class=`"sectionheader`" onclick=`"`$('#$DivCssId').toggle(500);`">$PreContent</span>`n"
            } else {
                Write-Output $PreContent
            }
        }

        Write-Verbose "DIV"
        if ($PSBoundParameters.ContainsKey('DivCSSClass')) {
            $temp = " class=`"$DivCSSClass`""
        } else {
            $temp = ""
        }
        if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
            $temp += " style=`"display:none;`""
        }
        Write-Output "<div id=`"$DivCSSID`"$temp>"
        if ($PSBoundParameters.ContainsKey('TableCSSClass')) {
            $css = "class=`"$TableCSSClass`""
        } else {
            $css = ""
        }

        Write-Verbose "TABLE"
        write-Output "<table $css id=`"$TableCSSID`">"

        $fragment = ''
        $wrote_first_line = $false
        $even_row = $false

        if ($properties -eq '*') {
            $all_properties = $true
        } else {
            $all_properties = $false
        }

    }
    PROCESS {

        foreach ($object in $inputobject) {
            Write-Verbose "Processing object"
            $datarow = ''
            $headerrow = ''

            if ($PSBoundParameters.ContainsKey('EvenRowCSSClass') -and $PSBoundParameters.ContainsKey('OddRowCssClass')) {
                if ($even_row) {
                    $row_css = $OddRowCSSClass
                    $even_row = $false
                    Write-Verbose "Even row"
                } else {
                    $row_css = $EvenRowCSSClass
                    $even_row = $true
                    Write-Verbose "Odd row"
                }
            } else {
                $row_css = ''
                Write-Verbose "No row CSS class"
            }

            if ($all_properties) {
                $properties = $object | Get-Member -MemberType Properties | Select -ExpandProperty Name
            }

            foreach ($prop in $properties) {
                Write-Verbose "Processing property"
                $name = $null
                $value = $null
                $cell_css = ''
                if ($prop -is [string]) {
                    Write-Verbose "Property $prop"
                    $name = $Prop
                    $value = $object.($prop)
                } elseif ($prop -is [hashtable]) {
                    Write-Verbose "Property hashtable"
                    if ($prop.ContainsKey('cssclass')) { $cell_css = $Object | ForEach $prop['cssclass'] }
                    if ($prop.ContainsKey('css')) { $cell_css = $Object | ForEach $prop['css'] }
                    if ($prop.ContainsKey('n')) { $name = $prop['n'] }
                    if ($prop.ContainsKey('name')) { $name = $prop['name'] }
                    if ($prop.ContainsKey('label')) { $name = $prop['label'] }
                    if ($prop.ContainsKey('l')) { $name = $prop['l'] }
                    if ($prop.ContainsKey('e')) { $value = $Object | ForEach $prop['e'] }
                    if ($prop.ContainsKey('expression')) { $value = $tObject | ForEach $prop['expression'] }
                    if ($name -eq $null -or $value -eq $null) {
                        Write-Error "Hashtable missing Name and/or Expression key"
                    }
                } else {
                    Write-Warning "Unhandled property $prop"
                }
                if ($As -eq 'table') {
                    Write-Verbose "Adding $name to header and $value to row"
                    $headerrow += "<th>$name</th>"
                    $datarow += "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$value</td>"
                } else {
                    $wrote_first_line = $true
                    $headerrow = ""
                    $datarow = "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$name :</td><td$(if ($css -ne '') { ' class="'+$css+'"' })>$value</td>"
                    Write-Output "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
                }
            }
            if (-not $wrote_first_line -and $as -eq 'Table') {
                Write-Verbose "Writing header row"
                Write-Output "<tr>$headerrow</tr><tbody>"
                $wrote_first_line = $true
            }
            if ($as -eq 'table') {
                Write-Verbose "Writing data row"
                Write-Output "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
            }
        }
    }
    END {
        Write-Verbose "PostContent"
        if ($PSBoundParameters.ContainsKey('PostContent')) {
            $fragment = "$PostContent`n$fragment"
        }
        Write-Verbose "Done"
        Write-Output "</tbody></table></div>"
    }
}