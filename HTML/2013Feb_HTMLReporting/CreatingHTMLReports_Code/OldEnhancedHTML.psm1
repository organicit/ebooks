
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
    $script = "<script type=`"text/javascript`" src=`"$jQueryURI`"></script>`n<script type=`"text/javascript`" src=`"$jQueryDataTableURI`"></script>"

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
        
        [string]$PreContent,

        [switch]$MakeHiddenSection,

        [string]$PostContent
    )
    BEGIN {
        $objects = @()
    }
    PROCESS {
        $objects += $InputObject
    }
    END {
        Write-Verbose "Converting objects to HTML fragment"
        $params = @{'Fragment'=$True;
                    'As'=$As}
        $fragment = $objects | ConvertTo-HTML @params
        Write-Debug "Done converting"

        Write-Verbose "Converting HTML to XML"
        [xml]$xml = $fragment

        Write-Verbose "Adding attributes to table"
        $table = $xml.SelectSingleNode('table')
        $temp = $xml.CreateAttribute('id')
        $temp.value = $TableCSSID
        $table.Attributes.Append($temp) | Out-Null
        if ($PSBoundParameters.ContainsKey('TableCSSClass')) {
            $temp = $xml.CreateAttribute('class')
            $temp.value = $TableCSSClass
            $table.Attributes.Append($temp) | Out-Null
        }
        Write-Debug "Done adding ID and CLASS to TABLE"

        Write-Verbose "Adding even/odd CSS class IDs"
        if ($PSBoundParameters.ContainsKey('EvenRowCSSClass') -and $PSBoundParameters.ContainsKey('OddRowCssClass')) {
            $classname = $OddRowCSSClass
            $count = 0
            foreach ($tr in $table.tr) {
                if ($count -gt 0) {
                    if ($classname -eq $EvenRowCSSClass) {
                        $classname = $OddRowCSSClass
                    } else {
                        $classname = $EvenRowCSSClass
                    }
                    $temp = $xml.CreateAttribute('class')
                    $temp.value = $classname
                    $tr.attributes.append($temp) | Out-null
                }
                $count++
            }
        }
        Write-Debug "Done adding CLASS to TRs"

        Write-Verbose "Outputting to string"
        $fragment = $xml.innerxml | out-string
        Write-Debug "Done creating new fragment"

        Write-Verbose "Wrapping table in DIV"
        if ($PSBoundParameters.ContainsKey('DivCSSClass')) {
            $temp = " class=`"$DivCSSClass`""
        } else {
            $temp = ""
        }
        if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
            $temp += " style=`"display:none;`""
        }
        $fragment = "<div id=`"$DivCSSID`"$temp>$fragment</div>"
        Write-Debug "Done adding DIV around fragment"

        Write-Verbose "Adding pre and post content"
        if ($PSBoundParameters.ContainsKey('PreContent')) {
            if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
                $fragment = "<span class=`"sectionheader`" onclick=`"`$('#$DivCssId').toggle(500);`">$PreContent</span>`n$fragment"
            } else {
                $fragment = "$PreContent`n$fragment"
            }
        }
        if ($PSBoundParameters.ContainsKey('PostContent')) {
            $fragment = "$PostContent`n$fragment"
        }
        Write-Debug "Done adding pre and post content"

        Write-Verbose "Done"
        Write-Output $fragment
    }
}