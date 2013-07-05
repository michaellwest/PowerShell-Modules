# Website: http://blogs.msdn.com/b/powershell/archive/2012/07/13/join-object.aspx

function AddItemProperties($item, $properties, $output) {
    if($item -ne $null) {
        foreach($property in $properties) {
            if($item -as [hashtable]) {
                $output."$($property)" = $item."$($property)"
                continue
            }
            $propertyHash = $property -as [hashtable]
            if($propertyHash -ne $null) {
                $hashName = $propertyHash["Name"] -as [string]
                if($hashName -eq $null) {
                    throw "there should be a string Name"  
                }
         
                $expression = $propertyHash["Expression"] -as [scriptblock]
                if($expression -eq $null) {
                    throw "there should be a ScriptBlock Expression"  
                }
         
                $_=$item
                $expressionValue = & $expression
         
                $output."$($hashName)" = $expressionValue
            } else {
                # .psobject.Properties allows you to list the properties of any object, also known as "reflection"
                foreach($itemProperty in $item.psobject.Properties) {
                    if ($itemProperty.Name -like $property) {
                        $output."$($itemProperty.Name)" = $itemProperty.Value
                    }
                }
            }
        }
    }
}
    
function WriteJoinObjectOutput($leftItem, $rightItem, $leftProperties, $rightProperties, $Type) {
    $output = New-Object PSObject

    foreach($property in @($leftProperties + $rightProperties)) {
        if($property -as [hashtable]) {
            $output | Add-Member -MemberType "NoteProperty" -Name $property.Name -Value $null
        } else {
            $output | Add-Member -MemberType "NoteProperty" -Name $property -Value $null
        }
    }

    if($Type -eq "AllInRight") {
        # This mix of rightItem with LeftProperties and vice versa is due to
        # the switch of Left and Right arguments for AllInRight
        AddItemProperties $rightItem $leftProperties $output
        AddItemProperties $leftItem $rightProperties $output
    } else {
        AddItemProperties $leftItem $leftProperties $output
        AddItemProperties $rightItem $rightProperties $output
    }
    $output
}

function Join-Object {
    <#
        .SYNOPSIS
           Joins two object collections.

        .DESCRIPTION
           Joins two object collections using the specified properties or expression.

        .PARAMETER Left
            Left side collection to join with the $Right collection

        .PARAMETER Right
            Right side collection to join with the $Left collection.

        .PARAMETER WhereScript
            ScriptBlock expression to determine when a match occurs.
                Example: {$whereLeft.Id -eq $whereRight.Id}
                Example: {$args[0].Id -eq $args[1].Id}

        .PARAMETER LeftProperties
            Properties from the $Left collection we want to return in the output.
            Each property can:
             - Be a plain property name like "Name"
             - Contain wildcards like "*"
             - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}. Name is the output property name
               and Expression is the property value. The same syntax is available in Select-Object and it is 
               important for Join-Object because joined lists could have a property with the same name

        .PARAMETER RightProperties
            Properties from $Right we want to return in the output.
            Like LeftProperties, each can be a plain name, wildcard or hashtable. See the LeftProperties comments.

        .PARAMETER LeftProperty
            Property from the $Left collection used to match with the $Right collection.
            This is used as an alternative to the $WhereScript.

        .PARAMETER RightProperty
            Property from the $Right collection used to match with the $Left collection.
            This is used as an alternative to the $WhereScript.

        .PARAMETER Type
            Type of join to perform on the $Left and $Right. 
            
            AllInLeft
             - All the $Left items are returned. Any match found in the $Right collection will be returned. 
            AllInRight
             - All the $Right items are returned. Any match found in the $Left collection will be returned.
            OnlyIfInBoth
             - All the $Left and $Right items are returned if a match is found.
            AllInBoth
             - All the $Left and $Right items are returned.

        .EXAMPLE
           PS C:\> $a = @(@{"Id"=1;"FirstName"="Michael"},@{"Id"=2;"FirstName"="John"})
           PS C:\> $b = @(@{"Id"=1;"LastName"="West"},@{"Id"=2;"LastName"="Smith"})
           PS C:\> Join-Object -Left $a -Right $b -LeftProperty "Id" -RightProperty "Id" -LeftProperties "FirstName" -RightProperties "LastName"

           FirstName LastName
           --------- --------
           Michael   West
           John      Smith 

           Join $a with $b using the property "Id", then return the properties "FirstName" and "LastName". Both collections are arrays of hashtable.

        .EXAMPLE
           PS C:\> $a = @([PSCustomObject]@{"Id"=1;"FirstName"="Michael";"LastName"="West"},[PSCustomObject]@{"Id"=2;"FirstName"="John";"LastName"="Smith"})
           PS C:\> $b = @([PSCustomObject]@{"Id"=1;"Location"="Home"},[PSCustomObject]@{"Id"=2;"Location"="Home"})
           PS C:\> Join-Object -Left $a -Right $b -LeftProperty "Id" -RightProperty "Id" -LeftProperties "FirstName","LastName" -RightProperties "Id",@{Name="Location Name";Expression={$_.Location}}

           FirstName LastName Id Location Name
           --------- -------- -- -------------
           Michael   West      1 Home
           John      Smith     2 Home

           Join $a with $b using the property "Id", then return the properties "FirstName", "LastName", "Id", and "Location" renamed to "Location Name". Both collections are arrays of PSCustomObject.

        .EXAMPLE
           PS C:\> $a = @([PSCustomObject]@{"Id"=1;"FirstName"="Michael";"LastName"="West"},[PSCustomObject]@{"Id"=2;"FirstName"="John";"LastName"="Smith"},[PSCustomObject]@{"Id"=3;"FirstName"="Jane";"LastName"="Doe"})
           PS C:\> $b = @([PSCustomObject]@{"Id"=1;"Location"="Spectrum"},[PSCustomObject]@{"Id"=2;"Location"="Spectrum"})
           PS C:\> Join-Object -Left $a -Right $b -WhereScript {param($whereLeft,$whereRight) $whereLeft.Id -eq $whereRight.Id} -LeftProperties "FirstName" -RightProperties "Location" -Type AllInLeft

           FirstName Location
           --------- --------
           Michael   Home
           John      Home
           Jane

           Join $a with $b using the ScriptBlock, then return the properties "FirstName" and "Location" including all items in $a. Both collections are arrays of PSCustomObject.
           Notice the use of local variables in the ScriptBlock. You could also use {$args[0].Id -eq $args[1].Id} to achieve the same result.
            
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [object[]]$Left,

        [Parameter(Mandatory=$true, Position=1)]
        [object[]]$Right,

        [Parameter(Mandatory=$true, Position=2, ParameterSetName="Expression")]
        [scriptblock]$WhereScript,

        [Parameter(Mandatory=$true, Position=3)]
        [object[]]$LeftProperties,

        [Parameter(Mandatory=$true, Position=4)]
        [object[]]$RightProperties,

        [Parameter(Mandatory=$true, ParameterSetName="Property")]
        [string]$LeftProperty,

        [Parameter(Mandatory=$true, ParameterSetName="Property")]
        [string]$RightProperty,

        [Parameter(Mandatory=$false)]
        [ValidateSet("AllInLeft","OnlyIfInBoth","AllInBoth", "AllInRight")]
        [string]$Type="OnlyIfInBoth"
    )

    begin {
        # a list of the matches in right for each object in left
        $leftMatchesInRight = New-Object System.Collections.Generic.List[object]

        # the count for all matches  
        $rightCount = $Right.Count
        $leftCount = $Left.Count
        $rightMatchesCount = New-Object "object[]" $rightCount
        for($i = 0; $i -lt $rightCount; $i++) {
            $rightMatchesCount[$i] = 0
        }

        $leftLookup = [Ordered]@{}
        foreach($leftItem in $Left) {
            if ($leftItem."$($LeftProperty)") {
                $leftLookup[[string]$leftItem."$($LeftProperty)"] += @($leftItem)
            }
        }
        $rightLookup = [Ordered]@{}
        foreach($rightItem in $Right) {
            if ($rightItem."$($RightProperty)") {
                $rightLookup[[string]$rightItem."$($RightProperty)"] += @($rightItem)
            }
        }
    }

    process {
        if($Type -eq "AllInRight") {
            # for AllInRight we just switch Left and Right
            $aux = $Left
            $Left = $Right
            $Right = $aux

            $aux = $leftCount
            $leftCount = $rightCount
            $rightCount = $aux

            $aux = $leftLookup
            $leftLookup = $rightLookup
            $rightLookup = $aux

            $aux = $LeftProperty
            $LeftProperty = $RightProperty
            $RightProperty = $aux
        }

        if($WhereScript) {
            # go over items in $Left and produce the list of matches
            foreach($leftItem in $Left) {
                $leftItemMatchesInRight = New-Object System.Collections.Generic.List[object]
                $leftMatchesInRight.Add($leftItemMatchesInRight) | Out-Null

                for($i = 0; $i -lt $rightCount; $i++) {
                    $rightItem = $right[$i]

                    if($Type -eq "AllInRight") {
                        # For AllInRight, we want $args[0] to refer to the left and $args[1] to refer to right,
                        # but since we switched left and right, we have to switch the where arguments
                        $whereLeft = $rightItem
                        $whereRight = $leftItem
                    } else {
                        $whereLeft = $leftItem
                        $whereRight = $rightItem
                    }

                    if(Invoke-Command -ScriptBlock $WhereScript -ArgumentList $whereLeft,$whereRight) {
                        $leftItemMatchesInRight.Add($rightItem) | Out-Null
                        $rightMatchesCount[$i]++
                    }            
                }
            }
        } else {
            foreach($leftItem in $Left) {
                $leftItemMatchesInRight = New-Object System.Collections.Generic.List[object]
                $leftMatchesInRight.Add($leftItemMatchesInRight) | Out-Null

                $val = $rightLookup[[string]$leftItem."$($LeftProperty)"]
                if($val) {
                    $leftItemMatchesInRight.AddRange($val) | Out-Null
                }
            }
        }
        # go over the list of matches and produce output
        for($i = 0; $i -lt $leftCount; $i++) {
            $leftItemMatchesInRight = $leftMatchesInRight[$i]
            $leftItem=$left[$i]
                               
            if($leftItemMatchesInRight.Count -eq 0) {
                if($Type -ne "OnlyIfInBoth") {
                    WriteJoinObjectOutput $leftItem  $null  $LeftProperties  $RightProperties $Type
                }

                continue
            }

            foreach($leftItemMatchInRight in $leftItemMatchesInRight) {
                WriteJoinObjectOutput $leftItem $leftItemMatchInRight  $LeftProperties  $RightProperties $Type
            }
        }
    }

    end {
        #produce final output for members of right with no matches for the AllInBoth option
        if($Type -eq "AllInBoth") {
            if ($WhereScript) {
                for($i=0; $i -lt $rightCount; $i++) {
                    $rightMatchCount = $rightMatchesCount[$i]
                    if($rightMatchCount -eq 0) {
                        $rightItem = $Right[$i]
                        WriteJoinObjectOutput $null $rightItem $LeftProperties $RightProperties $Type
                    }
                }
            } else {
                foreach($rightItem in $Right) {
                    $val = $leftLookup[$rightItem."$($rightProperty)"]
                    if(-not $val) {
                        WriteJoinObjectOutput $null $rightItem $LeftProperties $RightProperties $Type
                    }
                }
            }
        }
    }
}
