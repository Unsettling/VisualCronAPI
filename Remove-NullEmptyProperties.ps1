<#
.SYNOPSIS
Removes properties with $null values from custom-object copies of 
the input objects.

.DESCRIPTION
Note that output objects are custom objects that are copies of the input
objects with copies of only those input-object properties that are not $null.

CAVEAT: If you pipe multiple objects to this function, and these objects
        differ in what properties are non-$null-valued, the default output
        format will show only the non-$null-valued properties of the FIRST object.
        Use ... | ForEach-Object { Out-String -InputObject $_ } to avoid
        this problem.

.NOTES
Since the output objects are generally of a distinct type - [pscustomobject] -
and have only NoteProperty members, use of this function only makes sense
with plain-old data objects as input.

.EXAMPLE
> [pscustomobject] @{ one = 1; two = $null; three = 3 } | Remove-NullProperties

one three
--- -----
  1     3

.LINK
https://stackoverflow.com/a/44370806/444244

#>
function Remove-NullEmptyProperties {
  param(
    [parameter(Mandatory, ValueFromPipeline)]
    [psobject] $InputObject
  )

  process {
    # Create the initially empty output object.
    $obj = [PSCustomObject]::new()
    # Loop over all input-object properties.
    foreach ($prop in $InputObject.psobject.properties) {
      # If a property is non-$null, add it to the output object.
      $p = $InputObject.$($prop.Name)
      if ($null -ne $p -And '' -ne $p) {
        Add-Member -InputObject $obj -NotePropertyName $prop.Name -NotePropertyValue $prop.Value
      }
    }
    # Give the output object a type name that reflects the type of the input
    # object prefixed with 'NonNull.' - note that this is purely informational, unless
    # you define a custom output format for this type name.
    $obj.pstypenames.Insert(0, 'NonNull.' + $InputObject.GetType().FullName)
    # Output the output object.
    $obj
  }
}