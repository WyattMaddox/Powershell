function Search-FileXlsx ($Directory, $SearchString){
    #$Directory - String of the directory you want to look for .xlsx files in (recursivly)
    #$SearchString - The string you want to find (Not case sensitive)
 
    #Get a list of .xlsx files from directory
    $excel_files = Get-ChildItem -Path $Directory -Filter *.xlsx -Recurse | 
    Select-Object @{Name="Path";Expression={ ($_.FullName)}}, @{Name="Name";Expression={ ($_.Name)}}
    #Open Excel as a COM Object
    $Excel = New-Object -ComObject Excel.Application
    #loop over list of workbooks
    $excel_files | ForEach-Object{
        #Open each doc in read only mode
        $workbook = $Excel.workbooks.open( $_.Path, 2,$true );
        #For loop to search each sheet in the workbook
        foreach ($sheet in $workbook.Sheets){ 
            #Search the sheet for $searchstring
            $text_match = $workBook.Sheets.Item($sheet.Name).UsedRange.Find($SearchString)
            If($text_match.Text -match $SearchString)
                {
                #If search string found return workbook path and sheet name
                $workbook.FullName + " || " + $sheet.Name
                }
            }
        }
        #Close the Workbook and repeat loop
        $workbook.close()
    #Once Loop is finished close COM Object
    $Excel.Quit()
}

#Example run searches C:\Users recursively for the string "TombStoneX" 
Search-FileXlsx -Directory "C:\Users" -SearchString "TombStoneX"
