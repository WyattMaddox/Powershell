function Search-FileDocx ($Directory, $SearchString) { 
    #$Directory - String of the directory you want to look for .docx files in (recursivly)
    #$SearchString - The string you want to find

    #Get a list of .docx files from directory
    $word_files = Get-ChildItem -Path $Directory -Filter *.docx -Recurse | 
    Select-Object @{Name="Path";Expression={ ($_.FullName)}}, @{Name="Name";Expression={ ($_.Name)}}
    #Open Word as a COM Object
    $Word = New-Object -ComObject Word.Application
    #loop over list of word docs
    $word_files | ForEach-Object{
        #Open each doc in read only mode
        $doc = $word.documents.open( $_.Path, $false,$true );
        #Set the content of the file as a variable
        $range = $doc.content
        #Search the document content for the search string
        If ($range.Text -match $SearchString){
            #If the string is found return the document path
            $doc.FullName
        }
        #Close the Document and repeat loop
        $doc.close()
    }
    #Once Loop is finished close COM Object
    $Word.Quit()
}


#Example Search a directory recursively for a string
Search-FileDocx -Directory "C:\Users" -SearchString "TombStoneX"
