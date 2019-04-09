
SET rs = cmd.EXECUTE
SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")
SET objCSV = objFSO.createtextfile(FileName)

' create a header row and count the number of attributes
FOR EACH strAttribute in SPLIT(strAttributes,",")
	objcsv.write(comma & q & strAttribute & q)
	comma = "," ' all columns apart from the first column require a preceding comma
	i = i + 1
NEXT

' for each item returned by the Active Directory query
WHILE rs.eof <> TRUE AND rs.bof <> TRUE
	comma="" ' first column does not require a preceding comma
	objcsv.writeline ' Start a new line
	' For each column in the result set
	FOR j = 0 to (i - 1)
		SELECT CASE TYPENAME(rs(j).value)
		CASE "Null" ' handle null value
			objcsv.write(comma & q & q)
		CASE "Variant()" ' multi-valued attribute
			' Multi-valued attributes will be seperated by value specified in
			' "multivaluedsep" variable
			mvsep = "" 'No seperator required for first value
			objcsv.write(comma & q)
			FOR EACH strValue in rs(j).Value
				' Write value
				' single double quotes " are replaced by double double quotes ""
				objcsv.write(mvsep & REPLACE(strValue,q,q & q))
				mvsep = multivaluedsep ' seperator used when more than one value returned
			NEXT
			objcsv.write(q)
		CASE ELSE
			' Write value
			' single double quotes " are replaced by double double quotes ""
			objcsv.write(comma & q & REPLACE(rs(j).value,q,q & q) & q)
		END SELECT
		
		comma = "," ' all columns apart from the first column require a preceding comma
	NEXT
	rs.movenext
WEND

' Close csv file and ADO connection
cn.close
objCSV.Close

wscript.echo "Finished"