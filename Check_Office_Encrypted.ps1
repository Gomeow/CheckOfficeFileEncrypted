Function f_check_office($input_file)
{
	#Start App
	Try {
		$word = New-Object -ComObject word.application
		$word.Visible = $false
		$excel = New-Object -ComObject excel.application
		$excel.Visible = $false
		$ppt = New-Object -ComObject powerpoint.application
		#$ppt.Visible = $false
		echo "OFFICE Started Successfully"
	}
	Catch{
	echo "Connecting MS Office ERR!"
	}
	#Check App
	If (( $word.version -lt 14 ) -and ($excel.version -lt 14 ) -and ($ppt.version -lt 14 ) )
	{
		echo "OFFICE2010 or Higher Required!"
		continue #exit function
	}
	#Check File
	echo "Checking Office File ......"

	$filetype = 0
	$encrypted = "NO"
    echo $input_file
	if ( $input_file -like "*.doc*" ) { $filetype = 1 }
	if ( $input_file -like "*.xls*" ) { $filetype = 2 }
	if ( $input_file -like "*.ppt*" ) { $filetype = 3 }

	switch ($filetype)
	{
	1 { # 1 Word
		Try {
			$f = $word.Documents.Open($input_file, $null, $false, $null, "-","-")
			$f.Close()
			}
		catch {
			$e = $_.Exception.Tostring()
			if ( ($e -like "*password*") -or ($e -like "*密码*") -or ($e -like "*パス*") ) { $encrypted = "YES" }
			else { $encrypted = "ERR" }
			}
		$f = $null
		echo " $input_file $encrypted " #Make Out
	}
	2 { # 2 Excel
		Try {
			$f = $excel.WorkBooks.Open($input_file,0,1,5,"")
			$f.Close()
			}
			catch {
			$e = $_.Exception.Tostring()
			if ( $e -like "*CAPS LOCK*" ) { $encrypted = "YES" }
			else { $encrypted = "ERR" }
			}
			
		$f = $null
		echo " $input_file $encrypted " #Make Out
	}
	3 { # 3 PPT
		Try {
			$f = $ppt.Presentations.Open($input_file+"::-",1,$null,0)
			}
			catch {
				$e = $_.Exception.Tostring()
				if ( ($e -like "*password*") -or ($e -like "*密码*") -or ($e -like "*パス*") ) { $encrypted = "YES" }
				else { $encrypted = "ERR" }
			}
		Try { $f.Close() } Catch {  }
		$f = $null
		echo " $input_file $encrypted " #Make Out
	}
}

#Quit App
Try { $word.Quit() } Catch {}
Try { $excel.Quit() } Catch {}
Try { $ppt.Quit() } Catch {}

}

f_check_office "E:\Excel2007_.xlsx"
pause