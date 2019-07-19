#Require -Version 5.0
using namespace System.Collections

Add-Type -Path "C:\Program Files (x86)\Microsoft.NET\Primary Interop Assemblies\microsoft.mshtml.dll" 

${script:Table} = New-Object Collections.ArrayList

[String]${script:FFXIVBaseUrl} = "https://jp.finalfantasyxiv.com/"
[String]${script:PveTopUrl} = "jobguide/battle/"
[IO.FileInfo]${script:TmpFile} = New-TemporaryFile

${script:TopPage} = Invoke-WebRequest -Uri (${script:FFXIVBaseUrl} + ${script:PveTopUrl})

foreach (${local:JobUrl} in (${script:TopPage}.AllElements | Where-Object {$_.href -like "/jobguide/*/" -and $_.class -ne "jobguide__btn"} )) {
	${local:JobPage} = Invoke-WebRequest -Uri (${script:FFXIVBaseUrl} + ${local:JobUrl}.href)
	
	${script:Html} = New-Object -ComObject "HTMLFile"
	${script:Html}.IHTMLDocument2_write([System.Text.Encoding]::Utf8.GetString(${local:JobPage}.RawContentStream.GetBuffer()))

	foreach (${local:HtmlTable} in ${script:Html}.getElementsByTagName("table") | Where-Object {$_.className -eq "job__table"}){
		${local:HtmlTableHeader} = (${local:HtmlTable}.rows | Select-Object -First 1).Cells | ForEach-Object {$_.innerText -replace "\r\n","" }

		foreach (${local:Row} in ${local:HtmlTable}.rows | Select-Object -Skip 1) {
			${local:index} = 0
			${local:Values} = [ordered]@{}
			${local:Row}.Cells | ForEach-Object {
				${local:ToolTip} = ""
				if($_.innerHtml -match "data-tooltip=`"(?<class>.*?)`"") {
					${local:ToolTip} = $Matches["class"]
				}
				${local:Values} += @{${local:HtmlTableHeader}[${local:index}++] = (${local:ToolTip} + $_.innerText -replace "\r\n|クラスクエスト|ジョブクエスト"," ").Trim()}
			}
			if (![string]::IsNullOrWhiteSpace(${local:Row}.Id)) {
				${script:Table}.Add([pscustomobject]${local:Values}) | Out-Null
			}
		}
	}
}


foreach (${local:ColumnName} in (${script:Table} | ForEach-Object {($_ | Get-Member -MemberType NoteProperty | Select-Object Name)}) | Sort-Object -Property Name -Unique){
	if ((${script:Table}[0] | Get-Member -MemberType NoteProperty).Name -notcontains ${local:ColumnName}.Name){
		${script:Table}[0] | Add-Member -NotePropertyMembers @{${local:ColumnName}.Name = ""}
	}
}

${script:Table}.ToArray() | ConvertTo-Csv -NoTypeInformation | Set-Content -Encoding UTF8 FFXIV_JobActions.csv
