SystemUtil.Run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
While not Browser("Browser").Exist
	wait 1
Wend
Browser("Browser").Maximize
Browser("Browser").Navigate "https://www.google.com/" @@ hightlight id_;_1639802_;_script infofile_;_ZIP::ssf1.xml_;_

While not Browser("Browser").Page("Google").WebEdit("Buscar").Exist
	wait 3
Wend

'Seteo la palabra
Browser("Browser").Page("Google").WebEdit("Buscar").Set DataTable("p_palabra",dtGlobalSheet)
'Click Buscar
Browser("Browser").Page("Google").WebButton("Buscar con Google").Click

While not Browser("Browser").Page("Casa - Buscar con Google").WebButton("Preferencias").Exist
	wait 3
Wend

Reporter.ReportEvent micPass, "Buscar palabra", "Palabra encontrada"

Browser("Browser").Close



