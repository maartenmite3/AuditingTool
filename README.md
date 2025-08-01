Before you can use the CVE_Report_GUI.exe you need to have free NVD API key.
you can request one at : https://nvd.nist.gov/developers/request-an-api-key
past the api key on the first line of the nvd_api_key.txt

the source is include as well for those who want to use python instead of the executable.
or rebuilt it with pyinstaller from the source.

You can check the file integrity by using powershell.

Get-ChildItem "YOUR_FOLDER" -Recurse -File | Get-FileHash -Algorithm SHA256 

Algorithm       Hash                                                                   Path                                                                 
---------       ----                                                                   ----                                                                 
SHA256          901734591EB355562814ED33B34F20AFE11D927D5E3AFABD8490BC88924BD124       audit_template.docx                            
SHA256          D85BDAA8C60F33546D5C9B21C1356689227EAFDAEA05B70459B5256E0968F494       CIS_Report_Gui.exe                             
SHA256          D4EC0B2FA512F28688C1974EF13FAC80903D921AD0F6BD0AC51F86218D87C015       CIS_Report_Gui.py                              
SHA256          C4220B830DD2FBA664312A3729796BA91CDD63AE6576F632234D087AEE203FE8       CVE_Report_Gui.exe                             
SHA256          A689DB7A551AB214B87AEE9C916B4483CCC497B5129B67CE4A021C5798DACA6F       CVE_Report_Gui.py                              
SHA256          E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855       nvd_api_key.txt                                
                                      


