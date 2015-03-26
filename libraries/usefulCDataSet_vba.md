# VBA Project: **excelRestLibrary**
## VBA Module: **[usefulCDataSet](/libraries/usefulCDataSet.vba "source is here")**
### Type: StdModule  

This procedure list for repo (excelRestLibrary) was automatically created on 26/03/2015 10:03:57 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulCDataSet

---
VBA Procedure: **makeSheetFromJob**  
Type: **Function**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function makeSheetFromJob(job As cJobject, sheetName As String) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
sheetName|String|False||


---
VBA Procedure: **makeSheetHeadingsFromJob**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub makeSheetHeadingsFromJob(jo As cJobject, ds As cDataSet)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||


---
VBA Procedure: **rescurseSheetHeadersFromJob**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  
Description: ****  

*Private Function rescurseSheetHeadersFromJob(job As cJobject, jobHead As cJobject, Optional k As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
jobHead|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
k|String|True| vbNullString|


---
VBA Procedure: **addD3TreeItem**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  
Description: ****  

*Private Function addD3TreeItem(meOb As cJobject, ds As cDataSet, label As String, key As String, parentkey As String, Optional drd As cDataRow = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
meOb|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||
label|String|False||
key|String|False||
parentkey|String|False||
drd|[cDataRow](/libraries/cDataRow_cls.md "cDataRow")|True| Nothing|


---
VBA Procedure: **findD3Parent**  
Type: **Function**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Private**  
Description: ****  

*Private Function findD3Parent(ds As cDataSet, parentkey) As cDataRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||
parentkey|Variant|False||


---
VBA Procedure: **cleanDot**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function cleanDot(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **makeD3Tree**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  
Description: ****  

*Public Function makeD3Tree(meOb As cJobject, ds As cDataSet, dsOptions As cDataSet, Optional options As String = "options") As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
meOb|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||
dsOptions|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||
options|String|True| "options"|


---
VBA Procedure: **makeD3**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  
Description: ****  

*Private Function makeD3(meOb As cJobject, cj As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
meOb|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
cj|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
