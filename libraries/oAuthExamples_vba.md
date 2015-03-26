# VBA Project: **excelRestLibrary**
## VBA Module: **[oAuthExamples](/libraries/oAuthExamples.vba "source is here")**
### Type: StdModule  

This procedure list for repo (excelRestLibrary) was automatically created on 26/03/2015 10:03:57 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in oAuthExamples

---
VBA Procedure: **getGoogled**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Scope: **Public**  
Description: ****  

*Public Function getGoogled(scope As String, Optional replacementpackage As cJobject = Nothing, Optional clientID As String = vbNullString, Optional clientSecret As String = vbNullString, Optional complain As Boolean = True, Optional cloneFromeScope As String = vbNullString) As cOauth2*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
scope|String|False||
replacementpackage|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
clientID|String|True| vbNullString|
clientSecret|String|True| vbNullString|
complain|Boolean|True| True|
cloneFromeScope|String|True| vbNullString|


---
VBA Procedure: **testOauth2**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub testOauth2()*  

**no arguments required for this procedure**


---
VBA Procedure: **makeMyGoogleConsole**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Private**  
Description: ****  

*Private Function makeMyGoogleConsole() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **showLinkedinConsole**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub showLinkedinConsole()*  

**no arguments required for this procedure**


---
VBA Procedure: **every**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  
Description: ****  

*Private Function every(a As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **assert**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  
Description: ****  

*Private Function assert(what As Variant, message As cJobject, n As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
what|Variant|False||
message|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
n|String|False||


---
VBA Procedure: **setUpTest**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Private**  
Description: ****  

*Private Function setUpTest(dbtype As String, siloid As String, dbid As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dbtype|String|False||
siloid|String|False||
dbid|String|False||


---
VBA Procedure: **testDbAb**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function testDbAb()*  

**no arguments required for this procedure**


---
VBA Procedure: **lotsoftests**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub lotsoftests(handler As cDbAb, testData As cJobject)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
handler|[cDbAb](/libraries/cDbAb_cls.md "cDbAb")|False||
testData|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
