# VBA Project: **excelRestLibrary**
## VBA Module: **[cDbAb](/libraries/cDbAb.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (excelRestLibrary) was automatically created on 26/03/2015 10:03:57 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDbAb

---
VBA Procedure: **constraints**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function constraints(json As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json|String|False||


---
VBA Procedure: **maybeQuote**  
Type: **Function**  
Returns: **Variant**  
Scope: **Private**  
Description: ****  

*Private Function maybeQuote(jo As cJobject) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **setEndPoint**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setEndPoint(endPoint As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
endPoint|String|False||


---
VBA Procedure: **getEndPoint**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getEndPoint() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getResult**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function getResult() As cDbAbResult*  

**no arguments required for this procedure**


---
VBA Procedure: **setSiloId**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setSiloId(id As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
id|String|False||


---
VBA Procedure: **getSiloId**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getSiloId() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **setOauth2**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setOauth2(oauth2 As cOauth2) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
oauth2|[cOauth2](/libraries/cOauth2_cls.md "cOauth2")|False||


---
VBA Procedure: **setDbId**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setDbId(id As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
id|String|False||


---
VBA Procedure: **getDbId**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getDbId() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **setPeanut**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setPeanut(id As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
id|String|False||


---
VBA Procedure: **getQueryString**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function getQueryString(Optional queryJSON As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
queryJSON|String|True| vbNullString|


---
VBA Procedure: **getParamString**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function getParamString(Optional paramJSON As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
paramJSON|String|True| vbNullString|


---
VBA Procedure: **getAuthHeader**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function getAuthHeader() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **setDbName**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setDbName(dbName As String) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dbName|String|False||


---
VBA Procedure: **getDbName**  
Type: **Function**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Function getDbName() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **browser**  
Type: **Get**  
Returns: **[cBrowser](/libraries/cBrowser_cls.md "cBrowser")**  
Scope: **Public**  
Description: ****  

*Public Property Get browser() As cBrowser*  

**no arguments required for this procedure**


---
VBA Procedure: **setNoCache**  
Type: **Function**  
Returns: **[cDbAb](/libraries/cDbAb_cls.md "cDbAb")**  
Scope: **Public**  
Description: ****  

*Public Function setNoCache(noCache As Long) As cDbAb*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
noCache|Long|False||


---
VBA Procedure: **makeUrl**  
Type: **Function**  
Returns: **String**  
Scope: **Private**  
Description: ****  

*Private Function makeUrl(action As String, Optional noCache As Long = 0, Optional keepid As Boolean = False, Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
action|String|False||
noCache|Long|True| 0|
keepid|Boolean|True| False|
queryJSON|String|True| vbNullString|
paramsJSON|String|True| vbNullString|


---
VBA Procedure: **save**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function save(obs As cJobject) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
obs|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||the data to save


---
VBA Procedure: **query**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function query(Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString, Optional noCache As Long = 0, Optional keepid As Boolean = False) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
queryJSON|String|True| vbNullString|
paramsJSON|String|True| vbNullString|
noCache|Long|True| 0|
keepid|Boolean|True| False|


---
VBA Procedure: **update**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function update(keys As cJobject, obs As cJobject) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
keys|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
obs|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **remove**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function remove(Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
queryJSON|String|True| vbNullString|
paramsJSON|String|True| vbNullString|


---
VBA Procedure: **count**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function count(Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString, Optional noCache As Long = 0) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
queryJSON|String|True| vbNullString|
paramsJSON|String|True| vbNullString|
noCache|Long|True| 0|


---
VBA Procedure: **getObjects**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Public**  
Description: ****  

*Public Function getObjects(keys As cJobject, Optional noCache As Long = 0, Optional keepid As Boolean = False) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
keys|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
noCache|Long|True| 0|
keepid|Boolean|True| False|


---
VBA Procedure: **execute**  
Type: **Function**  
Returns: **[cDbAbResult](/libraries/cDbAbResult_cls.md "cDbAbResult")**  
Scope: **Private**  
Description: ****  

*Private Function execute(action As String, Optional method As String = "GET", Optional queryJSON As String = vbNullString, Optional paramsJSON As String = vbNullString, Optional data As cJobject = Nothing, Optional keys As cJobject = Nothing, Optional noCache As Long = 0, Optional keepid As Boolean = False) As cDbAbResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
action|String|False||
method|String|True| "GET"|
queryJSON|String|True| vbNullString|
paramsJSON|String|True| vbNullString|
data|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
keys|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
noCache|Long|True| 0|
keepid|Boolean|True| False|


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **tearDown**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function tearDown()*  

**no arguments required for this procedure**
