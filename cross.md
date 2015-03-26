# VBA Project: excelRestLibrary
This cross reference list for repo (excelRestLibrary) was automatically created on 26/03/2015 10:03:58 by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cBrowser||restLibrary|getRestLibrary
cBrowser||restLibrary|getAndMakeJobjectFromXML
cBrowser||restLibrary|makeJobjectFromXML
cBrowser||restLibrary|getAndMakeJobjectAuto
cCell||cDataRow|create
cDataColumn||cDataSet|create
cDataRow||cDataSet|filterOk
cDataRow||cDataSet|create
cDataSet||restLibrary|restQuery
cDataSet||restLibrary|createHeadingsFromKeys
cDataSets||cDataSet|populateData
cDbAb||oAuthExamples|setUpTest
cDbAbResult||cDbAb|remove
cDbAbResult||cDbAb|count
cDbAbResult||cDbAb|getObjects
cDbAbResult||cDbAb|execute
cDbAbResult||cDbAb|query
cHeadingRow||cDataSet|Class_Initialize
cJobject||restLibrary|restQuery
cJobject||restLibrary|createHeadingsFromKeys
cJobject||restLibrary|getRestLibrary
cJobject||restLibrary|createRestLibrary
cOauth2||oAuthExamples|getGoogled
cregXLib||regXLib|rxMakeRxLib
cStringChunker||cJobject|recurseSerialize
cStringChunker||cJobject|unSplitToString
cStringChunker||cJobject|serialize
cUAMeasure||UAMeasure|registerUA
oAuthExamples|getGoogled|restLibrary|restQuery
regXLib|rxGroup|cRest|dotsTail
regXLib|rxReplace|cRest|stripDots
regXLib|rxReplace|cRest|executeSingle
UAMeasure|registerUA|restLibrary|restQuery
usefulCDataSet|makeSheetFromJob|oAuthExamples|testDbAb
usefulcJobject|JSONParse|restLibrary|makeJobjectFromXML
usefulcJobject|xmlStringToJobject|restLibrary|getAndMakeJobjectAuto
usefulEncrypt|encryptSha1|UAMeasure|getUserHash
usefulSheetStuff|wholeSheet|restLibrary|restQuery
UsefulStuff|URLEncode|cRest|execute
UsefulStuff|URLEncode|cRest|encodedUri
