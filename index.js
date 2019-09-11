const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const https = require('https').Server(app);

var PORT = process.env.PORT || 5800
var excelData = [];


var strXml = "";

app.use('/upload',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			voucher();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});
app.use('/ledger',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			ledger();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});
app.use('/stock',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			stock();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});


app.use((req, res, next) => {
  res.status(404).render('404', { pageTitle: 'Page Not Found' });
});

app.listen(PORT);

console.log('listening on port ' + PORT);


 function voucher(){
  strXml = "";
strXml += "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
strXml += "<ENVELOPE>";
strXml += "<HEADER>";
strXml += "<TALLYREQUEST>Import Data</TALLYREQUEST>";
strXml += "</HEADER>";
strXml += "<BODY>";
strXml += "<IMPORTDATA>";
strXml += "<REQUESTDESC>";
strXml += "<REPORTNAME>All Masters</REPORTNAME>";
strXml += "</REQUESTDESC>";
strXml += "<REQUESTDATA>";
    

    for (var i = 0; i < excelData.length; i++)
                    {
                        let VOUCHERTYPE     = (excelData[i]["VOUCHERTYPE"]);
                        let DATE            = (excelData[i]["DATE"]);
                        let NARRATION       = (excelData[i]["NARRATION"]);
                        let VOUCHERNUMBER   = (excelData[i]["VOUCHERNUMBER"]);
                        let DRLEDGER        = (excelData[i]["DR.LEDGER"]);
                        let CRLEDGER        = (excelData[i]["CR.LEDGER"]);
                        let AMOUNT          = (excelData[i]["LEDGERAMOUNT"]);
                        let AMOUNT2          = (-(excelData[i]["LEDGERAMOUNT"]));
                        let bool            = "";
                        let bool1           = "";
                        let bool2           = "Yes";
                        if (VOUCHERTYPE == "Payment" || VOUCHERTYPE == "Journal"){
                            bool  = "Yes";
                            bool1 = "No";
                        }else {bool="No", bool1  = "Yes"};
                        
                        if (VOUCHERTYPE == "Receipt" || VOUCHERTYPE == "Contra"){
                            DRLEDGER        = (excelData[i]["CR.LEDGER"]);
                            CRLEDGER        = (excelData[i]["DR.LEDGER"]);
                        };
                      
                        
                        if (VOUCHERTYPE == "Journal"){
                             bool2           = "No";
                        };
                        if (VOUCHERTYPE == "Payment" || VOUCHERTYPE == "Journal"){
                            AMOUNT          = (-(excelData[i]["LEDGERAMOUNT"]));
                            AMOUNT2          = (excelData[i]["LEDGERAMOUNT"]);
                        };
                            
                        strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\"" + VOUCHERTYPE +  "\" ACTION=\"Create\">";
                        strXml += "<DATE>"+ DATE + "</DATE>";
                        strXml += "<GUID></GUID>";
                        strXml += "<NARRATION>" + NARRATION + "</NARRATION>";
                        strXml += "<VOUCHERTYPENAME>" + VOUCHERTYPE +" </VOUCHERTYPENAME>";
                        strXml += "<VOUCHERNUMBER>" + VOUCHERNUMBER + "</VOUCHERNUMBER>";
                        strXml += "<PARTYLEDGERNAME>" + DRLEDGER + "</PARTYLEDGERNAME>";
                        strXml += "<CSTFORMISSUETYPE/>";
                        strXml += "<CSTFORMRECVTYPE/>";
                        strXml += "<PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>";
                        strXml += "<VCHGSTCLASS/>";
                        strXml += "<HASCASHFLOW>Yes</HASCASHFLOW>";

                        strXml += "<ALLLEDGERENTRIES.LIST>";
                        strXml += "<LEDGERNAME>" + DRLEDGER + "</LEDGERNAME>";
                        strXml += "<GSTCLASS/>";

                        strXml += "<ISDEEMEDPOSITIVE>" + bool + "</ISDEEMEDPOSITIVE>";
                        strXml += "<ISPARTYLEDGER>Yes</ISPARTYLEDGER>";
                        strXml += "<ISLASTDEEMEDPOSITIVE>" + bool + "</ISLASTDEEMEDPOSITIVE>";
                        strXml += "<AMOUNT>" + AMOUNT + "</AMOUNT>";
                        strXml += "<VATEXPAMOUNT>" + AMOUNT + "</VATEXPAMOUNT>";
                        strXml += "</ALLLEDGERENTRIES.LIST>";
                        strXml += "<ALLLEDGERENTRIES.LIST>";
                        strXml += "<LEDGERNAME>" + CRLEDGER + "</LEDGERNAME>";
                        strXml += "<GSTCLASS/>";

                        strXml += "<ISDEEMEDPOSITIVE>" + bool1 + "</ISDEEMEDPOSITIVE>";
                        strXml += "<ISPARTYLEDGER>" + bool2 + "</ISPARTYLEDGER>";
                        strXml += "<ISLASTDEEMEDPOSITIVE>" + bool1 + "</ISLASTDEEMEDPOSITIVE>";
                        strXml += "<AMOUNT>" + AMOUNT2 + "</AMOUNT>";
                        strXml += "<VATEXPAMOUNT>" + AMOUNT2 + "</VATEXPAMOUNT>";
                        strXml += "</ALLLEDGERENTRIES.LIST>";
                        strXml += "</VOUCHER>";            
                        strXml += "</TALLYMESSAGE>";                      
                     };   
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};

 function ledger(){
  strXml = "";
strXml += "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
strXml += "<ENVELOPE>";
strXml += "<HEADER>";
strXml += "<TALLYREQUEST>Import Data</TALLYREQUEST>";
strXml += "</HEADER>";
strXml += "<BODY>";
strXml += "<IMPORTDATA>";
strXml += "<REQUESTDESC>";
strXml += "<REPORTNAME>All Masters</REPORTNAME>";
strXml += "</REQUESTDESC>";
strXml += "<REQUESTDATA>";
    

    for (var i = 0; i < excelData.length; i++)
                    {
                        let CUSTOMERCODE    = (excelData[i]["CUSTOMER CODE"]);
                        let NAME            = (excelData[i]["NAME"]);
                        let PARENT      	= (excelData[i]["PARENT"]);
                        let ADDRESS1  		= (excelData[i]["ADDRESS 1"]);
                        let ADDRESS2        = (excelData[i]["ADDRESS 2"]);
                        let ADDRESS3        = (excelData[i]["ADDRESS3"]);
                        let STATE         	= (excelData[i]["STATE"]);
                        let PIN          	= (excelData[i]["PIN"]);
                        let CONTACTPERSON  	= (excelData[i]["CONTACT PERSON"]);
						let TELEPHONE   	= (excelData[i]["TELEPHONE NO. "]);
						let MOBILE 			= (excelData[i]["MOBILE NO."]);
						let FAX  			= (excelData[i]["FAX"]);
						let EMAIL  			= (excelData[i]["E-MAIL"]);
						let PAN  			= (excelData[i]["PAN / IT NO."]);
						let TIN   			= (excelData[i]["TIN "]);
						let CSTNO  			= (excelData[i]["CST NO"]);
						let OPENING 		= (excelData[i]["OPENING DR/CR"]);
						let Amttype  		= (excelData[i]["Amt-type"]);
						
                        if (Amttype == "Cr" || Amttype == "CR" || Amttype == "cr"){
                            OPENING  = (OPENING * 1);
                            
                        }else {OPENING  = (OPENING * -1)};
                        

                            
                        strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\""+ NAME +"\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<ADDITIONALNAME.LIST>";
						strXml += "<ADDITIONALNAME> "+ CUSTOMERCODE +"/ADDITIONALNAME>";
						strXml += "</ADDITIONALNAME.LIST>";
                        strXml += "<GUID></GUID>";
                        strXml += "<PARENT>" + PARENT + "</PARENT>";
                        strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>" + NAME + "</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "<ADDRESS.LIST>";
                        strXml += "<ADDRESS>" + ADDRESS1 + " </ADDRESS>";
                        strXml += "<ADDRESS>" + ADDRESS2 + " </ADDRESS>";
                        strXml += "<ADDRESS>" + ADDRESS3 + " </ADDRESS>";
                        strXml += "</ADDRESS.LIST>";
	                    strXml += "<STATENAME> " + STATE + "</STATENAME>";
                        strXml += "<PINCODE> " + PIN + " </PINCODE>";
                        strXml += "<LEDGERCONTACT> " + CONTACTPERSON + " </LEDGERCONTACT>";
                        strXml += "<LEDGERPHONE> " + TELEPHONE + " </LEDGERPHONE>";
						strXml += "<LEDGERPHONE> " + MOBILE + " </LEDGERPHONE>";
                        strXml += "<LEDGERFAX> " + FAX + " </LEDGERFAX>";
                        strXml += "<EMAIL> " + EMAIL + " </EMAIL>";
                        strXml += "<INCOMETAXNUMBER> " + PAN + " </INCOMETAXNUMBER>";
                        strXml += "<VATTINNUMBER> " + TIN + " </VATTINNUMBER>";
                        strXml += "<SALESTAXNUMBER> " + CSTNO + " </SALESTAXNUMBER>";
                        strXml += "<OPENINGBALANCE> " + OPENING + " </OPENINGBALANCE>";
                        strXml += "<ISBILLWISEON>Yes</ISBILLWISEON>";
                        strXml += "<ISCOSTCENTRESON>No</ISCOSTCENTRESON>";
                        strXml += "<AFFECTSSTOCK>No</AFFECTSSTOCK>";
                        strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";                      
                     };   
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};

 function stock(){
  strXml = "";
strXml += "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
strXml += "<ENVELOPE>";
strXml += "<HEADER>";
strXml += "<TALLYREQUEST>Import Data</TALLYREQUEST>";
strXml += "</HEADER>";
strXml += "<BODY>";
strXml += "<IMPORTDATA>";
strXml += "<REQUESTDESC>";
strXml += "<REPORTNAME>All Masters</REPORTNAME>";
strXml += "</REQUESTDESC>";
strXml += "<REQUESTDATA>";
    

    for (var i = 0; i < excelData.length; i++)
                    {
                        let ALIAS    		= (excelData[i]["ALIAS"]);
                        let NAME            = (excelData[i]["NAME"]);
                        let PARENT      	= (excelData[i]["PARENT"]);
                        let CATEGORY  		= (excelData[i]["CATEGORY"]);
                        let UOM        		= (excelData[i]["UOM"]);
                        let COSTING        	= (excelData[i]["COSTING"]);
                        let OPENINGSTOCK 	= (excelData[i]["OPENING STOCK"]);
                        let OPENINGRATE     = (excelData[i]["OPENINGRATE"]);
						let Amttype  		= ((OPENINGSTOCK * OPENINGRATE)*-1 );
						
                        
                        

                            
                        strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<STOCKITEM NAME=\"" + NAME + "\" RESERVEDNAME=\"\">";
						strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
						strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
						strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
						strXml += "<ADDITIONALNAME.LIST>";
						strXml += "<ADDITIONALNAME> " + ALIAS + "</ADDITIONALNAME>";
						strXml += " </ADDITIONALNAME.LIST>";
    					strXml += " <PARENT>" + PARENT + "</PARENT>";
						strXml += " <CATEGORY>" + CATEGORY + "</CATEGORY>";
						strXml += " <OPENINGBALANCE> " + OPENINGSTOCK + "</OPENINGBALANCE>";
						strXml += " <OPENINGVALUE>  " + Amttype + "</OPENINGVALUE>";
						strXml += " <OPENINGRATE> " + OPENINGRATE + "</OPENINGRATE>";
						strXml += " <COSTINGMETHOD> " + COSTING + "</COSTINGMETHOD>";
						strXml += " <VALUATIONMETHOD>Avg. Price</VALUATIONMETHOD>";
						strXml += " <BASEUNITS> " + UOM + "</BASEUNITS>";
						strXml += " <ASORIGINAL>Yes</ASORIGINAL>";
						strXml += " <IGNOREPHYSICALDIFFERENCE>No</IGNOREPHYSICALDIFFERENCE>";
						strXml += " <IGNORENEGATIVESTOCK>No</IGNORENEGATIVESTOCK>";
						strXml += " <TREATSALESASMANUFACTURED>No</TREATSALESASMANUFACTURED>";
						strXml += " <TREATPURCHASESASCONSUMED>No</TREATPURCHASESASCONSUMED>";
						strXml += " <TREATREJECTSASSCRAP>No</TREATREJECTSASSCRAP>";
						strXml += " <HASMFGDATE>No</HASMFGDATE>";
						strXml += " <ALLOWUSEOFEXPIREDITEMS>No</ALLOWUSEOFEXPIREDITEMS>";
						strXml += " <LANGUAGENAME.LIST>";
						strXml += "  <NAME.LIST TYPE=\"String\">";
						strXml += "   <NAME> " + NAME + "</NAME>";
						strXml += "  </NAME.LIST>";
						strXml += "  <LANGUAGEID> 1033</LANGUAGEID>";
						strXml += " </LANGUAGENAME.LIST>";
						strXml += " </STOCKITEM>         ";        
                        strXml += "</TALLYMESSAGE>";                      
                     };   
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};