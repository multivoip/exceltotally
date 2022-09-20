const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const https = require('https').Server(app);

var PORT = process.env.PORT || 5500
var excelData = [];



var strXml = "";
//app.use('/',(req, res, next) =>  {

	
		
		//res.end('Hello World!');
		
    

          
   	
//});

app.use('/upload',(req, res, next) =>  {

	if (req.method == 'POST') {
		console.log(req);
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

app.use('/uploaddxb',(req, res, next) =>  {

	if (req.method == 'POST') {
		console.log(req);
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			voucherdxb();
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

app.use('/sales',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			sales();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});

app.use('/salesdxb',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			salesdxb();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});


app.use('/purchase',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			purchase();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});

app.use('/husk',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			husk();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});

app.use('/SalesIBW',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			SalesIBW();
			 res.header("Access-Control-Allow-Origin", "*");
			res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
			res.writeHead(200, 'OK', {'Content-Type': 'text/html'})
            res.end(strXml)
			
        });
    }	
});


app.use('/fees',(req, res, next) =>  {

	if (req.method == 'POST') {
        var jsonString = '';

        req.on('data', function (data) {
            jsonString += data;
        });

        req.on('end', function () {
            console.log(JSON.parse(jsonString));
			excelData = JSON.parse(jsonString);
			fees();
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
						strXml += "<UNIT NAME=\"" + UOM + "\" RESERVEDNAME=\"\">";
						strXml += "<ISSIMPLEUNIT>Yes</ISSIMPLEUNIT>";
						strXml += "</UNIT>";
						strXml += "</TALLYMESSAGE>";
						
                        strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
						strXml += "<STOCKCATEGORY NAME=\"" + CATEGORY + "\" RESERVEDNAME=\"\">";
						strXml += "<PARENT/>";
						strXml += " <LANGUAGENAME.LIST>";
						strXml += "  <NAME.LIST TYPE=\"String\">";
						strXml += "   <NAME> " + CATEGORY + "</NAME>";
						strXml += "  </NAME.LIST>";
						strXml += "  <LANGUAGEID> 1033</LANGUAGEID>";
						strXml += " </LANGUAGENAME.LIST>";
						strXml += "</STOCKCATEGORY>";
						strXml += "</TALLYMESSAGE>";
						
						strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
						strXml += "<STOCKGROUP NAME=\"" + PARENT + "\" RESERVEDNAME=\"\">";
						strXml += "<PARENT/>";
						strXml += " <LANGUAGENAME.LIST>";
						strXml += "  <NAME.LIST TYPE=\"String\">";
						strXml += "   <NAME> " + PARENT + "</NAME>";
						strXml += "  </NAME.LIST>";
						strXml += "  <LANGUAGEID> 1033</LANGUAGEID>";
						strXml += " </LANGUAGENAME.LIST>";
						strXml += "</STOCKGROUP>";
						strXml += "</TALLYMESSAGE>";
						
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

 function sales(){
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
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Sales 28\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sales Accounts</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Sales 28\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 28\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 28</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 28\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 14\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 14</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 14\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 14\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 14</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 14\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Sales 18\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sales Accounts</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Sales 18\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 18\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 18</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 18\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 9\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 9</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 9\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 9\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 9</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 9\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Sales 12\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sales Accounts</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Sales 12\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 12\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 12</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 12\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 6\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 6</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 6\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 6\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 6</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 6\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Sales 5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sales Accounts</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Sales 5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 5</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 2.5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 2.5</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 2.5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 2.5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 2.5</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 2.5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Round Off\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Indirect Expenses</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<SERVICECATEGORY>\&\#4\; Not Applicable</SERVICECATEGORY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Not Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Round Off</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";					
	

						
    for (var i = 0; i < excelData.length; i++)
                    {
						if ((excelData[i]["Date"]) == null ){
                          
                        }
						else {let Particulars     = (excelData[i]["Particulars"]);
						
						strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\""+ Particulars  +"\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sundry Debtors</PARENT>";
                        strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>" + Particulars  + "</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
						
                        let salDate    		= (excelData[i]["Date"]);
                        
                        let VoucherNo   	= (excelData[i]["Voucher No."]);
						
						let Value	=0;
						let igst     =0;
						let cgst     =0;
						let sgst     =0;
						let salesleg	="Sales 28%";
						let igstleg     ="IGST 28%";
						let cgstleg     ="CGST 14%";
						let sgstleg     ="SGST 14%";
						if ((excelData[i]["Sales 28%"]) !==null){
							Value  = (excelData[i]["Sales 28%"]);
							if ((excelData[i]["IGST 28%"]) !==null){
							igst  = (excelData[i]["IGST 28%"]);
						}else {cgst = (excelData[i]["CGST 14%"]);
								sgst=(excelData[i]["SGST 14%"]);
						}
						} else {if ((excelData[i]["Sales 18%"]) !==null){
							Value  = (excelData[i]["Sales 18%"]);
							salesleg	="Sales 18%";
							igstleg     ="IGST 18%";
							cgstleg     ="CGST 9%";
							sgstleg     ="SGST 9%";
								if ((excelData[i]["IGST 18%"]) !==null){
							igst  = (excelData[i]["IGST 18%"]);
						}else {cgst = (excelData[i]["CGST 9%"]);
								sgst=(excelData[i]["SGST 9%"]);
						}							
						} else {if ((excelData[i]["Sales 12%"]) !==null){
							Value  = (excelData[i]["Sales 12%"]);
							salesleg	="Sales 12%";
							igstleg     ="IGST 12%";
							cgstleg     ="CGST 6%";
							sgstleg     ="SGST 6%";
								if ((excelData[i]["IGST 12%"]) !==null){
							igst  = (excelData[i]["IGST 12%"]);
						}else {cgst = (excelData[i]["CGST 6%"]);
								sgst=(excelData[i]["SGST 6%"]);
						}														
						} else {if ((excelData[i]["Sales 5%"]) !==null){
							Value  = (excelData[i]["Sales 5%"]);
							salesleg	="Sales 5%";
							igstleg     ="IGST 5%";
							cgstleg     ="CGST 2.5%";
							sgstleg     ="SGST 2.5%";
								if ((excelData[i]["IGST 5%"]) !==null){
							igst  = (excelData[i]["IGST 5%"]);
						}else {cgst = (excelData[i]["CGST 2.5%"]);
								sgst=(excelData[i]["SGST 2.5%"]);
						}														
						}
						}
						}
						}
                        
						let TValue        	= (excelData[i]["Gross Total"]);
						
                        let ROUND 		= 0;
						if ((excelData[i]["ROUND OFF"]) !== null ){
							ROUND 		= (excelData[i]["ROUND OFF"])
											};
						
                     
						
						
                        
                        

                            
     strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
	 strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\"Sales\" ACTION=\"Create\" OBJVIEW=\"Invoice Voucher View\">"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<DATE>" + salDate + "</DATE>"; 
     strXml += "<GUID></GUID>"; 
     strXml += "<PARTYNAME>" + Particulars + "</PARTYNAME>"; 
     strXml += "<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>"; 
     strXml += "<VOUCHERNUMBER>" + VoucherNo + "</VOUCHERNUMBER>"; 
     strXml += "<PARTYLEDGERNAME>" + Particulars + "</PARTYLEDGERNAME>"; 
     strXml += "<BASICBASEPARTYNAME>" + Particulars + "</BASICBASEPARTYNAME>"; 
     strXml += "<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"; 
     strXml += "<DIFFACTUALQTY>No</DIFFACTUALQTY>"; 
     strXml += "<ISMSTFROMSYNC>No</ISMSTFROMSYNC>"; 
     strXml += "<ASORIGINAL>No</ASORIGINAL>"; 
     strXml += "<EFFECTIVEDATE>" + salDate + "</EFFECTIVEDATE>"; 
     strXml += "<ISISDVOUCHER>No</ISISDVOUCHER>"; 
     strXml += "<ISINVOICE>Yes</ISINVOICE>"; 
     strXml += "<ISVATDUTYPAID>Yes</ISVATDUTYPAID>"; 
     strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>" + Particulars + "</LEDGERNAME>"; 
     strXml += "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += "<REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>Yes</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>-"+ TValue+"</AMOUNT>"; 
     strXml += "<BILLALLOCATIONS.LIST>"; 
     strXml += "<NAME>" + VoucherNo + "</NAME>"; 
     strXml += "<BILLTYPE>New Ref</BILLTYPE>"; 
     strXml += "<TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>"; 
     strXml += "<AMOUNT>-"+ TValue+"</AMOUNT>"; 
     strXml += "</BILLALLOCATIONS.LIST>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ igstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>"+ igst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ igst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ cgstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>"+ cgst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ cgst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ sgstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>"+ sgst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ sgst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>";
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
     strXml += "</OLDAUDITENTRYIDS.LIST>";
     strXml += "<ROUNDTYPE>Upward Rounding</ROUNDTYPE>";
     strXml += "<LEDGERNAME>ROUND OFF</LEDGERNAME>";
     strXml += "<METHODTYPE>As Total Amount Rounding</METHODTYPE>";
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>";
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>";
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>";
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>";
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>";
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>";
     strXml += "<ROUNDLIMIT> 0.25</ROUNDLIMIT>";
     strXml += "<AMOUNT>"+ ROUND +"</AMOUNT>";
     strXml += "<VATEXPAMOUNT>"+ ROUND +"</VATEXPAMOUNT>";
	 strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<ALLINVENTORYENTRIES.LIST>"; 
	 for (var t = (i+1); t < excelData.length; t++)
                    {
						if ((excelData[t]["Date"]) == null){
                          

	                    let Quantity  		= (excelData[t]["Quantity"]);
                        let Rate       		= (excelData[t]["Rate"]);					
                        let itemParticulars     = (excelData[t]["Particulars"]);
						let Rvalue       		= (Quantity * Rate);
     
     strXml += "<STOCKITEMNAME>"+itemParticulars+"</STOCKITEMNAME>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<RATE>"+Rate+"/PC</RATE>"; 
     strXml += "<AMOUNT>"+Rvalue+"</AMOUNT>"; 
     strXml += "<ACTUALQTY> "+Quantity +" PC</ACTUALQTY>"; 
     strXml += "<BILLEDQTY> "+Quantity +" PC</BILLEDQTY>"; 
     strXml += "<BATCHALLOCATIONS.LIST>"; 
     strXml += "<GODOWNNAME>Main Location</GODOWNNAME>"; 
     strXml += "<BATCHNAME>Primary Batch</BATCHNAME>"; 
     strXml += "<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>"; 
     strXml += "<AMOUNT>"+Rvalue+"</AMOUNT>"; 
     strXml += "<ACTUALQTY> "+Quantity +" PC</ACTUALQTY>"; 
     strXml += "<BILLEDQTY> "+Quantity +" PC</BILLEDQTY>"; 
     strXml += "</BATCHALLOCATIONS.LIST>"; 
	 
	                         }
						else {
							t = excelData.length;
						};
					};
     strXml += "<ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ salesleg +"</LEDGERNAME>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>"+Value+"</AMOUNT>"; 
     strXml += "</ACCOUNTINGALLOCATIONS.LIST>"; 
	 strXml += "</ALLINVENTORYENTRIES.LIST>";	 
	 strXml += "</VOUCHER>";  	 
     strXml += "</TALLYMESSAGE>";   
						};	 
                     };   
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};

 function purchase(){
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
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Purchase 28\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Purchase Accounts</PARENT>";
						strXml += "<GSTAPPLICABLE>&#4; Applicable</GSTAPPLICABLE>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTTYPEOFSUPPLY>Goods</GSTTYPEOFSUPPLY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";     
	   strXml += "<GSTDETAILS.LIST>";
       strXml += "<HSNMASTERNAME/>";
       strXml += "<TAXABILITY>Taxable</TAXABILITY>";
       strXml += "<GSTNATUREOFTRANSACTION>Purchase Taxable</GSTNATUREOFTRANSACTION>";
       strXml += "<ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>";
       strXml += "<ISNONGSTGOODS>No</ISNONGSTGOODS>";
       strXml += "<GSTINELIGIBLEITC>No</GSTINELIGIBLEITC>";
       strXml += "<INCLUDEEXPFORSLABCALC>No</INCLUDEEXPFORSLABCALC>";
       strXml += "<STATEWISEDETAILS.LIST>";
       strXml += " <STATENAME>\&\#4\; Any</STATENAME>";
       strXml += " <RATEDETAILS.LIST>";
       strXml += "  <GSTRATEDUTYHEAD>Central Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 14</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 14</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Integrated Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 28</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess on Qty</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Quantity</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
       strXml += "</STATEWISEDETAILS.LIST>";
       strXml += "</GSTDETAILS.LIST>";						
	   strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Purchase 28\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 28\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 28</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 28\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 14\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 14</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 14\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 14\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 14</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 14\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Purchase 18\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Purchase Accounts</PARENT>";
						strXml += "<GSTAPPLICABLE>&#4; Applicable</GSTAPPLICABLE>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTTYPEOFSUPPLY>Goods</GSTTYPEOFSUPPLY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";     
	   strXml += "<GSTDETAILS.LIST>";
       strXml += "<HSNMASTERNAME/>";
       strXml += "<TAXABILITY>Taxable</TAXABILITY>";
       strXml += "<GSTNATUREOFTRANSACTION>Purchase Taxable</GSTNATUREOFTRANSACTION>";
       strXml += "<ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>";
       strXml += "<ISNONGSTGOODS>No</ISNONGSTGOODS>";
       strXml += "<GSTINELIGIBLEITC>No</GSTINELIGIBLEITC>";
       strXml += "<INCLUDEEXPFORSLABCALC>No</INCLUDEEXPFORSLABCALC>";
       strXml += "<STATEWISEDETAILS.LIST>";
       strXml += " <STATENAME>\&\#4\; Any</STATENAME>";
       strXml += " <RATEDETAILS.LIST>";
       strXml += "  <GSTRATEDUTYHEAD>Central Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 9</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 9</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Integrated Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 18</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess on Qty</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Quantity</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
       strXml += "</STATEWISEDETAILS.LIST>";
       strXml += "</GSTDETAILS.LIST>";						
	   strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Purchase 18\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 18\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 18</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 18\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 9\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 9</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 9\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 9\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 9</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 9\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Purchase 12\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Purchase Accounts</PARENT>";
						strXml += "<GSTAPPLICABLE>&#4; Applicable</GSTAPPLICABLE>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTTYPEOFSUPPLY>Goods</GSTTYPEOFSUPPLY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";     
	   strXml += "<GSTDETAILS.LIST>";
       strXml += "<HSNMASTERNAME/>";
       strXml += "<TAXABILITY>Taxable</TAXABILITY>";
       strXml += "<GSTNATUREOFTRANSACTION>Purchase Taxable</GSTNATUREOFTRANSACTION>";
       strXml += "<ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>";
       strXml += "<ISNONGSTGOODS>No</ISNONGSTGOODS>";
       strXml += "<GSTINELIGIBLEITC>No</GSTINELIGIBLEITC>";
       strXml += "<INCLUDEEXPFORSLABCALC>No</INCLUDEEXPFORSLABCALC>";
       strXml += "<STATEWISEDETAILS.LIST>";
       strXml += " <STATENAME>\&\#4\; Any</STATENAME>";
       strXml += " <RATEDETAILS.LIST>";
       strXml += "  <GSTRATEDUTYHEAD>Central Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 6</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 6</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Integrated Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 12</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess on Qty</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Quantity</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
       strXml += "</STATEWISEDETAILS.LIST>";
       strXml += "</GSTDETAILS.LIST>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Purchase 12\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"IGST 12\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Integrated Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 12</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>IGST 12\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 6\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 6</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 6\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 6\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 6</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 6\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Purchase 5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Purchase Accounts</PARENT>";
						strXml += "<GSTAPPLICABLE>&#4; Applicable</GSTAPPLICABLE>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTTYPEOFSUPPLY>Goods</GSTTYPEOFSUPPLY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<AFFECTSSTOCK>Yes</AFFECTSSTOCK>";     
	   strXml += "<GSTDETAILS.LIST>";
       strXml += "<HSNMASTERNAME/>";
       strXml += "<TAXABILITY>Taxable</TAXABILITY>";
       strXml += "<GSTNATUREOFTRANSACTION>Purchase Taxable</GSTNATUREOFTRANSACTION>";
       strXml += "<ISREVERSECHARGEAPPLICABLE>No</ISREVERSECHARGEAPPLICABLE>";
       strXml += "<ISNONGSTGOODS>No</ISNONGSTGOODS>";
       strXml += "<GSTINELIGIBLEITC>No</GSTINELIGIBLEITC>";
       strXml += "<INCLUDEEXPFORSLABCALC>No</INCLUDEEXPFORSLABCALC>";
       strXml += "<STATEWISEDETAILS.LIST>";
       strXml += " <STATENAME>\&\#4\; Any</STATENAME>";
       strXml += " <RATEDETAILS.LIST>";
       strXml += "  <GSTRATEDUTYHEAD>Central Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 2.50</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 2.50</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Integrated Tax</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += " <GSTRATE> 5</GSTRATE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>Cess on Qty</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Quantity</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
        strXml += "<RATEDETAILS.LIST>";
        strXml += " <GSTRATEDUTYHEAD>State Cess</GSTRATEDUTYHEAD>";
        strXml += " <GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>";
        strXml += "</RATEDETAILS.LIST>";
       strXml += "</STATEWISEDETAILS.LIST>";
       strXml += "</GSTDETAILS.LIST>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Purchase 5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
	strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"CGST 2.5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>Central Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 2.5</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>CGST 2.5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"SGST 2.5\%\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Duties &amp\; Taxes</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>GST</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<GSTDUTYHEAD>State Tax</GSTDUTYHEAD>";
						strXml += "<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>";
						strXml += "<RATEOFTAXCALCULATION> 2.5</RATEOFTAXCALCULATION>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>SGST 2.5\%</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\"Round Off\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Indirect Expenses</PARENT>";
						strXml += "<TAXCLASSIFICATIONNAME/>";
						strXml += "<TAXTYPE>Others</TAXTYPE>";
						strXml += "<GSTTYPE/>";
						strXml += "<APPROPRIATEFOR/>";
						strXml += "<SERVICECATEGORY>\&\#4\; Not Applicable</SERVICECATEGORY>";
						strXml += "<VATAPPLICABLE>\&\#4\; Not Applicable</VATAPPLICABLE>";
						strXml += "<ASORIGINAL>Yes</ASORIGINAL>";
						strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>Round Off</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";					
	

	

						
    for (var i = 0; i < excelData.length; i++)
                    {
						if ((excelData[i]["Date"]) == null ){
                          
                        }
						else {let Particulars     = (excelData[i]["Particulars"]);
						
						strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<LEDGER NAME=\""+ Particulars  +"\" RESERVEDNAME=\"\">";
                        strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
                        strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
                        strXml += "</OLDAUDITENTRYIDS.LIST>";
						strXml += "<GUID></GUID>";
                        strXml += "<PARENT>Sundry Creditors</PARENT>";
                        strXml += "<LANGUAGENAME.LIST>";
                        strXml += "<NAME.LIST TYPE=\"String\">";
                        strXml += "<NAME>" + Particulars  + "</NAME>";
                        strXml += "</NAME.LIST>";
                        strXml += "<LANGUAGEID> 1033</LANGUAGEID>";
                        strXml += "</LANGUAGENAME.LIST>";
	                    strXml += "</LEDGER>";	                    
                        strXml += "</TALLYMESSAGE>";
						
                        let salDate    		= (excelData[i]["Date"]);
                        
                        let VoucherNo   	= (excelData[i]["Voucher No."]);
						
						let Value	=0;
						let igst     =0;
						let cgst     =0;
						let sgst     =0;
						let salesleg	="Purchase 28%";
						let igstleg     ="IGST 28%";
						let cgstleg     ="CGST 14%";
						let sgstleg     ="SGST 14%";
						if ((excelData[i]["Purchase 28%"]) !==null){
							Value  = (excelData[i]["Purchase 28%"]);
							if ((excelData[i]["IGST 28%"]) !==null){
							igst  = (excelData[i]["IGST 28%"]);
						}else {cgst = (excelData[i]["CGST 14%"]);
								sgst=(excelData[i]["SGST 14%"]);
						}
						} else {if ((excelData[i]["Purchase 18%"]) !==null){
							Value  = (excelData[i]["Purchase 18%"]);
							salesleg	="Purchase 18%";
							igstleg     ="IGST 18%";
							cgstleg     ="CGST 9%";
							sgstleg     ="SGST 9%";
								if ((excelData[i]["IGST 18%"]) !==null){
							igst  = (excelData[i]["IGST 18%"]);
						}else {cgst = (excelData[i]["CGST 9%"]);
								sgst=(excelData[i]["SGST 9%"]);
						}							
						} else {if ((excelData[i]["Purchase 12%"]) !==null){
							Value  = (excelData[i]["Purchase 12%"]);
							salesleg	="Purchase 12%";
							igstleg     ="IGST 12%";
							cgstleg     ="CGST 6%";
							sgstleg     ="SGST 6%";
								if ((excelData[i]["IGST 12%"]) !==null){
							igst  = (excelData[i]["IGST 12%"]);
						}else {cgst = (excelData[i]["CGST 6%"]);
								sgst=(excelData[i]["SGST 6%"]);
						}														
						} else {if ((excelData[i]["Purchase 5%"]) !==null){
							Value  = (excelData[i]["Purchase 5%"]);
							salesleg	="Purchase 5%";
							igstleg     ="IGST 5%";
							cgstleg     ="CGST 2.5%";
							sgstleg     ="SGST 2.5%";
								if ((excelData[i]["IGST 5%"]) !==null){
							igst  = (excelData[i]["IGST 5%"]);
						}else {cgst = (excelData[i]["CGST 2.5%"]);
								sgst=(excelData[i]["SGST 2.5%"]);
						}														
						}
						}
						}
						}
                        
						let TValue        	= (excelData[i]["Gross Total"]);
						
                        let ROUND 		= 0;
						if ((excelData[i]["ROUND OFF"]) !== null ){
							ROUND 		= (excelData[i]["ROUND OFF"])
											};
						
						
                       
                        

                            
     strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
	 strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\" Purchase  \" ACTION=\"Create\" OBJVIEW=\"Invoice Voucher View\">"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += " </OLDAUDITENTRYIDS.LIST>"; 
     strXml += " <DATE>" + salDate + "</DATE>"; 
     strXml += " <GUID></GUID>"; 
     strXml += " <PARTYNAME>" + Particulars + "</PARTYNAME>"; 
     strXml += " <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>"; 
     strXml += " <VOUCHERNUMBER>" + VoucherNo + "</VOUCHERNUMBER>"; 
     strXml += " <PARTYLEDGERNAME>" + Particulars + "</PARTYLEDGERNAME>"; 
     strXml += " <BASICBASEPARTYNAME>" + Particulars + "</BASICBASEPARTYNAME>"; 
     strXml += " <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"; 
     strXml += " <DIFFACTUALQTY>No</DIFFACTUALQTY>"; 
     strXml += " <ISMSTFROMSYNC>No</ISMSTFROMSYNC>"; 
     strXml += " <ASORIGINAL>No</ASORIGINAL>"; 
     strXml += " <EFFECTIVEDATE>" + salDate + "</EFFECTIVEDATE>"; 
     strXml += " <ISISDVOUCHER>No</ISISDVOUCHER>"; 
     strXml += " <ISINVOICE>Yes</ISINVOICE>"; 
     strXml += " <ISVATDUTYPAID>Yes</ISVATDUTYPAID>"; 
     strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "  <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += " </OLDAUDITENTRYIDS.LIST>"; 
     strXml += " <LEDGERNAME>" + Particulars + "</LEDGERNAME>"; 
     strXml += " <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += " <LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += " <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += " <ISPARTYLEDGER>Yes</ISPARTYLEDGER>"; 
     strXml += " <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += " <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += " <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += " <AMOUNT>"+TValue+"</AMOUNT>"; 
     strXml += " <BILLALLOCATIONS.LIST>"; 
     strXml += " <NAME>" + VoucherNo + "</NAME>"; 
     strXml += "  <BILLTYPE>New Ref</BILLTYPE>"; 
     strXml += "  <TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>"; 
     strXml += "  <AMOUNT>"+TValue+"</AMOUNT>"; 
     strXml += " </BILLALLOCATIONS.LIST>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ igstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>-"+ igst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ igst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ cgstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>-"+ cgst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ cgst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<LEDGERNAME>"+ sgstleg +"</LEDGERNAME>"; 
     strXml += "<METHODTYPE>VAT</METHODTYPE>"; 
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>"; 
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "<AMOUNT>-"+ sgst +"</AMOUNT>"; 
     strXml += "<VATEXPAMOUNT>"+ sgst +"</VATEXPAMOUNT>"; 
     strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<LEDGERENTRIES.LIST>";
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
     strXml += "</OLDAUDITENTRYIDS.LIST>";
     strXml += "<ROUNDTYPE>Upward Rounding</ROUNDTYPE>";
     strXml += "<LEDGERNAME>ROUND OFF</LEDGERNAME>";
     strXml += "<METHODTYPE>As Total Amount Rounding</METHODTYPE>";
     strXml += "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>";
     strXml += "<LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += "<REMOVEZEROENTRIES>Yes</REMOVEZEROENTRIES>";
     strXml += "<ISPARTYLEDGER>No</ISPARTYLEDGER>";
     strXml += "<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>";
     strXml += "<ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>";
     strXml += "<ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>";
     strXml += "<ROUNDLIMIT> 0.25</ROUNDLIMIT>";
     strXml += "<AMOUNT>-"+ ROUND +"</AMOUNT>";
     strXml += "<VATEXPAMOUNT>"+ ROUND +"</VATEXPAMOUNT>";
	 strXml += "</LEDGERENTRIES.LIST>"; 
	 strXml += "<ALLINVENTORYENTRIES.LIST>";
	 for (var t = (i+1); t < excelData.length; t++)
                    {
						if ((excelData[t]["Date"]) == null ){
                          

	                    let Quantity  		= (excelData[t]["Quantity"]);
                        let Rate       		= (excelData[t]["Rate"]);					
                        let itemParticulars     = (excelData[t]["Particulars"]);
						let Rvalue       		= (Quantity * Rate);
      
     strXml += " <STOCKITEMNAME>"+itemParticulars+"</STOCKITEMNAME>"; 
     strXml += " <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "<ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += " <RATE>"+Rate+"/PC</RATE>"; 
     strXml += " <AMOUNT>-"+Rvalue+"</AMOUNT>"; 
     strXml += " <ACTUALQTY> "+Quantity +" PC</ACTUALQTY>"; 
     strXml += " <BILLEDQTY> "+Quantity +" PC</BILLEDQTY>"; 
     strXml += " <BATCHALLOCATIONS.LIST>"; 
     strXml += " <GODOWNNAME>Main Location</GODOWNNAME>"; 
     strXml += "  <BATCHNAME>Primary Batch</BATCHNAME>"; 
     strXml += "  <DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>"; 
     strXml += "  <AMOUNT>-"+Value+"</AMOUNT>"; 
     strXml += "  <ACTUALQTY> "+Quantity +" PC</ACTUALQTY>"; 
     strXml += " <BILLEDQTY> "+Quantity +" PC</BILLEDQTY>"; 
     strXml += "</BATCHALLOCATIONS.LIST>"; 
	                         }
						else {
							t = excelData.length;
						};
					};
     strXml += " <ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += " <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "  <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += " <LEDGERNAME>"+ salesleg +"</LEDGERNAME>"; 
     strXml += " <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "  <AMOUNT>-"+Value+"</AMOUNT>"; 
     strXml += " </ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += " </ALLINVENTORYENTRIES.LIST>      ";     
	 strXml += "</VOUCHER>";  	 
     strXml += "</TALLYMESSAGE>";   
						};	 
                     };   
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};


 function voucherdxb(){
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
						strXml += "<SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>";
						strXml += "  <BANKALLOCATIONS.LIST>";
        strXml += "<DATE>"+ DATE + "</DATE>";
        strXml += "<INSTRUMENTDATE>"+ DATE + "</INSTRUMENTDATE>";
        strXml += "<BANKERSDATE>"+ DATE + "</BANKERSDATE>";
        strXml += "<NAME></NAME>";
        strXml += "<TRANSACTIONTYPE>Cheque/DD</TRANSACTIONTYPE>";
        strXml += "<PAYMENTFAVOURING>" + DRLEDGER + "</PAYMENTFAVOURING>";
        strXml += "<CHEQUECROSSCOMMENT>A/c Payee</CHEQUECROSSCOMMENT>";
        strXml += "<TRANSFERMODE>RTGS</TRANSFERMODE>";
        strXml += "<UNIQUEREFERENCENUMBER></UNIQUEREFERENCENUMBER>";
        strXml += "<STATUS>No</STATUS>";
        strXml += "<PAYMENTMODE>Transacted</PAYMENTMODE>";
        strXml += "<SECONDARYSTATUS/>";
        strXml += "<BANKPARTYNAME>" + DRLEDGER + "</BANKPARTYNAME>";
        strXml += "<ISCONNECTEDPAYMENT>No</ISCONNECTEDPAYMENT>";
        strXml += "<ISSPLIT>No</ISSPLIT>";
		strXml += " <ISCONTRACTUSED>No</ISCONTRACTUSED>";
		strXml += " <ISACCEPTEDWITHWARNING>No</ISACCEPTEDWITHWARNING>";
        strXml += "<ISTRANSFORCED>No</ISTRANSFORCED>";
        strXml += "<CHEQUEPRINTED> 1</CHEQUEPRINTED>";
        strXml += "<AMOUNT>" + AMOUNT2 + "</AMOUNT>";
        strXml += "<CONTRACTDETAILS.LIST>        </CONTRACTDETAILS.LIST>";
        strXml += "<BANKSTATUSINFO.LIST>        </BANKSTATUSINFO.LIST>";
		strXml += "</BANKALLOCATIONS.LIST>";
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

 function salesdxb(){
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


						
    for (var i = 0; i < excelData.length; i++){
						let Particulars   = (excelData[i]["Particulars"]);					
                        let salDate    			= (excelData[i]["Date"]);
                        let VoucherNo   		= (excelData[i]["Voucher No."]);
						let VATDEALERTYPE   	= (excelData[i]["VATDEALERTYPE"]);
						let PLACEOFSUPPLY   	= (excelData[i]["PLACEOFSUPPLY"]);
						let SalesLEDGERNAME   	= (excelData[i]["Sales LEDGERNAME"]);
						let VatLEDGERNAME   	= (excelData[i]["Vat LEDGERNAME"]);
						let SalesAmt   			= (excelData[i]["Sales Amt"]);
						let Vat   				= (excelData[i]["Vat"]);
						let Total   			= (-(excelData[i]["Total"]));
						
                                                 
     strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
	 strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\"Sales\" ACTION=\"Create\" OBJVIEW=\"Invoice Voucher View\">"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "</OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<DATE>" + salDate + "</DATE>"; 
     strXml += "<GUID></GUID>"; 
	 strXml += "<VATDEALERTYPE>" + VATDEALERTYPE + "</VATDEALERTYPE>"; 
	 strXml += "<STATENAME>" + PLACEOFSUPPLY + "</STATENAME>"; 
	 strXml += "<COUNTRYOFRESIDENCE>UAE</COUNTRYOFRESIDENCE>";
     strXml += "<PLACEOFSUPPLY>" + PLACEOFSUPPLY + "</PLACEOFSUPPLY>"; 
     strXml += "<PARTYNAME>" + Particulars + "</PARTYNAME>"; 
     strXml += "<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>"; 
     strXml += "<VOUCHERNUMBER>" + VoucherNo + "</VOUCHERNUMBER>"; 
     strXml += "<PARTYLEDGERNAME>" + Particulars + "</PARTYLEDGERNAME>"; 
     strXml += "<BASICBASEPARTYNAME>" + Particulars + "</BASICBASEPARTYNAME>"; 
     strXml += "<CSTFORMISSUETYPE/>";
     strXml += "<CSTFORMRECVTYPE/>";
     strXml += "<FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>";
     strXml += "<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>";
     strXml += "<BASICBUYERNAME>"+ Particulars + "</BASICBUYERNAME>"; 
     strXml += "<PLACEOFSUPPLYCOUNTRY>UAE</PLACEOFSUPPLYCOUNTRY>";
     strXml += "<EMIRATEPOS>" + PLACEOFSUPPLY + "</EMIRATEPOS>"; 
     strXml += "<PARTYMAILINGNAME>"+ Particulars + "</PARTYMAILINGNAME>"; 
     strXml += "<CONSIGNEEMAILINGNAME>"+ Particulars + "</CONSIGNEEMAILINGNAME>"; 
     strXml += "<CONSIGNEECOUNTRYNAME>UAE</CONSIGNEECOUNTRYNAME>";
     strXml += "<VCHGSTCLASS/>";
     strXml += "<CONSIGNEESTATENAME>" + PLACEOFSUPPLY + "</CONSIGNEESTATENAME>"; 
     strXml += "<VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>";
     strXml += "<DIFFACTUALQTY>No</DIFFACTUALQTY>";
     strXml += "<ISMSTFROMSYNC>No</ISMSTFROMSYNC>";
     strXml += "<ASORIGINAL>No</ASORIGINAL>";
     strXml += "<AUDITED>No</AUDITED>";
     strXml += "<FORJOBCOSTING>No</FORJOBCOSTING>";
     strXml += "<ISOPTIONAL>No</ISOPTIONAL>";
     strXml += "<EFFECTIVEDATE>" + salDate + "</EFFECTIVEDATE>"; 
     strXml += "<USEFOREXCISE>No</USEFOREXCISE>";
     strXml += "<ISFORJOBWORKIN>No</ISFORJOBWORKIN>";
     strXml += "<ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>";
     strXml += "<USEFORINTEREST>No</USEFORINTEREST>";
     strXml += "<USEFORGAINLOSS>No</USEFORGAINLOSS>";
     strXml += "<USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>";
     strXml += "<USEFORCOMPOUND>No</USEFORCOMPOUND>";
     strXml += "<USEFORSERVICETAX>No</USEFORSERVICETAX>";
     strXml += "<ISDELETED>No</ISDELETED>";
     strXml += "<ISONHOLD>No</ISONHOLD>";
     strXml += "<ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>";
     strXml += "<ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>";
     strXml += "<EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>";
     strXml += "<USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>";
     strXml += "<IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>";
     strXml += "<EXCISEOPENING>No</EXCISEOPENING>";
     strXml += "<USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>";
     strXml += "<ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>";
     strXml += "<ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>";
     strXml += "<ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>";
     strXml += "<INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>";
     strXml += "<ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>";
     strXml += "<ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>";
     strXml += "<IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>";
     strXml += "<ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>";
     strXml += "<ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>";
     strXml += "<ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>";
     strXml += "<ISISDVOUCHER>No</ISISDVOUCHER>";
     strXml += "<ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>";
     strXml += "<ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>";
     strXml += "<ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>";
     strXml += "<GSTNOTEXPORTED>No</GSTNOTEXPORTED>";
     strXml += "<IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>";
     strXml += "<ISGSTREFUND>No</ISGSTREFUND>";
     strXml += "<ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>";
     strXml += "<ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>";
     strXml += "<IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>";
     strXml += "<IRNJSONEXPORTED>No</IRNJSONEXPORTED>";
     strXml += "<IRNCANCELLED>No</IRNCANCELLED>";
     strXml += "<ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>";
     strXml += "<ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>";
     strXml += "<ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>";
     strXml += "<ISCANCELLED>No</ISCANCELLED>";
     strXml += "<HASCASHFLOW>No</HASCASHFLOW>";
     strXml += "<ISPOSTDATED>No</ISPOSTDATED>";
     strXml += "<USETRACKINGNUMBER>No</USETRACKINGNUMBER>";
     strXml += "<ISINVOICE>Yes</ISINVOICE>";
     strXml += "<MFGJOURNAL>No</MFGJOURNAL>";
     strXml += "<HASDISCOUNTS>No</HASDISCOUNTS>";
     strXml += "<ASPAYSLIP>No</ASPAYSLIP>";
     strXml += "<ISCOSTCENTRE>No</ISCOSTCENTRE>";
     strXml += "<ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>";
     strXml += "<ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>";
     strXml += "<ISBLANKCHEQUE>No</ISBLANKCHEQUE>";
     strXml += "<ISVOID>No</ISVOID>";
     strXml += "<ORDERLINESTATUS>No</ORDERLINESTATUS>";
     strXml += "<VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>";
     strXml += "<VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>";
     strXml += "<ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>";
     strXml += "<VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>";
     strXml += "<ISVATDUTYPAID>Yes</ISVATDUTYPAID>";
     strXml += "<ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>";
     strXml += "<ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>";
     strXml += "<ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>";
     strXml += "<CHANGEVCHMODE>No</CHANGEVCHMODE>";
     strXml += "<RESETIRNQRCODE>No</RESETIRNQRCODE>";
     strXml += "<ALTERID></ALTERID>";
     strXml += "<MASTERID></MASTERID>";
     strXml += "<VOUCHERKEY></VOUCHERKEY>";
     strXml += "<EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>";
     strXml += "<EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>";
     strXml += "<OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>";
     strXml += "<ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>";
     strXml += "<AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>";
	 strXml += "<DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>";
     strXml += "<ALLINVENTORYENTRIES.LIST>      </ALLINVENTORYENTRIES.LIST>";
     strXml += "<SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>";
     strXml += "<IRNERRORLIST.LIST>      </IRNERRORLIST.LIST>";
     strXml += "<INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>";
     strXml += "<INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>";
     strXml += "<INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>";
     strXml += "<ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>";
     strXml += "<ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>";
     strXml += "<INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>";
     strXml += "<LEDGERENTRIES.LIST>";
     strXml += " <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
     strXml += "  <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
     strXml += " </OLDAUDITENTRYIDS.LIST>";
     strXml += " <LEDGERNAME>"+ Particulars + "</LEDGERNAME>"; 
     strXml += " <GSTCLASS/>";
     strXml += " <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>";
     strXml += " <LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += " <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>";
     strXml += " <ISPARTYLEDGER>Yes</ISPARTYLEDGER>";
     strXml += " <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>";
     strXml += " <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>";
     strXml += " <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>";
     strXml += " <AMOUNT>"+ Total + "</AMOUNT>";
     strXml += " <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>";
     strXml += " <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>";
     strXml += " <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>";
     strXml += " <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>";
     strXml += " <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>";
     strXml += " <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>";
     strXml += " <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>";
     strXml += " <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>";
     strXml += " <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>";
     strXml += " <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>";
     strXml += " <RATEDETAILS.LIST>       </RATEDETAILS.LIST>";
     strXml += " <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>";
     strXml += " <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>";
     strXml += " <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>";
     strXml += " <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>";
     strXml += " <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>";
     strXml += " <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>";
     strXml += " <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>";
     strXml += " <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>";
     strXml += " <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>";
     strXml += " <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>";
     strXml += " <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>";
     strXml += " <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>";
     strXml += "</LEDGERENTRIES.LIST>";
     strXml += "<LEDGERENTRIES.LIST>";
     strXml += " <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
     strXml += "  <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
     strXml += " </OLDAUDITENTRYIDS.LIST>";
     strXml += " <LEDGERNAME>" + SalesLEDGERNAME + "</LEDGERNAME>"; 
     strXml += " <STATNATURENAME>Domestic Taxable Supplies</STATNATURENAME>";
     strXml += " <GSTCLASS/>";
     strXml += " <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>";
     strXml += " <LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += " <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>";
     strXml += " <ISPARTYLEDGER>No</ISPARTYLEDGER>";
     strXml += " <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>";
     strXml += " <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>";
     strXml += " <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>";
     strXml += " <AMOUNT>" + SalesAmt + "</AMOUNT>";
     strXml += " <VATEXPAMOUNT>" + SalesAmt + " </VATEXPAMOUNT>";
     strXml += " <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>";
     strXml += " <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>";
     strXml += " <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>";
     strXml += " <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>";
     strXml += " <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>";
     strXml += " <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>";
     strXml += " <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>";
     strXml += " <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>";
     strXml += " <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>";
     strXml += " <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>";
     strXml += " <RATEDETAILS.LIST>       </RATEDETAILS.LIST>";
     strXml += " <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>";
     strXml += " <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>";
     strXml += " <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>";
     strXml += " <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>";
     strXml += " <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>";
     strXml += " <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>";
     strXml += " <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>";
     strXml += " <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>";
     strXml += " <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>";
     strXml += " <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>";
     strXml += " <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>";
     strXml += " <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>";
     strXml += "</LEDGERENTRIES.LIST>";
     strXml += "<LEDGERENTRIES.LIST>";
     strXml += " <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">";
     strXml += "  <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>";
     strXml += " </OLDAUDITENTRYIDS.LIST>";
     strXml += " <LEDGERNAME>"+ VatLEDGERNAME + "</LEDGERNAME>";
     strXml += " <GSTCLASS/>";
     strXml += " <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>";
     strXml += " <LEDGERFROMITEM>No</LEDGERFROMITEM>";
     strXml += " <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>";
     strXml += " <ISPARTYLEDGER>No</ISPARTYLEDGER>";
     strXml += " <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>";
     strXml += " <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>";
     strXml += " <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>";
     strXml += " <AMOUNT>" + Vat + "</AMOUNT>";
     strXml += " <VATEXPAMOUNT>" + Vat + "</VATEXPAMOUNT>";
     strXml += " <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>";
     strXml += " <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>";
     strXml += " <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>";
     strXml += " <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>";
     strXml += " <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>";
     strXml += " <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>";
     strXml += " <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>";
     strXml += " <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>";
     strXml += " <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>";
     strXml += " <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>";
     strXml += " <RATEDETAILS.LIST>       </RATEDETAILS.LIST>";
     strXml += " <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>";
     strXml += " <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>";
     strXml += " <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>";
     strXml += " <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>";
     strXml += " <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>";
     strXml += " <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>";
     strXml += " <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>";
     strXml += " <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>";
     strXml += " <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>";
     strXml += " <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>";
     strXml += " <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>";
     strXml += " <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>";
     strXml += "</LEDGERENTRIES.LIST>";
     strXml += "<PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>";
     strXml += "<ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>";
     strXml += "<GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>";
     strXml += "<GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>";
     strXml += "<TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>";
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

 function fees(){
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
                       
                        let DATE            = (excelData[i]["DATE"]);
                        let NARRATION       = (excelData[i]["NARRATION"]);
						let STOCKITEMNAME   = (excelData[i]["NAME"]);
                        let VOUCHERTYPE     = "Stock Journal"                   
                        let AMOUNT          = (excelData[i]["AMOUNT"]);
                       
                         
                        strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
                        strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\"" + VOUCHERTYPE +  "\" ACTION=\"Create\">";
                        strXml += "<DATE>"+ DATE + "</DATE>";
                        strXml += "<GUID></GUID>";
                        strXml += "<NARRATION>" + NARRATION + "</NARRATION>";
                        strXml += "<VOUCHERTYPENAME>" + VOUCHERTYPE +" </VOUCHERTYPENAME>";
                        strXml += "<VOUCHERNUMBER></VOUCHERNUMBER>";
                        strXml += "<CSTFORMISSUETYPE/>";
                        strXml += "<CSTFORMRECVTYPE/>";
                        strXml += "<PERSISTEDVIEW>Consumption Voucher View</PERSISTEDVIEW>";
                        strXml += "<VCHGSTCLASS/>";
						strXml += "<VCHENTRYMODE>Use for Stock Journal</VCHENTRYMODE>";
						strXml += " <DIFFACTUALQTY>No</DIFFACTUALQTY>";
						strXml += "<ISMSTFROMSYNC>No</ISMSTFROMSYNC>";
						strXml += "<ASORIGINAL>No</ASORIGINAL>";
						strXml += "<AUDITED>No</AUDITED>";
						strXml += "<FORJOBCOSTING>No</FORJOBCOSTING>";
						strXml += "<EFFECTIVEDATE>"+ DATE + "</EFFECTIVEDATE>";
						strXml += "<ISCANCELLED>No</ISCANCELLED>";
						strXml += "<HASCASHFLOW>No</HASCASHFLOW>";
						strXml += "<ISPOSTDATED>No</ISPOSTDATED>";
 						strXml += "<USETRACKINGNUMBER>No</USETRACKINGNUMBER>";
 						strXml += "<ISINVOICE>No</ISINVOICE>";
 						strXml += "<MFGJOURNAL>No</MFGJOURNAL>";
 						strXml += "<HASDISCOUNTS>No</HASDISCOUNTS>";
 						strXml += "<ASPAYSLIP>No</ASPAYSLIP>";
 						strXml += "<ISCOSTCENTRE>No</ISCOSTCENTRE>";
 						strXml += "<ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>";
 						strXml += "<ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>";
						strXml += "<ISBLANKCHEQUE>No</ISBLANKCHEQUE>";
						strXml += "<ISVOID>No</ISVOID>";
 						strXml += "<ORDERLINESTATUS>No</ORDERLINESTATUS>";
 						strXml += "<VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>";
 						strXml += "<VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>";
 						strXml += "<ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>";
 						strXml += "<VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>";
 						strXml += "<ISVATDUTYPAID>Yes</ISVATDUTYPAID>";
 						strXml += "<ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>";
 						strXml += "<ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>";
						strXml += "<ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>";
 						strXml += "<CHANGEVCHMODE>No</CHANGEVCHMODE>";
 						strXml += "<RESETIRNQRCODE>No</RESETIRNQRCODE>";
						strXml += "<ALTERID> </ALTERID>";
						strXml += "<MASTERID> </MASTERID>";
						strXml += "<VOUCHERKEY></VOUCHERKEY>";
						strXml += "<INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>";
 						strXml += "<INVENTORYENTRIESIN.LIST>";
 						strXml += " <STOCKITEMNAME>"+ STOCKITEMNAME +"</STOCKITEMNAME>";
						strXml += " <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>";
						strXml += " <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>";
						strXml += " <ISAUTONEGATE>No</ISAUTONEGATE>";
 						strXml += " <ISCUSTOMSCLEARANCE>No</ISCUSTOMSCLEARANCE>";
 						strXml += " <ISTRACKCOMPONENT>No</ISTRACKCOMPONENT>";
 						strXml += " <ISTRACKPRODUCTION>No</ISTRACKPRODUCTION>";
 						strXml += " <ISPRIMARYITEM>No</ISPRIMARYITEM>";
 						strXml += " <ISSCRAP>No</ISSCRAP>";
						strXml += " <AMOUNT>-" + AMOUNT +  "</AMOUNT>";
 						strXml += " <BATCHALLOCATIONS.LIST>";
						strXml += "  <GODOWNNAME>Main Location</GODOWNNAME>";
 						strXml += "  <BATCHNAME>Primary Batch</BATCHNAME>";
 						strXml += "  <INDENTNO/>";
						strXml += "  <ORDERNO/>";
 						strXml += "  <TRACKINGNUMBER/>";
 						strXml += "  <DYNAMICCSTISCLEARED>No</DYNAMICCSTISCLEARED>";
 						strXml += "  <AMOUNT>-" + AMOUNT +  "</AMOUNT>";
 						strXml += "  <ADDITIONALDETAILS.LIST>        </ADDITIONALDETAILS.LIST>";
 						strXml += "  <VOUCHERCOMPONENTLIST.LIST>        </VOUCHERCOMPONENTLIST.LIST>";
 						strXml += " </BATCHALLOCATIONS.LIST>";
 						strXml += " <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>";
 						strXml += " <SUPPLEMENTARYDUTYHEADDETAILS.LIST>       </SUPPLEMENTARYDUTYHEADDETAILS.LIST>";
 						strXml += " <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>";
						strXml += " <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>";
						strXml += " <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>";
						strXml += " <EXCISEALLOCATIONS.LIST>       </EXCISEALLOCATIONS.LIST>";
						strXml += " <EXPENSEALLOCATIONS.LIST>       </EXPENSEALLOCATIONS.LIST>";
						strXml += "</INVENTORYENTRIESIN.LIST>";
						strXml += "<INVENTORYENTRIESOUT.LIST>      </INVENTORYENTRIESOUT.LIST>";
						strXml += "<PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>";
						strXml += "<ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>";
						strXml += "<GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>";
 						strXml += "<GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>";
 						strXml += "<TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>";
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

function husk(){
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
						let PartyName 	    = (excelData[i]["Party Name"]);
						let Dated 			= (excelData[i]["DATE"]);
						let Material		= (excelData[i]["Material"]);
						let netwt 			= (excelData[i]["NET WT"]);
						let Narration 		= (excelData[i]["narration"]);
						let Rate 			= (excelData[i]["RATE"]);
						let Amount 			= (excelData[i]["AMOUNT"]);
						let Srno 			= (excelData[i]["Sr no"]);
						let Purchaseac    	= (excelData[i]["Purchase AC"]);              

                            
     strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
	 strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\" Purchase  \" ACTION=\"Create\" OBJVIEW=\"Invoice Voucher View\">"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += " </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<DATE>"+ Dated + "</DATE>"; 
     strXml += " <PARTYNAME>"+ PartyName + "</PARTYNAME>"; 
     strXml += " <PARTYLEDGERNAME>"+ PartyName + "</PARTYLEDGERNAME>"; 
     strXml += " <VOUCHERNUMBER>"+ Srno + "</VOUCHERNUMBER>"; 
     strXml += " <BASICBASEPARTYNAME>"+ PartyName + "</BASICBASEPARTYNAME>"; 
     strXml += " <CSTFORMISSUETYPE/>"; 
     strXml += " <CSTFORMRECVTYPE/>"; 
     strXml += " <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>"; 
     strXml += " <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"; 
     strXml += " <PARTYMAILINGNAME>"+ PartyName + "</PARTYMAILINGNAME>"; 
     strXml += " <VCHGSTCLASS/>"; 
     strXml += " <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>"; 
     strXml += " <DIFFACTUALQTY>No</DIFFACTUALQTY>"; 
     strXml += " <ISMSTFROMSYNC>No</ISMSTFROMSYNC>"; 
     strXml += " <ASORIGINAL>No</ASORIGINAL>"; 
     strXml += " <AUDITED>No</AUDITED>"; 
     strXml += "<FORJOBCOSTING>No</FORJOBCOSTING>"; 
     strXml += " <ISOPTIONAL>No</ISOPTIONAL>"; 
     strXml += " <EFFECTIVEDATE>"+ Dated + "</EFFECTIVEDATE>"; 
     strXml += " <USEFOREXCISE>No</USEFOREXCISE>"; 
     strXml += " <ISFORJOBWORKIN>No</ISFORJOBWORKIN>"; 
     strXml += " <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>"; 
     strXml += " <USEFORINTEREST>No</USEFORINTEREST>"; 
     strXml += " <USEFORGAINLOSS>No</USEFORGAINLOSS>"; 
     strXml += " <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>"; 
     strXml += " <USEFORCOMPOUND>No</USEFORCOMPOUND>"; 
     strXml += " <USEFORSERVICETAX>No</USEFORSERVICETAX>"; 
     strXml += " <ISDELETED>No</ISDELETED>"; 
     strXml += " <ISONHOLD>No</ISONHOLD>"; 
     strXml += " <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>"; 
     strXml += " <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>"; 
     strXml += " <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>"; 
     strXml += " <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>"; 
     strXml += " <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>"; 
     strXml += " <EXCISEOPENING>No</EXCISEOPENING>"; 
     strXml += " <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>"; 
     strXml += " <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>"; 
     strXml += " <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>"; 
     strXml += " <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>"; 
     strXml += " <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>"; 
     strXml += " <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>"; 
     strXml += " <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>"; 
     strXml += " <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>"; 
     strXml += " <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>"; 
     strXml += " <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>";
     strXml += " <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>"; 
     strXml += " <ISISDVOUCHER>No</ISISDVOUCHER>"; 
     strXml += " <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>"; 
     strXml += " <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>"; 
     strXml += " <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>"; 
     strXml += " <GSTNOTEXPORTED>No</GSTNOTEXPORTED>"; 
     strXml += " <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>"; 
     strXml += " <ISGSTREFUND>No</ISGSTREFUND>"; 
     strXml += " <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>"; 
     strXml += " <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>"; 
     strXml += " <IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>";
     strXml += " <IRNJSONEXPORTED>No</IRNJSONEXPORTED>"; 
     strXml += " <IRNCANCELLED>No</IRNCANCELLED>"; 
     strXml += " <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>"; 
	 strXml += "<ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>"; 
     strXml += " <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>"; 
     strXml += " <ISCANCELLED>No</ISCANCELLED>"; 
     strXml += " <HASCASHFLOW>No</HASCASHFLOW>"; 
     strXml += " <ISPOSTDATED>No</ISPOSTDATED>"; 
     strXml += " <USETRACKINGNUMBER>No</USETRACKINGNUMBER>"; 
     strXml += " <ISINVOICE>Yes</ISINVOICE>"; 
     strXml += " <MFGJOURNAL>No</MFGJOURNAL>"; 
     strXml += " <HASDISCOUNTS>No</HASDISCOUNTS>"; 
     strXml += " <ASPAYSLIP>No</ASPAYSLIP>"; 
     strXml += " <ISCOSTCENTRE>No</ISCOSTCENTRE>"; 
     strXml += " <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>"; 
     strXml += " <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>"; 
     strXml += " <ISBLANKCHEQUE>No</ISBLANKCHEQUE>"; 
     strXml += " <ISVOID>No</ISVOID>"; 
     strXml += " <ORDERLINESTATUS>No</ORDERLINESTATUS>"; 
     strXml += " <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>"; 
     strXml += " <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>"; 
     strXml += " <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>"; 
     strXml += " <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>"; 
     strXml += " <ISVATDUTYPAID>Yes</ISVATDUTYPAID>"; 
     strXml += " <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>"; 
     strXml += " <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>"; 
     strXml += " <ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>"; 
     strXml += " <CHANGEVCHMODE>No</CHANGEVCHMODE>"; 
     strXml += " <RESETIRNQRCODE>No</RESETIRNQRCODE>"; 
     strXml += " <ALLINVENTORYENTRIES.LIST>"; 
     strXml += "  <STOCKITEMNAME>"+ Material + "/STOCKITEMNAME>";
     strXml += "  <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISAUTONEGATE>No</ISAUTONEGATE>"; 
     strXml += "  <ISCUSTOMSCLEARANCE>No</ISCUSTOMSCLEARANCE>"; 
     strXml += "  <ISTRACKCOMPONENT>No</ISTRACKCOMPONENT>"; 
     strXml += "  <ISTRACKPRODUCTION>No</ISTRACKPRODUCTION>"; 
     strXml += "  <ISPRIMARYITEM>No</ISPRIMARYITEM>"; 
     strXml += "  <ISSCRAP>No</ISSCRAP>"; 
     strXml += "  <RATE>"+ Rate + "</RATE>"; 
     strXml += "  <AMOUNT>-"+ Amount + "</AMOUNT>"; 
     strXml += "  <ACTUALQTY>"+ netwt + "</ACTUALQTY>"; 
     strXml += "  <BILLEDQTY>"+ netwt + "</BILLEDQTY>"; 
     strXml += "  <ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += "   <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "    <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "   </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "   <LEDGERNAME>"+ Purchaseac + "</LEDGERNAME>"; 
     strXml += "   <GSTCLASS/>"; 
     strXml += "   <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "   <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "   <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "   <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "   <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "   <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "   <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "   <AMOUNT>-"+ Amount + "</AMOUNT>"; 
     strXml += "  </ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += " </ALLINVENTORYENTRIES.LIST>"; 
     strXml += " <LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "  <LEDGERNAME>"+ PartyName + "</LEDGERNAME>"; 
     strXml += "  <GSTCLASS/>"; 
     strXml += "  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>Yes</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "  <AMOUNT>"+ Amount + "</AMOUNT>"; 
     strXml += " </LEDGERENTRIES.LIST>"; 
     strXml += " </VOUCHER>"; 
     strXml += "</TALLYMESSAGE>";   
	 					};	 
                    
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};


function SalesIBW(){
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
						let PartyName 	    = (excelData[i]["Party Name"]);
						let Dated 			= (excelData[i]["DATE"]);
						let Material		= (excelData[i]["Material"]);
						let netwt 			= (excelData[i]["NET WT"]);
						let Narration 		= (excelData[i]["Narration"]);
						let Rate 			= (excelData[i]["RATE"]);
						let Amount 			= (excelData[i]["AMOUNT"]);
						let qttype 			= (excelData[i]["QT TYPE"]);
						let Srno 			= (excelData[i]["Sr no"]);
						let Salesac	    	= (excelData[i]["Sales AC"]);
						let vat		    	= (excelData[i]["VAT AC"]);
						let upkar	    	= (excelData[i]["Upkar AC"]);
						let addtax    		= (excelData[i]["Add Tax AC"]);
						let Roundoffac    	= (excelData[i]["Roundoff AC"]);
						let vatamt	    	= (excelData[i]["VAT AMT"]);
						let upkaramt	   	= (excelData[i]["Upkar AMT"]);
						let addtaxamt	   	= (excelData[i]["Add TAX AMT"]);
						let roundamt	   	= (excelData[i]["Roundoff AMT"]);

                            
     strXml += "<TALLYMESSAGE xmlns:UDF=\"TallyUDF\">";
	 strXml += "<VOUCHER REMOTEID=\"\" VCHKEY=\"\" VCHTYPE=\" Sales \" ACTION=\"Create\" OBJVIEW=\"Invoice Voucher View\">"; 
     strXml += "<OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "<OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += " </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "<DATE>"+ Dated + "</DATE>"; 
     strXml += " <PARTYNAME>"+ PartyName + "</PARTYNAME>"; 
     strXml += " <PARTYLEDGERNAME>"+ PartyName + "</PARTYLEDGERNAME>"; 
     strXml += " <VOUCHERNUMBER>"+ Srno + "</VOUCHERNUMBER>"; 
     strXml += " <BASICBASEPARTYNAME>"+ PartyName + "</BASICBASEPARTYNAME>"; 
     strXml += " <CSTFORMISSUETYPE/>"; 
     strXml += " <CSTFORMRECVTYPE/>"; 
     strXml += " <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>"; 
     strXml += " <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>"; 
     strXml += " <PARTYMAILINGNAME>"+ PartyName + "</PARTYMAILINGNAME>"; 
     strXml += " <VCHGSTCLASS/>"; 
     strXml += " <VCHENTRYMODE>Item Invoice</VCHENTRYMODE>"; 
     strXml += " <DIFFACTUALQTY>No</DIFFACTUALQTY>"; 
     strXml += " <ISMSTFROMSYNC>No</ISMSTFROMSYNC>"; 
     strXml += " <ASORIGINAL>No</ASORIGINAL>"; 
     strXml += " <AUDITED>No</AUDITED>"; 
     strXml += "<FORJOBCOSTING>No</FORJOBCOSTING>"; 
     strXml += " <ISOPTIONAL>No</ISOPTIONAL>"; 
     strXml += " <EFFECTIVEDATE>"+ Dated + "</EFFECTIVEDATE>"; 
     strXml += " <USEFOREXCISE>No</USEFOREXCISE>"; 
     strXml += " <ISFORJOBWORKIN>No</ISFORJOBWORKIN>"; 
     strXml += " <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>"; 
     strXml += " <USEFORINTEREST>No</USEFORINTEREST>"; 
     strXml += " <USEFORGAINLOSS>No</USEFORGAINLOSS>"; 
     strXml += " <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>"; 
     strXml += " <USEFORCOMPOUND>No</USEFORCOMPOUND>"; 
     strXml += " <USEFORSERVICETAX>No</USEFORSERVICETAX>"; 
     strXml += " <ISDELETED>No</ISDELETED>"; 
     strXml += " <ISONHOLD>No</ISONHOLD>"; 
     strXml += " <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>"; 
     strXml += " <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>"; 
     strXml += " <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>"; 
     strXml += " <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>"; 
     strXml += " <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>"; 
     strXml += " <EXCISEOPENING>No</EXCISEOPENING>"; 
     strXml += " <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>"; 
     strXml += " <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>"; 
     strXml += " <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>"; 
     strXml += " <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>"; 
     strXml += " <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>"; 
     strXml += " <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>"; 
     strXml += " <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>"; 
     strXml += " <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>"; 
     strXml += " <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>"; 
     strXml += " <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>";
     strXml += " <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>"; 
     strXml += " <ISISDVOUCHER>No</ISISDVOUCHER>"; 
     strXml += " <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>"; 
     strXml += " <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>"; 
     strXml += " <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>"; 
     strXml += " <GSTNOTEXPORTED>No</GSTNOTEXPORTED>"; 
     strXml += " <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>"; 
     strXml += " <ISGSTREFUND>No</ISGSTREFUND>"; 
     strXml += " <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>"; 
     strXml += " <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>"; 
     strXml += " <IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>";
     strXml += " <IRNJSONEXPORTED>No</IRNJSONEXPORTED>"; 
     strXml += " <IRNCANCELLED>No</IRNCANCELLED>"; 
     strXml += " <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>"; 
	 strXml += "<ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>"; 
     strXml += " <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>"; 
     strXml += " <ISCANCELLED>No</ISCANCELLED>"; 
     strXml += " <HASCASHFLOW>Yes</HASCASHFLOW>"; 
     strXml += " <ISPOSTDATED>No</ISPOSTDATED>"; 
     strXml += " <USETRACKINGNUMBER>No</USETRACKINGNUMBER>"; 
     strXml += " <ISINVOICE>Yes</ISINVOICE>"; 
     strXml += " <MFGJOURNAL>No</MFGJOURNAL>"; 
     strXml += " <HASDISCOUNTS>No</HASDISCOUNTS>"; 
     strXml += " <ASPAYSLIP>No</ASPAYSLIP>"; 
     strXml += " <ISCOSTCENTRE>No</ISCOSTCENTRE>"; 
     strXml += " <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>"; 
     strXml += " <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>"; 
     strXml += " <ISBLANKCHEQUE>No</ISBLANKCHEQUE>"; 
     strXml += " <ISVOID>No</ISVOID>"; 
     strXml += " <ORDERLINESTATUS>No</ORDERLINESTATUS>"; 
     strXml += " <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>"; 
     strXml += " <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>"; 
     strXml += " <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>"; 
     strXml += " <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>"; 
     strXml += " <ISVATDUTYPAID>Yes</ISVATDUTYPAID>"; 
     strXml += " <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>"; 
     strXml += " <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>"; 
     strXml += " <ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>"; 
     strXml += " <CHANGEVCHMODE>No</CHANGEVCHMODE>"; 
     strXml += " <RESETIRNQRCODE>No</RESETIRNQRCODE>"; 
     strXml += " <ALLINVENTORYENTRIES.LIST>"; 
     strXml += "  <STOCKITEMNAME>"+ Material + "/STOCKITEMNAME>";
     strXml += "  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISAUTONEGATE>No</ISAUTONEGATE>"; 
     strXml += "  <ISCUSTOMSCLEARANCE>No</ISCUSTOMSCLEARANCE>"; 
     strXml += "  <ISTRACKCOMPONENT>No</ISTRACKCOMPONENT>"; 
     strXml += "  <ISTRACKPRODUCTION>No</ISTRACKPRODUCTION>"; 
     strXml += "  <ISPRIMARYITEM>No</ISPRIMARYITEM>"; 
     strXml += "  <ISSCRAP>No</ISSCRAP>"; 
     strXml += "  <RATE>"+ Rate + "</RATE>"; 
     strXml += "  <AMOUNT>"+ Amount + "</AMOUNT>"; 
     strXml += "  <ACTUALQTY>"+ netwt + "</ACTUALQTY>"; 
     strXml += "  <BILLEDQTY>"+ netwt + "</BILLEDQTY>"; 
     strXml += "  <ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += "   <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "    <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "   </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "   <LEDGERNAME>"+ Salesac + "</LEDGERNAME>"; 
     strXml += "   <GSTCLASS/>"; 
     strXml += "   <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "   <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "   <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "   <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "   <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "   <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "   <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "   <AMOUNT>"+ Amount + "</AMOUNT>"; 
     strXml += "  </ACCOUNTINGALLOCATIONS.LIST>"; 
     strXml += " </ALLINVENTORYENTRIES.LIST>"; 
     strXml += " <LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "  <LEDGERNAME>"+ vat + "</LEDGERNAME>"; 
     strXml += "  <GSTCLASS/>"; 
     strXml += "  <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>Yes</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "  <AMOUNT>-"+ vatamt + "</AMOUNT>"; 
	 strXml += " <VATEXPAMOUNT> "+ vatamt +" </VATEXPAMOUNT>"; 
     strXml += " </LEDGERENTRIES.LIST>"; 
	 strXml += " <LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "  <LEDGERNAME>"+ upkar + "</LEDGERNAME>"; 
     strXml += "  <GSTCLASS/>"; 
     strXml += "  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "  <AMOUNT>"+ upkaramt + "</AMOUNT>";
	 strXml += " <VATEXPAMOUNT> "+ upkaramt +" </VATEXPAMOUNT>"; 
     strXml += " </LEDGERENTRIES.LIST>"; 
	 strXml += " <LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
     strXml += "  <LEDGERNAME>"+ addtax + "</LEDGERNAME>"; 
     strXml += "  <GSTCLASS/>"; 
     strXml += "  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
     strXml += "  <AMOUNT>"+ addtaxamt + "</AMOUNT>";
	 strXml += " <VATEXPAMOUNT> "+ addtaxamt +" </VATEXPAMOUNT>"; 
     strXml += " </LEDGERENTRIES.LIST>";
	 strXml += " <LEDGERENTRIES.LIST>"; 
     strXml += "  <OLDAUDITENTRYIDS.LIST TYPE=\"Number\">"; 
     strXml += "   <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>"; 
     strXml += "  </OLDAUDITENTRYIDS.LIST>"; 
	 strXml += "  <ROUNDTYPE>Normal Rounding</ROUNDTYPE>"; 
     strXml += "  <LEDGERNAME>"+ Roundoffac + "</LEDGERNAME>"; 
     strXml += "  <GSTCLASS/>"; 
     strXml += "  <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>"; 
     strXml += "  <LEDGERFROMITEM>No</LEDGERFROMITEM>"; 
     strXml += "  <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>"; 
     strXml += "  <ISPARTYLEDGER>No</ISPARTYLEDGER>"; 
     strXml += "  <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>"; 
     strXml += "  <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>"; 
     strXml += "  <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>"; 
	 strXml += "  <ROUNDLIMIT> 1</ROUNDLIMIT>"; 
     strXml += "  <AMOUNT>"+ roundamt + "</AMOUNT>";
	 strXml += " <VATEXPAMOUNT> "+ roundamt +" </VATEXPAMOUNT>"; 
     strXml += " </LEDGERENTRIES.LIST>";
     strXml += " </VOUCHER>"; 
     strXml += "</TALLYMESSAGE>";   
	 					};	 
                    
strXml += "</REQUESTDATA>";
strXml += "</IMPORTDATA>";
strXml += "</BODY>";
strXml += "</ENVELOPE>";
 
    //console.log(strXml);
	return strXml;
};