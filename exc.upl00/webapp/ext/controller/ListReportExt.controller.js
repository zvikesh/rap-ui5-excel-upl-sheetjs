sap.ui.define(
  [
    "sap/ui/core/Fragment",
    "sap/m/MessageToast",
    "zvks/exc/upl00/ext/utils/xlsx.full.min", //More: https://docs.sheetjs.com/docs/getting-started/installation/amd#sap-ui5
  ],
  function (Fragment, MessageToast, xlsx) {
    //Keep, xlsx as small. Making it Caps caused me lot of trouble.
    "use strict";

    return {
      //Initialization
      excelSheetsData: [], //Excel data
      pDialog: null, //Pop up dialog

      onExcelUploadButton: function (oEvent) {
        //VKS? This view is not used anywhere
        var oView = this.getView();

        if (!this.pDialog) {
          Fragment.load({
            id: "idExcelUploadFragment", //Assign the ID
            name: "zvks.exc.upl00.ext.fragment.ExcelUpload", //Path to instantiate the object of Fragment control
            type: "XML", //Default
            controller: this,
          })
            .then((oDialog) => {
              /*--- Async ---*/
              var oFileUploaderSet = Fragment.byId(
                "idExcelUploadFragment",
                "idFileUploaderSet"
              );
              //oFileUploaderSet.removeAllItems();

              this.pDialog = oDialog;
              this.pDialog.open();
            })
            .catch((error) => alert(error.message));
        } else {
          //Pop Up dialog already instantiated
          var oFileUploaderSet = Fragment.byId(
            "idExcelUploadFragment",
            "idFileUploaderSet"
          );
          oFileUploaderSet.removeAllItems(); //Clear the existing items
          this.pDialog.open();
        }
      },

      onBeforeUploadStarts: function (oEvent) {
        console.log("File Before Upload Event Fired!!!");
        /* TODO: check for file upload count */
        var oFileUploaderSet = Fragment.byId(
          "idExcelUploadFragment",
          "idFileUploaderSet"
        );
        oFileUploaderSet.removeAllItems(); //Clear the existing items, since once file can be uploaded at at time
        this.excelSheetsData = []; //Clear the runtime data, if any
      },

      onUploadCompleted: function (oEvent) {
        console.log("File Uploaded!!!");

        var that = this;

        /*--- SheetJS > Common Use Cases > Data Import ---*/

        var reader = new FileReader();

        reader.onload = (oEvent) => {
          //Getting the binary excel file content

          let oExcelContent = oEvent.currentTarget.result;
          //Let oExcelData = oEvent.Target.result;         //Docs

          let oWorkbook = XLSX.read(oExcelContent, { type: "binary" });

          //Here reading only the excel file sheet- Sheet1
          var oWorksheet = oWorkbook.Sheets[oWorkbook.SheetNames[0]]; //Prefered
          //var oWorksheet = oWorkbook.Sheets["Sheet1"];

          //var oExcelData = XLSX.utils.sheet_to_row_object_array(oWorksheet);

          var oExcelData = XLSX.utils.sheet_to_json(oWorksheet); //Pass parameter {header: 1} if first row is data row

          oWorkbook.SheetNames.forEach(function (sheetName) {
            // appending the excel file data to the global variable
            that.excelSheetsData.push(
              XLSX.utils.sheet_to_json(oWorkbook.Sheets[sheetName])
            );
          });
          console.log("Excel Data", oExcelData);
          console.log("Excel Sheets Data", this.excelSheetsData);
        };

        /* TODO: Read excel file data*/
        //Getting the UploadSet Control reference
        var oFileUploaderSet = Fragment.byId(
          "idExcelUploadFragment",
          "idFileUploaderSet"
        );

        var oFileUploaderItem = oFileUploaderSet.getItems()[0]; //Since, we will be uploading only 1 file. So, reading the first item.

        oFileUploaderItem.setVisibleEdit(false); //visibleEdit="false" property is not working

        var oFileBlob = oFileUploaderItem.getFileObject(); //File object
        reader.readAsArrayBuffer(oFileBlob);

        MessageToast.show("Upload Successful");
      },

      onAfterItemRemoved: function (oEvent) {
        console.log("File Remove/delete Event Fired!!!");
        /* TODO: Clear the already read excel file data */
        this.excelSheetsData = []; //Clear the runtime data, if any
      },

      onDownloadTemplate: function (oEvent) {
        console.log("Template Download Button Clicked!!!");

        /* TODO: Excel file template download */

        //Get the oData model binded to this application
        var oModel = this.getView().getModel();

        console.log(
          oModel.getServiceMetadata().dataServices.schema[0].entityType
        );

        //Get the property list of the entity for which we need to download the template
        var oBuilding = oModel
          .getServiceMetadata()
          .dataServices.schema[0].entityType.find(
            (x) => x.name === "TableUploadType"
          );

        //Set the list of entity property, that has to be present in excel file template
        var propertyList = [
          "Materiala",
          "Materialb",
          "Movementype",
          "Vendor",
          "Poref",
          "Plant",
          "Storageloc",
          "Trpoststatus",
        ];

        var colList = {};

        //Finding the property description corresponding to the Property ID
        propertyList.forEach((value, index) => {
          let property = oBuilding.property.find((x) => x.name === value);
          colList[property.extensions.find((x) => x.name === "label").value] =
            "";
        });

        var excelColumnList = [];
        excelColumnList.push(colList); //

        const oWorkSheet = XLSX.utils.json_to_sheet(excelColumnList); //Initializing the excel work sheet
        const oWorkBook = XLSX.utils.book_new(); //Creating the new excel work book
        XLSX.utils.book_append_sheet(oWorkBook, oWorkSheet, "Sheet1"); //Set the file value
        XLSX.writeFile(oWorkBook, "Template.xlsx"); //Download the created excel file

        MessageToast.show("Downloading the Template...");
      },

      onUploadSet: function (oEvent) {
        console.log("Upload Button Clicked!!!");

        /* TODO:Call to OData */

        //Checking if excel file contains data or not
        if (!this.excelSheetsData.length) {
          MessageToast.show("Select file to Upload");
          return;
        }

        var that = this;
        var oSource = oEvent.getSource();

        //Creating a promise as the extension api accepts odata call in form of promise only
        var fnAddMessage = function () {
          return new Promise((fnResolve, fnReject) => {
            that.callOdata(fnResolve, fnReject);
            /*
            that.callOdata()
            .then((fnResolve)=>{console.log("Create Success")})
            .catch((fnReject)=>{console.log(fnReject.getText())});
          */
          });
        };

        var mParameters = {
          sActionLabel: oSource.getText(), // or "Your custom text"
        };

        //Calling the oData service using extension API
        this.extensionAPI.securedExecution(fnAddMessage, mParameters);
        this.pDialog.close();
      },

      onCloseDialog: function (oEvent) {
        this.pDialog.close();
      },

      //Helper method to call OData
      callOdata: function (fnResolve, fnReject) {
        
        console.log(">>>>>>>>>>>>>>>>>>>>>");

        //Intializing the message manager for displaying the odata response messages
        var oModel = this.getView().getModel();

        //Creating odata payload object for Building entity
        var payload = {};

        this.excelSheetsData[0].forEach((value, index) => {
          //Setting the payload data
          /*
          payload = {
            Materiala: value["Material A"].toString(),
            Materialb: value["Material B"].toString(),
            Movementype: value["Movement Type"].toString(),
            Vendor: value["Supplier/Vendor"].toString(),
            Poref: value["PO reference"].toString(),
            Plant: value["Plant"].toString(),
            Storageloc: value["Storage Location"].toString(),
            Trpoststatus: value["Transfer Post Status"].toString(),
          };
          */

          //Setting the payload data
          payload = { 
            Materiala: value["Materiala"].toString(),
            Materialb: value["Materialb"].toString(),
            Movementype: value["Movementype"].toString(),
            Vendor: value["Vendor"].toString(),
            Poref: value["Poref"].toString(),
            Plant: value["Plant"].toString(),
            Storageloc: value["Storageloc"].toString(),
            Trpoststatus: value["Trpoststatus"].toString(),
          };

          console.log("This is payload");
          console.log(payload);

          //Setting excel file row number for identifying the exact row in case of error or success
          payload.ExcelRowNumber = index + 1;

          //Calling the odata service
          oModel.create("/TableUpload", payload, {
            success: (result) => {
              
              console.log(">>>>>>>>>>>>>>>>>");
              console.log(result);

              var oMessageManager = sap.ui.getCore().getMessageManager();

              var oMessage = new sap.ui.core.message.Message({
                message: "Inspectation Created with ID: ",
                persistent: true, // create message as transition message
                type: sap.ui.core.MessageType.Success,
              });

              oMessageManager.addMessages(oMessage);
              fnResolve();
            },

            error: fnReject,
          });
        });
      },
    };
  }
);
