function onOpen(e) {
  var copySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3")
  var rangeApprove = copySheet.getRange('A2:B2')
  var rangeUpdate = copySheet.getRange('A5')
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1") 
      
  for (var r=4; sheet.getRange(r,2).getValue() != "" || sheet.getRange(r,3).getValue() != "" ; r+=1) {
    if (sheet.getRange(r,2).getValue() == "") { //check for new program 
      // Extract name from email
      var userMail = sheet.getRange(r,3).getValue()
      var userName = userMail.substring(0,userMail.indexOf(".")) + " " + userMail.substring(userMail.indexOf(".")+1,userMail.indexOf("@"))
      sheet.getRange(r,2).setValue("=PROPER(\""+userName+"\")")
      
      //Select process
      var marker = sheet.getRange(r,7).getValue()
      var bioassay = sheet.getRange(r,8).getValue()
      var donorNum = sheet.getRange(r,10).getValue()
      var traitType = sheet.getRange(r,9).getValue()
      
      if (marker == "YES") {
        sheet.getRange(r,8).setValue("YES") //if "marker" is available, also bioassay.
        sheet.getRange(r,9).setValue("Dominant") //if "marker" is available, genomic trait is dominant.
        if (donorNum == "1 Donor") {
          sheet.getRange(r,22).setValue("A")
          sheet.getRange(r,26).setBackground('#bdbdbd')
          sheet.getRange(r,28).setBackground('#bdbdbd')
          sheet.getRange(r,30).setBackground('#bdbdbd')
          sheet.getRange(r,32).setBackground('#bdbdbd')
        }
        else {
          sheet.getRange(r,22).setValue("C")
        }
      }
      else {
        if (bioassay == "YES") {
          sheet.getRange(r,10).setValue("1 Donor")   //if have no marker but have bioassay, donor parent will be accept only 1
          if (traitType == "Dominant") {
            sheet.getRange(r,22).setValue("B2")
            sheet.getRange(r,26).setBackground('#bdbdbd')
            sheet.getRange(r,28).setBackground('#bdbdbd')
            sheet.getRange(r,30).setBackground('#bdbdbd')
            sheet.getRange(r,32).setBackground('#bdbdbd')
          }
          else if (traitType == "Recessive") {
            sheet.getRange(r,22).setValue("B1")
          }          
        }
      }
      if (sheet.getRange(r,22).getValue() == "C" & sheet.getRange(r,22).getValue() != "") {
        // Create drop down for approve and update
        var spreadsheet = SpreadsheetApp.getActive();
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet3'), true);
        spreadsheet.getRange('A2:B2').activate();
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Form Responses 1'), true);
        sheet.getRange(r,17).activate();
        spreadsheet.getRange('Sheet3!A2:B2').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet3'), true);
        spreadsheet.getRange('A5').activate();
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Form Responses 1'), true);
        sheet.getRange(r,23).activate();
        spreadsheet.getRange('Sheet3!A5').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        sheet.getRange(r,24).activate();
        spreadsheet.getRange('Sheet3!A5').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      }
      sheet.getRange(r,2,r,40).activate()
      sheet.getActiveRangeList().setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      sheet.getRange(r,1).activate()
    }
  }
}


function submitUpdate() {
  // Get the sheet where the data is, in sheet 'system' 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1") 
  
  var rng = sheet.getRange(4,1,200,34)
  var data = rng.getValues();
 
  for (var i=0; i<data.length; i+=1) {
    var col = data[i]
    
    // Send e-mail to K.Darush for approving
    var manager2 = "amika.yawichai@eastwestseed.com"
    var manager1 = "ekkachai.inchonnabot@eastwestseed.com"
    var approve = col[16]
    if (approve == "Approved") {
      var rangeCell = sheet.getRange(i+4,17)
      rangeCell.setDataValidation(null)
      data[i][16] = "Approved, " + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
      //rangeCell.setBackground('#d9ead3')
      rng.setValues(data)
      
      var emailAddress = manager2
      var subject = "New TI project need to be approved"
      MailApp.sendEmail({
        to: emailAddress,
        cc:manager1,
        subject: subject,
        htmlBody: "The request has been created by " +data[i][1]+" on "+ Utilities.formatDate(data[i][0],"GMT+1", "dd/MM/yyyy")+'<br /><br />'
        + "<b>Crop: </b>" + col[4] + '<br />'
        + "<b>Trait: </b>" + col[5]+'<br />'
        + "<b>Genomic trait: </b>" + col[8]+'<br />'
        + "<b>Marker available: </b>" + col[6]+'<br />'
        + "<b>Bioassay available: </b>" + col[7]+'<br />'
        + "<b>Donor parents: </b>" + col[10]+'<br />'
        + "<b>Donor seed quantities: </b>" + col[11]+'<br />'
        + "<b>Recurrent parents: </b>" + col[12]+'<br />'
        + "<b>Recurrent seed quantities: </b>" + col[13]+'<br />' 
        + "<b>Remark: </b>" + col[15]+'<br /><br />'
        + "Please find more details and approve this project in the "
        +"<a href=\"https://docs.google.com/spreadsheets/d/1pN8UMD-brOwOCp0gpNtvtVewaNT2rXjbdDDj_ck-T2c/edit#gid=161182431&range=Q15\">link</a>"+'<br /><br />'
        +"Regrads, <br />"
        +"<b>TRAIT INTROGRESSION GROUP</b>"
      })
    }

    // Send e-mail to user when project was approved by K. Darush
    var emailUser = col[2]
    var approve2 = col[17]
    if (approve2 == "Approved") {
      var rangeCell = sheet.getRange(i+4,18)
      rangeCell.setDataValidation(null)
      data[i][17] = "Approved, " + Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
      //rangeCell.setBackground('#d9ead3')
      rng.setValues(data)
      
      var emailAddress = emailUser
      var subject = "Your TI project has been approved"
      MailApp.sendEmail({
        to: emailAddress,
        cc:manager1,
        subject: subject,
        htmlBody: "Please find the details of your project below or in the "
        +"<a href=\"https://docs.google.com/spreadsheets/d/1vJukcajnN5yWhajUQe1XbMPiWUc1PK4RHcNOr_L67pI/edit?usp=sharing\">link</a>."+'<br /><br />'
        + "<b>Crop: </b>" + col[4] + '<br />'
        + "<b>Trait: </b>" + col[5]+'<br />'
        + "<b>Genomic trait: </b>" + col[8]+'<br />'
        + "<b>Marker available: </b>" + col[6]+'<br />'
        + "<b>Bioassay available: </b>" + col[7]+'<br />'
        + "<b>Donor parents: </b>" + col[10]+'<br />'
        + "<b>Donor seed quantities: </b>" + col[11]+'<br />'
        + "<b>Recurrent parents: </b>" + col[12]+'<br />'
        + "<b>Recurrent seed quantities: </b>" + col[13]+'<br />' 
        + "<b>Remark: </b>" + col[15]+'<br /><br />'
        + "The project will be started as soon as posible and we will inform you immediately.<br /><br />"
        +"Regrads, <br />"
        +"<b>TRAIT INTROGRESSION GROUP</b>"
      })
    }
    
    // Send e-mail to user for progressive notification
    for (var j = 33 ; j>=22 ; j-=1) {
      var subject = "Your TI project has beed updated!!"
      var result = col[j]
      if (j== 23 || j==22) {
        if (col[j] == "Done") {
          var rangeCell = sheet.getRange(i+4,j+1)
          rangeCell.setDataValidation(null)
          rangeCell.setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"))
          if (j == 22) {  //start project date is the same as P1/P2 submit date
            sheet.getRange(i+4,20).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"))
          }
          
           MailApp.sendEmail({
             to: emailUser,
             cc:manager1,
             subject: subject,
             htmlBody: "Please find the details of your project below or in the "
             +"<a href=\"https://docs.google.com/spreadsheets/d/1vJukcajnN5yWhajUQe1XbMPiWUc1PK4RHcNOr_L67pI/edit?usp=sharing\">link</a>."+'<br /><br />'
             + "<b>Crop: </b>" + col[4] + '<br />'
             + "<b>Trait: </b>" + col[5]+'<br />'
             + "<b>Genomic trait: </b>" + col[8]+'<br />'
             + "<b>Process status:</b> " + sheet.getRange(3,j+1).getValue()+'<br /><br />'
             +"Regrads, <br />"
             +"<b>TRAIT INTROGRESSION GROUP</b>"
           })
        }
      }
      else if (result != "" & result.indexOf(",") == -1) { //Check wheather a result was already updated  
        var rangeCell = sheet.getRange(i+4,j+1)
        var resultAll = rangeCell.getFormula()
        var resultLink = resultAll.substring(resultAll.indexOf("(")+2,resultAll.indexOf(",")-1)
        var result = resultAll.substring(resultAll.indexOf(",")+2,resultAll.indexOf("%")+1)
        rangeCell.setValue("=HYPERLINK(\""+resultLink+"\", \""+result+", "+Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")+"\")")
        rangeCell.activate()
        sheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        //rangeCell.setDataValidation(null)
        //data[i][j] = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
        //rng.setValues(data)
        
        var subject = "Your TI project has beed updated!!"
        MailApp.sendEmail({
          to: emailUser,
          cc:manager1,
          subject: subject,
          htmlBody: "Please find the details of your project below or in the "
          +"<a href=\"https://docs.google.com/spreadsheets/d/1vJukcajnN5yWhajUQe1XbMPiWUc1PK4RHcNOr_L67pI/edit?usp=sharing\">link</a>."+'<br /><br />'
          + "<b>Crop: </b>" + col[4] + '<br />'
          + "<b>Trait: </b>" + col[5]+'<br />'
          + "<b>Genomic trait: </b>" + col[8]+'<br />'
          + "<b>Process status:</b> " +sheet.getRange(3,j+1).getValue()+" = "
          + "<a href="+resultLink+">"+result+"</a>"+'<br /><br />'
          +"Regrads, <br />"
          +"<b>TRAIT INTROGRESSION GROUP</b>"
        })
      }
    }
      
  }  
  
}
