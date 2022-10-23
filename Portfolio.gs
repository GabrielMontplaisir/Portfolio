// ID for template slide. TODO: Make it selectable by user.
var templateSlideID = "1NvIum_IB-2wUZSCeVFlcOTj5xV4e7cysBcnOifkhZj4"

// The column name for the student grades.
var comment = PropertiesService.getScriptProperties().getProperty('commentCol');
// Logger.log(comment)

function exportPortfolio() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();

  // Try to get the Responses to a linked form. If form is not linked, grab the responses from the sheet instead, between the email and the comment column.
  try {
    var formResponses = FormApp.openByUrl(sh.getFormUrl()).getResponses();
    // Logger.log("Found Form")
  } catch (e) {
    Logger.log("Did not find form")
    return "Could not find a linked Google Form. Please link a form first, then try again."
  }

  // Check if all responses have an associated email address to Document Properties.
  var docPropsKeys = PropertiesService.getDocumentProperties().getKeys();
  // Logger.log(docPropsKeys)

  // Add email to Doc Props
  for (var e in formResponses) {
    if (docPropsKeys.includes(formResponses[e].getRespondentEmail())) {
      // Logger.log("found key")
      continue
    }
    PropertiesService.getDocumentProperties().setProperty(formResponses[e].getRespondentEmail(), "");
  }

  //Refresh Document Properties.
  docPropsKeys = PropertiesService.getDocumentProperties().getKeys();
  //Logger.log(docPropsKeys)

  if (sh.getRange(1,1).getValue() != 'Exported') {
    sh.insertColumnBefore(1).setColumnWidth(1,60);
    sh.getRange(1,1).setValue('Exported');
  }

  // Get values from Spreadsheet
  var data = sh.getDataRange().getValues();
  // Logger.log(data)

  // Start sorting comments
  sortComments(docPropsKeys, formResponses, data);

  // If everything completed successfully, return this.
  return "Exported successfully to portfolios."
}


/* Function to create a portfolio for the student.
* Grab the student ID, name, and their row in the Portfolio tab.
* In theory, the row # is dependant on whether their ID matches the ID from the form. You could switch their spot, and it wouldn't affect the script in a negative way.
* If the script can't get the URL for the document (because it's in the trash or otherwise), the script will create a new portfolio for the person.
*/
function createStudentPortfolio(student){
  // Get Student's name. Requires the Admin SDK API.
  try {
    var name = AdminDirectory.Users.get(student, {viewType:'domain_public', fields:'name'});
    var fullName = name.name.fullName;
  } catch(err) {
    var fullName = student.substring(0, student.lastIndexOf("@"));;
  }
  // Logger.log(fullName);

  // Create folders for the User to put all the portfolios in one place. Calls on the createPortfolioFolder() function.
  var classFolderID = DriveApp.getFolderById(createPortfolioFolder());

  // Create a new Portfolio with the student and gives them editor access. Move it to the user's portfolio folder. **TODO: Make it selectable by the user?**
  var newPortfolio = SlidesApp
    .create(fullName+' Portfolio')
    .addEditor(student);
  DriveApp.getFileById(newPortfolio.getId()).moveTo(classFolderID);

  // Set Document Properties to eventually replace the Portfolio tab.
  PropertiesService.getDocumentProperties().setProperty(student, newPortfolio.getUrl())

  // When you create a new Slide, the first slide will be the default. These next commands are to remove the default first slide, replace it with the first slide from the template, then name the Portfolio.
  newPortfolio.getSlides()[0].remove();
  newPortfolio.appendSlide(SlidesApp.openById(templateSlideID).getSlides()[0]);
  newPortfolio.getSlides()[0].replaceAllText("{{Portfolio Name}}", fullName+' Portfolio');

  // Return the Portfolio URL if everything worked.
  return newPortfolio.getUrl();
}

/* 
* The main function to the whole operation. This is where the magic happens. I will do my best to break everything down.
*/
function sortComments(docPropsKeys, formResponses, data) {
  var sh = SpreadsheetApp.getActive().getActiveSheet();
  // Logger.log(docPropsKeys);
  // Logger.log(formResponses);
  // Logger.log(data);
  try {
    var lowercased = data[0].map(name => name.toLowerCase());
    var commentIndex = lowercased.indexOf(comment);  // Get Comment column, selected by user.
  } catch (err) {
    return "Cannot find Comments column. Please select the Comments column by using the hamburger icon on the top right."
  }
  // Logger.log(commentIndex)

  // For all the students who filled a response...
  for (var l in docPropsKeys) {
    // Logger.log('Current docProp: '+docPropsKeys[l])

    // Get the Portfolio URL
    var portfolioURL = PropertiesService.getDocumentProperties().getProperty(docPropsKeys[l]);

    try {
      var studentPortfolio = SlidesApp.openByUrl(portfolioURL);
    } catch (e) {
      Logger.log("Cannot find Portfolio for "+docPropsKeys[l]+". Creating new one.");
      var studentPortfolio = SlidesApp.openByUrl(createStudentPortfolio(docPropsKeys[l]));
    }

    try {
      var row = sh.createTextFinder(docPropsKeys[l]).matchEntireCell(true).findNext().getRow();
    } catch (e) {
      Logger.log("Cannot find student in list. Moving on.")
      continue
    }

    if (!sh.getRange(row,1).isChecked()) {

      // For all the students who answered the form...
      for (var s in formResponses){
        var email = formResponses[s].getRespondentEmail();
        // Logger.log(email)

        // If the current student is the one associated to the current Doc Prop...
        if (email == docPropsKeys[l]) {

          // Find Comment
          var studentComment = data.find((r) => {
            return r.includes(email)
          });

          // Logger.log(email+' - '+docPropsKeys[l]+' - '+studentComment[commentIndex]);

          try {
            var formResponse = formResponses[s].getItemResponses();
            // Logger.log(formResponse)
          } catch (e) {
            Logger.log("Did not find Form responses.")
          }

          var currentSlide = studentPortfolio.appendSlide(SlidesApp.openById(templateSlideID).getSlides()[1]);
          currentSlide.replaceAllText("{{Title}}", "Commentaire de la réponse "+SpreadsheetApp.getActiveSheet().getName());
          var responsePlaceholderArray = [];
          for (var r in formResponse) {
            responsePlaceholderArray.push("{{Response "+r+"}}\n");
          }
          currentSlide.replaceAllText("{{Response}}", responsePlaceholderArray.join(''))
          for (var r in responsePlaceholderArray) {
            try {
              currentSlide.replaceAllText("{{Response "+r+"}}", "Question: "+formResponse[r].getItem().getTitle()+"\n"+formResponse[r].getResponse());
            } catch (e) {
              Logger.log("No Form Response - Cannot replace questions/answers")
            }
          }

          try {
            currentSlide.replaceAllText("{{Comment}}", studentComment[commentIndex]);
            sh.getRange(row,1).insertCheckboxes();
            sh.getRange(row,1).check();
          } catch (err) {
            Logger.log("No comment - "+email+" did not fill out form")
          }
          break
        };
      };
    };
  };
}