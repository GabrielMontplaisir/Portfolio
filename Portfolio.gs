// ID for template slide. TODO: Make it selectable by user.
var templateSlideID = "1NvIum_IB-2wUZSCeVFlcOTj5xV4e7cysBcnOifkhZj4"

// The column name to identify students. This should be a unique identifier, such as an email.
// var email = PropertiesService.getScriptProperties().getProperty("IDCol");

// The column name for the student grades.
var comment = PropertiesService.getScriptProperties().getProperty('commentCol');

function exportPortfolio() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();


// Get values from Spreadsheet
  var data = sh.getDataRange().getValues();
  // var emailIndex = data[0].indexOf(email); // Get Email column, selected by the user.
  var commentIndex = data[0].indexOf(comment);  // Get Comment column, selected by user.

  // Try to get the Responses to a linked form. If form is not linked, grab the responses from the sheet instead, between the email and the comment column.
  try {
    var formResponses = FormApp.openByUrl(sh.getFormUrl()).getResponses();
    // Logger.log("Found Form")
  } catch (e) {
    // Logger.log("Did not find form")
    // var formResponses = sh.getRange(1,emailIndex+2, sh.getLastRow(), commentIndex-2).getValues();
    return "Could not find a linked Google Form. Please link a form first, then try again."
  }

  // Check if all responses have an associated email address to Document Properties.
  var docPropsKeys = PropertiesService.getDocumentProperties().getKeys();
  for (var e in formResponses) {
    if (docPropsKeys.includes(formResponses[e].getRespondentEmail())) {
      Logger.log("found key")
      continue
    }
    PropertiesService.getDocumentProperties().setProperty(formResponses[e].getRespondentEmail(), "")
  }

  // Place all values in an easily retrievable Array that will be passed to sortComments(). If there's an error, it's likely because the email and comment columns were not set properly.
  try {
    var studentComments = [
      // sh.getRange(2,emailIndex+1,sh.getLastRow()-1).getValues(),
      formResponses,
      sh.getRange(2,commentIndex+1,sh.getLastRow()-1).getValues()
    ];
    // Logger.log(studentComments)
  } catch (e) {
    // return "Identifier doesn't match or can't be found in portfolio. Select identifiers through the hamburger icon in the top right, or make sure they're correct in the portfolio tab."
    return "Comment Column can't be found in portfolio. Make sure the Comment Column exists, and select it through the hamburger icon in the top right."
  }

  // Start sorting comments
  sortComments(docPropsKeys, studentComments);

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
  var name = AdminDirectory.Users.get(student, {viewType:'domain_public', fields:'name'});
  var fullName = name.name.fullName;
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
function sortComments(docPropsKeys, studentComments) {
  // Logger.log(docPropsKeys)
  // Logger.log(studentComments)
  // For all the students in the Portfolio Tab...
  for (var l in docPropsKeys) {
    Logger.log(docPropsKeys[l]);

    // For all the students who filled a response...
    for (var s in studentComments[0]){
      var email = studentComments[0][s].getRespondentEmail();
      // Logger.log(email)
      if (email == docPropsKeys[l]) {

        // Get the Portfolio URL
        var portfolioURL = PropertiesService.getDocumentProperties().getProperty(docPropsKeys[l]);
        // Logger.log(docPropsKeys[l]+' - '+studentComments[0][s]+' - '+studentComments[1][s]);

        try {
          var formResponse = studentComments[0][s].getItemResponses();
          // Logger.log(formResponse)
        } catch (e) {
          Logger.log("Did not find Form responses.")
          // formResponse = studentComments[1][s];
        }
      
        try {
          // Logger.log(docPropsKeys[l])
          var studentPortfolio = SlidesApp.openByUrl(portfolioURL);
        } catch (e) {
          // Logger.log("Cannot find student Portfolio. Creating new one.");
          var studentPortfolio = SlidesApp.openByUrl(createStudentPortfolio(docPropsKeys[l]));
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
            Logger.log("No form Response")
            // currentSlide.replaceAllText("{{Response "+r+"}}", "Question: "+studentComments[0][0][r]+"\n"+studentComments[0][s+1][r]);
          }
        }
        currentSlide.replaceAllText("{{Comment}}", studentComments[1][s]);

        // Logger.log(studentComments);
        break
      };
    };
  };
}