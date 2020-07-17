/**
 * Mails Summary Table to all the team members.
 */

/**
 * Returns an array of objects containing milestone data
 * Fetches data of each milestone from the Summary Sheet.
 * @return {object[]} An array of objects containing milestone data
 */
function getMilestoneData() {
  var countColumns = sendMilestoneCount();
  // Milestone Data stored from Row 6, Column C onwards.
  var firstEngineerRow = 6;
  var firstMilestoneColumn = 2;
  var values =
      SpreadsheetApp.getActive()
          .getSheetByName("Summary")
          .getRange(firstEngineerRow, firstMilestoneColumn, 10, countColumns)
          .getValues();
  var milestones = [];
  for (var j = 0; j < countColumns; j++) {
    var milestone = {};
    // Milestone Number
    milestone.number = j + 1;
    // Desired Launch Date
    milestone.desiredLaunchDate = getFormattedDate(values[0][j]);
    // Estimated Work Days
    milestone.estimatedWorkDays = values[1][j];
    // Work Days Completed
    milestone.workDaysCompleted = values[2][j];
    // Remaining Work Days
    milestone.remainingWorkDays = values[3][j];
    // Start Date
    milestone.startDate = getFormattedDate(values[4][j]);
    // Estimated Launch Date
    milestone.estimatedLaunchDate = getFormattedDate(values[5][j]);
    // Tasks with no assigned Days
    milestone.tasksWithNoAssignedDays = values[6][j];
    // Days ahead / behind of schedule (+/-)
    milestone.daysSchedule = values[7][j];
    // Remaining Weeks in Iteration
    milestone.remainingWeeksInIteration = values[8][j];
    // Number of engineers assigned to the milestone
    milestone.engineerCount = values[9][j];
    milestones.push(milestone);
  }
  return milestones;
}

/**
 * Returns string of the email addresses at which
 * the email is to be sent separated by ","
 * @return {String} String of email addresses
 */
function getMailAddress() {
  var countTeam = teamSize();
  var milestoneNumber = sendMilestoneCount();
  // Team Data stored from Row 6 onwards
  var firstEngineerRow = 6;
  var contactColumn = 10 + milestoneNumber;
  var teamMail = SpreadsheetApp.getActive()
                     .getSheetByName('Team')
                     .getRange(firstEngineerRow, contactColumn, countTeam)
                     .getValues();
  var emailTo = '';
  var countMail = 0;
  for (var i = 0; i < teamMail.length; i++) {
    if (teamMail[i] == '') {
      continue;
    }
    if (countMail) {
      emailTo += ",";
    }
    emailTo += teamMail[i];
    countMail++;
  }
  return emailTo;
}

/**
 * Returns the title of the project
 * @return {String} The title of the project
 */
function getProjectTitle() {
  var projectTitle = SpreadsheetApp.getActive()
                         .getSheetByName('Tasks')
                         .getRange('B1')
                         .getValue();
  return projectTitle;
}

/**
 * Returns the HTML body of the email
 * @param {object[]} Array of objects containing milestone data
 * @return {String} The HTML body of the email
 */
function getEmailHtml(milestoneData) {
  var htmlTemplate = HtmlService.createTemplateFromFile("mail_template");
  htmlTemplate.milestones = milestoneData;
  var htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}

/**
 * Sends an email message containing summary table to all the team members.
 * Method sendEmail(recipient, subject, body, options) of Class MailApp is used.
 */
function sendEmail() {
  var milestoneData = getMilestoneData();
  var htmlBody = getEmailHtml(milestoneData);
  var recipients = getMailAddress();
  var projectTitle = getProjectTitle();
  MailApp.sendEmail({
    to : recipients,
    subject : 'Weekly Updates (' + projectTitle + ')',
    htmlBody : htmlBody
  });
}