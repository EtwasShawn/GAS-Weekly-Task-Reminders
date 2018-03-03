function onOpen() {
	var ui = SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
	ui.createMenu("Test")
	.addItem("Send Reminders", "controller")
	.addToUi();
	ui.createMenu("Settings")
	.addItem("Authorize", "authorize")
	.addToUi();
}

//easily authorize the script to run from the menu
function authorize(){
    var respEmail = Session.getActiveUser().getEmail();
    MailApp.sendEmail(respEmail,"Form Authorizer", "Your form has now been authorized to send you emails");
}
//Controller to tie all the functions together
function controller(){
	try {
		//gets spreadsheet data and organizes it into table like format
		var results = getReminders();
		var tasks = results[0]
		var headers = getHeaders(results[1]);
		//builds html body and sends off email
		sendEmail(tasks,headers);
	}
	catch (e) {
		MailApp.sendEmail(Session.getActiveUser().getEmail(), "Error report",
		"rnMessage: " + e.message
		+ "rnFile: " + e.fileName
		+ "rnLine: " + e.lineNumber);
	}
}

function sendEmail(tasks,headers){
	var tableStart = "<br><br><html><body><table border="1">";
	var tableEnd = "</table></body></html>";
	var header = addHTML(headers);
	var emailSub = "Weekly Reminder";
    for (i in tasks){
		var emailMsg = "Here is your weekly task reminder.  n"
		var email = tasks[i][0];
		var body = addHTML(tasks[i]);
		body = tableStart  + header + body + tableEnd;
		emailMsg = emailMsg + body;
		sendMail(email,emailSub, emailMsg);
		Logger.log(body);
	}
}

//function to get headers
function getHeaders(values){
	var headers = values[0];
	return headers;
}

//function to add html table elements to headers and rows/cells
function addHTML(plainMSG){
	plainMSG.shift();
	var rowStart = "<tr>";
	var rowEnd = "</tr>";
	var cellStart = "<td>";
	var cellEnd = "</td>";

	for (i in plainMSG){
		plainMSG[i] = rowStart + cellStart + plainMSG[i] + cellEnd + rowEnd;
	}
	plainMSG  = plainMSG.join("");
	return plainMSG;
}

//function to send out mail
function sendMail(emailAddr,emailSub, emailMsg){
	MailApp.sendEmail(emailAddr,emailSub,"", {htmlBody: emailMsg});
}

//gets spreadsheet date and builds a unique list of people and tasks
function getReminders() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getSheetByName("Tasks");
	var col = 2;
	var row = sheet.getLastRow();
	var range = sheet.getRange(1,1,row,col);
	var values = range.getValues();
	var names = values.slice(0);
	names.shift();
	names = getNames(names);
	var people = [];
	for (i in names){
		var person = [];
		person.push(names[i][0]);
			for(j in values){
				if(names[i][0] == values[j][0]){
				person.push(values[j][1]);
				}
			}
		people.push(person);
	}
	return [people, values];
}

//function used to get unique names
function getNames(values){
	var noDups = values;
	noDups.sort();
	for(var i = noDups.length-1; i > 0; i--){
		if(noDups[i][0].trim() == noDups[i-1][0].trim()){
		noDups.splice(i-1,1);
		}
	}
	return noDups;
}