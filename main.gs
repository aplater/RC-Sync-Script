function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('————> Turn on Script <————')
        .addItem('Run Now',
            'createOnEditTrigger')
        .addToUi();
}

function syncActive() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'Tracker';

    if (ss.getName() == sheetName) {
        syncSheet(ss);
    } else {
        ss.toast("To run this feature, you must work from the 'Tracker' sheet only.", "Warning!", 25);
    }
}


function create_event(e) {

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    var ssa = SpreadsheetApp.getActive();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var lastRow = sheet.getLastRow();
  
  /* Google Calendar ID, employee name and employee email sit within the Settings sheet and allow the publishing of events and importing of the employee's details into the invite where specified by string. */

    var calendarID = ss.getSheetByName('Settings (Do Not Delete)').getRange(7, 4).getValue();
    console.log(calendarID);
    var calendar = CalendarApp.getCalendarById(calendarID);

    var employee_name = ss.getSheetByName('Settings (Do Not Delete)').getRange(8, 4).getValue();
    var employee_email = ss.getSheetByName('Settings (Do Not Delete)').getRange(9, 4).getValue();

    var index = e.range.getRow();
    var status = sheet.getRange(index, 2).getValue();
    var candidate_name = sheet.getRange(index, 3).getValue();
    var source = sheet.getRange(index, 4).getValue();
    var candidate_phone = sheet.getRange(index, 5).getValue();
    var candidate_email = sheet.getRange(index, 6).getValue();
    var recruiter = sheet.getRange(index, 7).getValue();
    var hiring_manager = sheet.getRange(index, 8).getValue();

    var job_link = sheet.getRange(index, 9).getValue();
    var code_pair = sheet.getRange(index, 10).getValue();

    var interview_type = sheet.getRange(index, 11).getValue(); // Dropdown dictates template of invite
    var make_event = sheet.getRange(index, 12);
    var event_link = sheet.getRange(index, 13); 
    var interview_date = sheet.getRange(index, 15);

    // Checks if the job link and code pair columns are blank. If so, adds or removes.
  
    var isJobBlank = ssa.getSheetByName("Tracker").getRange(index, 9).isBlank();
    if (ssa.getSheetByName("Tracker").getRange(index, 9).isBlank()) {
        isJobBlank = "";
    } else {
        isJobBlank = "Job Description";
    }
  
    var isCodePairBlank = ssa.getSheetByName("Tracker").getRange(index, 10).isBlank();
    if (ssa.getSheetByName("Tracker").getRange(index, 10).isBlank()) {
        isCodePairBlank = "";
    } else {
        isCodePairBlank = "Code Pair Link";
    }

    // Adds a separator | if both of the above options are present
  
    var doBothExist = isJobBlank + isCodePairBlank;
    if (ssa.getSheetByName("Tracker").getRange(index, 9).isBlank() || ssa.getSheetByName("Tracker").getRange(index, 10).isBlank()) {
        doBothExist = "";
    } else {
        doBothExist = " | ";
    }
  

    if (e.range.getColumn() == 12) {

        if (make_event.isChecked() == true) {


            /* // Only allow the creation of an event based on set parameters (not recommended)
            var fromDate = new Date(2020,0,1,0,0,0);
            var toDate = new Date(2050,1,1,0,0,0); 
            var events = calendar.getEvents(fromDate, toDate, {search:candidate_name}); */

            if (ssa.getSheetByName("Tracker").getRange(index, 13).isBlank()) {

                if (ssa.getSheetByName("Tracker").getRange(index, 11).isBlank()) {
                    ss.toast("You must choose an Interview Type.", "WARNING!", 25);
                } else {

                    var event_title = "";
                    var desc = "";
                    var today = new Date();

                    var thirtyMinutes = new Date(today);
                    thirtyMinutes.setMinutes(today.getMinutes() + 30);

                    // Event parameters: all day, private, red color
                  
                    var event = calendar.createAllDayEvent(event_title, today, {
                        description: desc
                    }).setVisibility(CalendarApp.Visibility.PRIVATE).setColor("11"); // 11 = RED
                    console.log('Event ID: ' + event.getId());

                    var eventId = event.getId().split('@');
                    console.log(eventId);

                    // Hyperlinks the event generate in the "invite link" column

                    var url = "https://www.google.com/calendar/u/0/r/eventedit/" + Utilities.base64Encode(eventId[0] + " " + calendarID).replace("==", '');
                    console.log(url);
                    event_link.setFormula('=HYPERLINK("' + url + '","View Hiring Team Invite")');

                    // Fills in the interview date in the "interview date" column
                  
                    interview_date.setValue(event.getStartTime());

                  
                        // Automatically generates and attaches a Google Meet link (where necessary, e.g. Google Meet Interview)
                  
                        var tmpEvent = {
                          conferenceData:{
                            createRequest:{
                              conferenceSolutionKey:{
                                type: "hangoutsMeet"
                              },
                            requestId: charIdGenerator()
                            }
                          }
                        }
                        
                        var gMeetIdSplit = event.getId().replace("@google.com","");
                        console.log(gMeetIdSplit);

/* Removed column where dropdown "urgent" is present, injects [URGENT] in front of the event title.

                    if (priority == 'Urgent' || priority == 'Reschedule Urgent') {
                        event.setTitle("[URGENT] " + String(interview_type) + " - " + String(candidate_name));
                    }
*/



                  
                  /* "Interview Type" column as specified in vars in beginning will contain a dropdown (through strict data validaton). Each dropdown option will represent a template that can be designed below.

                    // +++++++++++++++++++++++++++++++++ PHONE INTERVIEW TEMPLATE

                    var description_phone = '<strong>This is a phone interview.</strong><br /><br /></strong><strong>RC:</strong> Please email ' + String(employee_name) + ‘ at ' + String(employee_email) + 'for all scheduling updates<br /><strong>Hiring Manager:</strong>' + String(hiring_manager) + '<br /><strong>Recruiter:</strong>' + String(recruiter) + '<br /><a href="' + String(job_link) + '">' + String(isJobBlank) + '</a>'+ String(doBothExist) +'<a href="' + String(code_pair) + '">' + String(isCodePairBlank) + '</a><br/><br /><strong>Interview Feedback Form</strong><em>Placeholder text in italics.</em><br /><br /><strong>Candidate Info:</strong><br />Source: ' + String(source) + '<br />' + String(candidate_phone) + '<br />' + String(candidate_email) + '<br /><strong><br /></strong><strong>Additional Resources:</strong><ul><li><a href=“http://google.com/“>Google</a></li></ul><em>Warning note.</em>';

                    if (interview_type == 'Phone Interview') { // These words must match what is set in the data validation in "Interview Type" column
                        event.setTitle("Phone Interview - " + String(candidate_name)); // Title of event
                        event.setDescription(description_phone + ""); // Sets description to the above var description_phone (make sure var being called on matches)
                        event.setLocation("Please call the candidate at " + String(candidate_phone)); // Location field of event
                        ss.toast("Reminder to attach the resume.", "Invite Created!", 25); // Specific pop up if this template is utilized
                    }


                    // +++++++++++++++++++++++++++++++++ GOOGLE MEET INTERVIEW TEMPLATE

                    var description_phone = '<strong>This is a Google Meet interview.</strong><br /><br /></strong><strong>RC:</strong> Please email ' + String(employee_name) + ‘ at ' + String(employee_email) + 'for all scheduling updates<br /><strong>Hiring Manager:</strong>' + String(hiring_manager) + '<br /><strong>Recruiter:</strong>' + String(recruiter) + '<br /><a href="' + String(job_link) + '">' + String(isJobBlank) + '</a>'+ String(doBothExist) +'<a href="' + String(code_pair) + '">' + String(isCodePairBlank) + '</a><br/><br /><strong>Interview Feedback Form</strong><em>Placeholder text in italics.</em><br /><br /><strong>Candidate Info:</strong><br />Source: ' + String(source) + '<br />' + String(candidate_phone) + '<br />' + String(candidate_email) + '<br /><strong><br /></strong><strong>Additional Resources:</strong><ul><li><a href=“http://google.com/“>Google</a></li></ul><em>Warning note.</em>';

                    if (interview_type == 'Google Meet Interview') { 
                        event.setTitle("Google Meet Interview - " + String(candidate_name)); 
                        event.setDescription(description_google_meet + "");
                        event.setLocation("Please join through the Google Meet link"); 
                        ss.toast("Don't forget to attach the resume.", "Invite Created!", 25); 
                        eventResource = Calendar.Events.patch(tmpEvent, calendarID, gMeetIdSplit, {conferenceDataVersion:1}); // Automatically generates and attaches a Google Meet for this chosen template
                    }



                    // In this scenario, if the following template is chosen, two separately designed events are created and hyperlinked for a blind internal interviewing process (hiring team receives an invite with private details and candidate receives a separate invite with a matching Google Meet link)

                    // +++++++++++++++++++++++++++++++++ DOUBLE GOOGLE MEET INTERVIEW TEMPLATE

                    var hiring_team_description_internal_google_meet = '<strong>This is a Google Meet interview.</strong><br /><br /></strong><strong>RC:</strong> Please email ' + String(employee_name) + ‘ at ' + String(employee_email) + 'for all scheduling updates<br /><strong>Hiring Manager:</strong>' + String(hiring_manager) + '<br /><strong>Recruiter:</strong>' + String(recruiter) + '<br /><a href="' + String(job_link) + '">' + String(isJobBlank) + '</a>'+ String(doBothExist) +'<a href="' + String(code_pair) + '">' + String(isCodePairBlank) + '</a><br/><br /><strong>Interview Feedback Form</strong><em>Placeholder text in italics.</em><br /><br /><strong>Candidate Info:</strong><br />Source: ' + String(source) + '<br />' + String(candidate_phone) + '<br />' + String(candidate_email) + '<br /><strong><br /></strong><strong>Additional Resources:</strong><ul><li><a href=“http://google.com/“>Google</a></li></ul><em>Warning note.</em>';

                    var candidate_description_internal_google_meet = '<strong>RC:</strong>&nbsp;Please email&nbsp;' + String(employee_name) + '&nbsp;at ' + String(employee_email) + '&nbsp;for all scheduling updates<br /><strong>Hiring Manager:&nbsp;</strong>' + String(hiring_manager) + '<br /><strong>Recruiter:&nbsp;</strong>' + String(recruiter) + '<br /><strong>Interviewer(s): </strong><i>Insert here.</i><br /><br /><strong>NOTE: </strong>Special note here.';


                    if (interview_type == 'Internal Google Meet Interview') { 
                        event.setTitle("Google Meet Interview - " + String(candidate_name));
                        event.setDescription(hiring_team_description_internal_google_meet + ""); 
                        event.setLocation("Please join through the Google Meet link"); 
                        ss.toast("Don't forget to attach the resume.", "Invites Created!", 30); 

                        // Generates and attaches a Google Meet link
                        
                        eventResource = Calendar.Events.patch(tmpEvent, calendarID, gMeetIdSplit, {conferenceDataVersion:1});
                        console.log(eventResource.conferenceData);
                        
                       
                              /* Event #2, created simultaneously */
                              
                              var event2 = calendar.createAllDayEvent(event_title, today, {
                                  description: desc
                              }).setVisibility(CalendarApp.Visibility.PRIVATE).setColor("6"); // 6 = ORANGE (color differs from event #1, which is red).

                              var eventId2 = event2.getId().split('@');
                              console.log(eventId2);

                              event2.setTitle("Google Meet Interview - " + String(candidate_name)); // 2nd event's title
                              event2.setDescription(candidate_description_internal_google_meet + ""); // 2nd event's description (in the var above).
                              event2.setLocation("Your interviewer(s) will join you through the Google Meet link"); // 2nd event's location

                              var url2 = "https://www.google.com/calendar/u/0/r/eventedit/" + Utilities.base64Encode(eventId2[0] + " " + calendarID).replace("==", '');
                              console.log(url2);
                              event_link.offset(0,1).setFormula('=HYPERLINK("' + url2 + '","View Candidate Invite")');

                              // Matches the Google Meet link of event #1
                  
                              var gMeetIdSplit2  = event2.getId().replace("@google.com","");
                              
                              var tmpEvent2 = {
                                conferenceData: eventResource.conferenceData
                              }

                              Calendar.Events.patch(tmpEvent2,calendarID,gMeetIdSplit2,{conferenceDataVersion:1});
                    }

              
                    // Stripped down template, feel free to customize or replicate as needed

                    // +++++++++++++++++++++++++++++++++ ONSITE INTERVIEW

                    var description_onsite = "Onsite Interview";

                    if (interview_type == 'Onsite Interview') {
                        event.setDescription(description_onsite + "");
                        event.setLocation("Onsite Interview");
                    }


              
             
           /* TEMPLATES END HERE */
              
              
              

            // If invite link is already hyperlinked in row, checkbox will serve as an interview date grabber by analyzing invite link.
            // When checkbox is unchecked and checked again, checks "if blank" of original if statement for event creation (above) is false. If false, rather than creating new event, reads the hyperlinked event.
              
            } if (!ssa.getSheetByName("Tracker").getRange(index, 13).isBlank()) {

                const viewInvite = sheet.getRange(index, 13).getFormula();
                const getEventId = viewInvite.match(/eventedit\/(\w+)/);

                if (getEventId && getEventId.length == 2) {
                    const splitEventId = Utilities.newBlob(Utilities.base64Decode(getEventId[1])).getDataAsString().split(" ");
                    const timeOfEvent = calendar.getEventById(splitEventId[0]).getStartTime();
                    console.log(splitEventId[0]);
                  
                  // When grabbing the date of the analyzed link, replaces old date with new in "Interview Date" column
                  
                  interview_date.setValue(timeOfEvent);
                  
                }
            }
        }
    }
}

// onEdit/checkbox click event creation function

function createOnEditTrigger(e) {
    var triggers = ScriptApp.getProjectTriggers();
    var shouldCreateTrigger = true;
    triggers.forEach(function(trigger) {

        if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === "create_event") {
            shouldCreateTrigger = false;
        }
    });

    if (shouldCreateTrigger) {
        ScriptApp.newTrigger("create_event").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
    }
}


// Generates a random string for the automatic Google Meet link creation and attachment process

function charIdGenerator() {
    var charId = "";
    for (var i = 1; i < 10; i++) {
        charId += String.fromCharCode(97 + Math.random() * 10);
    }
    return charId;
}
