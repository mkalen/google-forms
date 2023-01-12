/**
 *  ----------------------
 *   Google Forms Limiter
 *  ----------------------
 * 
 *  Set response limits for Google Forms and close form automatically.
 * 
 *  license: MIT
 *  language: Google Apps Script
 * 
 *  Original author: Amit Agarwal
 *  email: amit@labnol.org
 *  web: https://digitalinspiration.com/
 *
 *  Modifications for week-occurring triggers by: Martin Kalén (mkalen @ GitHub)
 */


/* Set the form's open and close day / time in JSON object format with keys:
 * day: ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
 * time: "HH:MM" (local time for the actual date)
 */
const FORM_OPEN = { "day": "Friday", "time": "16:00" };
const FORM_CLOSE = { "day": "Wednesday", "time": "08:00" };

/* Set the RESPONSE_COUNT equal to the total number of entries 
that you would like to receive after which the form is closed automatically. 
If you would not like to set a limit, set this value to blank. */
RESPONSE_COUNT = "";

const weekdays = [ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ];

/* Initialize the form, setup time based triggers */
function Initialize() {
    weeklyReInit();
}

/* Delete all existing Script Triggers */
function deleteTriggers_() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i in triggers) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
}

/* Send a mail to the form owner when the form status changes */
function informUser_(subject) {
    var formURL = FormApp.getActiveForm().getPublishedUrl();
    MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, formURL);
}

function weeklyReInit() {
    deleteTriggers_();

    const dateTime = new Date();
    if (FORM_OPEN !== "") {
        const nextOpenDateTime = parseDate_(dateTime, FORM_OPEN);
        if (dateTime.getTime() < nextOpenDateTime.getTime()) {
            closeForm();
            ScriptApp.newTrigger("openForm")
                .timeBased()
                .at(nextOpenDateTime)
                .create();
        }
    }

    if (FORM_CLOSE !== "") {
        const nextCloseDateTime = parseDate_(dateTime, FORM_CLOSE);
        if (dateTime.getTime() < nextCloseDateTime.getTime()) {
            ScriptApp.newTrigger("closeForm")
                .timeBased()
                .at(nextCloseDateTime)
                .create();
        }
    }

    if (RESPONSE_COUNT !== "") {
        ScriptApp.newTrigger("checkLimit")
            .forForm(FormApp.getActiveForm())
            .onFormSubmit()
            .create();
    }

    // Re-init in a week
    dateTime.setDate(dateTime.getDate() + 7);
    ScriptApp.newTrigger("weeklyReInit")
        .timeBased()
        .at(dateTime)
        .create();
}

/* Allow Google Form to Accept Responses */
function openForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(true);
    informUser_("Your Google Form is now accepting responses");
}

/* Close the Google Form, Stop Accepting Reponses */
function closeForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(false);
    informUser_("Your Google Form is no longer accepting responses");
}

/* If Total # of Form Responses >= Limit, Close Form */
function checkLimit() {
    if (FormApp.getActiveForm().getResponses().length >= RESPONSE_COUNT) {
        closeForm();
    }
}

/* Parse the Date for creating Time-Based Triggers */
function parseDate_(baseDate, nextTrigger) {
    // E.g: { "day": "Friday", "time": "16:00" }
    //                                  01234
    const nextWeekdayName = nextTrigger.day;
    const nextWeekdayTime = nextTrigger.time;
    const nextDayOfWeek = weekdays.indexOf(nextWeekdayName);
    // TODO: Check index -1 and report invalid weekday name
    const hours = nextWeekdayTime.substr(0, 2);
    const minutes = nextWeekdayTime.substr(3, 2);
    const nextDate = new Date(baseDate.getTime());
    nextDate.setSeconds(0);
    nextDate.setMilliseconds(0);
    nextDate.setDate(nextDate.getDate() + (nextDayOfWeek + 7 - nextDate.getDay()) % 7);
    nextDate.setHours(hours);
    nextDate.setMinutes(minutes);
    return nextDate;
}
