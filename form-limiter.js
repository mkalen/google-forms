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
 *  Modifications for week-occurring triggers
 *  By:     MK (mkalen @ GitHub)
 *  Web:    https://github.com/mkalen/google-forms
 */

const weekdays = [ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" ];
const DO_NOT_MAIL = 0;
const MAIL_ON_OPEN = 1;
const MAIL_ON_CLOSE = 2;
const MAIL_ON_LIMIT = 4;
const ALL_MAILS = MAIL_ON_OPEN | MAIL_ON_CLOSE | MAIL_ON_LIMIT;

/**
 * User setting.
 * Set the form's open and close day / time in JSON object format with keys:
 * day: ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
 * time: "HH:MM" (local time for the actual date)
 */
const FORM_OPEN = { "day": "Sunday", "time": "11:00" };
const FORM_CLOSE = { "day": "Sunday", "time": "13:00" };

/**
 * User setting.
 * Set the RESPONSE_COUNT equal to the total number of entries 
 * that you would like to receive after which the form is closed automatically.
 * If you would not like to set a limit, set this value to blank.
 */
const RESPONSE_COUNT = "";

/**
 * User setting, choose e-mail notifications.
 * Set value to either of (or use "|" to combine):
 *  DO_NOT_MAIL = no e-mails
 *  MAIL_ON_OPEN = mail on form open by time trigger
 *  MAIL_ON_CLOSE = mail on form open by time trigger
 *  MAIL_ON_LIMIT = mail on form open by response limit
 *  ALL_MAILS = all e-mails
 */
const MAIL_NOTIFY = ALL_MAILS;

/* Initialize the form, setup time based triggers */
function Initialize() {
    weeklyReInit();
}

/* Delete all existing Script Triggers */
function deleteTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
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

    var controlOpen = false;
    const dateTime = new Date();
    var nextOpenDateTime;

    if (FORM_OPEN !== "") {
        controlOpen = true;
        nextOpenDateTime = parseDate_(dateTime, FORM_OPEN);
        if (dateTime.getTime() < nextOpenDateTime.getTime()) {
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

        // If we control both open and close, check that current interval state is fulfilled
        if (controlOpen) {
            var shouldBeOpen = nextCloseDateTime.getTime() <= nextOpenDateTime.getTime();
            if (shouldBeOpen && !isFormAcceptingResponses()) {
                openForm();
            } else if (!shouldBeOpen && isFormAcceptingResponses()) {
                closeForm();
            }
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

/**
 * Returns whether the form is currently open or not.
 * @return boolean form is currently open or not
 */
function isFormAcceptingResponses() {
    var form = FormApp.getActiveForm();
    return form.isAcceptingResponses();
}

/**
 * Allow Google Form to Accept Responses.
 */
function openForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(true);
    if (MAIL_NOTIFY & MAIL_ON_OPEN) {
        informUser_("Your Google Form is now accepting responses");
    }
}

/**
 * Close the Google Form, Stop Accepting Reponses.
 */
function closeForm() {
    var form = FormApp.getActiveForm();
    form.setAcceptingResponses(false);
    if (MAIL_NOTIFY & MAIL_ON_CLOSE) {
        informUser_("Your Google Form is no longer accepting responses");
    }
}

/* If Total # of Form Responses >= Limit, Close Form */
function checkLimit() {
    if (FormApp.getActiveForm().getResponses().length >= RESPONSE_COUNT) {
        if (MAIL_NOTIFY & MAIL_ON_LIMIT) {
            informUser_("Your Google Form reached the response limit");
        }
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
    // If time has already passed in comparison to base date, add a week's time
    if (nextDate.getTime() <= baseDate.getTime()) {
      nextDate.setDate(nextDate.getDate() + 7);
    }
    return nextDate;
}
