var moment = Moment.load();

var GLOBAL = {
  
  formId: "1y4lWV8FrIFj1G6gLvI8NlQPsP9B0tl23hiwTyVqP_5E",
  calendarId: "allenisd.org_o995gikjf8qut9emhhr5g5nggc@group.calendar.google.com",
  
  formMap : {
    eventTitle: "Your Name",
    startTime : "Start Time",
    date : "Date",
    endTime: "End Time",
    description: "Extra notes",
    
    //location: "Event Location",
    email: "Add Guests",
    
  },
};

function onFormSubmit() {
  var eventObject = getFormResponse();
  var event = createCalendarEvent(eventObject);
}

function getFormResponse() {
  // Get a form object by opening the form using the
  // form id stored in the GLOBAL variable object
  var form = FormApp.openById(GLOBAL.formId),
      //Get all responses from the form. 
      //This method returns an array of form responses
      responses = form.getResponses(),
      //find the length of the responses array
      length = responses.length,
      //find the index of the most recent form response
      //since arrays are zero indexed, the last response 
      //is the total number of responses minus one
      lastResponse = responses[length-1],
      //get an array of responses to every question item 
      //within the form for which the respondent provided an answer
      itemResponses = lastResponse.getItemResponses(),
      //create an empty object to store data from the last 
      //form response
      //that will be used to create a calendar event
      eventObject = {};
  //Loop through each item response in the item response array
  eventObject.email = lastResponse.getRespondentEmail();
  for (var i = 0, x = itemResponses.length; i<x; i++) {
    //Get the title of the form item being iterated on
    var thisItem = itemResponses[i].getItem().getTitle(),
        //get the submitted response to the form item being
        //iterated on
        thisResponse = itemResponses[i].getResponse();
    //based on the form question title, map the response of the 
    //item being iterated on into our eventObject variable
    //use the GLOBAL variable formMap sub object to match 
    //form question titles to property keys in the event object
    switch (thisItem) {
      case GLOBAL.formMap.eventTitle:
        eventObject.title = thisResponse;
        break;
      case GLOBAL.formMap.startTime:
        eventObject.startTime = thisResponse;
        break;
      case GLOBAL.formMap.endTime:
        eventObject.endTime = thisResponse;
        break; 
      case GLOBAL.formMap.description:
        eventObject.description = thisResponse;
        break;
      case GLOBAL.formMap.date:
        eventObject.date = thisResponse;
      //case GLOBAL.formMap.location:
        //eventObject.location = thisResponse;
        //break;
      //case GLOBAL.formMap.email:
        //eventObject.email = thisResponse;
        //break;
    } 
  }
  //eventObject.email = form.getRespondentEmail();
  eventObject.startTime = eventObject.date + " " + eventObject.startTime + ":00";
  eventObject.endTime = eventObject.date + " " + eventObject.endTime + ":00";
  return eventObject;
}

function createCalendarEvent(eventObject) {
  //Get a calendar object by opening the calendar using the
  //calendar id stored in the GLOBAL variable object
  var calendar = CalendarApp.getCalendarById(GLOBAL.calendarId),
      //The title for the event that will be created
      title = eventObject.title,
      //The start time and date of the event that will be created
      startTime = moment(eventObject.startTime).toDate(),
      //The end time and date of the event that will be created
      endTime = moment(eventObject.endTime).toDate();
  //an options object containing the description and guest list
  //for the event that will be created
  var options = {
    description : eventObject.description,
    //guests : eventObject.email,
    //location: eventObject.location,
  };
  
  //PUT THE TESTING FOR OVERLAPPING APPOINTMENTS HERE
  var startTime2 = startTime;
  var endTime2 = endTime;
  startTime2.setSeconds(startTime2.getSeconds() + 1);
  endTime2.setSeconds(endTime2.getSeconds() - 1);
  var overlaptest = calendar.getEvents(startTime2, endTime2);
  
  
  if(!Array.isArray(overlaptest) || !overlaptest.length) {
  
  try {
    //create a calendar event with given title, start time,
    //end time, and description and guests stored in an 
    //options argument
    var event = calendar.createEvent(title, startTime, 
                                     endTime, options)
    } catch (e) {
      //delete the guest property from the options variable, 
      //as an invalid email address with cause this method to 
      //throw an error.
      delete options.guests
      //create the event without including the guest
      var event = calendar.createEvent(title, startTime, 
                                       endTime, options)
      }
  
    return event;   }
  
  else{
    MailApp.sendEmail(eventObject.email, "Overlapping Reservations [AUTOMATED RESPONSE]", "Your recent submission for a computer lab reservation overlaps with another existing reservation."); 
  }
  return null;
}
