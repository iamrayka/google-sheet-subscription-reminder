// Reads the input from Google sheet and sends subscription renewal email notification
function sendReminders() {
    // Read the values from tracking sheet, e.g. a sheet with 7 columns 
    const trackingsheet = SpreadsheetApp.getActive().getSheetByName("the title of the sheet")
    // The following method will be acepting as its input
    // getrange() = (row, column, numRows, numColumns)
    
    // In case you'd like to send the email directly to your cutomer,
    // add a column between renewal "secondReminderDay" and "comments" columns
    // label it as you desire, e.g. Customer e-mail
    const data = trackingsheet.getRange(2, 1, trackingsheet.getLastRow() - 1, 7).getValues()


    // Looping through all rows
    data.forEach(function(row) {
        const customerName = row[0]
        const vendorName = row[1]
        const serviceName = row[2]
        const renewalDate = row[3]
        const firstReminderDay = row[4]
        const secondReminderDay = row[5]
        const recipientEmail = row [6]
        const daysLeft = getDate(renewalDate)


        // First notification e-mail
        if (daysLeft == firstReminderDay) {
            const emailSubject = `Time to renew your ${vendorName} ${serviceName} subscription for ${customerName}`;
            
            const emailBody =
                `Hello There,
                It's time to renew ${vendorName}'s ${serviceName} subscription for ${customerName}. 
                Please look at the details here-: 
                ${trackingsheet.getParent().getUrl()}
                Many thanks`
                
            MailApp.sendEmail(recipientEmail, emailSubject, emailBody)
        }


        // Second notification e-mail
        if (daysLeft == secondReminderDay) {
            const emailSubject = `Time to renew your ${vendorName} ${serviceName} subscription for ${customerName}`;
            
            const emailBody =
                `Hello There,
                It's time to renew ${vendorName}'s ${serviceName} subscription for ${customerName}.
                Please look at the details here-: 
                ${trackingsheet.getParent().getUrl()}
                Note : this is the last reminder; please take an action now by either renewing or cancelling your subscription.
                Thank You.`
            
            MailApp.sendEmail(recipientEmail, emailSubject, emailBody)
        }




    })
}

// getDate() takes a date and return: currentDate - givenDate [i.e. renewalDate in this case]
function getDate(providedDate) {
    var todaysDate = new Date();
    var creationDate = new Date(providedDate);
    //Math.abs() returns positive number only 
    const diffTime = Math.abs(creationDate - todaysDate);
    // 1 seconds = 1000 milliseconds
    // 1 day = 24 hours x 60 minutes x 60 seconds 
    const diffDays = Math.ceil(diffTime / (24 * 60 * 60 * 1000));
    return diffDays
}