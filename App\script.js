/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
	};

    return app;
})();


// This function is called when Office.js is ready to start your Add-in
Office.initialize = function (reason) {
	$(document).ready(function () {

		app.initialize();

		var myOm;
		var myItem;
		var myBody;

		myOm = Office.context.mailbox;
		myItem = myOm.item;

		if (myItem.itemType == Office.MailboxEnums.ItemType.Message) {

			if (Office.context.mailbox.item.body.getAsync !== undefined) {
                Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                    myBody = asyncResult.value;
					var temp = myBody.replace(/\n/g,"<br>&gt; ");
					myItem.displayReplyAllForm
	        		("<html><head><head><body>\n\n" + temp + "\n</body></html>");
                });
            }
            else { // Method not available
                app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
            }
	    }
	});
};