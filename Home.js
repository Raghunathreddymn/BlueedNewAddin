
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");
            
            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $('#Respond').click(send);
            $('#sharedoc').click(share);
            $('#btnreview').click(share);
            $('#btnreview').click(send);
            
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // This variable will keep the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
    function test()
    {
        //showNotification("Error:", error);
        window.open("Respond.html", "hello", `toolbar=no,directories=no,titlebar=no,scrollbars=no,resizable=no,status=no,location=no,toolbar=no,menubar=no,
width=300,height=300,left=300,top=300`);
    }
    function share() {
        //showNotification("Error:", error);
        window.open("ShareDocument.html", "hello", `toolbar=no,directories=no,titlebar=no,scrollbars=no,resizable=no,status=no,location=no,toolbar=no,menubar=no,
width=500,height=500,left=300,top=300`);
    }
    function send() {
        console.log("before");
       // alert("before Record Successfuly1");
        var DocumentActionUserID = 2;
        var Activity = "Review";
        var ActivityDate="29"
        var CreatedBy = CreateGuid();
        var CreatedDate = new Date().toLocaleString();
        var ModifiedBy = CreateGuid();
        var ModifiedDate = new Date().toLocaleString();
       // alert("before Record Successfuly");
            //if (txtid.length != 0 || txtname.length != 0 || txtsalary.length != 0 || txtcity.length != 0) {
                var rs = new ActiveXObject("ADODB.Recordset");
                var connection = new ActiveXObject("ADODB.Connection");
                var connectionstring = "jdbc:sqlserver://blueed.database.windows.net:1433;database=BTService;user=BlueedAdmin@blueed;password={your_password_here};encrypt=true;trustServerCertificate=false;hostNameInCertificate=*.database.windows.net;loginTimeout=30;a Source=.;Initial Catalog=EmpDetail;Persist Security Info=True;User ID=sa;Password=****;Provider=SQLOLEDB";
                connection.Open(connectionstring);
                //var rs = new ActiveXObject("ADODB.Recordset");
                 rs.Open("insert into DocumentActionUserActivityID values('" + DocumentActionUserID + "','" + Activity + "','"
                + ActivityDate + "','" + CreatedBy + "','" + CreatedDate + "','" + ModifiedBy + "','" + ModifiedDate + "')", connection);
              //  alert("Insert Record Successfuly");
               
                connection.close();
            //}
            //else {
            //    alert("Please Enter Employee \n Id \n Name \n Salary \n City ");
            //}
          
    }
    function CreateGuid() {
        function _p8(s) {
            var p = (Math.random().toString(16) + "000000000").substr(2, 8);
            return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
        }
        return _p8() + _p8(true) + _p8(true) + _p8();
    }

    
})();
