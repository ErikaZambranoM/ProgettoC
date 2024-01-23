function LogToServiceWorker(Message) {
    const stackTrace = new Error().stack.split('\n');
    const caller = stackTrace.slice(1).find(line => !line.includes('at LogToServiceWorker'));
    const lineNumber = caller.match(/:(\d+):\d+/)[1];
    const relevantLine = stackTrace[1];
    const SourceFileName = relevantLine.substring(relevantLine.lastIndexOf('/') + 1, relevantLine.indexOf('.js') + 3);
    chrome.runtime.sendMessage({ action: 'message', SourceFileName: SourceFileName, SourceLine: lineNumber, message: Message });
};

// Fetch authentication data from Local Storage
chrome.runtime.sendMessage({ action: 'message', message: 'Retrieving session authentication...' });
var keys = Object.keys(localStorage);
var LocalStorage_FilteredKeys = keys.filter(key => key.includes('https://service.flow.microsoft.com//user_impersonation https://service.flow.microsoft.com//.default'));
var LocalStorage_PA_Session = null;
if (LocalStorage_FilteredKeys.length > 0) {
    var LocalStorage_PA_Session_String = localStorage.getItem(LocalStorage_FilteredKeys[0]);
    var LocalStorage_PA_Session_Object = JSON.parse(LocalStorage_PA_Session_String);
    // Send the token object back to the background script
    chrome.runtime.sendMessage({ action: 'LocalStorage_PA_Session_Retrieved', LocalStorage_PA_Session_Object: LocalStorage_PA_Session_Object });
    //chrome.runtime.sendMessage({ action: 'LocalStorage_PA_Session_Retrieved', LocalStorage_PA_Session_Object: 'LocalStorage_PA_Session_Object' }); // For testing null token
}