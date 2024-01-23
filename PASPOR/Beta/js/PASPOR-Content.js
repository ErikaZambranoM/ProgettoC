function LogToServiceWorker(Message) {
    const stackTrace = new Error().stack.split('\n');
    const caller = stackTrace.slice(1).find(line => !line.includes('at LogToServiceWorker'));
    const lineNumber = caller.match(/:(\d+):\d+/)[1];
    const relevantLine = stackTrace[1];
    const SourceFileName = relevantLine.substring(relevantLine.lastIndexOf('/') + 1, relevantLine.indexOf('.js') + 3);
    chrome.runtime.sendMessage({ action: 'message', SourceFileName: SourceFileName, SourceLine: lineNumber, message: Message });
};

// Retrieve local storage keys
var keys = Object.keys(localStorage);

// Fetch Power Automate authentication data from Local Storage filtering the keys that only include the Power Automate session resources
LogToServiceWorker('Retrieving Power Automate session authentication...');
var PA_LocalStorage_FilteredKeys = keys.filter(key => key.includes('https://service.flow.microsoft.com//user_impersonation https://service.flow.microsoft.com//.default'));
var LocalStorage_PA_Session = null;
if (PA_LocalStorage_FilteredKeys.length > 0) {
    var LocalStorage_PA_Session_String = localStorage.getItem(PA_LocalStorage_FilteredKeys[0]);
    var LocalStorage_PA_Session_Object = JSON.parse(LocalStorage_PA_Session_String);
}

// Retrieve Organization ID from Local Storage
//! add error handling for when the Organization ID or local key is not found
LogToServiceWorker('Retrieving Dynamics instanceApiUrl...');
var PA_Env_Key_String = localStorage.getItem("powerautomate-environments");
if (PA_Env_Key_String.length > 0) {
    var PA_Env_Object = JSON.parse(PA_Env_Key_String);
    var Dynamics_InstanceApiUrl = PA_Env_Object.value[0].properties.linkedEnvironmentMetadata.instanceApiUrl;
}
//LogToServiceWorker('Dynamics instanceApiUrl: ' + Dynamics_InstanceApiUrl);

// Fetch Dynamics authentication data from Local Storage filtering the keys that only include the Dynamics session resources
LogToServiceWorker('Retrieving Dynamics session authentication...');
var Dynamics_LocalStorage_FilteredKeys = keys.filter(key => key.includes(`${Dynamics_InstanceApiUrl}/user_impersonation ${Dynamics_InstanceApiUrl}/.default`));
var LocalStorage_Dynamics_Session = null;
if (Dynamics_LocalStorage_FilteredKeys.length > 0) {
    var LocalStorage_Dynamics_Session_String = localStorage.getItem(Dynamics_LocalStorage_FilteredKeys[0]);
    var LocalStorage_Dynamics_Session_Object = JSON.parse(LocalStorage_Dynamics_Session_String);
}

// Merge the token objects into a single object
var LocalStorage_Session_Object = {
    PA_Session_Object: LocalStorage_PA_Session_Object,
    Dynamics_Session_Object: LocalStorage_Dynamics_Session_Object,
    Dynamics_InstanceApiUrl: Dynamics_InstanceApiUrl
};

// Send the token object back to the background script
chrome.runtime.sendMessage({ action: 'LocalStorage_Session_Retrieved', LocalStorage_Session_Object: LocalStorage_Session_Object }); // To test a null token, simply replace the token object with a string
//LogToServiceWorker(LocalStorage_Session_Object);  // Debug only, print the token object to the console