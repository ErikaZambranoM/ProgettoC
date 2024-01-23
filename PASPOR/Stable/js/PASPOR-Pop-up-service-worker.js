// Start the service worker
console.log('Starting Service Worker...');

// Declare variables
const FileName = self.location.pathname.split('/').pop();
let LocalStorage_PA_Session_Token = null;
const inMemoryCache = {};

// Logging function that includes the file name and line number of the caller
function Log({
    SourceFileName = null,
    SourceLine = null,
    Message = null
}) {
    let CurrentDateTime = new Date().toLocaleString();

    // If message is from within the service worker itself, use the service worker's line number
    if (SourceFileName === null) {
        SourceFileName = FileName;
        const stackTrace = new Error().stack.split('\n').slice(1);
        const caller = stackTrace.find(line => !line.includes('at Log'));
        SourceLine = caller.match(/:(\d+):\d+/)[1];
    }
    console.log(CurrentDateTime + ' - [' + SourceFileName + ':' + SourceLine + ']', Message);
};


// Create a listener to receive messages from the popup
Log({ Message: 'Starting console output Listener...' });
self.addEventListener('message', event => {
    let SourceFileName = event.source.url.split('/').pop()
    // Print the message to the service worker console
    Log({ SourceFileName: SourceFileName, SourceLine: event.data.SourceLine, Message: event.data.message });
});

// Create a listener to receive messages from content scripts
Log({ Message: 'Starting action event Listener...' });
chrome.runtime.onMessage.addListener((request) => {

    // Check if the message is the token
    if (request.action === 'LocalStorage_PA_Session_Retrieved') {
        // Store the token
        LocalStorage_PA_Session_Token = request.LocalStorage_PA_Session_Object.secret;

        // Send the token to the popup
        //console.log("Received token: ", LocalStorage_PA_Session_Token);
        chrome.runtime.sendMessage({ action: 'Send_Token_To_Popup', token: LocalStorage_PA_Session_Token });
    }

    // Simply print a message to the service worker console
    else if (request.action === 'message') {
        let SourceFileName = request.SourceFileName;
        Log({
            SourceFileName: SourceFileName, SourceLine: request.SourceLine, Message: request.message
        });
    }
});


// Create a cache
Log({ Message: 'Starting cache listener...' });
self.addEventListener('fetch', event => {
    // Exclude requests with a chrome-extension scheme
    if (event.request.url.startsWith('chrome-extension://')) {
        return;
    }

    // Cache other requests
    event.respondWith(
        caches.match(event.request).then(response => {
            if (typeof response === 'undefined') {
                Log({ Message: 'Caching data for site: ' + event.request.url });
            } else {
                Log({ Message: 'Retrieving data from cache for site: ' + event.request.url });
            }
            return response || fetch(event.request).then(async fetchResponse => {
                if (!fetchResponse.ok) { // Check if status code is OK
                    throw new Error(`Server responded with error: ${fetchResponse.statusText}`);
                }
                const cache = await caches.open('PASPOR');
                cache.put(event.request, fetchResponse.clone());
                return fetchResponse;
            });
        }).catch(error => {
            Log({ Message: `Objects retrieval failed: ${error.message}.` });
            self.clients.matchAll().then(clients => {
                clients.forEach(client => {
                    client.postMessage({
                        type: 'ERROR',
                        message: `Fetch event handling failed: ${error.message}.`
                    });
                });
            });
            return new Response(JSON.stringify({ error: 'An error occurred', details: error.message }), {
                headers: { 'Content-Type': 'application/json' },
                status: 500,
                statusText: 'Internal Server Error'
            });
        })
    );
});


//self.addEventListener('fetch', event => {
//
//    // Exclude requests with a chrome-extension scheme
//    if (event.request.url.startsWith('chrome-extension://')) {
//        return;
//    }
//
//    // Cache other requests
//    event.respondWith(
//        caches.match(event.request).then(response => {
//
//            // If the request is in the normal cache because of type 'basic', return it from normal cache
//            if (response) {
//                Log({ Message: 'Retrieving data from cache for site: ' + event.request.url });
//                return response;
//            }
//
//            // If the request is in the in-memory cache because not of type 'basic', return it from in-memory cache
//            else if (inMemoryCache[event.request.url]) {
//                Log({ Message: 'Retrieving data from in-memory cache for site: ' + event.request.url });
//                return new Response(inMemoryCache[event.request.url]);
//            }
//
//            // If the request is not in any cache, fetch it and cache it based on type
//            else {
//                Log({ Message: `Fetching for site: ${event.request.url} ...` });
//                return fetch(event.request).then(fetchResponse => {
//                    Log({ Message: 'Respons type: ' + fetchResponse.type });
//
//                    if (fetchResponse.type !== 'basic') {
//                        fetchResponse.clone().text().then(content => {
//                            inMemoryCache[event.request.url] = content;
//                        });
//                        Log({ Message: 'Caching data on in-memory cache (will be lost when extension reloads)...' });
//                        return fetchResponse;
//                    }
//                    else {
//                        return response || fetch(event.request).then(async fetchResponse => {
//                            const cache = await caches.open('PASPOR');
//                            cache.put(event.request, fetchResponse.clone());
//                            Log({ Message: 'Caching data on normal cache...' });
//                            return fetchResponse;
//                        });
//                    }
//                });
//            }
//        })
//    );
//});

// Service worker is loaded
Log({ Message: 'Service Worker Loaded.' });