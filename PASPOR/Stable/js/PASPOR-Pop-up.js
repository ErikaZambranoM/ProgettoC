/*
    ToDo:

    - add on hover desc with shortcut

    //1
    //- Show the number of retrieved Lists and Libraries
    //    - During filter, show numberOfFiltered(Lists and Libraries) / Total(Lists and Libraries)


    2
    - Add Div under SearchBox to Show Buttons for Main Site (same properties as Lists and Libraries linkButtons)
        - Refresh (de-cache (for specific site?))
    - Remember to Add exception if user has no permissions to access a page


    3
    - Add SubDiv on downArrow click for every retrieved object
        - Show the number of elements (N/A if no permissions)
        - Buttons
            - Open
            - Open in new tab
            - Object Permissions
              - (https://tecnimont.sharepoint.com/sites/DDWave2/_layouts/15/user.aspx?obj=%7B guid %7D,list&List=%7B guid %7D)
                  - es. https://tecnimont.sharepoint.com/sites/DDWave2/_layouts/15/user.aspx?obj=%7B867F4D65-217D-400F-B4C7-41106761889D%7D,list&List=%7B867F4D65-217D-400F-B4C7-41106761889D%7D
            - Object Settings
            - Add to Favorites ObectName
    - Remember to Add exception if user has no permissions to access a page


    4
        - Add Favorites view
        - Also permit more then one custom favorite view (es.: MyDay, Transmittals, FTAStuffs)


    5
    - Keyboard shortcuts (Check availability and bypass)
        - Bind Ctrl + F to search box
        - Bind Enter to Open the first object in the list if only one container is visible
        - Bind Ctrl Enter to Open the last object in the list if only one container is visible
        - Bind Ctrl + ArrowKeys to navigate between objects
            - Or simply Arrow (and ctrl + arrow to scroll (override default behavior))
        - Bind Alt + Click on Object for multiple selection
        - Bind Ctrl + Number to open first 9 favorites
        - Main Site Buttons
            - Ctrl + H for HomePage ? or Ctrl + M for MainSite ?
            - Ctrl + S for SiteSettings ?
            - Ctrl + P for SitePermissions ?
            - Ctrl + R for RecycleBin ?
            - Ctrl + C for SiteContents ?
            - Ctrl + R for Refresh ?
        - SubDiv Buttons
            - (If (collapsedObjects.Count -gt 1 && isNotMultipleSelection){Bind to last uncollapsed div (make it clear to the user))
                - Keep Alt pressed during the whole operation?
            - Ctrl + Shift + O for Open ?
            - Ctrl + Shift + N for Open in new tab ?
            - Ctrl + Shift + P for Object Permissions ?
            - Ctrl + Shift + S for Object Settings ?


    6
    - Add sorting functionality to the lists and libraries (alphabetical (default), last created, last modified)

    7
        - Add a title to the popup (Site Title)

    8
        - Provide User Settings page with following options:
            - Tutorial
            - Verbose output for debug
            - Site Independent Automatic Scan Delay  (PickList: Seconds, Minutes, Hours) (Max: 4h; Always Scan = 0)
            - Default link tab behaviour (Same tab, new tab, new window)
            - Button colors
            - Font size or style
            - Clear cache on browser close
            - Favorites objects
            - keybinding options
            - permit user to exchange some buttons' position
            - enable sites history
            - Generic cache clean for extension

    9
        - If not on SPO url:
            - Show PrjSearchBox
            - Show GoToSiteLinkButton
                GoToSiteLinkButton(WebBrowseIcon) = Goes to site //Made with same logic of retrieved objects
                PrjSearchBox.OnConfirm = API Search for site, GetObjects, Resize windows and show (just as if the extension is clicked while on SPO page)

    10
        - Runs
            - "https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/Default-7cc91888-5aa0-49e5-a836-22cda2eae0fc/flows/5dda561c-66d8-421f-83a9-4a8a939f82ee/runs?api-version=2016-11-01"
                - URI TO OBTAIN TRIGGER COLUMNS: Internal.Properties.Trigger.outputsLink.Uri.Body?
                - API TO GET ERROR DETAILS: /flows/5dda561c-66d8-421f-83a9-4a8a939f82ee/runs/08585058704050060459102033732CU50?$expand=properties%2Factions,properties%2Fflow&api-version=2016-11-01&include=repetitionCount

*/

// #region Variables

//let listsCount = 0;
//let librariesCount = 0;

// #endregion Variables

// #region Functions

// Function to log messages to the service worker
function LogToServiceWorker(Message, lineNumber) {

    if (typeof lineNumber === 'undefined') {
        const stackTrace = new Error().stack.split('\n').slice(1);
        const caller = stackTrace.find(line => !line.includes('at LogToServiceWorker'));
        lineNumber = caller.match(/:(\d+):\d+/)[1];
    }
    navigator.serviceWorker.controller.postMessage({ SourceLine: lineNumber, message: Message });
};

// Pause function to sleep for 'ms' milliseconds (only for testing purposes)
function Pause(milliseconds) {
    const dt = new Date();
    while ((new Date()) - dt <= milliseconds) { /* Do nothing, just wait */ }
}

// Function to set intitial window size and show loading animation
function ShowLoading() {
    document.body.style.width = '50px';
    document.body.style.height = '50px';
    document.getElementById('Loading').classList.remove('Hidden');
}

// Function that checks if the URL is a SharePoint Online Site
async function CheckIfSharePointSite(url) {
    return new Promise((resolve, reject) => {
        LogToServiceWorker('Checking if SPO Url: ' + url);
        const result = url.includes('.sharepoint.com');
        resolve(result);
    });
}

// Function that checks if the URL is a Power Automate site
async function CheckIfPowerAutomateSite(url) {
    return new Promise((resolve, reject) => {
        LogToServiceWorker('Checking if PA Url: ' + url);
        const result = url.includes('make.powerautomate.com');
        resolve(result);
    });
}

// Function that extracts from an URL the SharePoint Main Site URL
const GetSharePointMainSiteUrl = url => {
    const urlObject = new URL(url);
    const pathParts = urlObject.pathname.split('/');
    let siteUrl = urlObject.origin;
    for (let i = 0; i < pathParts.length; i++) {
        if (pathParts[i] === 'sites' || pathParts[i] === 'teams') {
            siteUrl += `/${pathParts[i]}/${pathParts[i + 1]}`;
            break;
        }
    }
    return siteUrl;
};

// Function that extracts from an URL the Power Automate Environment ID
const GetPowerAutomateEnvironment = url => {
    const regex = /https:\/\/make\.powerautomate\.com\/environments\/([a-zA-Z0-9-]+)\//i;
    const match = url.match(regex);
    const environmentId = match[1];
    return environmentId;
};

// Function that retrieves SharePoint Online Site Objects
async function GetSharePointSiteObjects(SPOMainSiteUrl) {
    try {
        LogToServiceWorker('Trying to retrieve SharePoint Online Site Objects...');

        // Create a URL object from the passed SharePoint Online Site URL
        const urlObject = new URL(SPOMainSiteUrl);

        // Extract the SharePoint tenant URL from the Main Site URL
        const SPOTenantUrl = urlObject.origin;

        // Declare variables
        let listsCount = 0;
        let librariesCount = 0;
        const ListsContainer = document.getElementById('ListsContainer');
        const LibrariesContainer = document.getElementById('LibrariesContainer');
        const ListsButton = document.getElementById('ListsButton');
        const LibrariesButton = document.getElementById('LibrariesButton');
        const ShouldBustCache = localStorage.getItem('ShouldBustCache');


        const tempListenerPromise = new Promise(async (resolve, reject) => {

            // Retrieve all Lists and Document Libraries from the SharePoint Online Site
            API_Url = `${SPOMainSiteUrl}/_api/web/lists?$filter=Hidden eq false&$expand=DefaultViewUrl&$select=Title,DefaultViewUrl,BaseType`;
            if (ShouldBustCache === 'true') {
                LogToServiceWorker('Cache buster requested.');
                const CacheBuster = new Date().getTime();
                API_Url = `${API_Url}&?cacheBuster=${CacheBuster}`;
                localStorage.setItem('ShouldBustCache', 'false');
            }
            const API_Call = await fetch(API_Url, {
                headers: { 'Accept': 'application/json;odata=verbose' }
            });
            const Response = await API_Call.json();
            //LogToServiceWorker('Retrieved flows: ' + JSON.stringify(Response));

            if (Response.error) {
                return;
            }

            //Pause(5000); // Pause for 5 seconds (only for testing purposes)

            // Create a link for each list and library
            Response.d.results.forEach(item => {

                // Create an anchor tag to be used as a link inside Lists/Libraries container
                const SPO_Object_Link = document.createElement('a');
                SPO_Object_Link.href = `${SPOTenantUrl}${item.DefaultViewUrl}`;
                SPO_Object_Link.textContent = item.Title;
                SPO_Object_Link.className = 'SPO_Object_Link';
                SPO_Object_Link.style.textDecoration = 'none';
                SPO_Object_Link.style.display = 'block';
                SPO_Object_Link.id = item.Title.replace(/[^a-zA-Z0-9]/g, '_');

                // #region SPO_Object_Link Event Listeners

                // Add event listeners to the anchor tag for different mouse clicks
                let IsCtrlPressed = false;

                // Add event listener for Ctrl key press
                SPO_Object_Link.addEventListener('mousedown', function (event) {
                    if (event.button === 0 && !event.shiftKey && event.ctrlKey) {
                        // Ctrl key was pressed
                        IsCtrlPressed = true;
                    }
                });

                // Add event listener for left mouse button click variations
                SPO_Object_Link.addEventListener('click', function (event) {
                    // If Ctrl key was pressed, open the link in a new tab
                    if (IsCtrlPressed) {
                        IsCtrlPressed = false; // Reset flag
                        chrome.tabs.create({ url: SPO_Object_Link.href, active: false });
                        LogToServiceWorker('User opened on new tab: ' + SPO_Object_Link.href);
                    }

                    // Prevent the default behavior
                    event.preventDefault();

                    // If Shift key was pressed, open the link in a new window
                    if (event.shiftKey) {
                        window.open(SPO_Object_Link.href, "_blank", "noopener,noreferrer");
                        LogToServiceWorker('User opened in a new window: ' + SPO_Object_Link.href);
                    }
                    // If no modifier key was pressed, open the link in the same tab
                    else if (!event.ctrlKey) {
                        chrome.tabs.update({ url: SPO_Object_Link.href });
                        window.close();
                        LogToServiceWorker('User opened in the same tab: ' + SPO_Object_Link.href);
                    }
                });

                // Add event listener for middle mouse button click
                SPO_Object_Link.addEventListener('auxclick', function (event) {
                    if (event.button === 1) {
                        chrome.tabs.create({ url: SPO_Object_Link.href, active: false });
                        event.preventDefault();
                        LogToServiceWorker('User opened on new tab: ' + SPO_Object_Link.href);
                    }
                });

                // #endregion SPO_Object_Link Event Listeners

                // Add the anchor tag to the corresponding container
                if (item.BaseType === 0) {
                    ListsContainer.appendChild(SPO_Object_Link);
                    listsCount++;
                } else if (item.BaseType === 1) {
                    LibrariesContainer.appendChild(SPO_Object_Link);
                    librariesCount++;
                }
            });
            resolve();

        });
        // Await the promise for the tempListener to complete
        await tempListenerPromise;

        ListsButton.childNodes[0].textContent = `${ListsButton.childNodes[0].textContent} (${listsCount})`;
        LibrariesButton.childNodes[0].textContent = `${LibrariesButton.childNodes[0].textContent} (${librariesCount})`;

        // #region SiteButtons Event Listeners

        const ButtonIds = [
            "HomepageButton",
            "SiteSettingsButton",
            "SitePermissionsButton",
            "RecycleBinButton",
            "SiteContentsButton"
        ];

        ButtonIds.forEach(SiteButtonId => {

            const SiteButton = document.getElementById(SiteButtonId);

            // Determine the URL based on the button's ID
            let url;
            switch (SiteButtonId) {
                case "HomepageButton":
                    url = SPOMainSiteUrl;
                    break;
                case "SiteSettingsButton":
                    url = `${SPOMainSiteUrl}/_layouts/15/settings.aspx`;
                    break;
                case "SitePermissionsButton":
                    url = `${SPOMainSiteUrl}/_layouts/15/user.aspx`;
                    break;
                case "RecycleBinButton":
                    url = `${SPOMainSiteUrl}/_layouts/15/AdminRecycleBin.aspx?view=5`;
                    break;
                case "SiteContentsButton":
                    url = `${SPOMainSiteUrl}/_layouts/15/viewlsts.aspx?view=14`;
                    break;
                // Add more cases here
                default:
                    url = "#";
            }
            SiteButton.href = url;

            // Add event listeners to the anchor tag for different mouse clicks
            let IsCtrlPressed = false;

            // Add event listener for Ctrl key press
            SiteButton.addEventListener('mousedown', function (event) {
                if (event.button === 0 && !event.shiftKey && event.ctrlKey) {
                    // Ctrl key was pressed
                    IsCtrlPressed = true;
                }
            });

            // Add event listener for left mouse button click variations
            SiteButton.addEventListener('click', function (event) {
                // If Ctrl key was pressed, open the link in a new tab
                if (IsCtrlPressed) {
                    IsCtrlPressed = false; // Reset flag
                    chrome.tabs.create({ url: SiteButton.href, active: false });
                    LogToServiceWorker('User opened on new tab: ' + SiteButton.href);
                }

                // Prevent the default behavior
                event.preventDefault();

                // If Shift key was pressed, open the link in a new window
                if (event.shiftKey) {
                    window.open(SiteButton.href, "_blank", "noopener,noreferrer");
                    LogToServiceWorker('User opened in a new window: ' + SiteButton.href);
                }
                // If no modifier key was pressed, open the link in the same tab
                else if (!event.ctrlKey) {
                    chrome.tabs.update({ url: SiteButton.href });
                    window.close();
                    LogToServiceWorker('User opened in the same tab: ' + SiteButton.href);
                }
            });

            // Add event listener for middle mouse button click
            SiteButton.addEventListener('auxclick', function (event) {
                if (event.button === 1) {
                    chrome.tabs.create({ url: SiteButton.href, active: false });
                    event.preventDefault();
                    LogToServiceWorker('User opened on new tab: ' + SiteButton.href);
                }
            });

        });

        // #endregion SiteButtons Event Listeners

        LogToServiceWorker('SharePoint Online Site Objects retrieved.');

    } catch (error) {
        throw error;
    };
}

// Function that retrieves Power Automate Objects
async function GetPowerAutomateSiteObjects(environmentId, tabId) {
    try {
        LogToServiceWorker('Trying to retrieve Power Automate Environment Objects...');
        let filteredData = null;
        let LocalStorage_PA_Session_Token = null;
        const FlowsContainer = document.getElementById('FlowsContainer');
        const FlowsButton = document.getElementById('FlowsButton');
        const ShouldBustCache = localStorage.getItem('ShouldBustCache');


        // Temporary listener for the 'Send_Token_To_Popup' action
        const tempListenerPromise = new Promise(async (resolve, reject) => {
            const tempListener = async (request, sender, sendResponse) => {
                if (request.action === 'Send_Token_To_Popup') {
                    LocalStorage_PA_Session_Token = request.token;
                    //LogToServiceWorker('Received Token: ' + LocalStorage_PA_Session_Token);
                    chrome.runtime.onMessage.removeListener(tempListener);  // Remove listener

                    // Check if the token is null
                    if (LocalStorage_PA_Session_Token) {

                        const API_Url = `https://emea.api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${environmentId}/flows?$filter=*&$include=includeSolutionCloudFlows&api-version=2016-11-01`; //search(%27team%27)
                        if (ShouldBustCache === 'true') {
                            LogToServiceWorker('Cache buster requested.');
                            const CacheBuster = new Date().getTime();
                            API_Url = `${API_Url}&?cacheBuster=${CacheBuster}`;
                            localStorage.setItem('ShouldBustCache', 'false');
                        }
                        const API_Call = await fetch(API_Url, {
                            method: 'GET',
                            headers: {
                                'Authorization': `Bearer ${LocalStorage_PA_Session_Token}`,
                                'Content-Type': 'application/json'
                            }
                        });
                        const Response = await API_Call.json();
                        //LogToServiceWorker('Retrieved flows: ' + JSON.stringify(Response));

                        // throw an error if no data is returned from the fetch
                        if (Response.error) {
                            return;
                        }

                        // Create a new object to store the desired properties
                        filteredData = Response.value.map(item => ({
                            FlowDisplayName: item.properties.displayName,
                            FlowID: item.name
                        }));

                        // Sort the filteredData array by FlowDisplayName
                        filteredData.sort((a, b) => a.FlowDisplayName.localeCompare(b.FlowDisplayName));
                        //LogToServiceWorker('Retrieved flows: ' + JSON.stringify(filteredData));

                        // Create a link for each flow
                        filteredData.forEach(item => {

                            // Create an anchor tag to be used as a link inside Lists/Libraries container
                            const PAFlow_Object_Link = document.createElement('a');
                            PAFlow_Object_Link.href = `https://make.powerautomate.com/environments/${environmentId}/flows/${item.FlowID}/details`;
                            PAFlow_Object_Link.textContent = item.FlowDisplayName;
                            PAFlow_Object_Link.className = 'PAFlow_Object_Link';
                            PAFlow_Object_Link.style.textDecoration = 'none';
                            PAFlow_Object_Link.style.display = 'block';
                            PAFlow_Object_Link.id = item.FlowDisplayName.replace(/[^a-zA-Z0-9]/g, '_');

                            // #region PAFlow_Object_Link Event Listeners

                            // Add event listeners to the anchor tag for different mouse clicks
                            let IsCtrlPressed = false;

                            // Add event listener for Ctrl key press
                            PAFlow_Object_Link.addEventListener('mousedown', function (event) {
                                if (event.button === 0 && !event.shiftKey && event.ctrlKey) {
                                    // Ctrl key was pressed
                                    IsCtrlPressed = true;
                                }
                            });

                            // Add event listener for left mouse button click variations
                            PAFlow_Object_Link.addEventListener('click', function (event) {
                                // If Ctrl key was pressed, open the link in a new tab
                                if (IsCtrlPressed) {
                                    IsCtrlPressed = false; // Reset flag
                                    chrome.tabs.create({ url: PAFlow_Object_Link.href, active: false });
                                    LogToServiceWorker('User opened on new tab: ' + PAFlow_Object_Link.href);
                                }

                                // Prevent the default behavior
                                event.preventDefault();

                                // If Shift key was pressed, open the link in a new window
                                if (event.shiftKey) {
                                    window.open(PAFlow_Object_Link.href, "_blank", "noopener,noreferrer");
                                    LogToServiceWorker('User opened in a new window: ' + PAFlow_Object_Link.href);
                                }
                                // If no modifier key was pressed, open the link in the same tab
                                else if (!event.ctrlKey) {
                                    chrome.tabs.update({ url: PAFlow_Object_Link.href });
                                    window.close();
                                    LogToServiceWorker('User opened in the same tab: ' + PAFlow_Object_Link.href);
                                }
                            });

                            // Add event listener for middle mouse button click
                            PAFlow_Object_Link.addEventListener('auxclick', function (event) {
                                if (event.button === 1) {
                                    chrome.tabs.create({ url: PAFlow_Object_Link.href, active: false });
                                    event.preventDefault();
                                    LogToServiceWorker('User opened on new tab: ' + PAFlow_Object_Link.href);
                                }
                            });

                            // #endregion PAFlow_Object_Link Event Listeners

                            // Add the anchor tag to the corresponding container
                            FlowsContainer.appendChild(PAFlow_Object_Link);
                        });
                        resolve();
                    } else {
                        //LogToServiceWorker('Token is null. Aborting.');
                        reject('Token is null. Aborting.');
                    };
                };
            };

            // Add the temporary listener
            chrome.runtime.onMessage.addListener(tempListener);
        });

        chrome.scripting.executeScript({
            target: { tabId: tabId },
            files: ['js/PASPOR-Content.js']
        }, (result) => {
            if (chrome.runtime.lastError) {
                LogToServiceWorker(`Script execution error: ${chrome.runtime.lastError.message}`);
                alert(`Script execution error: ${chrome.runtime.lastError.message}`);
                setTimeout(() => window.close(), 10);
            }
        });

        // Await the promise for the tempListener to complete
        try {
            await tempListenerPromise;
        } catch (error) {
            throw new Error(error);
        }

        //FlowsButton.textContent = FlowsButton.textContent;
        FlowsButton.childNodes[0].textContent = `${FlowsButton.childNodes[0].textContent} (${filteredData.length})`;
        LogToServiceWorker('Power Automate Objects retrieved.');
    } catch (error) {
        throw error;
    };
}

// Function that shows the content of the popup (to be called after objectes have been retrieved)
function Show_Objects_Containers(siteType) {
    LogToServiceWorker('Showing retrieved objects...');
    document.getElementById('MainContainer').classList.remove('Hidden');
    document.getElementById('ButtonsContainer').classList.remove('Hidden');
    document.getElementById('SearchBoxContainer').style.display = 'block';
    document.getElementById('SearchBox').focus();
    if (siteType === "PA") {
        document.getElementById('FlowsButton').classList.remove('Hidden');
    } else if (siteType === "SPO") {
        document.getElementById('ListsButton').classList.remove('Hidden');
        document.getElementById('LibrariesButton').classList.remove('Hidden');
        document.getElementById('SiteButtonsContainer').classList.remove('Hidden');
    }
    else {
        LogToServiceWorker('Error: Invalid site type "' + siteType + '"');
    }
}

function ResizeAndShowElements(siteType) {

    /*
      Maximum allowed size for the whole popup is 1000px width and 3024px height.

      The height of the Main Container will be set to the window scroll height.
      The width of the Main Container will be set to automatically fit the content.

      Width calculations start from the wider container (Lists or Libraries), whose width is already set to automatically
      fit the object with the longest name.
      Both containers' width will be set equal to the greater width amongst them (even the wider one, otherwise it will loose size during filter).

      Then, also the width of the buttons will be set equal to the width of the wider container.
  */

    // Enable the elements to be displayed
    HideLoading();

    Show_Objects_Containers(siteType);
    LogToServiceWorker('Resizing elements...');

    // Get the Main Container
    const Main_Container = document.querySelector('.Main_Container');

    // Get Main Container padding to be added to the height
    const MainContainerPaddingTop = parseFloat(getComputedStyle(Main_Container).paddingTop);

    // Set the display of the Main Container to inline-block to allow the width to be adapted to the content
    Main_Container.style.display = 'inline-block';

    if (siteType === "SPO") {
        // Get the width of the wider container
        const SPO_Objects_Container_Lists = document.querySelector('.SPO_Objects_Container_Lists');
        const SPO_Objects_Container_Libraries = document.querySelector('.SPO_Objects_Container_Libraries');
        const Wider_SPOContainerWidth = Math.max(SPO_Objects_Container_Lists.offsetWidth, SPO_Objects_Container_Libraries.offsetWidth);

        // Set the width of both containers to the width of the Wider container
        SPO_Objects_Container_Lists.style.width = `${Wider_SPOContainerWidth}px`;
        SPO_Objects_Container_Libraries.style.width = `${Wider_SPOContainerWidth}px`;

        // Set the width of both buttons to the width of the Wider container
        const ListsButton = document.getElementById('ListsButton');
        const LibrariesButton = document.getElementById('LibrariesButton');
        ListsButton.style.width = `${Wider_SPOContainerWidth + 1}px`; // !Fix: +1 to compensate for the border (set dynamically)
        LibrariesButton.style.width = `${Wider_SPOContainerWidth - 0.5}px`; // !Fix: -0.5 to compensate for the border (set dynamically)

    } else if (siteType === "PA") {
        // Get the width of the container
        const PA_Objects_Container = document.querySelector('.PA_Objects_Container');
        const PAContainerOffsetWidth = PA_Objects_Container.offsetWidth;

        //LogToServiceWorker('PAContainerScrollWidth: ' + PAContainerOffsetWidth);

        // Set the width of both containers to the width of the Wider container
        PA_Objects_Container.style.width = `${PAContainerOffsetWidth}px`;

        // Set the width of the button to the width of the container
        const FlowsButton = document.getElementById('FlowsButton');
        FlowsButton.style.width = `${PAContainerOffsetWidth}px`;
    }
    else {
        LogToServiceWorker('Error: Invalid site type "' + siteType + '"');
    }

    // Set the height of the Main Container to the window scroll height
    Main_Container.style.height = `${document.documentElement.scrollHeight + MainContainerPaddingTop}px`;
};

// Function that hides the loading animation and restores auto width and height to the body
function HideLoading() {
    document.getElementById('Loading').classList.add('Hidden');
    document.body.style.width = 'auto';
    document.body.style.height = 'auto';
}

// Function applied to the search box that filters individual items inside the containers
function FilterIndividualItems(query) {

    let HigherContainerHeight = 0;
    const SPO_Object_Containers = document.querySelectorAll('.SPO_Objects_Container_Lists, .SPO_Objects_Container_Libraries, .PA_Objects_Container');
    // Get the Main Container
    const Main_Container = document.querySelector('.Main_Container');
    // Get Main Container padding to be added to the height
    const MainContainerPaddingTop = parseFloat(getComputedStyle(Main_Container).paddingTop);

    // Replace * with .*
    query = query.replace(/\*/g, '.*');
    // Replace % with .
    query = query.replace(/%/g, '.');
    const regex = new RegExp(query, 'i');


    // Reset all containers and items to be visible
    SPO_Object_Containers.forEach(SPO_Object_Container => {
        SPO_Object_Container.style.display = 'block';
        SPO_Object_Container.querySelectorAll('.SPO_Object_Link, .PAFlow_Object_Link').forEach(SPO_Object => {
            SPO_Object.style.display = 'block';
        });
    });

    const hasFilters = query.trim() !== '';
    SPO_Object_Containers.forEach(SPO_Object_Container => {
        let hasVisibleItems = false;
        let totalItemCount = 0;
        let visibleItemCount = 0;
        const SPO_Object_Links = SPO_Object_Container.querySelectorAll('.SPO_Object_Link, .PAFlow_Object_Link');

        SPO_Object_Links.forEach(SPO_Object => {
            const SPO_Object_Name = SPO_Object.innerText.toLowerCase();
            totalItemCount++;
            if (regex.test(SPO_Object_Name)) {
                hasVisibleItems = true;
                visibleItemCount++;
            } else {
                SPO_Object.style.display = 'none';
            }
        });

        if (!hasVisibleItems) {
            SPO_Object_Container.style.display = 'none';
            const ContainerButton = SPO_Object_Container.parentNode.childNodes[0];
            if (ContainerButton) {
                const containerButtonText = ContainerButton.textContent.trim();
                const containerButtonTextWithoutCount = containerButtonText.replace(/\s*\(\d+\/?\d*\)\s*$/, '');
                ContainerButton.textContent = `${containerButtonTextWithoutCount} (${hasFilters ? '0' : totalItemCount})`;
            }
        } else {
            const ContainerButton = SPO_Object_Container.parentNode.childNodes[0];
            if (ContainerButton) {
                const containerButtonText = ContainerButton.textContent.trim();
                const containerButtonTextWithoutCount = containerButtonText.replace(/\s*\(\d+\/?\d*\)\s*$/, '');
                ContainerButton.textContent = `${containerButtonTextWithoutCount} (${visibleItemCount}/${totalItemCount})`;
            }
            if (SPO_Object_Container.offsetHeight > HigherContainerHeight) {
                HigherContainerHeight = SPO_Object_Container.offsetHeight;
                DistanceFromTop = parseFloat(document.getElementById(SPO_Object_Container.id).getBoundingClientRect().top + window.scrollY);
            }
        }
    });
    Main_Container.style.height = `${HigherContainerHeight + DistanceFromTop + MainContainerPaddingTop}px`;
};


// #endregion Functions

// #region Window Event Listeners

// Popup is loaded
document.addEventListener('DOMContentLoaded', () => {

    LogToServiceWorker('Loading window...');
    ShowLoading();

    // Get the current tab
    chrome.tabs.query({ active: true, currentWindow: true }, async tabs => {

        try {

            // Get the current tab URL
            const CurrentTabUrl = tabs[0].url;

            // If the URL is a SharePoint Online Site, then proceed showing retrieved SharePoint Online Site Objects
            if (await CheckIfSharePointSite(CurrentTabUrl)) {

                // Extract from URL SharePoint main site URL
                LogToServiceWorker('Valid SharePoint Url.');
                const SPOMainSiteUrl = GetSharePointMainSiteUrl(CurrentTabUrl);

                // Retrieve SharePoint Online Site Objects
                await GetSharePointSiteObjects(SPOMainSiteUrl);

                // Resize and show retrieved content
                ResizeAndShowElements('SPO');
            }
            // If the URL is a Power Automate site, then proceed showing retrieved Power Automate site Objects
            else if (await CheckIfPowerAutomateSite(CurrentTabUrl)) {
                LogToServiceWorker('Valid Power Automate Url.');
                const PAEnvironment = GetPowerAutomateEnvironment(CurrentTabUrl);

                // Retrieve Power Automate Objects
                await GetPowerAutomateSiteObjects(PAEnvironment, tabs[0].id);

                // Resize and show retrieved content
                ResizeAndShowElements('PA');
            }
            // Else, show 'Not a SharePoint Site' message
            else {
                LogToServiceWorker('Not a SharePoint Online or Power Automate site.');
                displayNotSharePointOrPASiteMessage(); //! Show 'Not a SharePoint Site' message. Change this to a div and css class in hidden state. Just change the display to block when needed.
            }
        } catch (error) {
            let ErrorRow = error.stack.split('\n').slice(1).find(line => line.includes('at chrome-extension')).match(/:(\d+):\d+/)[1]; //"!include"->"include"
            let ErrorSource = error.stack.split('\n').slice(1).find(line => line.includes('at chrome-extension')).match(/([^\/]*\.js:\d+:\d+)/)[1]; //"!include"->"include"
            LogToServiceWorker(`ERROR: ${error}`, ErrorRow);
            // Display the error in a popup
            alert(`An error occurred at ${ErrorSource}:\n${error}`); //.message
            setTimeout(() => window.close(), 10);
            //throw error;
        }
    });
});

// Refresh button listener (ignores cache)
document.addEventListener('DOMContentLoaded', function () {
    const refreshButton = document.getElementById('RefreshButton');

    refreshButton.addEventListener('click', function () {
        // Enable cache-busting
        localStorage.setItem('ShouldBustCache', 'true');

        // Reload the popup
        location.reload(true);
    });
});

// Real-time search functionality
document.addEventListener('DOMContentLoaded', () => {
    const searchBox = document.getElementById('SearchBox');
    if (searchBox) {
        searchBox.addEventListener('input', function () {
            const query = this.value.toLowerCase();
            FilterIndividualItems(query);
        });
    }
});

// Listener for messages from the service worker
navigator.serviceWorker.addEventListener('message', event => {
    if (event.data.type === 'ERROR') {
        alert(event.data.message);
        setTimeout(() => window.close(), 10);
    }
});

// #endregion Window Event Listeners







// check order from here:
//! Function that show the output for unsupported sites  //TODO: Remove this function and create a div and css class for the message in hidden state. Just change the display to block when needed.
function displayNotSharePointOrPASiteMessage() {
    HideLoading(); // Hide loading animation
    const messageDiv = document.createElement('div');
    messageDiv.textContent = 'Not a SharePoint Online or Power Automate site';
    messageDiv.style.textAlign = 'center';
    messageDiv.style.padding = '10px';
    messageDiv.style.color = '#d9534f';
    messageDiv.style.whiteSpace = 'nowrap'; // Prevent text wrapping
    messageDiv.style.fontWeight = 'bold'; // Set the text as bold
    document.body.appendChild(messageDiv);
    document.body.style.width = 'auto';
    document.body.style.height = 'auto';
}