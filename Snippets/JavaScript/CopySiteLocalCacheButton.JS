/*

Cause of the error:
"
Something went wrong
Invalid statement, ResultCode: 1
"

Local cache:
nucleus.b0803988-95b6-4920-a5ef-dbffec8b6d86.firstContactCount: 1
nucleus.b0803988-95b6-4920-a5ef-dbffec8b6d86.firstContact: true

the UID seems taken from odSuiteNavdata.data.ClientData
*/

// Create a new button element
let copyButton = document.createElement('button');
copyButton.textContent = 'Copy Local Storage Data';
copyButton.style.position = 'fixed';
copyButton.style.left = '10px';
copyButton.style.top = '10px';
copyButton.style.zIndex = '1000';
copyButton.style.padding = '10px';
copyButton.style.backgroundColor = '#4CAF50';
copyButton.style.color = 'white';
copyButton.style.border = 'none';
copyButton.style.cursor = 'pointer';

// Function to copy data to clipboard
// Function to copy data to clipboard
function copyLocalStorageData() {
    let localStorageData = {};
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        let value = localStorage.getItem(key);

        // Try parsing the JSON data. If it fails, use the raw value.
        try {
            value = JSON.parse(value);
        } catch (e) {
            // Value is not JSON, use it as is
        }

        localStorageData[key] = value;
    }

    // Convert the data to a prettified JSON string
    const prettifiedDataStr = JSON.stringify(localStorageData, null, 2); //test with 4

    navigator.clipboard.writeText(prettifiedDataStr).then(function () {
        console.log('Prettified Local Storage data copied to clipboard');
    }, function (err) {
        console.error('Could not copy text: ', err);
    });
}

// Attach event listener to the button
copyButton.addEventListener('click', copyLocalStorageData);

// Append the button to the body (or any other visible element)
document.body.appendChild(copyButton);