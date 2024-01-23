//to stop, delete TextToFind and set again;
const TextToFind = "Mar 6,"; //"4245-BH-DM-PIB1360101-IS00";

// Search for the text using window.find
let searchResult = window.find(TextToFind);

// Loop until the text is found or all elements have been checked
let elements = document.getElementsByTagName("*");
while (!searchResult && elements.length) {

	// If the text was not found, check if there is a "Show more" button
	for (let i = 0; i < elements.length; i++) {

		if (elements[i].textContent === "Show more") {

			elements[i].scrollIntoView();
			elements[i].click();
			break;

		}

	}

	// Wait for the page to reload
	await new Promise(resolve => setTimeout(resolve, 2000));

	// Get updated list of elements
	elements = document.getElementsByTagName("*");

	// Search for the text again
	searchResult = window.find(TextToFind);

}

// If the text was found, highlight it
if (searchResult) {

	console.log(`Text '${TextToFind}' found`)
	document.designMode = "on";
	document.execCommand("HiliteColor", false, "yellow");
	document.designMode = "off";

}