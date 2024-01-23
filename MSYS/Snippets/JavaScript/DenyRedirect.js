// Run this command on the browser developer console to deny redirecting to another page.
window.onbeforeunload = function () { return 'Leave page?'; };