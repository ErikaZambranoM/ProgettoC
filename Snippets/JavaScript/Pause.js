// Pause function to sleep for 'ms' milliseconds (only for testing purposes)
function Pause(milliseconds) {
    const dt = new Date();
    while ((new Date()) - dt <= milliseconds) { /* Do nothing, just wait */ }
}
// Pause(10000); // Sleep for 10 second