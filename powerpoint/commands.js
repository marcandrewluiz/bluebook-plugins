/******/ (function() { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
/* global Office */
Office.onReady(function () {
  // Office.js is ready to be called.
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
// Register the function with Office.
Office.actions.associate("action", action);
/******/ })()
;
//# sourceMappingURL=commands.js.map