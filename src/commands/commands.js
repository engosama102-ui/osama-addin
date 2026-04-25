/* global Office */

Office.onReady(() => {});

function action(event) {
  event.completed();
}

Office.actions.associate("action", action);
